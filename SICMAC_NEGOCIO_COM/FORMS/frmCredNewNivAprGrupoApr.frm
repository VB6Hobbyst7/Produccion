VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredNewNivAprGrupoApr 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7935
   Icon            =   "frmCredNewNivAprGrupoApr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabNivApr 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Niveles de Aprobación"
      TabPicture(0)   =   "frmCredNewNivAprGrupoApr.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "feNivApr"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtNivApr"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraPorCargo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FraPorUsuario"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelaNivApr"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdGuardarNivApr"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkCorrigeSug"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkValidaAg"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSubir"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdBajar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdEliminarNivApr"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdEditarNivApr"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdCerrarNivApr"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "OptCargoUsu(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "OptCargoUsu(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).ControlCount=   16
      Begin VB.OptionButton OptCargoUsu 
         Caption         =   "Por Usuario"
         Height          =   255
         Index           =   2
         Left            =   4560
         TabIndex        =   45
         Top             =   950
         Width           =   1215
      End
      Begin VB.OptionButton OptCargoUsu 
         Caption         =   "Por Cargo"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   44
         Top             =   950
         Width           =   1095
      End
      Begin VB.CommandButton cmdCerrarNivApr 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         TabIndex        =   43
         Top             =   5520
         Width           =   1170
      End
      Begin VB.CommandButton cmdEditarNivApr 
         Caption         =   "Editar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   42
         Top             =   5520
         Width           =   1170
      End
      Begin VB.CommandButton cmdEliminarNivApr 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   41
         Top             =   5520
         Width           =   1170
      End
      Begin VB.CommandButton cmdBajar 
         Caption         =   "Bajar Orden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6480
         TabIndex        =   40
         Top             =   4440
         Width           =   1050
      End
      Begin VB.CommandButton cmdSubir 
         Caption         =   "Subir Orden"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6480
         TabIndex        =   39
         Top             =   3960
         Width           =   1050
      End
      Begin VB.CheckBox chkValidaAg 
         Caption         =   "Valida Agencia"
         Height          =   255
         Left            =   2280
         TabIndex        =   37
         Top             =   3400
         Width           =   1575
      End
      Begin VB.CheckBox chkCorrigeSug 
         Caption         =   "Corrige Sugerencia"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   3400
         Width           =   1815
      End
      Begin VB.CommandButton cmdGuardarNivApr 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5160
         TabIndex        =   35
         Top             =   3360
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancelaNivApr 
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
         Height          =   345
         Left            =   6360
         TabIndex        =   34
         Top             =   3360
         Width           =   1170
      End
      Begin VB.Frame FraPorUsuario 
         Enabled         =   0   'False
         Height          =   2295
         Left            =   4440
         TabIndex        =   19
         Top             =   960
         Width           =   3135
         Begin VB.ComboBox cboTipoPorUsuario 
            Height          =   315
            ItemData        =   "frmCredNewNivAprGrupoApr.frx":0326
            Left            =   1320
            List            =   "frmCredNewNivAprGrupoApr.frx":0330
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   300
            Width           =   735
         End
         Begin VB.CommandButton cmdAgregarPorUsuario 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2205
            TabIndex        =   26
            Top             =   300
            Width           =   375
         End
         Begin VB.CommandButton cmdEliminarPorUsuario 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2640
            TabIndex        =   25
            Top             =   300
            Width           =   375
         End
         Begin SICMACT.TxtBuscar txtBuscarUsuario 
            Height          =   315
            Left            =   120
            TabIndex        =   28
            Top             =   300
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   556
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SICMACT.FlexEdit fePorUsuario 
            Height          =   1095
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   2865
            _ExtentX        =   5054
            _ExtentY        =   1931
            Cols0           =   4
            HighLight       =   1
            EncabezadosNombres=   "-CodUsu-Usuario-Tipo"
            EncabezadosAnchos=   "300-0-1100-1100"
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
            ColumnasAEditar =   "X-X-X-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-C-C"
            FormatosEdit    =   "0-1-0-1"
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            ColWidth0       =   300
            RowHeight0      =   300
         End
         Begin SICMACT.EditMoney txtNumFirmasPorUsuario 
            Height          =   300
            Left            =   2400
            TabIndex        =   33
            Top             =   1875
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label Label4 
            Caption         =   "Nº Firmas Digitales :"
            Height          =   255
            Left            =   840
            TabIndex        =   31
            Top             =   1920
            Width           =   1455
         End
      End
      Begin VB.Frame FraPorCargo 
         Enabled         =   0   'False
         Height          =   2295
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   4215
         Begin VB.CommandButton cmdEliminarPorCargo 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3600
            TabIndex        =   23
            Top             =   300
            Width           =   375
         End
         Begin VB.CommandButton cmdAgregarPorCargo 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3150
            TabIndex        =   22
            Top             =   300
            Width           =   375
         End
         Begin VB.ComboBox cboTipoPorCargo 
            Height          =   315
            ItemData        =   "frmCredNewNivAprGrupoApr.frx":033A
            Left            =   1920
            List            =   "frmCredNewNivAprGrupoApr.frx":0344
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   300
            Width           =   735
         End
         Begin SICMACT.TxtBuscar txtBuscarCargo 
            Height          =   315
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   1455
            _ExtentX        =   1720
            _ExtentY        =   556
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SICMACT.FlexEdit fePorCargo 
            Height          =   1095
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   1931
            Cols0           =   4
            HighLight       =   1
            EncabezadosNombres=   "-CargoCod-Cargo-Tipo"
            EncabezadosAnchos=   "300-0-2400-880"
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
            ColumnasAEditar =   "X-X-X-X"
            ListaControles  =   "0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-C"
            FormatosEdit    =   "0-1-1-0"
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            ColWidth0       =   300
            RowHeight0      =   300
         End
         Begin SICMACT.EditMoney txtNumFirmasPorCargo 
            Height          =   300
            Left            =   3480
            TabIndex        =   32
            Top             =   1875
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label Label3 
            Caption         =   "Nº Firmas Digitales :"
            Height          =   255
            Left            =   1920
            TabIndex        =   30
            Top             =   1920
            Width           =   1455
         End
      End
      Begin VB.TextBox txtNivApr 
         Height          =   300
         Left            =   960
         TabIndex        =   16
         Top             =   480
         Width           =   2955
      End
      Begin SICMACT.FlexEdit feNivApr 
         Height          =   1455
         Left            =   120
         TabIndex        =   38
         Top             =   3960
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   2566
         Cols0           =   6
         HighLight       =   1
         EncabezadosNombres=   "Orden-NivelCod-Nivel-Valida Ag?-Tipo-Valores"
         EncabezadosAnchos=   "600-0-2600-1020-900-800"
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
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C-C"
         FormatosEdit    =   "0-1-1-0-0-0"
         TextArray0      =   "Orden"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   600
         RowHeight0      =   300
      End
      Begin VB.Label Label2 
         Caption         =   "Nivel :"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   510
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTabGrupos 
      Height          =   6015
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Registrar/Editar Niveles de Aprobación"
      TabPicture(0)   =   "frmCredNewNivAprGrupoApr.frx":034E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "feGrupos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtGrupo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdGrabarGrupoApr"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCancelarGrupoApr"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCerrar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdEditarGrupo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdEliminarGrupo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton cmdEliminarGrupo 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   10
         Top             =   5520
         Width           =   1170
      End
      Begin VB.CommandButton cmdEditarGrupo 
         Caption         =   "Editar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   9
         Top             =   5520
         Width           =   1170
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         TabIndex        =   11
         Top             =   5520
         Width           =   1170
      End
      Begin VB.CommandButton cmdCancelarGrupoApr 
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
         Height          =   345
         Left            =   6000
         TabIndex        =   7
         Top             =   3480
         Width           =   1170
      End
      Begin VB.CommandButton cmdGrabarGrupoApr 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4800
         TabIndex        =   6
         Top             =   3480
         Width           =   1170
      End
      Begin VB.Frame Frame2 
         Caption         =   " Aplicable a Sub Productos "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   480
         TabIndex        =   15
         Top             =   960
         Width           =   3255
         Begin VB.CheckBox chkTodosSubProductos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   160
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
         Begin VB.ListBox lstSubProd 
            Height          =   1635
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   4
            Top             =   555
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " Aplicable a Agencias "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3960
         TabIndex        =   14
         Top             =   960
         Width           =   3255
         Begin VB.ListBox lstAgencias 
            Height          =   1635
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   555
            Width           =   3015
         End
         Begin VB.CheckBox chkTodosAgencia 
            Caption         =   "Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtGrupo 
         Height          =   300
         Left            =   1080
         TabIndex        =   0
         Top             =   480
         Width           =   6195
      End
      Begin SICMACT.FlexEdit feGrupos 
         Height          =   1455
         Left            =   120
         TabIndex        =   8
         Top             =   3960
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   2566
         Cols0           =   5
         HighLight       =   1
         EncabezadosNombres=   "-GrupoCod-Grupo-SubProdts.-Agencias"
         EncabezadosAnchos=   "300-0-4320-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C"
         FormatosEdit    =   "0-1-1-0-0"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   300
         RowHeight0      =   300
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo :"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   510
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCredNewNivAprGrupoApr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredNewNivAprGrupoApr
'** Descripción : Formulario para la Administracion de los Grupos de Niveles de Aprobacion y
'**               Registro de los Niveles de Aprobacion creado segun RFC110-2012
'** Creación : JUEZ, 20121128 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim fbNuevo As Boolean
Dim fbActualiza As Boolean
Dim fnTipoReg As Integer
Dim fnRowNivApr As Integer

Public Sub InicioGrupoAprobacion()
    Me.Caption = "Grupo de Aprobación de Créditos"
    Me.SSTabGrupos.Visible = True
    Me.SSTabNivApr.Visible = False
    fbNuevo = True
    fbActualiza = False
    CargaDatosGruposApr
    ListarAgencias
    ListarSubProductos
    feGrupos.TopRow = 1
    feGrupos.Row = 1
    cmdCerrar.Cancel = True
    cmdCerrarNivApr.Cancel = False
    Me.Show 1
End Sub

Public Sub InicioRegistroNiveles()
    Me.Caption = "Registro de Niveles de Aprobación"
    Me.SSTabGrupos.Visible = False
    Me.SSTabNivApr.Visible = True
    fbNuevo = True
    fbActualiza = False
    fnTipoReg = 0
    CargaDatosNiveles
    feNivApr.TopRow = 1
    feNivApr.Row = 1
    cmdCerrar.Cancel = False
    cmdCerrarNivApr.Cancel = True
    Me.Show 1
End Sub

Private Sub cmdCancelarGrupoApr_Click()
    Dim nRow As Integer
    nRow = feGrupos.Row
    Call LimpiaControles(Me, True, False)
    chkTodosAgencia.value = 0
    ListarAgencias
    chkTodosSubProductos.value = 0
    ListarSubProductos
    fbNuevo = True
    fbActualiza = False
    cmdEditarGrupo.Enabled = True
    cmdEliminarGrupo.Enabled = True
    feGrupos.Enabled = True
    CargaDatosGruposApr
    feGrupos.TopRow = nRow
    feGrupos.Row = nRow
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub ListarAgencias()
    Dim oAge As COMDConstantes.DCOMAgencias
    Dim rsAgencias As ADODB.Recordset
    Set oAge = New COMDConstantes.DCOMAgencias
        Set rsAgencias = oAge.ObtieneAgencias()
    Set oAge = Nothing
    If rsAgencias Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        lstAgencias.Clear
        With rsAgencias
            Do While Not rsAgencias.EOF
                lstAgencias.AddItem rsAgencias!nConsValor & " " & Trim(rsAgencias!cConsDescripcion)
                rsAgencias.MoveNext
            Loop
        End With
        lstAgencias.Selected(0) = True
    End If
End Sub

Private Sub ListarSubProductos()
    Dim oSubProd As COMDConstantes.DCOMConstantes
    Dim rsSubProd As ADODB.Recordset
    Set oSubProd = New COMDConstantes.DCOMConstantes
        Set rsSubProd = oSubProd.RecuperaConstantes(3033, 1)
    Set oSubProd = Nothing
    If rsSubProd Is Nothing Then
        MsgBox " No se encuentran los Sub Productos de Creditos ", vbInformation, " Aviso "
    Else
        lstSubProd.Clear
        With rsSubProd
            Do While Not rsSubProd.EOF
                lstSubProd.AddItem rsSubProd!nConsValor & " " & Trim(rsSubProd!cConsDescripcion)
                
                rsSubProd.MoveNext
            Loop
        End With
        lstSubProd.Selected(0) = True
    End If
End Sub

Private Sub chkTodosAgencia_Click()
    Call CheckLista(IIf(chkTodosAgencia.value = 1, True, False), lstAgencias)
End Sub

Private Sub chkTodosSubProductos_Click()
    Call CheckLista(IIf(chkTodosSubProductos.value = 1, True, False), lstSubProd)
End Sub

Private Sub CheckLista(ByVal bCheck As Boolean, ByVal lstLista As ListBox)
    Dim i As Integer
    For i = 0 To lstLista.ListCount - 1
        lstLista.Selected(i) = bCheck
    Next i
End Sub

Private Sub CargaDatosGruposApr()
    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    Set oNiv = New COMDCredito.DCOMNivelAprobacion
    
    Set rs = oNiv.RecuperaGruposApr()
    Set oNiv = Nothing
    Call LimpiaFlex(feGrupos)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feGrupos.AdicionaFila
            lnFila = feGrupos.Row
            feGrupos.TextMatrix(lnFila, 1) = rs!cGrupoCod
            feGrupos.TextMatrix(lnFila, 2) = rs!cGrupoDesc
            feGrupos.TextMatrix(lnFila, 3) = "Ver"
            feGrupos.TextMatrix(lnFila, 4) = "Ver"
            rs.MoveNext
        Loop
        feGrupos.TopRow = 1
    Else
        cmdEditarGrupo.Enabled = False
        cmdEliminarGrupo.Enabled = False
        feGrupos.Enabled = False
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdEditarGrupo_Click()
    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer, nRow As Integer
    nRow = feGrupos.Row
    fbActualiza = True
    fbNuevo = False
    cmdEditarGrupo.Enabled = False
    cmdEliminarGrupo.Enabled = False
    feGrupos.Enabled = False
    
    Set oNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oNiv.RecuperaGruposApr(feGrupos.TextMatrix(feGrupos.Row, 1))
    Set oNiv = Nothing

    txtGrupo.Text = rs!cGrupoDesc

    lstSubProd.Clear
    Call LlenaListas(lstSubProd, 1)
    lstAgencias.Clear
    Call LlenaListas(lstAgencias, 2)
    feGrupos.TopRow = nRow
    feGrupos.Row = nRow
End Sub

Private Sub cmdEliminarGrupo_Click()
    If feGrupos.TextMatrix(feGrupos.Row, 0) <> "" Then
        If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feGrupos.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Dim oNivApr As COMNCredito.NCOMNivelAprobacion
            Set oNivApr = New COMNCredito.NCOMNivelAprobacion
            Call oNivApr.dEliminaGruposApr(feGrupos.TextMatrix(feGrupos.Row, 1))
            feGrupos.EliminaFila feGrupos.Row
            CargaDatosGruposApr
        End If
    End If
End Sub

Private Sub LlenaListas(ByRef Lista As ListBox, ByVal pnTipoLista As Integer)
    Dim oLista As COMDCredito.DCOMNivelAprobacion
    Dim oLCred As COMDConstantes.DCOMConstantes
    Dim oAge As COMDConstantes.DCOMAgencias
    Dim rs As ADODB.Recordset
    Dim rsLista As ADODB.Recordset
    Dim i As Integer, J As Integer
    
    Set oLista = New COMDCredito.DCOMNivelAprobacion
    Set rs = oLista.RecuperaGruposAprDetalle(feGrupos.TextMatrix(feGrupos.Row, 1), pnTipoLista)
    Set oLista = Nothing

    If pnTipoLista = 1 Then
        Set oLCred = New COMDConstantes.DCOMConstantes
        Set rsLista = oLCred.RecuperaConstantes(3033, 1)
        Set oLCred = Nothing
    Else
        Set oAge = New COMDConstantes.DCOMAgencias
        Set rsLista = oAge.ObtieneAgencias()
        Set oAge = Nothing
    End If
    
    For i = 0 To rsLista.RecordCount - 1
        Lista.AddItem rsLista!nConsValor & " " & Trim(rsLista!cConsDescripcion)
        rs.MoveFirst
        For J = 0 To rs.RecordCount - 1
            If Trim(rsLista!nConsValor) = Trim(rs!cAgeSubProdValor) Then
                Lista.Selected(i) = True
                Exit For
            End If
            rs.MoveNext
        Next J
        rsLista.MoveNext
    Next i
End Sub

Private Sub cmdEliminarPorUsuario_Click()
    If fePorUsuario.TextMatrix(fePorUsuario.Row, 0) <> "" Then
        If MsgBox("¿Eliminar los datos de la fila " + CStr(fePorUsuario.Row) + " de la lista de Usuarios?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            fePorUsuario.EliminaFila fePorUsuario.Row
            Call VerificaNumFirmas(False, txtNumFirmasPorUsuario, fePorUsuario)
        End If
    End If
End Sub

Private Sub cmdGrabarGrupoApr_Click()
    If ValidaDatosGruposApr Then
        Dim oNivApr As COMNCredito.NCOMNivelAprobacion
        Dim MatTpoProd() As String
        Dim MatAgencias() As String
        Dim nRow As Integer
        nRow = feGrupos.Row
        ReDim MatTpoProd(DevuelveCantidadCheckList(lstSubProd), 1)
        ReDim MatAgencias(DevuelveCantidadCheckList(lstAgencias), 1)
        
        MatTpoProd = LlenaMatriz(lstSubProd)
        MatAgencias = LlenaMatriz(lstAgencias)
        
        Set oNivApr = New COMNCredito.NCOMNivelAprobacion
        
        If oNivApr.VerificaDatosGrupoApr(IIf(fbNuevo, "", feGrupos.TextMatrix(feGrupos.Row, 1)), Trim(txtGrupo.Text), MatTpoProd, MatAgencias) = False Then
            If MsgBox("¿Está seguro de " + IIf(fbNuevo = True, "registrar", "actualizar") + " los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            If fbNuevo Then
                Call oNivApr.dInsertaGruposApr(Trim(txtGrupo.Text), MatTpoProd, MatAgencias)
            ElseIf fbActualiza Then
                Call oNivApr.dActualizaGruposApr(feGrupos.TextMatrix(feGrupos.Row, 1), Trim(txtGrupo.Text), _
                                                    MatTpoProd, MatAgencias)
            End If
        Else
            MsgBox "Existe un Grupo ya ingresado con los mismos datos, favor de verificar", vbInformation, "Aviso"
        End If
        MsgBox "Los datos se " & IIf(fbNuevo = True, "registraron", "actualizaron") & " correctamente", vbInformation, "Aviso"
        Call cmdCancelarGrupoApr_Click
        CargaDatosGruposApr
        feGrupos.TopRow = nRow
        feGrupos.Row = nRow
    End If
End Sub

Private Function LlenaMatriz(ByVal lstLista As ListBox) As Variant
    Dim MatLista() As String
    Dim i As Integer
    Dim nTamano As Integer

    ReDim MatLista(DevuelveCantidadCheckList(lstLista), 1)
    nTamano = 1
    For i = 1 To lstLista.ListCount
        If lstLista.Selected(i - 1) = True Then
            MatLista(nTamano, 0) = Trim(Left(lstLista.List(i - 1), 3))
            nTamano = nTamano + 1
        End If
    Next
    LlenaMatriz = MatLista()
End Function

Private Function ValidaDatosGruposApr() As Boolean
    Dim i As Integer
    Dim CTpoProd As Integer
    Dim CAgencia As Integer
    ValidaDatosGruposApr = False
    
    CTpoProd = DevuelveCantidadCheckList(lstSubProd)
    CAgencia = DevuelveCantidadCheckList(lstAgencias)
    
    If Trim(txtGrupo.Text) = "" Then
        MsgBox "Debe detallar el nombre del Grupo", vbInformation, "Aviso"
        txtGrupo.SetFocus
        ValidaDatosGruposApr = False
        Exit Function
    End If
    If CTpoProd = 0 Then
        MsgBox "Debe seleccionar al menos un Tipo de Producto", vbInformation, "Aviso"
        lstSubProd.SetFocus
        ValidaDatosGruposApr = False
        Exit Function
    End If
    If CAgencia = 0 Then
        MsgBox "Debe seleccionar al menos una Agencia", vbInformation, "Aviso"
        lstAgencias.SetFocus
        ValidaDatosGruposApr = False
        Exit Function
    End If
    
    ValidaDatosGruposApr = True
End Function

Private Sub feGrupos_Click()
    If feGrupos.TextMatrix(feGrupos.Row, feGrupos.Col) <> "" Then
        MuestraDatosGridGrupo
    End If
End Sub

Private Sub feGrupos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If feGrupos.TextMatrix(feGrupos.Row, feGrupos.Col) <> "" Then
            MuestraDatosGridGrupo
        End If
    End If
End Sub

Private Sub MuestraDatosGridGrupo()
    Dim oSubProd As COMDConstantes.DCOMConstantes
    Dim oAge As COMDConstantes.DCOMAgencias
    Dim oLista As COMDCredito.DCOMNivelAprobacion
    Dim rsLista As ADODB.Recordset, rsDatos As ADODB.Recordset
    
    If feGrupos.Col = 3 Then
        Set oSubProd = New COMDConstantes.DCOMConstantes
            Set rsLista = oSubProd.RecuperaConstantes(3033, 1)
        Set oSubProd = Nothing
        Set oLista = New COMDCredito.DCOMNivelAprobacion
            Set rsDatos = oLista.RecuperaGruposAprDetalle(feGrupos.TextMatrix(feGrupos.Row, 1), 1)
        Set oLista = Nothing
        frmCredListaDatos.Inicio "Sub Productos", rsDatos, rsLista, 1
    ElseIf feGrupos.Col = 4 Then
        Set oAge = New COMDConstantes.DCOMAgencias
            Set rsLista = oAge.ObtieneAgencias()
        Set oAge = Nothing
        Set oLista = New COMDCredito.DCOMNivelAprobacion
            Set rsDatos = oLista.RecuperaGruposAprDetalle(feGrupos.TextMatrix(feGrupos.Row, 1), 2)
        Set oLista = Nothing
        frmCredListaDatos.Inicio "Agencias", rsDatos, rsLista, 1
    End If
End Sub

'***********************************************************************************************************************

Private Sub CargaDatosNiveles()
    Dim oConst As COMDConstantes.DCOMConstantes
    Set oConst = New COMDConstantes.DCOMConstantes
    txtBuscarCargo.lbUltimaInstancia = False
    txtBuscarCargo.psRaiz = "CARGOS DISPONIBLES PARA LOS NIVELES DE APROBACION"
    txtBuscarCargo.rs = oConst.ObtenerCargosArea
    txtBuscarUsuario.lbUltimaInstancia = False
    txtBuscarUsuario.psRaiz = "USUARIOS DISPONIBLES PARA LOS NIVELES DE APROBACION"
    txtBuscarUsuario.rs = oConst.ObtenerUsuariosArea
    
    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    Set oNiv = New COMDCredito.DCOMNivelAprobacion
    
    Set rs = oNiv.RecuperaNivApr()
    Set oNiv = Nothing
    Call LimpiaFlex(feNivApr)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feNivApr.AdicionaFila
            lnFila = feNivApr.Row
            feNivApr.TextMatrix(lnFila, 1) = rs!cNivAprCod
            feNivApr.TextMatrix(lnFila, 2) = rs!cNivAprDesc
            feNivApr.TextMatrix(lnFila, 3) = rs!cValidaAg
            feNivApr.TextMatrix(lnFila, 4) = rs!cTipoReg
            feNivApr.TextMatrix(lnFila, 5) = "Ver"
            rs.MoveNext
        Loop
    Else
        cmdEditarNivApr.Enabled = False
        cmdEliminarNivApr.Enabled = False
        feNivApr.Enabled = False
        cmdSubir.Enabled = False
        cmdBajar.Enabled = False
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub feNivApr_DblClick()
    Me.feNivApr.Row = 1
End Sub

Private Sub feNivApr_GotFocus()
    Me.feNivApr.Row = 1
End Sub

Private Sub OptCargoUsu_Click(Index As Integer)
    Select Case Index
    Case 1
        txtBuscarUsuario = ""
        cboTipoPorUsuario.ListIndex = -1
        Call LimpiaFlex(fePorUsuario)
        txtNumFirmasPorUsuario.value = 0
        FraPorCargo.Enabled = True
        FraPorUsuario.Enabled = False
        fnTipoReg = 1
    Case 2
        txtBuscarCargo = ""
        cboTipoPorCargo.ListIndex = -1
        Call LimpiaFlex(fePorCargo)
        txtNumFirmasPorCargo.value = 0
        FraPorCargo.Enabled = False
        FraPorUsuario.Enabled = True
        fnTipoReg = 2
    End Select
End Sub

Private Sub txtBuscarCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTipoPorCargo.SetFocus
    End If
End Sub

Private Sub txtBuscarUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTipoPorUsuario.SetFocus
    End If
End Sub

Private Sub txtGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkTodosSubProductos.SetFocus
    End If
End Sub

Private Sub txtGrupo_LostFocus()
    txtGrupo.Text = UCase(txtGrupo.Text)
End Sub

Private Sub cmdAgregarPorCargo_Click()
    Dim oConst As COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    If Trim(txtBuscarCargo) = "" Then
        MsgBox "Falta ingresar el cargo", vbInformation, "Aviso"
        txtBuscarCargo.SetFocus
        Exit Sub
    End If
    If Trim(cboTipoPorCargo.Text) = "" Then
        MsgBox "Falta ingresar el tipo", vbInformation, "Aviso"
        cboTipoPorCargo.SetFocus
        Exit Sub
    End If
    
    For i = 1 To fePorCargo.Rows - 1
        If fePorCargo.TextMatrix(i, 0) <> "" Then
            If Trim(fePorCargo.TextMatrix(i, 1)) = Right(Trim(txtBuscarCargo), 6) Then
                MsgBox "El cargo ya fue ingresado", vbInformation, "Aviso"
                txtBuscarCargo.SetFocus
                Exit Sub
            End If
        End If
    Next i
    
    Set oConst = New COMDConstantes.DCOMConstantes
    Set rs = oConst.ObtenerCargosArea(Right(Trim(txtBuscarCargo), 6))
    
    If Not rs.EOF Then
        fePorCargo.AdicionaFila
        fePorCargo.TextMatrix(fePorCargo.Row, 1) = Right(Trim(txtBuscarCargo), 6)
        fePorCargo.TextMatrix(fePorCargo.Row, 2) = rs!Descripcion
        fePorCargo.TextMatrix(fePorCargo.Row, 3) = IIf(Trim(cboTipoPorCargo.Text) = "N", "Nesesario", "Opcional")
        Call VerificaNumFirmas(False, txtNumFirmasPorCargo, fePorCargo)
        txtBuscarCargo = ""
        Me.cboTipoPorCargo.ListIndex = -1
    Else
        MsgBox "El codigo ingresado no existe", vbInformation, "Aviso"
        txtBuscarCargo.SetFocus
    End If
End Sub

Private Sub cmdEliminarPorCargo_Click()
    If fePorCargo.TextMatrix(fePorCargo.Row, 0) <> "" Then
        If MsgBox("¿Eliminar los datos de la fila " + CStr(fePorCargo.Row) + " de la lista de Cargos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            fePorCargo.EliminaFila fePorCargo.Row
            Call VerificaNumFirmas(False, txtNumFirmasPorCargo, fePorCargo)
        End If
    End If
End Sub

Private Sub cboTipoPorCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregarPorCargo.SetFocus
    End If
End Sub

Private Sub cboTipoPorUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregarPorUsuario.SetFocus
    End If
End Sub

Private Sub cmdAgregarPorUsuario_Click()
    Dim oConst As COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    If Trim(txtBuscarUsuario) = "" Then
        MsgBox "Falta ingresar el usuario", vbInformation, "Aviso"
        txtBuscarUsuario.SetFocus
        Exit Sub
    End If
    If Trim(cboTipoPorUsuario.Text) = "" Then
        MsgBox "Falta ingresar el tipo", vbInformation, "Aviso"
        cboTipoPorUsuario.SetFocus
        Exit Sub
    End If
    
    For i = 1 To fePorUsuario.Rows - 1
        If fePorUsuario.TextMatrix(i, 0) <> "" Then
            If Trim(fePorUsuario.TextMatrix(i, 1)) = Right(Trim(txtBuscarUsuario), 6) Then
                MsgBox "El usuario ya fue ingresado", vbInformation, "Aviso"
                txtBuscarUsuario.SetFocus
                Exit Sub
            End If
        End If
    Next i
    
    Set oConst = New COMDConstantes.DCOMConstantes
    Set rs = oConst.ObtenerUsuariosArea(Right(Trim(txtBuscarUsuario), 4))
    
    If Not rs.EOF Then
        fePorUsuario.AdicionaFila
        fePorUsuario.TextMatrix(fePorUsuario.Row, 1) = Right(Trim(txtBuscarUsuario), 4)
        fePorUsuario.TextMatrix(fePorUsuario.Row, 2) = rs!Descripcion
        fePorUsuario.TextMatrix(fePorUsuario.Row, 3) = IIf(Trim(cboTipoPorUsuario.Text) = "N", "Nesesario", "Opcional")
        Call VerificaNumFirmas(False, txtNumFirmasPorUsuario, fePorUsuario)
        txtBuscarUsuario = ""
        cboTipoPorUsuario.ListIndex = -1
    Else
        MsgBox "El codigo ingresado no existe", vbInformation, "Aviso"
        txtBuscarUsuario.SetFocus
    End If
End Sub

Private Sub txtNivApr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        OptCargoUsu(1).SetFocus
    End If
End Sub

Private Sub txtNivApr_LostFocus()
    txtNivApr.Text = UCase(txtNivApr.Text)
End Sub

Private Sub txtNumFirmasPorCargo_Change()
    If CDbl(txtNumFirmasPorCargo.Text) > 100 Then
        txtNumFirmasPorCargo.Text = Replace(Mid(txtNumFirmasPorCargo.Text, 1, Len(txtNumFirmasPorCargo.Text) - 1), ",", "")
    End If
End Sub

Private Sub txtNumFirmasPorCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 0
    End If
    If KeyAscii = 13 Then
        chkCorrigeSug.SetFocus
    End If
End Sub

Private Sub txtNumFirmasPorCargo_LostFocus()
    Call VerificaNumFirmas(True, txtNumFirmasPorCargo, fePorCargo)
End Sub

Private Sub VerificaNumFirmas(ByVal pbVerificaReg As Boolean, ByVal pNroFirmas As EditMoney, ByVal pFlex As FlexEdit)
    If pbVerificaReg = False Then
        Dim fnNes As Integer, fnOpc As Integer, i As Integer
        fnNes = 0
        fnOpc = 0
        For i = 1 To pFlex.Rows - 1
            If pFlex.TextMatrix(i, 0) <> "" Then
                If Left(Trim(pFlex.TextMatrix(i, 3)), 1) = "N" Then
                    fnNes = fnNes + 1
                Else
                    fnOpc = fnOpc + 1
                End If
            End If
        Next i
        If fnNes = 0 And fnOpc = 0 Then
            pNroFirmas.Text = 0
        ElseIf fnNes = 0 And fnOpc <> 0 Then
            If CInt(pNroFirmas.Text) > 0 Then
                pNroFirmas.Text = pNroFirmas.Text
            Else
                pNroFirmas.Text = 1
            End If
        ElseIf fnNes <> 0 Then
            If CInt(pNroFirmas.Text) < fnNes Then
                pNroFirmas.Text = fnNes
            End If
        End If
    Else
        If pNroFirmas.value > pFlex.Rows - 1 Then
            MsgBox "El nro de Firmas no puede ser mayor a la cantidad de registros de la lista", vbInformation, "Aviso"
        End If
    End If
    pNroFirmas.Text = CInt(pNroFirmas)
End Sub

Private Sub cmdGuardarNivApr_Click()
    If ValidaDatosNivApr Then
        Dim oNivApr As COMNCredito.NCOMNivelAprobacion
        Dim MatValores() As String, i As Integer
        fnRowNivApr = feNivApr.Row
        ReDim MatValores(IIf(fnTipoReg = 1, fePorCargo.Rows, fePorUsuario.Rows) - 1, 2)
        For i = 1 To IIf(fnTipoReg = 1, fePorCargo.Rows, fePorUsuario.Rows) - 1
            MatValores(i - 1, 0) = Trim(IIf(fnTipoReg = 1, fePorCargo.TextMatrix(i, 1), fePorUsuario.TextMatrix(i, 1)))
            MatValores(i - 1, 1) = Left(Trim(IIf(fnTipoReg = 1, fePorCargo.TextMatrix(i, 3), fePorUsuario.TextMatrix(i, 3))), 1)
        Next
        
        Set oNivApr = New COMNCredito.NCOMNivelAprobacion
        
        If oNivApr.VerificaDatosNivApr(IIf(fbNuevo, "", feNivApr.TextMatrix(feNivApr.Row, 1)), Trim(txtNivApr.Text)) = False Then
            If MsgBox("¿Está seguro de " + IIf(fbNuevo = True, "registrar", "actualizar") + " los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
            If fbNuevo Then
                Call oNivApr.dInsertaNivApr(Trim(txtNivApr.Text), fnTipoReg, IIf(fnTipoReg = 1, txtNumFirmasPorCargo.Text, txtNumFirmasPorUsuario.Text), _
                                            chkCorrigeSug.value, chkValidaAg.value, MatValores)
            ElseIf fbActualiza Then
                Call oNivApr.dActualizaNivApr(feNivApr.TextMatrix(feNivApr.Row, 1), txtNivApr.Text, fnTipoReg, _
                                              IIf(fnTipoReg = 1, txtNumFirmasPorCargo.Text, txtNumFirmasPorUsuario.Text), _
                                              chkCorrigeSug.value, chkValidaAg.value, MatValores)
            End If
        Else
            MsgBox "Existe un Nivel ya ingresado con los mismos datos, favor de verificar", vbInformation, "Aviso"
            Exit Sub
        End If
        MsgBox "Los datos se " & IIf(fbNuevo = True, "registraron", "actualizaron") & " correctamente", vbInformation, "Aviso"
        Call cmdCancelaNivApr_Click
        CargaDatosNiveles
        feNivApr.TopRow = fnRowNivApr
        feNivApr.Row = fnRowNivApr
    End If
End Sub

Private Sub txtNumFirmasPorUsuario_Change()
    If CDbl(txtNumFirmasPorUsuario.Text) > 100 Then
        txtNumFirmasPorUsuario.Text = Replace(Mid(txtNumFirmasPorUsuario.Text, 1, Len(txtNumFirmasPorUsuario.Text) - 1), ",", "")
    End If
End Sub

Private Sub txtNumFirmasPorUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkCorrigeSug.SetFocus
    End If
End Sub

Private Sub txtNumFirmasPorUsuario_LostFocus()
    Call VerificaNumFirmas(True, txtNumFirmasPorUsuario, fePorCargo)
End Sub

Private Function ValidaDatosNivApr() As Boolean
    ValidaDatosNivApr = False
    
    If Trim(txtNivApr.Text) = "" Then
        MsgBox "Debe detallar el nombre del Nivel de Aprobación", vbInformation, "Aviso"
        txtNivApr.SetFocus
        ValidaDatosNivApr = False
        Exit Function
    End If
    If OptCargoUsu(1).value = 0 And OptCargoUsu(2).value = 0 Then
        MsgBox "Falta elegir si el registro será por Cargo o Por Usuario", vbInformation, "Aviso"
        OptCargoUsu(1).SetFocus
        ValidaDatosNivApr = False
        Exit Function
    End If
    
    If fnTipoReg = 1 Then
        Call VerificaNumFirmas(False, txtNumFirmasPorCargo, fePorCargo)
        If ValidaDatosGrid(fePorCargo, "Debe ingresar al menos un cargo", "Faltan datos en la lista de Cargos", 4) = False Then
            ValidaDatosNivApr = False
            Exit Function
        End If
    Else
        Call VerificaNumFirmas(False, txtNumFirmasPorUsuario, fePorUsuario)
        If ValidaDatosGrid(fePorUsuario, "Debe ingresar al menos un usuario", "Faltan datos en la lista de Usuarios", 4) = False Then
            ValidaDatosNivApr = False
            Exit Function
        End If
    End If
    ValidaDatosNivApr = True
End Function

Private Sub cmdCancelaNivApr_Click()
    fnRowNivApr = feNivApr.Row
    txtNivApr.Text = ""
    If OptCargoUsu(1).value = True Then
        Call LimpiaFlex(fePorCargo)
        FraPorCargo.Enabled = False
    Else
        Call LimpiaFlex(fePorUsuario)
        FraPorUsuario.Enabled = False
    End If
    OptCargoUsu(1).value = False
    OptCargoUsu(2).value = False
    txtBuscarCargo = ""
    txtBuscarUsuario = ""
    cboTipoPorCargo.ListIndex = -1
    cboTipoPorUsuario.ListIndex = -1
    txtNumFirmasPorCargo.Text = 0
    txtNumFirmasPorUsuario.Text = 0
    chkCorrigeSug.value = 0
    chkValidaAg.value = 0
    fbNuevo = True
    fbActualiza = False
    cmdEditarNivApr.Enabled = True
    cmdEliminarNivApr.Enabled = True
    feNivApr.Enabled = True
    cmdSubir.Enabled = True
    cmdBajar.Enabled = True
    CargaDatosNiveles
    feNivApr.TopRow = fnRowNivApr
    feNivApr.Row = fnRowNivApr
End Sub

Private Sub cmdCerrarNivApr_Click()
    Unload Me
End Sub

Private Sub cmdEditarNivApr_Click()
    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    fnRowNivApr = feNivApr.Row
    fbActualiza = True
    fbNuevo = False
    cmdEditarNivApr.Enabled = False
    cmdEliminarNivApr.Enabled = False
    feNivApr.Enabled = False
    cmdSubir.Enabled = False
    cmdBajar.Enabled = False
    
    Set oNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oNiv.RecuperaNivApr(feNivApr.TextMatrix(feNivApr.Row, 1))
    Set oNiv = Nothing

    txtNivApr.Text = rs!cNivAprDesc
    OptCargoUsu(rs!nTipoReg).value = True
    Call LimpiaFlex(IIf(rs!nTipoReg = 1, fePorCargo, fePorUsuario))
    LlenaGrid (IIf(rs!nTipoReg = 1, fePorCargo, fePorUsuario))
    If rs!nTipoReg = 1 Then
        txtNumFirmasPorCargo = rs!nNumCantFirmas
    Else
        txtNumFirmasPorUsuario = rs!nNumCantFirmas
    End If
    chkCorrigeSug.value = rs!nCorrigeSug
    chkValidaAg.value = IIf(rs!cValidaAg = "SI", 1, 0)
    feNivApr.TopRow = fnRowNivApr
    feNivApr.Row = fnRowNivApr
End Sub

Private Sub LlenaGrid(ByVal pFlex As FlexEdit)
    Dim oLista As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset, lnFila As Integer
    Set oLista = New COMDCredito.DCOMNivelAprobacion
    Set rs = oLista.RecuperaNivAprValores(feNivApr.TextMatrix(feNivApr.Row, 1))
    Set oLista = Nothing
    Do While Not rs.EOF
        pFlex.AdicionaFila
        lnFila = pFlex.Row
        pFlex.TextMatrix(lnFila, 1) = rs!cValorCod
        pFlex.TextMatrix(lnFila, 2) = rs!cValorDesc
        pFlex.TextMatrix(lnFila, 3) = IIf(rs!cTipoCod = "N", "Necesario", "Opcional")
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdEliminarNivApr_Click()
    If feNivApr.TextMatrix(feNivApr.Row, 0) <> "" Then
        If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feNivApr.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Dim oNivApr As COMNCredito.NCOMNivelAprobacion
            Set oNivApr = New COMNCredito.NCOMNivelAprobacion
            Call oNivApr.dEliminaNivApr(feNivApr.TextMatrix(feNivApr.Row, 1))
            feNivApr.EliminaFila feNivApr.Row
            CargaDatosNiveles
        End If
    End If
End Sub

Private Sub cmdSubir_Click()
    Dim oNiveles As COMNCredito.NCOMNivelAprobacion
    Dim nRow As Integer
    Set oNiveles = New COMNCredito.NCOMNivelAprobacion
    nRow = feNivApr.Row
    If feNivApr.TextMatrix(feNivApr.Row, 0) <> 1 Then
        Call oNiveles.dActualizaOrdenNivel(feNivApr.TextMatrix(feNivApr.Row, 1), feNivApr.TextMatrix(feNivApr.Row - 1, 1), _
                                        CInt(feNivApr.TextMatrix(feNivApr.Row, 0)) - 1, CInt(feNivApr.TextMatrix(feNivApr.Row, 0)))
        CargaDatosNiveles
        feNivApr.TopRow = nRow - 1
        feNivApr.Row = nRow - 1
    End If
    Set oNiveles = Nothing
End Sub

Private Sub cmdBajar_Click()
    Dim oNiveles As COMNCredito.NCOMNivelAprobacion
    Dim oDNiv As COMDCredito.DCOMNivelAprobacion
    Dim psMaxOrden As Integer, nRow As Integer
    Set oNiveles = New COMNCredito.NCOMNivelAprobacion
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    psMaxOrden = oDNiv.ObtenerUltimoOrdenNivel()
    nRow = feNivApr.Row
    If feNivApr.TextMatrix(feNivApr.Row, 0) < psMaxOrden Then
        Call oNiveles.dActualizaOrdenNivel(feNivApr.TextMatrix(feNivApr.Row, 1), feNivApr.TextMatrix(feNivApr.Row + 1, 1), _
                                            CInt(feNivApr.TextMatrix(feNivApr.Row, 0)) + 1, CInt(feNivApr.TextMatrix(feNivApr.Row, 0)))
        CargaDatosNiveles
        feNivApr.TopRow = nRow + 1
        feNivApr.Row = nRow + 1
    End If
    Set oNiveles = Nothing
    Set oDNiv = Nothing
End Sub

Private Sub feNivApr_Click()
    If feNivApr.TextMatrix(feNivApr.Row, feNivApr.Col) <> "" Then
        MuestraDatosGridNivApr
    End If
End Sub

Private Sub feNivApr_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If feNivApr.TextMatrix(feNivApr.Row, feNivApr.Col) <> "" Then
            MuestraDatosGridNivApr
        End If
    End If
End Sub

Private Sub MuestraDatosGridNivApr()
    Dim oLista As COMDCredito.DCOMNivelAprobacion
    Dim rsDatos As ADODB.Recordset
    Dim MatTitulos() As String
    ReDim MatTitulos(1, 2)
    MatTitulos(0, 0) = feNivApr.TextMatrix(feNivApr.Row, 4)
    MatTitulos(0, 1) = "Tipo"
    If feNivApr.Col = 5 Then
        Set oLista = New COMDCredito.DCOMNivelAprobacion
            Set rsDatos = oLista.RecuperaNivAprValores(feNivApr.TextMatrix(feNivApr.Row, 1))
        Set oLista = Nothing
        frmCredListaDatos.Inicio feNivApr.TextMatrix(feNivApr.Row, 4), rsDatos, , 2, MatTitulos
    End If
End Sub
