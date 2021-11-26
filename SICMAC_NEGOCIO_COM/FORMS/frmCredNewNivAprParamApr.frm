VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredNewNivAprParamApr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de Aprobación de Crédito por grupo, nivel y monto"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "frmCredNewNivAprParamApr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Clientes ordinarios"
      TabPicture(0)   =   "frmCredNewNivAprParamApr.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FlexEdit1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "feParam"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdGrabar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdEliminar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdEditar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdCancelar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Clientes Preferenciales"
      TabPicture(1)   =   "frmCredNewNivAprParamApr.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPCancelar"
      Tab(1).Control(1)=   "cmdPEditar"
      Tab(1).Control(2)=   "cmdPEliminar"
      Tab(1).Control(3)=   "cmdPGrabar"
      Tab(1).Control(4)=   "Frame12"
      Tab(1).Control(5)=   "Frame11"
      Tab(1).Control(6)=   "Frame10"
      Tab(1).Control(7)=   "Frame9"
      Tab(1).Control(8)=   "Frame8"
      Tab(1).Control(9)=   "Frame1"
      Tab(1).Control(10)=   "fePParam"
      Tab(1).Control(11)=   "FlexEdit3"
      Tab(1).ControlCount=   12
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
         Height          =   345
         Left            =   6480
         TabIndex        =   51
         Top             =   5280
         Width           =   930
      End
      Begin VB.CommandButton cmdPCancelar 
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
         Left            =   -68520
         TabIndex        =   50
         Top             =   5280
         Width           =   930
      End
      Begin VB.CommandButton cmdPEditar 
         Caption         =   "Editar"
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
         Height          =   345
         Left            =   -74760
         TabIndex        =   49
         Top             =   5280
         Width           =   930
      End
      Begin VB.CommandButton cmdPEliminar 
         Caption         =   "Eliminar"
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
         Height          =   345
         Left            =   -73800
         TabIndex        =   48
         Top             =   5280
         Width           =   930
      End
      Begin VB.CommandButton cmdPGrabar 
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
         Left            =   -67560
         TabIndex        =   0
         Top             =   5280
         Width           =   930
      End
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -74730
         TabIndex        =   46
         Top             =   3240
         Width           =   590
         Begin VB.Label Label12 
            Caption         =   " Nivel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   47
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -74130
         TabIndex        =   44
         Top             =   3240
         Width           =   2080
         Begin VB.Label Label11 
            Caption         =   "Firmas"
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
            Left            =   720
            TabIndex        =   45
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   " Niveles a Aprobar "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74760
         TabIndex        =   39
         Top             =   1200
         Width           =   3255
         Begin VB.ListBox lstPNiveles 
            Height          =   1185
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   41
            Top             =   480
            Width           =   3015
         End
         Begin VB.CheckBox chkPTodosNivel 
            Caption         =   "Todos"
            Height          =   200
            Left            =   160
            TabIndex        =   40
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   " Riesgo 1"
         Height          =   1815
         Left            =   -71280
         TabIndex        =   34
         Top             =   1200
         Width           =   2235
         Begin SICMACT.EditMoney txtPR1Desde 
            Height          =   300
            Left            =   360
            TabIndex        =   35
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
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
         End
         Begin SICMACT.EditMoney txtPR1Hasta 
            Height          =   300
            Left            =   360
            TabIndex        =   36
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
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
         Begin VB.Label Label10 
            Caption         =   "Hasta :"
            Height          =   255
            Left            =   150
            TabIndex        =   38
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label9 
            Caption         =   "Desde :"
            Height          =   255
            Left            =   150
            TabIndex        =   37
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   " Riesgo 2 "
         Height          =   1815
         Left            =   -68835
         TabIndex        =   29
         Top             =   1200
         Width           =   2235
         Begin SICMACT.EditMoney txtPR2Desde 
            Height          =   300
            Left            =   360
            TabIndex        =   30
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
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
         End
         Begin SICMACT.EditMoney txtPR2Hasta 
            Height          =   300
            Left            =   360
            TabIndex        =   31
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
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
         Begin VB.Label Label8 
            Caption         =   "Hasta :"
            Height          =   255
            Left            =   180
            TabIndex        =   33
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Desde :"
            Height          =   255
            Left            =   180
            TabIndex        =   32
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Grupo de Crédito : "
         Height          =   615
         Left            =   -74760
         TabIndex        =   27
         Top             =   480
         Width           =   8295
         Begin VB.ComboBox cboPGrupoApr 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmCredNewNivAprParamApr.frx":0342
            Left            =   120
            List            =   "frmCredNewNivAprParamApr.frx":034C
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   210
            Width           =   8055
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
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
         Height          =   345
         Left            =   240
         TabIndex        =   24
         Top             =   5280
         Width           =   930
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
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
         Height          =   345
         Left            =   1200
         TabIndex        =   23
         Top             =   5280
         Width           =   930
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   870
         TabIndex        =   21
         Top             =   3240
         Width           =   2080
         Begin VB.Label Label3 
            Caption         =   "Firmas"
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
            Left            =   720
            TabIndex        =   22
            Top             =   120
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   270
         TabIndex        =   19
         Top             =   3240
         Width           =   590
         Begin VB.Label Label1 
            Caption         =   " Nivel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   20
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdGrabar 
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
         Left            =   7440
         TabIndex        =   18
         Top             =   5280
         Width           =   930
      End
      Begin VB.Frame Frame4 
         Caption         =   " Niveles a Aprobar "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   3255
         Begin VB.CheckBox chkTodosNivel 
            Caption         =   "Todos"
            Height          =   200
            Left            =   160
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
         Begin VB.ListBox lstNiveles 
            Height          =   1185
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   16
            Top             =   480
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   " Riesgo 1"
         Height          =   1815
         Left            =   3720
         TabIndex        =   10
         Top             =   1200
         Width           =   2235
         Begin SICMACT.EditMoney txtDesdeR1 
            Height          =   300
            Left            =   360
            TabIndex        =   11
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
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
         End
         Begin SICMACT.EditMoney txtHastaR1 
            Height          =   300
            Left            =   360
            TabIndex        =   12
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
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
            Caption         =   "Desde :"
            Height          =   255
            Left            =   150
            TabIndex        =   14
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Hasta :"
            Height          =   255
            Left            =   150
            TabIndex        =   13
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   " Riesgo 2 "
         Height          =   1815
         Left            =   6160
         TabIndex        =   5
         Top             =   1200
         Width           =   2235
         Begin SICMACT.EditMoney txtDesdeR2 
            Height          =   300
            Left            =   360
            TabIndex        =   6
            Top             =   600
            Width           =   1575
            _ExtentX        =   2778
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
         End
         Begin SICMACT.EditMoney txtHastaR2 
            Height          =   300
            Left            =   360
            TabIndex        =   7
            Top             =   1200
            Width           =   1575
            _ExtentX        =   2778
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
         Begin VB.Label Label5 
            Caption         =   "Desde :"
            Height          =   255
            Left            =   180
            TabIndex        =   9
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Hasta :"
            Height          =   255
            Left            =   180
            TabIndex        =   8
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   " Grupo de Crédito : "
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   8175
         Begin VB.ComboBox cboGrupoApr 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmCredNewNivAprParamApr.frx":03A4
            Left            =   120
            List            =   "frmCredNewNivAprParamApr.frx":03AE
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   210
            Width           =   7935
         End
      End
      Begin SICMACT.FlexEdit feParam 
         Height          =   1695
         Left            =   240
         TabIndex        =   25
         Top             =   3435
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2990
         Cols0           =   8
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "-cParamCod-Firmas-Desde-Hasta-Desde-Hasta-Aux"
         EncabezadosAnchos=   "600-0-2100-1300-1300-1300-1300-0"
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
         ColumnasAEditar =   "X-X-X-3-4-5-6-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-2-2-2-2-0"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit FlexEdit1 
         Height          =   1095
         Left            =   240
         TabIndex        =   26
         Top             =   3120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1931
         Cols0           =   4
         ScrollBars      =   0
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "-Firmas-Riesgo 1 (S/.)-Riesgo 2 (S/.)"
         EncabezadosAnchos=   "600-2100-2610-2610"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C"
         FormatosEdit    =   "0-0-0-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit fePParam 
         Height          =   1695
         Left            =   -74760
         TabIndex        =   42
         Top             =   3440
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   2990
         Cols0           =   8
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "-cParamCod-Firmas-Desde-Hasta-Desde-Hasta-Aux"
         EncabezadosAnchos=   "600-0-2100-1300-1300-1300-1300-0"
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
         ColumnasAEditar =   "X-X-X-3-4-5-6-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-R-R-R-R-R-C"
         FormatosEdit    =   "0-0-0-2-2-2-2-0"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit FlexEdit3 
         Height          =   1095
         Left            =   -74760
         TabIndex        =   43
         Top             =   3120
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   1931
         Cols0           =   4
         ScrollBars      =   0
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "-Firmas-Riesgo 1 (S/.)-Riesgo 2 (S/.)"
         EncabezadosAnchos=   "600-2100-2610-2610"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C"
         FormatosEdit    =   "0-0-0-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdCerrar 
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
      Left            =   7800
      TabIndex        =   1
      Top             =   6000
      Width           =   930
   End
End
Attribute VB_Name = "frmCredNewNivAprParamApr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredNewNivAprParamApr
'** Descripción : Formulario para la Administracion de los Paramámetros de Aprobación creado segun RFC110-2012
'** Creación : JUEZ, 20121203 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oNiveles As COMDCredito.DCOMNivelAprobacion
Dim oNNiv As COMNCredito.NCOMNivelAprobacion
Dim rs As ADODB.Recordset
Dim fbEdita As Boolean

'RECO20140129 ERS173-2014**********'
Dim nClientePreferencial As Integer
Dim nClienteOrdinario As Integer
'RECO FIN**************************'

Private Sub cboPGrupoApr_Click()
    Call cboControl_Click(cboPGrupoApr, fePParam, txtPR1Desde, txtPR2Desde, lstPNiveles, cmdPEditar, cmdPEliminar, nClientePreferencial)
End Sub
Private Sub chkPTodosNivel_Click()
    Call CheckLista(IIf(chkPTodosNivel.value = 1, True, False), lstPNiveles)
End Sub
Private Sub chkTodosNivel_Click()
    Call CheckLista(IIf(chkTodosNivel.value = 1, True, False), lstNiveles)
End Sub
Private Sub CheckLista(ByVal bCheck As Boolean, ByVal lstLista As ListBox)
    Dim i As Integer
    For i = 0 To lstLista.ListCount - 1
        lstLista.Selected(i) = bCheck
    Next i
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaControles
End Sub

Private Sub CmdEditar_Click()
    Call Editar(lstNiveles, feParam.TextMatrix(feParam.row, 1), txtDesdeR1, txtHastaR1, txtDesdeR2, txtHastaR2, cmdEditar, feParam, cboGrupoApr)
    cmdEliminar.Enabled = False
End Sub
Private Sub cmdGrabar_Click()
    If ValidaDatos(lstNiveles, cboGrupoApr, txtDesdeR1, txtHastaR1, txtDesdeR2, txtHastaR2) Then
        Call Grabar(nClienteOrdinario, lstNiveles, feParam.TextMatrix(feParam.row, 1), Trim(Right(cboGrupoApr.Text, 10)), _
                    txtDesdeR1.Text, txtHastaR1.Text, txtDesdeR2.Text, txtHastaR2.Text)
    End If
End Sub
Private Sub LimpiaControles()
    Call ListaNiveles(lstNiveles)
    Call ListaNiveles(lstPNiveles)
    fbEdita = False
    feParam.Enabled = True
    fePParam.Enabled = True
    txtHastaR1.Text = Format(0, "#,##0.00")
    txtHastaR2.Text = Format(0, "#,##0.00")
    txtPR1Hasta.Text = Format(0, "#,##0.00")
    txtPR2Hasta.Text = Format(0, "#,##0.00")
    cmdEditar.Enabled = True
    cmdPEditar.Enabled = True
    feParam.Enabled = True
    fePParam.Enabled = True
    cboPGrupoApr.Enabled = True
    cboGrupoApr.Enabled = True
    Call cboPGrupoApr_Click
    Call cboGrupoApr_Click
    cmdEliminar.Enabled = True
    cmdPEliminar.Enabled = True
End Sub

Private Sub cmdPCancelar_Click()
    Call LimpiaControles
End Sub

Private Sub cmdPEditar_Click()
    Call Editar(lstPNiveles, fePParam.TextMatrix(fePParam.row, 1), txtPR1Desde, txtPR1Hasta, txtPR2Desde, txtPR2Hasta, cmdPEditar, fePParam, cboPGrupoApr)
    cmdPEliminar.Enabled = False
End Sub

Private Sub cmdPEliminar_Click()
    Call EliminarClick(fePParam)
    Call cboControl_Click(cboPGrupoApr, fePParam, txtPR1Desde, txtPR2Desde, lstPNiveles, cmdPEditar, cmdPEliminar, nClientePreferencial)
End Sub

Private Sub cmdPGrabar_Click()
    If ValidaDatos(lstPNiveles, cboPGrupoApr, txtPR1Desde, txtPR1Hasta, txtPR2Desde, txtPR2Hasta) Then
        Call Grabar(nClientePreferencial, lstPNiveles, fePParam.TextMatrix(fePParam.row, 1), Trim(Right(cboPGrupoApr.Text, 10)), _
                    txtPR1Desde.Text, txtPR1Hasta.Text, txtPR2Desde.Text, txtPR2Hasta.Text)
    End If
End Sub

Private Sub feParam_Click()
    Call Flex_Click(feParam)
End Sub
Private Sub feParam_KeyPress(KeyAscii As Integer)
    Call Flex_KeyPress(feParam)
End Sub
Private Sub fePParam_Click()
    Call Flex_Click(fePParam)
End Sub
Private Sub fePParam_KeyPress(KeyAscii As Integer)
    Call Flex_KeyPress(fePParam)
End Sub
Private Sub Form_Load()
    'RECO20140129 ERS173-2014***************************
    nClientePreferencial = 2
    nClienteOrdinario = 1
    'RECO FIN*******************************************
    Call CargarCombo(cboGrupoApr, lstNiveles)
    Call CargarCombo(cboPGrupoApr, lstPNiveles)
    fbEdita = False
End Sub
Private Sub cboGrupoApr_Click()
    Call cboControl_Click(cboGrupoApr, feParam, txtDesdeR1, txtDesdeR2, lstNiveles, cmdEditar, cmdEliminar, nClienteOrdinario)
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    Call EliminarClick(feParam)
    Call cboControl_Click(cboGrupoApr, feParam, txtDesdeR1, txtDesdeR2, lstNiveles, cmdEditar, cmdEliminar, nClienteOrdinario)
End Sub
Private Sub cmdNuevo_Click()
    Call Nuevo_Click(feParam)
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Call cmdPCancelar_Click
    Call cmdCancelar_Click
End Sub

Private Sub txtDesdeR1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtHastaR1.SetFocus
    End If
End Sub
Private Sub txtDesdeR2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtHastaR2.SetFocus
    End If
End Sub
Private Sub txtHastaR1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtHastaR2.SetFocus
        txtHastaR2.MarcaTexto
    End If
End Sub
Private Sub txtHastaR2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    End If
End Sub
Private Function ValidaDatos(ByVal plstControl As ListBox, ByVal pcboControl As ComboBox, ByVal ptxtR1Desde As EditMoney, _
                             ByVal ptxtR1Hasta As EditMoney, ByVal ptxtR2Desde As EditMoney, ByVal ptxtR2Hasta As EditMoney) As Boolean
    Dim CNiv As Integer
    ValidaDatos = False
    
    CNiv = DevuelveCantidadCheckList(plstControl)
    
    If Trim(pcboControl.Text) = "" Then
        MsgBox "Debe elegir un Grupo de Aprobación", vbInformation, "Aviso"
        cboGrupoApr.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If CNiv = 0 Then
        MsgBox "Debe seleccionar al menos un Nivel a aprobar", vbInformation, "Aviso"
        plstControl.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    If CDbl(ptxtR1Hasta.Text) <= 0 Then
        MsgBox "Debe ingresar el tamaño maximo del Riesgo 1", vbInformation, "Aviso"
        'ptxtR1Hasta.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If CDbl(ptxtR2Hasta.Text) <= 0 Then
        MsgBox "Debe ingresar el tamaño maximo del Riesgo 2", vbInformation, "Aviso"
        'ptxtR2Hasta.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If CDbl(ptxtR1Hasta.Text) < CDbl(ptxtR1Desde.Text) Then
        MsgBox "El valor Hasta debe ser mayor al valor Desde en el Riesgo 1", vbInformation, "Aviso"
        'ptxtR1Hasta.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    If CDbl(ptxtR2Hasta.Text) < CDbl(ptxtR2Desde.Text) Then
        MsgBox "El valor Hasta debe ser mayor al valor Desde en el Riesgo 2", vbInformation, "Aviso"
        'ptxtR2Hasta.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    ValidaDatos = True
End Function

'RECO20150129 ERS173-2014*********************************************
Public Sub Grabar(ByVal pnTpoCliente As Integer, ByVal plstControl As ListBox, ByVal psParamCod As String, ByVal psGrupoCod As String, _
                    ByVal pnR1Desde As Double, ByVal pnR1Hasta As Double, ByVal pnR2Desde As Double, ByVal pnR2Hasta As Double)
    Dim psNiveles As String, i As Integer
    Set oNNiv = New COMNCredito.NCOMNivelAprobacion
        
    psNiveles = "'"
    For i = 0 To plstControl.ListCount - 1
        If plstControl.Selected(i) Then
            psNiveles = psNiveles & Trim(plstControl.List(i)) & ","
        End If
    Next i
    psNiveles = Mid(psNiveles, 1, Len(psNiveles) - 1) & "'"
        
    If MsgBox("¿Está seguro de registrar los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
    If fbEdita Then
        Call oNNiv.dActualizaParamGruposNiveles(psParamCod, psGrupoCod, psNiveles, pnR1Desde, pnR1Hasta, pnR2Desde, pnR2Hasta, pnTpoCliente)
    Else
        Call oNNiv.dInsertaParamGruposNiveles(psGrupoCod, psNiveles, pnR1Desde, pnR1Hasta, pnR2Desde, pnR2Hasta, pnTpoCliente)
    End If
    MsgBox "Los datos se registraron con éxito", vbInformation, "Aviso"
    LimpiaControles
    cboGrupoApr_Click
    cboPGrupoApr_Click
End Sub
Public Sub Editar(ByVal lstControl As ListBox, ByVal psParamCod As String, ByVal R1Desde As Control, ByVal R1Hasta As Control, _
                  ByVal R2Desde As Control, ByVal R2Hasta As Control, ByVal cmdControl As CommandButton, ByVal feControl As FlexEdit, _
                  ByVal cboControl As ComboBox)
    Dim rsLista As ADODB.Recordset, rsDatos As ADODB.Recordset
    Dim i As Integer, J As Integer
    
    Set oNiveles = New COMDCredito.DCOMNivelAprobacion
        Set rsLista = oNiveles.RecuperaNivAprListaParam(1)
        Set rsDatos = oNiveles.RecuperaParametrosGrupoNivelesApr(psParamCod)
    Set oNiveles = Nothing
    fbEdita = True
    If Not rsDatos.EOF Then
        lstControl.Clear
        For i = 0 To rsLista.RecordCount - 1
            lstControl.AddItem rsLista!cNivAprCod & " " & Trim(rsLista!cNivAprDesc)
            rsDatos.MoveFirst
            For J = 0 To rsDatos.RecordCount - 1
                If Trim(rsLista!cNivAprDesc) = Trim(rsDatos!cNivAprDesc) Then
                    lstControl.Selected(i) = True
                    Exit For
                End If
                rsDatos.MoveNext
            Next J
            rsLista.MoveNext
        Next i
        Set rsLista = Nothing
        rsDatos.MoveLast
        R1Desde.Text = Format(rsDatos!nDesdeR1, "#,##0.00")
        R1Hasta.Text = Format(rsDatos!nHastaR1, "#,##0.00")
        R2Desde.Text = Format(rsDatos!nDesdeR2, "#,##0.00")
        R2Hasta.Text = Format(rsDatos!nHastaR2, "#,##0.00")
        cmdControl.Enabled = False
        feControl.Enabled = False
        cboControl.Enabled = False
    End If
End Sub
Public Sub Flex_Click(ByVal pfeControl As FlexEdit)
    Dim rsLista As ADODB.Recordset, rsDatos As ADODB.Recordset, cDatos As String
    If pfeControl.TextMatrix(pfeControl.row, 0) <> "" Then
        If pfeControl.Col = 2 Then
            Set oNiveles = New COMDCredito.DCOMNivelAprobacion
                Set rsLista = oNiveles.RecuperaNivAprListaParam(1)
                Set rsDatos = oNiveles.RecuperaParametrosGrupoNivelesApr(pfeControl.TextMatrix(pfeControl.row, 1))
            Set oNiveles = Nothing
            frmCredListaDatos.Inicio "Niveles a Aprobar", rsDatos, rsLista, 1
        End If
    End If
End Sub
Public Sub Flex_KeyPress(ByVal pfeControl As FlexEdit)

    'If KeyAscii = 13 Then
        Dim rsLista As ADODB.Recordset, rsDatos As ADODB.Recordset, cDatos As String
        If pfeControl.TextMatrix(pfeControl.row, 0) <> "" Then
            If pfeControl.Col = 2 Then
                Set oNiveles = New COMDCredito.DCOMNivelAprobacion
                    Set rsLista = oNiveles.RecuperaNivApr()
                    Set rsDatos = oNiveles.RecuperaParametrosGrupoNivelesApr(pfeControl.TextMatrix(pfeControl.row, 1))
                Set oNiveles = Nothing
                frmCredListaDatos.Inicio "Niveles a Aprobar", rsDatos, rsLista, 1
            End If
        End If
    'End If
End Sub
Public Sub CargarCombo(ByVal pcboControl As ComboBox, ByVal plstControl As ListBox)
    Set oNiveles = New COMDCredito.DCOMNivelAprobacion
    Set rs = oNiveles.RecuperaGruposApr()
    Set oNiveles = Nothing
    pcboControl.Clear
    While Not rs.EOF
        pcboControl.AddItem rs.Fields(1) & Space(500) & rs.Fields(0)
        rs.MoveNext
    Wend
    Set rs = Nothing
    Call ListaNiveles(plstControl)
End Sub
Private Sub ListaNiveles(ByVal plstControl As ListBox)
    Set oNiveles = New COMDCredito.DCOMNivelAprobacion
    Set rs = oNiveles.RecuperaNivAprListaParam(1)
    Set oNiveles = Nothing
    If rs.EOF Then
        MsgBox " No se encuentran los Niveles de Aprobación ", vbInformation, " Aviso "
    Else
        plstControl.Clear
        With rs
            Do While Not rs.EOF
                plstControl.AddItem rs!cNivAprCod & " " & Trim(rs!cNivAprDesc)
                rs.MoveNext
            Loop
        End With
        plstControl.Selected(0) = True
    End If
    Set rs = Nothing
End Sub
Public Sub cboControl_Click(ByVal pcboControl As ComboBox, ByVal pfeControl As Control, ByVal pR1Desde As EditMoney, ByVal pR2Desde As EditMoney, _
                            ByVal plstControl As ListBox, ByVal pcmdEdit As CommandButton, ByVal pcmdElimin As CommandButton, ByVal pnTpoCliente As Integer)
    Dim lnFila As Integer
    Set oNiveles = New COMDCredito.DCOMNivelAprobacion
    Set rs = oNiveles.RecuperaParametrosGrupoApr(Trim(Right(pcboControl.Text, 10)), pnTpoCliente)
    Set oNiveles = Nothing
    Call LimpiaFlex(pfeControl)
    If Not rs.EOF Then
        Do While Not rs.EOF
            pfeControl.AdicionaFila
            lnFila = pfeControl.row
            pfeControl.TextMatrix(lnFila, 1) = rs!cParamCod
            pfeControl.TextMatrix(lnFila, 2) = "..."
            pfeControl.TextMatrix(lnFila, 3) = Format(rs!nDesdeR1, "#,##0.00")
            pfeControl.TextMatrix(lnFila, 4) = Format(rs!nHastaR1, "#,##0.00")
            pfeControl.TextMatrix(lnFila, 5) = Format(rs!nDesdeR2, "#,##0.00")
            pfeControl.TextMatrix(lnFila, 6) = Format(rs!nHastaR2, "#,##0.00")
            rs.MoveNext
            pcmdEdit.Enabled = True
            pcmdElimin.Enabled = True
        Loop
        pfeControl.TopRow = 1
        rs.MoveLast
        pR1Desde.Text = Format(CDbl(rs!nHastaR1) + 0.01, "#,##0.00")
        pR2Desde.Text = Format(CDbl(rs!nHastaR2) + 0.01, "#,##0.00")
    Else
        pR1Desde.Text = Format(0, "#,##0.00")
        pR2Desde.Text = Format(0, "#,##0.00")
        pcmdEdit.Enabled = False
        pcmdElimin.Enabled = False
    End If
    Call ListaNiveles(plstControl)
    plstControl.SetFocus
End Sub
Public Sub EliminarClick(ByVal pfeControl As FlexEdit)
    If pfeControl.TextMatrix(pfeControl.row, 0) <> "" Then
        If MsgBox("¿Está seguro de eliminar los datos?  (Se eliminara la última fila)", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Dim oNivApr As COMNCredito.NCOMNivelAprobacion
            Set oNivApr = New COMNCredito.NCOMNivelAprobacion
            Call oNivApr.dEliminaParamApr(pfeControl.TextMatrix(pfeControl.Rows - 1, 1))
            'pfeControl.EliminaFila pfeControl.row - 1
        End If
    End If
End Sub
Public Sub Nuevo_Click(ByVal pfeControl As FlexEdit)
    pfeControl.AdicionaFila
    If pfeControl.TextMatrix(feParam.Rows - 1, 0) <> "1" Then
        If pfeControl.TextMatrix(pfeControl.Rows - 2, 4) <> "" Or pfeControl.TextMatrix(pfeControl.Rows - 2, 6) <> "" Then
            pfeControl.TextMatrix(pfeControl.Rows - 1, 3) = Format(CDbl(pfeControl.TextMatrix(pfeControl.Rows - 2, 4)) + 0.01, "#,##0.00")
            pfeControl.TextMatrix(pfeControl.Rows - 1, 5) = Format(CDbl(pfeControl.TextMatrix(pfeControl.Rows - 2, 6)) + 0.01, "#,##0.00")
        Else
            pfeControl.TextMatrix(pfeControl.Rows - 1, 3) = Format(0, "#,##0.00")
            pfeControl.TextMatrix(pfeControl.Rows - 1, 5) = Format(0, "#,##0.00")
        End If
    Else
        pfeControl.TextMatrix(pfeControl.Rows - 1, 3) = Format(0, "#,##0.00")
        pfeControl.TextMatrix(pfeControl.Rows - 1, 5) = Format(0, "#,##0.00")
    End If
    pfeControl.TextMatrix(pfeControl.Rows - 1, 2) = "..."
    'pfeControl.SetFocus
    cmdEliminar.Enabled = True
End Sub

Private Sub txtPR1Desde_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
        txtPR1Hasta.SetFocus
    End If
End Sub

Private Sub txtPR1Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPR2Hasta.SetFocus
        txtPR2Hasta.MarcaTexto
    End If
End Sub

Private Sub txtPR2Hasta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdPGrabar.SetFocus
    End If
End Sub
'RECO FIN***************************************************j**********
