VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProyectadoVsEjecutadoPorAgencia 
   Caption         =   "Proyectado Vs Ejecutado por Agencia"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9705
   Icon            =   "frmProyectadoVsEjecutadoPorAgencia.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   7200
      TabIndex        =   9
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   8520
      TabIndex        =   8
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Dólares Americanos"
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
      Height          =   3255
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   9495
      Begin TabDlg.SSTab SSTab1 
         Height          =   2535
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4471
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   520
         TabCaption(0)   =   "Semana 1"
         TabPicture(0)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "feMESemana(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Semana 2"
         TabPicture(1)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "feMESemana(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Semana 3"
         TabPicture(2)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":0342
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "feMESemana(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Semana 4"
         TabPicture(3)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":035E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "feMESemana(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Semana 5"
         TabPicture(4)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":037A
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "feMESemana(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Semana 6"
         TabPicture(5)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":0396
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "feMESemana(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Resumen"
         TabPicture(6)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":03B2
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "feMESemana(6)"
         Tab(6).ControlCount=   1
         Begin SICMACT.FlexEdit feMESemana 
            Height          =   1695
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMESemana 
            Height          =   1695
            Index           =   1
            Left            =   -74760
            TabIndex        =   23
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMESemana 
            Height          =   1695
            Index           =   2
            Left            =   -74760
            TabIndex        =   24
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMESemana 
            Height          =   1695
            Index           =   3
            Left            =   -74760
            TabIndex        =   25
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMESemana 
            Height          =   1695
            Index           =   4
            Left            =   -74760
            TabIndex        =   26
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMESemana 
            Height          =   1695
            Index           =   5
            Left            =   -74760
            TabIndex        =   27
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMESemana 
            Height          =   1695
            Index           =   6
            Left            =   -74760
            TabIndex        =   28
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nuevo Soles"
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
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   9495
      Begin TabDlg.SSTab TabSoles 
         Height          =   2535
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4471
         _Version        =   393216
         Tabs            =   7
         TabsPerRow      =   7
         TabHeight       =   520
         TabCaption(0)   =   "Semana 1"
         TabPicture(0)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":03CE
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "feMNSemana(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Semana 2"
         TabPicture(1)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":03EA
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "feMNSemana(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Semana 3"
         TabPicture(2)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":0406
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "feMNSemana(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Semana 4"
         TabPicture(3)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":0422
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "feMNSemana(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Semana 5"
         TabPicture(4)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":043E
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "feMNSemana(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Semana 6"
         TabPicture(5)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":045A
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "feMNSemana(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Resumen"
         TabPicture(6)   =   "frmProyectadoVsEjecutadoPorAgencia.frx":0476
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "feMNSemana(6)"
         Tab(6).ControlCount=   1
         Begin SICMACT.FlexEdit feMNSemana 
            Height          =   1695
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMNSemana 
            Height          =   1695
            Index           =   1
            Left            =   -74760
            TabIndex        =   17
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMNSemana 
            Height          =   1695
            Index           =   2
            Left            =   -74760
            TabIndex        =   18
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMNSemana 
            Height          =   1695
            Index           =   3
            Left            =   -74760
            TabIndex        =   19
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMNSemana 
            Height          =   1695
            Index           =   4
            Left            =   -74760
            TabIndex        =   20
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMNSemana 
            Height          =   1695
            Index           =   5
            Left            =   -74760
            TabIndex        =   21
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin SICMACT.FlexEdit feMNSemana 
            Height          =   1695
            Index           =   6
            Left            =   -74760
            TabIndex        =   22
            Top             =   480
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2990
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Concepto-Proyeccion (A)-Ejecutado (B)-Diferencia A y B-% Cumplimiento"
            EncabezadosAnchos=   "0-2000-1500-1500-1500-1500"
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
            EncabezadosAlineacion=   "C-L-R-R-R-R"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione"
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9495
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtAnio 
         Height          =   315
         Left            =   4920
         MaxLength       =   4
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmProyectadoVsEjecutadoPorAgencia.frx":0492
         Left            =   2880
         List            =   "frmProyectadoVsEjecutadoPorAgencia.frx":04BA
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   495
         Left            =   6000
         TabIndex        =   1
         Top             =   345
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Año:"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmProyectadoVsEjecutadoPorAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmProyectadoVsEjecutadoPorAgencia
'***     Descripcion:       Permite ver el versus entre lo proyectado y ejecutado por agencias
'***     Creado por:        FRHU
'***     Fecha-Tiempo:         24/04/2014 01:00:00 PM
'*****************************************************************************************
'FRHU 20140514 Observacion: Se agrego la validacion
Option Explicit
Private Function ValidarGrabar() As Boolean
   ValidarGrabar = True
   If cboAgencia.ListIndex = -1 Then
        MsgBox "Debe Seleccionar la Agencia"
        ValidarGrabar = False
   ElseIf cboMes.ListIndex = -1 Then
        MsgBox "Debe Seleccionar el Mes"
         ValidarGrabar = False
   ElseIf txtAnio.Text = "" Then
        MsgBox "Debe Seleccionar el Año"
         ValidarGrabar = False
   ElseIf Len(txtAnio.Text) <> 4 Then
        MsgBox "Ingrese un año correcto"
        ValidarGrabar = False
   End If
End Function
'FIN FRHU 20140514
Private Sub cmdMostrar_Click()
    Dim oCred As New COMDCredito.DCOMCredito
    Dim oRs As ADODB.Recordset
    Dim nSemana As Integer
    Dim i As Integer
    
    If Not ValidarGrabar Then Exit Sub 'FRHU 20140514 Observacion
    Screen.MousePointer = 11
    Set oRs = oCred.ObtenerProyectadoEjecutadoXAgencia(Right(Me.cboAgencia.Text, 2), Me.cboMes.ListIndex + 1, Me.txtAnio.Text)
    Screen.MousePointer = 0
    If Not oRs.EOF And Not oRs.BOF Then
        If oRs!valor = 0 Then
            MsgBox "No Hay datos que Mostrar", vbInformation
            Exit Sub
        End If
        Do While Not oRs.EOF
            i = (oRs!nSemana - 1)
            If i = 6 Then
            feMNSemana(i).TextMatrix(1, 2) = Format(oRs!nDesembMN, "#,##0.00")
            feMNSemana(i).TextMatrix(2, 2) = Format(oRs!nCrecimMN, "#,##0.00")
            'feMNSemana(i).TextMatrix(3, 2) = oRS!nCarteraMN
            feMNSemana(i).TextMatrix(1, 3) = Format(oRs!DesembolsoMN, "#,##0.00")
            feMNSemana(i).TextMatrix(2, 3) = Format(oRs!CrecimientoMN, "#,##0.00")
            'feMNSemana(i).TextMatrix(3, 3) = oRS!CarteraAtrasadaMN
            feMNSemana(i).TextMatrix(1, 4) = Format(oRs!DifDeseMN, "#,##0.00")
            feMNSemana(i).TextMatrix(2, 4) = Format(oRs!DifCrecMN, "#,##0.00")
            'feMNSemana(i).TextMatrix(3, 4) = oRS!DifCarteMN
            feMNSemana(i).TextMatrix(1, 5) = Format(oRs!DivDeseMN, "0.00%")
            feMNSemana(i).TextMatrix(2, 5) = Format(oRs!DivCrecMN, "0.00%")
            'feMNSemana(i).TextMatrix(3, 5) = Format(oRS!DivCarteMN, "0.00%")
            
            feMESemana(i).TextMatrix(1, 2) = Format(oRs!nDesembME, "#,##0.00")
            feMESemana(i).TextMatrix(2, 2) = Format(oRs!nCrecimME, "#,##0.00")
            'feMESemana(i).TextMatrix(3, 2) = oRS!nCarteraME
            feMESemana(i).TextMatrix(1, 3) = Format(oRs!DesembolsoME, "#,##0.00")
            feMESemana(i).TextMatrix(2, 3) = Format(oRs!CrecimientoME, "#,##0.00")
            'feMESemana(i).TextMatrix(3, 3) = oRS!CarteraAtrasadaME
            feMESemana(i).TextMatrix(1, 4) = Format(oRs!DifDeseME, "#,##0.00")
            feMESemana(i).TextMatrix(2, 4) = Format(oRs!DifCrecME, "#,##0.00")
            'feMESemana(i).TextMatrix(3, 4) = oRS!DifCarteME
            feMESemana(i).TextMatrix(1, 5) = Format(oRs!DivDeseME, "0.00%")
            feMESemana(i).TextMatrix(2, 5) = Format(oRs!DivCrecME, "0.00%")
            'feMESemana(i).TextMatrix(3, 5) = Format(oRS!DivCarteME, "0.00%")
            Else
            feMNSemana(i).TextMatrix(1, 2) = Format(oRs!nDesembMN, "#,##0.00")
            feMNSemana(i).TextMatrix(2, 2) = Format(oRs!nCrecimMN, "#,##0.00")
            feMNSemana(i).TextMatrix(3, 2) = Format(oRs!nCarteraMN, "#,##0.00")
            feMNSemana(i).TextMatrix(1, 3) = Format(oRs!DesembolsoMN, "#,##0.00")
            feMNSemana(i).TextMatrix(2, 3) = Format(oRs!CrecimientoMN, "#,##0.00")
            feMNSemana(i).TextMatrix(3, 3) = Format(oRs!CarteraAtrasadaMN, "#,##0.00")
            feMNSemana(i).TextMatrix(1, 4) = Format(oRs!DifDeseMN, "#,##0.00")
            feMNSemana(i).TextMatrix(2, 4) = Format(oRs!DifCrecMN, "#,##0.00")
            feMNSemana(i).TextMatrix(3, 4) = Format(oRs!DifCarteMN, "#,##0.00")
            feMNSemana(i).TextMatrix(1, 5) = Format(oRs!DivDeseMN, "0.00%")
            feMNSemana(i).TextMatrix(2, 5) = Format(oRs!DivCrecMN, "0.00%")
            feMNSemana(i).TextMatrix(3, 5) = Format(oRs!DivCarteMN, "0.00%")
            
            feMNSemana(6).TextMatrix(3, 2) = Format(oRs!nCarteraMN, "#,##0.00")
            feMNSemana(6).TextMatrix(3, 3) = Format(oRs!CarteraAtrasadaMN, "#,##0.00")
            feMNSemana(6).TextMatrix(3, 4) = Format(oRs!DifCarteMN, "#,##0.00")
            feMNSemana(6).TextMatrix(3, 5) = Format(oRs!DivCarteMN, "0.00%")
            
            feMESemana(i).TextMatrix(1, 2) = Format(oRs!nDesembME, "#,##0.00")
            feMESemana(i).TextMatrix(2, 2) = Format(oRs!nCrecimME, "#,##0.00")
            feMESemana(i).TextMatrix(3, 2) = Format(oRs!nCarteraME, "#,##0.00")
            feMESemana(i).TextMatrix(1, 3) = Format(oRs!DesembolsoME, "#,##0.00")
            feMESemana(i).TextMatrix(2, 3) = Format(oRs!CrecimientoME, "#,##0.00")
            feMESemana(i).TextMatrix(3, 3) = Format(oRs!CarteraAtrasadaME, "#,##0.00")
            feMESemana(i).TextMatrix(1, 4) = Format(oRs!DifDeseME, "#,##0.00")
            feMESemana(i).TextMatrix(2, 4) = Format(oRs!DifCrecME, "#,##0.00")
            feMESemana(i).TextMatrix(3, 4) = Format(oRs!DifCarteME, "#,##0.00")
            feMESemana(i).TextMatrix(1, 5) = Format(oRs!DivDeseME, "0.00%")
            feMESemana(i).TextMatrix(2, 5) = Format(oRs!DivCrecME, "0.00%")
            feMESemana(i).TextMatrix(3, 5) = Format(oRs!DivCarteME, "0.00%")
            
            feMESemana(6).TextMatrix(3, 2) = Format(oRs!nCarteraME, "#,##0.00")
            feMESemana(6).TextMatrix(3, 3) = Format(oRs!CarteraAtrasadaME, "#,##0.00")
            feMESemana(6).TextMatrix(3, 4) = Format(oRs!DifCarteME, "#,##0.00")
            feMESemana(6).TextMatrix(3, 5) = Format(oRs!DivCarteME, "0.00%")
            End If
            oRs.MoveNext
        Loop
        Me.cmdMostrar.Enabled = False
        Me.cmdExportar.Enabled = True
    Else
        MsgBox "No Hay datos que Mostrar", vbInformation
        Exit Sub
    End If
    Me.cboAgencia.Enabled = False
    Me.cboMes.Enabled = False
    Me.txtAnio.Enabled = False
    Set oRs = Nothing
End Sub
Private Sub Form_Load()
    Call CargarComboAgencia
    Call CargarConcepto
End Sub
Private Sub CargarComboAgencia()
    Dim oAge As New COMDConstantes.DCOMAgencias
    Dim rs As New ADODB.Recordset

    Set rs = oAge.ObtieneAgencias
    Do While Not rs.EOF
        Me.cboAgencia.AddItem (rs!cConsDescripcion & Space(50) & rs!nConsValor)
        rs.MoveNext
    Loop
End Sub
Private Sub CargarConcepto()
Dim i As Integer
    For i = 0 To Me.feMNSemana.Count - 1
        feMNSemana(i).AdicionaFila
        Me.feMNSemana(i).TextMatrix(1, 1) = "Desembolso"
        feMNSemana(i).AdicionaFila
        Me.feMNSemana(i).TextMatrix(2, 1) = "Crecimiento"
        feMNSemana(i).AdicionaFila
        Me.feMNSemana(i).TextMatrix(3, 1) = "Cartera Atrasada"
        
        feMESemana(i).AdicionaFila
        Me.feMESemana(i).TextMatrix(1, 1) = "Desembolso"
        feMESemana(i).AdicionaFila
        Me.feMESemana(i).TextMatrix(2, 1) = "Crecimiento"
        feMESemana(i).AdicionaFila
        Me.feMESemana(i).TextMatrix(3, 1) = "Cartera Atrasada"
    Next i
End Sub
Private Sub cmdCancelar_Click()
    Me.cmdMostrar.Enabled = True
    Me.cmdExportar.Enabled = False
    Me.txtAnio.Text = ""
    Dim i As Integer
    For i = 0 To Me.feMNSemana.Count - 1
        feMNSemana(i).TextMatrix(1, 2) = ""
        feMNSemana(i).TextMatrix(2, 2) = ""
        feMNSemana(i).TextMatrix(3, 2) = ""
        feMNSemana(i).TextMatrix(1, 3) = ""
        feMNSemana(i).TextMatrix(2, 3) = ""
        feMNSemana(i).TextMatrix(3, 3) = ""
        feMNSemana(i).TextMatrix(1, 4) = ""
        feMNSemana(i).TextMatrix(2, 4) = ""
        feMNSemana(i).TextMatrix(3, 4) = ""
        feMNSemana(i).TextMatrix(1, 5) = ""
        feMNSemana(i).TextMatrix(2, 5) = ""
        feMNSemana(i).TextMatrix(3, 5) = ""
            
        feMESemana(i).TextMatrix(1, 2) = ""
        feMESemana(i).TextMatrix(2, 2) = ""
        feMESemana(i).TextMatrix(3, 2) = ""
        feMESemana(i).TextMatrix(1, 3) = ""
        feMESemana(i).TextMatrix(2, 3) = ""
        feMESemana(i).TextMatrix(3, 3) = ""
        feMESemana(i).TextMatrix(1, 4) = ""
        feMESemana(i).TextMatrix(2, 4) = ""
        feMESemana(i).TextMatrix(3, 4) = ""
        feMESemana(i).TextMatrix(1, 5) = ""
        feMESemana(i).TextMatrix(2, 5) = ""
        feMESemana(i).TextMatrix(3, 5) = ""
    Next i
    Me.cboAgencia.Enabled = True
    Me.cboMes.Enabled = True
    Me.txtAnio.Enabled = True
    Me.txtAnio.SetFocus
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdExportar_Click()
    Dim xlsAplicacion As Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim fs As Scripting.FileSystemObject
    Dim lsArchivo As String
    Dim lsNomHoja  As String
    Dim lsArchivo1 As String
    Dim lbExisteHoja As Boolean
    Dim nSaltoContador As Integer
    '---
    Dim oCred As New COMDCredito.DCOMCredito
    Dim oRs As ADODB.Recordset
    Dim i As Integer
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    lsArchivo = "ProyectadoVsEjecutadoXAgencia"
    'Primera Hoja ******************************************************
    lsNomHoja = "HOJA"
    '*******************************************************************
    lsArchivo1 = "\spooler\RepProyectadoVsEjecutadoXAgencia" & gsCodUser & "_" & Format(gdFecha, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsArchivo & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsArchivo & ".xlsx")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    xlHoja1.Cells(4, 3) = Left(Me.cboAgencia.Text, 30)
    xlHoja1.Cells(4, 10) = Me.cboMes.Text
    Dim fila1 As Integer
    fila1 = 8
    For i = 0 To Me.feMNSemana.Count - 1
            'Proyectado
            xlHoja1.Cells(fila1, 4) = feMNSemana(i).TextMatrix(1, 2)
            xlHoja1.Cells(fila1 + 1, 4) = feMNSemana(i).TextMatrix(2, 2)
            xlHoja1.Cells(fila1 + 2, 4) = feMNSemana(i).TextMatrix(3, 2)
            'Ejecutado
            xlHoja1.Cells(fila1, 5) = feMNSemana(i).TextMatrix(1, 3)
            xlHoja1.Cells(fila1 + 1, 5) = feMNSemana(i).TextMatrix(2, 3)
            xlHoja1.Cells(fila1 + 2, 5) = feMNSemana(i).TextMatrix(3, 3)
            'Diferencia
            xlHoja1.Cells(fila1, 6) = feMNSemana(i).TextMatrix(1, 4)
            xlHoja1.Cells(fila1 + 1, 6) = feMNSemana(i).TextMatrix(2, 4)
            xlHoja1.Cells(fila1 + 2, 6) = feMNSemana(i).TextMatrix(3, 4)
            'Porcentaje
            xlHoja1.Cells(fila1, 7) = feMNSemana(i).TextMatrix(1, 5)
            xlHoja1.Cells(fila1 + 1, 7) = feMNSemana(i).TextMatrix(2, 5)
            xlHoja1.Cells(fila1 + 2, 7) = feMNSemana(i).TextMatrix(3, 5)
            
            'Proyectado
            xlHoja1.Cells(fila1, 8) = feMESemana(i).TextMatrix(1, 2)
            xlHoja1.Cells(fila1 + 1, 8) = feMESemana(i).TextMatrix(2, 2)
            xlHoja1.Cells(fila1 + 2, 8) = feMESemana(i).TextMatrix(3, 2)
            'Ejecutado
            xlHoja1.Cells(fila1, 9) = feMESemana(i).TextMatrix(1, 3)
            xlHoja1.Cells(fila1 + 1, 9) = feMESemana(i).TextMatrix(2, 3)
            xlHoja1.Cells(fila1 + 2, 9) = feMESemana(i).TextMatrix(3, 3)
            'Diferencia
            xlHoja1.Cells(fila1, 10) = feMESemana(i).TextMatrix(1, 4)
            xlHoja1.Cells(fila1 + 1, 10) = feMESemana(i).TextMatrix(2, 4)
            xlHoja1.Cells(fila1 + 2, 10) = feMESemana(i).TextMatrix(3, 4)
            'Porcentaje
            xlHoja1.Cells(fila1, 11) = feMESemana(i).TextMatrix(1, 5)
            xlHoja1.Cells(fila1 + 1, 11) = feMESemana(i).TextMatrix(2, 5)
            xlHoja1.Cells(fila1 + 2, 11) = feMESemana(i).TextMatrix(3, 5)
            
            fila1 = fila1 + 3
    Next i
    
    xlHoja1.SaveAs App.Path & lsArchivo1
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
    
End Sub
Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        Me.cmdMostrar.SetFocus
    End If
End Sub
