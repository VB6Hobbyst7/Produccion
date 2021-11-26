VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogComprobanteRegistroNew 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10515
   Icon            =   "frmLogComprobanteRegistroNew.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
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
      Left            =   9360
      TabIndex        =   37
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar2 
      Caption         =   "Cancelar"
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
      Left            =   8220
      TabIndex        =   36
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
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
      TabIndex        =   35
      Top             =   7320
      Width           =   1095
   End
   Begin TabDlg.SSTab sstOrigen 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Buscar"
      TabPicture(0)   =   "frmLogComprobanteRegistroNew.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblDescripcion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtBuscar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdNuevaValorizacion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdRegistrarComprobante"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdMostrar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cbotpoOrigen"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fraItemComprobante"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.Frame fraItemComprobante 
         Caption         =   "Items a Registrar en Comprobante"
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
         Height          =   1900
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   9960
         Begin Sicmact.FlexEdit feContratoAdqBienes 
            Height          =   1575
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2778
            Cols0           =   10
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Alt--Objeto-Descripción-Unidad-Cantidad-P. Unit-cAgeCod-CtaCont"
            EncabezadosAnchos=   "0-0-450-1200-3000-1200-1200-1200-0-0"
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
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-4-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin Sicmact.FlexEdit feContratoObra 
            Height          =   1575
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2778
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#--Monto-Descripción-CtaCont-cAgeCod-cObjeto"
            EncabezadosAnchos=   "0-450-1200-6000-0-0-1"
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
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-X-X-X-X"
            ListaControles  =   "0-4-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-L-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin Sicmact.FlexEdit feContrato 
            Height          =   1575
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2778
            Cols0           =   11
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-cCtaContCod--N° de Pago-Fecha de Pago-Moneda-Monto-Tipo-Estado-cAgeCod-cObjeto"
            EncabezadosAnchos=   "0-0-450-1000-1200-1000-1200-1800-1800-0-0"
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
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-2-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-C-C-C-R-C-L-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-2-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   7
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
         Begin Sicmact.FlexEdit feOrden 
            Height          =   1575
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   2778
            Cols0           =   12
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-nMovNro-nMovItem-cCtaContCod--Ag.Destino-Objeto-Descripcion-Unidad-Solicitado-P.Unitario-SubTotal"
            EncabezadosAnchos=   "0-0-0-0-450-1000-1400-2000-1000-900-1100-1100"
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
            ColumnasAEditar =   "X-X-X-X-4-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-4-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-C-C-C-C-C-L-C-C-R-R"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-2-0-2-2"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   7
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.ComboBox cbotpoOrigen 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
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
         Left            =   9240
         TabIndex        =   6
         Top             =   600
         Width           =   910
      End
      Begin VB.CommandButton cmdRegistrarComprobante 
         Caption         =   "&Registrar Comprobante"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   2175
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
         Left            =   2450
         TabIndex        =   13
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdNuevaValorizacion 
         Caption         =   "&Nueva Valorización"
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
         Left            =   8400
         TabIndex        =   14
         Top             =   2880
         Width           =   1815
      End
      Begin Sicmact.TxtBuscar txtBuscar 
         Height          =   315
         Left            =   3960
         TabIndex        =   4
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   6
         sTitulo         =   ""
         EnabledText     =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Documento Origen"
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
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Doc. Origen"
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
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   680
         Width           =   1815
      End
      Begin VB.Label lblDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5850
         TabIndex        =   5
         Top             =   600
         Width           =   3375
      End
   End
   Begin TabDlg.SSTab sstComprobante 
      Height          =   3615
      Left            =   120
      TabIndex        =   15
      Top             =   3600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Comprobante"
      TabPicture(0)   =   "frmLogComprobanteRegistroNew.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame3 
         Caption         =   "Información General"
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
         Height          =   1350
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   5850
         Begin VB.TextBox txtObservacion 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            TabIndex        =   24
            Top             =   960
            Width           =   4400
         End
         Begin VB.Label lblAreaAgeNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   22
            Top             =   600
            Width           =   3210
         End
         Begin VB.Label lblAreaAgeCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1320
            TabIndex        =   21
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label lblProveedorNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2520
            TabIndex        =   19
            Top             =   240
            Width           =   3210
         End
         Begin VB.Label lblProveedorCod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1320
            TabIndex        =   18
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label5 
            Caption         =   "Observaciones:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Área Usuaria:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Proveedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Comprobante"
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
         Height          =   1365
         Left            =   6120
         TabIndex        =   25
         Top             =   360
         Width           =   4050
         Begin VB.TextBox txtComprobanteSerie 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            TabIndex        =   29
            Top             =   600
            Width           =   795
         End
         Begin VB.TextBox txtComprobanteNro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   30
            Top             =   600
            Width           =   2145
         End
         Begin VB.ComboBox cboTpoComprobante 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   240
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker txtComprobanteFecEmision 
            Height          =   285
            Left            =   840
            TabIndex        =   32
            Top             =   915
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   240058369
            CurrentDate     =   41586
         End
         Begin VB.Label Label11 
            Caption         =   "Tipo:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   375
         End
         Begin VB.Label Label8 
            Caption         =   "N°:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   255
         End
         Begin VB.Label Label9 
            Caption         =   "Emisión:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Detalle del Comprobante"
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
         Height          =   1695
         Left            =   240
         TabIndex        =   33
         Top             =   1800
         Width           =   10005
         Begin Sicmact.FlexEdit feComprobanteDet 
            Height          =   1335
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   2355
            Cols0           =   10
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Ag.Dest-Objeto-Descripcion-Unidad-Solicitado-P.Unitario-SubTotal-CtaContCod-nItem"
            EncabezadosAnchos=   "400-800-1400-2800-950-950-1100-1100-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
            TextStyleFixed  =   3
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-C-R-R-R-L-C"
            FormatosEdit    =   "0-0-0-0-0-3-2-2-0-0"
            CantEntero      =   12
            CantDecimales   =   3
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            Appearance      =   0
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
End
Attribute VB_Name = "frmLogComprobanteRegistroNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'Nombre : frmLogComprobanteRegistroNew
'Descripcion:Formulario para el Registro de Comprobantes
'Creacion: PASIERS0772014
'*****************************

'Variables Tab Buscar *************************************
Option Explicit
Dim gsopecod As String
Dim fnMovNro As Long
Dim fsMatrizDatos() As String
Dim fsMatrizDatObra() As String
Dim fsAreaAgeCod As String
Dim fsCtaContCodOC As String, fsCtaContCodOS As String
Dim fnTipo As Integer
Dim fnMoneda As Integer
Dim fnFormTamanioIni As Integer
Dim fnFormTamanioActiva As Integer
Dim olog As DLogGeneral
'***********************************************************
'Tab Comprobante - Variables Ordenes ***********************
Dim fsCtaContCodProv As String
Dim fsProveedorIFICod As String
Dim fsProveedorIFICtaCod As String
Dim fsDocTpo As String
Dim fsDocNro As String
Dim fnTipoPago As Integer
Dim fnTpoCambio As Integer
Dim fnTipoPagoContrato As Integer
'***********************************************************
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Public Function Inicio(ByVal psOpeCod As String)
    gsopecod = psOpeCod
    fnMoneda = Mid(psOpeCod, 3, 1)
    Me.Show 1
End Function
Private Sub cboTpoComprobante_Click()
    txtComprobanteSerie.SetFocus
End Sub
Private Sub cbotpoOrigen_Click()
    fnTipo = Trim(Right(cbotpoOrigen, 100))
    MuestraGrid
    EstadoControles 0
    txtBuscar.SetFocus
End Sub
Private Sub cmdCancelar_Click()
    EstadoInicial False
    cbotpoOrigen.SetFocus
End Sub
Private Sub cmdCancelar2_Click()
    LimpiarDetcomprobante
    EstadoInicial True
    cbotpoOrigen.SetFocus
End Sub
Private Sub cmdCerrar_Click() 'PASI20141229
    EstadoInicial True
    Unload Me
End Sub
Private Sub cmdMostrar_Click()
    Dim rs As ADODB.Recordset
    If Len(txtBuscar.Text) = 0 Or txtBuscar.Text = "0000000000" Then
        MsgBox "No se ha seleccionado ningún Documento de Origen.", vbOKOnly + vbExclamation, "Aviso."
        If cbotpoOrigen.ListIndex = -1 Then
            cbotpoOrigen.SetFocus
        Else
            txtBuscar.SetFocus
        End If
        Exit Sub
    Else
        Select Case fnTipo
            Case LogTipoDocOrigenComprobante.OrdenCompra
                Set rs = olog.ListaOrdenCompraDetxRegistroComprobante(fsMatrizDatos(1, 1), fsAreaAgeCod, fsCtaContCodOC)
                EstadoControles 4
            Case LogTipoDocOrigenComprobante.OrdenServicio
                Set rs = olog.ListaOrdenServicioDetxRegistroComprobante(fsMatrizDatos(1, 1), fsAreaAgeCod, fsCtaContCodOS)
                EstadoControles 4
            Case LogTipoDocOrigenComprobante.ContratoServicio, LogTipoDocOrigenComprobante.ContratoArrendamiento
                If olog.ExisteAdendaContratos(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1)) > 0 Then
                    Set rs = olog.ListaContratoSADetxRegistrarComprobante(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), olog.ExisteAdendaContratos(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1)))
                Else
                    Set rs = olog.ListaContratoSADetxRegistrarComprobante(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), 0)
                End If
                EstadoControles 1
            Case LogTipoDocOrigenComprobante.ContratoObra
                    Set rs = olog.ListaContratoObraDetxRegistrarComprobante(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), olog.ExisteAdendaContratos(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1)))
                EstadoControles 2
                cmdNuevaValorizacion.SetFocus
            Case LogTipoDocOrigenComprobante.ContratoAdqBienes
                If olog.ExisteAdendaContratos(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1)) > 0 Then
                    Set rs = olog.ListaContratoABDetxRegistrarComprobante(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), olog.ExisteAdendaContratos(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1)))
                Else
                    Set rs = olog.ListaContratoABDetxRegistrarComprobante(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), 0)
                End If
                EstadoControles 1
        End Select
        cbotpoOrigen.Enabled = False
        txtBuscar.Enabled = False
        LimpiarFlex
        LlenaGrid rs
    End If
End Sub
Private Sub LimpiarFlex()
    Call LimpiaFlex(feOrden)
    Call LimpiaFlex(feContrato)
    Call LimpiaFlex(feContratoObra)
    Call LimpiaFlex(feContratoAdqBienes)
End Sub
Private Sub MuestraGrid()
    LimpiarFlex
    If cbotpoOrigen.ListIndex = -1 Then
        feOrden.Visible = True
        feContrato.Visible = False
        feContratoObra.Visible = False
        feContratoAdqBienes.Visible = False
    Else
        Select Case fnTipo
            Case LogTipoDocOrigenComprobante.OrdenCompra, LogTipoDocOrigenComprobante.OrdenServicio
                feOrden.Visible = True
                feContrato.Visible = False
                feContratoObra.Visible = False
                feContratoAdqBienes.Visible = False
            Case LogTipoDocOrigenComprobante.ContratoServicio, LogTipoDocOrigenComprobante.ContratoArrendamiento
                feOrden.Visible = False
                feContrato.Visible = True
                feContratoObra.Visible = False
                feContratoAdqBienes.Visible = False
            Case LogTipoDocOrigenComprobante.ContratoObra
                feOrden.Visible = False
                feContrato.Visible = False
                feContratoObra.Visible = True
                feContratoAdqBienes.Visible = False
            Case LogTipoDocOrigenComprobante.ContratoAdqBienes
'                feOrden.Visible = False
'                feContrato.Visible = False
'                feContratoObra.Visible = False
'                feContratoAdqBienes.Visible = True
                feOrden.Visible = True
                feContrato.Visible = False
                feContratoObra.Visible = False
                feContratoAdqBienes.Visible = False
        End Select
    End If
End Sub
Private Sub LlenaGrid(ByVal rs As ADODB.Recordset)
Dim row As Long
        Select Case fnTipo
            Case LogTipoDocOrigenComprobante.OrdenCompra, LogTipoDocOrigenComprobante.OrdenServicio
                Do While Not rs.EOF
                    feOrden.AdicionaFila
                    row = feOrden.row
                    feOrden.TextMatrix(row, 1) = rs!nMovNro
                    feOrden.TextMatrix(row, 2) = rs!nMovItem
                    feOrden.TextMatrix(row, 3) = rs!cCtaContCod
                    feOrden.TextMatrix(row, 5) = rs!cAgeCod
                    feOrden.TextMatrix(row, 6) = rs!cObjeto
                    feOrden.TextMatrix(row, 7) = rs!cDescripcion
                    feOrden.TextMatrix(row, 8) = rs!cUnidad
                    feOrden.TextMatrix(row, 9) = rs!nSolicitado
                    feOrden.TextMatrix(row, 10) = Format(rs!nPrecioUnitario, gsFormatoNumeroView)
                    feOrden.TextMatrix(row, 11) = Format(rs!nSubTotal, gsFormatoNumeroView)
                    rs.MoveNext
                Loop
                feOrden.SetFocus
                SendKeys "{Right}"
            Case LogTipoDocOrigenComprobante.ContratoServicio, LogTipoDocOrigenComprobante.ContratoArrendamiento
                Do While Not rs.EOF
                    feContrato.AdicionaFila
                    row = feContrato.row
                    feContrato.TextMatrix(row, 1) = rs!cCtaContCod
                    feContrato.TextMatrix(row, 3) = rs!nNPago
                    feContrato.TextMatrix(row, 4) = Format(rs!dFecPago, gsFormatoFechaView)
                    feContrato.TextMatrix(row, 5) = rs!cMoneda
                    feContrato.TextMatrix(row, 6) = Format(rs!nMonto, gsFormatoNumeroView)
                    feContrato.TextMatrix(row, 7) = rs!cTipo
                    feContrato.TextMatrix(row, 8) = rs!cEstado
                    feContrato.TextMatrix(row, 9) = rs!cAgeCod
                    feContrato.TextMatrix(row, 10) = rs!cObjeto
                    fnTipoPagoContrato = rs!nTipoPago
                    rs.MoveNext
                Loop
                feContrato.SetFocus
                SendKeys "{Right}"
            Case LogTipoDocOrigenComprobante.ContratoObra
                Do While Not rs.EOF
                    feContratoObra.AdicionaFila
                    row = feContratoObra.row
                    feContratoObra.TextMatrix(row, 4) = rs!cCtaContCod
                    feContratoObra.TextMatrix(row, 5) = rs!cAgeCod
                    feContratoObra.TextMatrix(row, 6) = rs!cObjeto
                    rs.MoveNext
                Loop
            Case LogTipoDocOrigenComprobante.ContratoAdqBienes
                Do While Not rs.EOF
'                    feContratoAdqBienes.AdicionaFila
'                    row = feContratoAdqBienes.row
'                    feContratoAdqBienes.TextMatrix(row, 1) = "-"
'                    feContratoAdqBienes.TextMatrix(row, 3) = rs!Objeto
'                    feContratoAdqBienes.TextMatrix(row, 4) = rs!Descripcion
'                    feContratoAdqBienes.TextMatrix(row, 5) = rs!unidad
'                    feContratoAdqBienes.TextMatrix(row, 6) = rs!Cant
'                    feContratoAdqBienes.TextMatrix(row, 7) = rs!PreUnit
'                    feContratoAdqBienes.TextMatrix(row, 8) = rs!cAgeCod
'                    feContratoAdqBienes.TextMatrix(row, 9) = rs!cCtaContCod
'                    rs.MoveNext
                feOrden.AdicionaFila
                    row = feOrden.row
                    feOrden.TextMatrix(row, 1) = 0
                    feOrden.TextMatrix(row, 2) = rs!nMovItem
                    feOrden.TextMatrix(row, 3) = rs!cCtaContCod
                    feOrden.TextMatrix(row, 5) = rs!cAgeCod
                    feOrden.TextMatrix(row, 6) = rs!cObjeto
                    feOrden.TextMatrix(row, 7) = rs!cDescripcion
                    feOrden.TextMatrix(row, 8) = rs!cUnidad
                    feOrden.TextMatrix(row, 9) = rs!nSolicitado
                    feOrden.TextMatrix(row, 10) = Format(rs!nPrecioUnitario, gsFormatoNumeroView)
                    feOrden.TextMatrix(row, 11) = Format(rs!nSubTotal, gsFormatoNumeroView)
                    rs.MoveNext
                Loop
                'feContratoAdqBienes.SetFocus
                SendKeys "{Right}"
        End Select
End Sub
Private Sub cmdNuevaValorizacion_Click()
    fsMatrizDatObra = frmLogValorizaComproObra.Inicio(0)
    If fsMatrizDatObra(1, 1) <> "" Then
        'Call LimpiaFlex(Me.feContratoObra)
        'feContratoObra.AdicionaFila
        feContratoObra.TextMatrix(1, 2) = fsMatrizDatObra(1, 1)
        feContratoObra.TextMatrix(1, 3) = fsMatrizDatObra(2, 1)
        cmdRegistrarComprobante.Enabled = True
        EstadoControles 1
        feContratoObra.SetFocus
        SendKeys "{Right}"
    Else
        cmdNuevaValorizacion.Enabled = True
    End If
End Sub
Private Function ValidaDetComprobante() As Boolean
    Dim nMonto As Currency
    Dim I As Integer
    
    If Len(fsProveedorIFICod) = 0 Or Len(fsProveedorIFICtaCod) = 0 Then
        MsgBox "Ud. debe verificar que el proveedor " & UCase(lblProveedorNombre.Caption) & Chr(10) & "se encuentre registrado en la BD de Proveedores de Logìstica, además tenga" & Chr(10) & "configurado una cuenta en " & UCase(IIf(fnMoneda = 1, "SOLES", "DOLARES")) & "para poder continuar con el proceso.", vbInformation, "Aviso"
        ValidaDetComprobante = False
        Exit Function
    End If
    If Len(txtObservacion.Text) = 0 Then
        MsgBox "No se ha ingresado las Observaciones.", vbOKOnly + vbExclamation, "Aviso"
        txtObservacion.SetFocus
        ValidaDetComprobante = False
        Exit Function
    End If
    If cboTpoComprobante.ListIndex = -1 Then
        MsgBox "No se ha seleccionado el tipo de Documento.", vbOKOnly + vbExclamation, "Aviso"
        cboTpoComprobante.SetFocus
        ValidaDetComprobante = False
        Exit Function
    End If
    If Len(txtComprobanteSerie.Text) = 0 Then
        MsgBox "No se ha completado el número de comprobante.", vbOKOnly + vbExclamation, "Aviso"
        txtComprobanteSerie.SetFocus
        ValidaDetComprobante = False
        Exit Function
    End If
    If Len(txtComprobanteNro.Text) = 0 Then
        MsgBox "No se ha completado el número de comprobante.", vbOKOnly + vbExclamation, "Mensaje"
        txtComprobanteNro.SetFocus
        ValidaDetComprobante = False
        Exit Function
    End If
    If (CInt(Trim(Right(cbotpoOrigen, 100))) = LogTipoContrato.ContratoServicio And fnTipoPagoContrato = 2) Or CInt(Trim(Right(cbotpoOrigen, 100))) = LogTipoContrato.ContratoObra Then
        nMonto = 0
        For I = 1 To feComprobanteDet.Rows - 1
            nMonto = nMonto + CCur(feComprobanteDet.TextMatrix(I, 7))
        Next
        If Not olog.ExisteSaldoContrato(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), nMonto) Then
                MsgBox "El Monto del Comprobante supera el Saldo del Contrato, no se puede realizar el registro.", vbOKOnly + vbExclamation, "Mensaje"
                feComprobanteDet.SetFocus
                ValidaDetComprobante = False
                Exit Function
        End If
    End If
    ValidaDetComprobante = True
End Function
Private Sub cmdRegistrar_Click()
    Dim olog As NLogGeneral
    Dim oDLog As DLogGeneral
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim DatosOrden() As TComprobanteOrden
    Dim DatosContratoxCron() As TComprobanteContratoxCronograma
    Dim DatosContratoBien() As TComprobanteContratoxBien
    Dim DatosContratoxObra() As TComprobanteContratoxObra
    Dim indexMat, I As Integer
    Dim rsAreas As ADODB.Recordset
    Dim lsValidaMovReg As String 'vapa
    Dim cadena As String
    On Error GoTo ErrRegComprobante
    
    If Not ValidaDetComprobante Then Exit Sub
    
    indexMat = 0
    ReDim DatosOrden(indexMat)
    ReDim DatosContratoxCron(indexMat)
    ReDim TComprobanteContratoxBien(indexMat)
    ReDim DatosContratoxObra(indexMat)
    
    Select Case fnTipo
        Case LogTipoDocOrigenComprobante.OrdenCompra, LogTipoDocOrigenComprobante.OrdenServicio
            For indexMat = 1 To feComprobanteDet.Rows - 1
                ReDim Preserve DatosOrden(indexMat)
                DatosOrden(indexMat).nMovItem = feComprobanteDet.TextMatrix(indexMat, 9)
                DatosOrden(indexMat).sCtaContCod = feComprobanteDet.TextMatrix(indexMat, 8)
                DatosOrden(indexMat).sObjeto = feComprobanteDet.TextMatrix(indexMat, 2)
                DatosOrden(indexMat).sDescripcion = feComprobanteDet.TextMatrix(indexMat, 3)
                DatosOrden(indexMat).nCantidad = CDbl(feComprobanteDet.TextMatrix(indexMat, 5))
                DatosOrden(indexMat).nTotal = feComprobanteDet.TextMatrix(indexMat, 7)
            Next
        
        Case LogTipoDocOrigenComprobante.ContratoServicio _
            , LogTipoDocOrigenComprobante.ContratoArrendamiento
            
            If fnTipo = LogTipoDocOrigenComprobante.ContratoServicio Then
                I = 0
                For indexMat = 1 To feContrato.Rows - 1
                    If feContrato.TextMatrix(indexMat, 2) = "." Then
                        I = I + 1
                        ReDim Preserve DatosContratoxCron(I)
                        DatosContratoxCron(I).nNPago = feContrato.TextMatrix(indexMat, 3)
                    End If
                Next
                For indexMat = 1 To feComprobanteDet.Rows - 1
                    ReDim Preserve DatosOrden(indexMat)
                    DatosOrden(indexMat).nMovItem = feComprobanteDet.TextMatrix(indexMat, 9)
                    DatosOrden(indexMat).sCtaContCod = feComprobanteDet.TextMatrix(indexMat, 8)
                    DatosOrden(indexMat).sObjeto = feComprobanteDet.TextMatrix(indexMat, 2)
                    DatosOrden(indexMat).sDescripcion = feComprobanteDet.TextMatrix(indexMat, 3)
                    DatosOrden(indexMat).nCantidad = CDbl(feComprobanteDet.TextMatrix(indexMat, 5))
                    DatosOrden(indexMat).nTotal = feComprobanteDet.TextMatrix(indexMat, 7)
                Next
            Else
                 For indexMat = 1 To feComprobanteDet.Rows - 1
                    ReDim Preserve DatosContratoxCron(indexMat)
                    DatosContratoxCron(indexMat).nNPago = feComprobanteDet.TextMatrix(indexMat, 9)
                    DatosContratoxCron(indexMat).sCtaContCod = feComprobanteDet.TextMatrix(indexMat, 8)
                    DatosContratoxCron(indexMat).sDescripcion = feComprobanteDet.TextMatrix(indexMat, 3)
                    DatosContratoxCron(indexMat).nMonto = feComprobanteDet.TextMatrix(indexMat, 7)
                Next
            End If
           
        Case LogTipoDocOrigenComprobante.ContratoAdqBienes
'            For indexMat = 1 To feComprobanteDet.Rows - 1
'                ReDim Preserve DatosContratoxItem(indexMat)
'                DatosContratoxItem(indexMat).sBSCod = feComprobanteDet.TextMatrix(indexMat, 2)
'                DatosContratoxItem(indexMat).sCtaContCod = feComprobanteDet.TextMatrix(indexMat, 8)
'                DatosContratoxItem(indexMat).sDescripcion = feComprobanteDet.TextMatrix(indexMat, 3)
'                DatosContratoxItem(indexMat).nCantidad = feComprobanteDet.TextMatrix(indexMat, 5)
'                DatosContratoxItem(indexMat).nMonto = feComprobanteDet.TextMatrix(indexMat, 7)
'             Next
            For indexMat = 1 To feComprobanteDet.Rows - 1
                ReDim Preserve DatosContratoBien(indexMat)
                DatosContratoBien(indexMat).nMovItem = feComprobanteDet.TextMatrix(indexMat, 9)
                DatosContratoBien(indexMat).sCtaContCod = feComprobanteDet.TextMatrix(indexMat, 8)
                DatosContratoBien(indexMat).sObjeto = feComprobanteDet.TextMatrix(indexMat, 2)
                DatosContratoBien(indexMat).sDescripcion = feComprobanteDet.TextMatrix(indexMat, 3)
                DatosContratoBien(indexMat).nCantidad = feComprobanteDet.TextMatrix(indexMat, 5)
                DatosContratoBien(indexMat).nTotal = feComprobanteDet.TextMatrix(indexMat, 7)
            Next
        Case LogTipoDocOrigenComprobante.ContratoObra
            For indexMat = 1 To feComprobanteDet.Rows - 1
                ReDim Preserve DatosContratoxObra(indexMat)
                DatosContratoxObra(indexMat).sCtaContCod = feComprobanteDet.TextMatrix(indexMat, 8)
                DatosContratoxObra(indexMat).sDescripcion = feComprobanteDet.TextMatrix(indexMat, 3)
                DatosContratoxObra(indexMat).nMonto = feComprobanteDet.TextMatrix(indexMat, 7)
            Next
    End Select
    Select Case fnTipo
        Case LogTipoDocOrigenComprobante.OrdenCompra _
            , LogTipoDocOrigenComprobante.ContratoAdqBienes
            
            fsCtaContCodProv = fsCtaContCodOC
            
        Case LogTipoDocOrigenComprobante.OrdenServicio _
            , LogTipoDocOrigenComprobante.ContratoServicio _
            , LogTipoDocOrigenComprobante.ContratoArrendamiento _
            , LogTipoDocOrigenComprobante.ContratoObra
        
            fsCtaContCodProv = fsCtaContCodOS
    End Select
    
    If fsCtaContCodProv = "" Then
        MsgBox "No se ha definido cuenta contable de Proveedor, consulte al Dpto. de TI", vbInformation, "Aviso"
        Exit Sub
    End If
    
'    If MsgBox("¿Esta seguro de realizar el Registro de Comprobante?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
'        Exit Sub
'    End If
    
    Set olog = New NLogGeneral
    'Screen.MousePointer = 11
   ' lsValidaMovReg = olog.ValidaComprobanteReg(Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, CDate(txtComprobanteFecEmision.value)) 'vapa
    lsValidaMovReg = olog.ValidaComprobanteReg(Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, lblProveedorCod.Caption) ' CDate(txtComprobanteFecEmision.value)) 'vapa
    If lsValidaMovReg = "no" Then
        If MsgBox("¿Esta seguro de realizar el Registro de Comprobante?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
        End If
    Screen.MousePointer = 11
    Select Case fnTipo 'Insercion de Comprobante de acuerdo al tipo
        Case LogTipoDocOrigenComprobante.OrdenCompra _
            , LogTipoDocOrigenComprobante.OrdenServicio
            lnMovNro = olog.GrabarComprobanteOrden(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsopecod, "Registro del Comprobante Nº " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, lblProveedorCod.Caption, lblAreaAgeCod.Caption, fnMoneda, fnTipo, Trim(txtObservacion.Text), Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, CDate(txtComprobanteFecEmision.value), fnTipoPago, fsProveedorIFICod, fsProveedorIFICtaCod, DatosOrden, fsCtaContCodProv, fnTpoCambio, lsMovNro, fsMatrizDatos(1, 1))
        Case LogTipoDocOrigenComprobante.ContratoServicio
            lnMovNro = olog.GrabaComprobanteContratoxServicio(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsopecod, fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), "Registro del Comprobante Nº " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, lblProveedorCod.Caption, lblAreaAgeCod.Caption, fnMoneda, fnTipo, Trim(txtObservacion.Text), Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, CDate(txtComprobanteFecEmision.value), fnTipoPago, fsProveedorIFICod, fsProveedorIFICtaCod, DatosContratoxCron, DatosOrden, fsCtaContCodProv, fnTpoCambio, lsMovNro)
        Case LogTipoDocOrigenComprobante.ContratoArrendamiento
            lnMovNro = olog.GrabaComprobanteContratoxCronograma(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsopecod, fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), "Registro del Comprobante Nº " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, lblProveedorCod.Caption, lblAreaAgeCod.Caption, fnMoneda, fnTipo, Trim(txtObservacion.Text), Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, CDate(txtComprobanteFecEmision.value), fnTipoPago, fsProveedorIFICod, fsProveedorIFICtaCod, DatosContratoxCron, fsCtaContCodProv, fnTpoCambio, lsMovNro)
        Case LogTipoDocOrigenComprobante.ContratoAdqBienes
            lnMovNro = olog.GrabaComprobanteContratoxBienes(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsopecod, fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), "Registro del Comprobante Nº " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, lblProveedorCod.Caption, lblAreaAgeCod.Caption, fnMoneda, fnTipo, Trim(txtObservacion.Text), Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, CDate(txtComprobanteFecEmision.value), fnTipoPago, fsProveedorIFICod, fsProveedorIFICtaCod, DatosContratoBien, fsCtaContCodProv, fnTpoCambio, lsMovNro)
        Case LogTipoDocOrigenComprobante.ContratoObra
            lnMovNro = olog.GrabaComprobanteContratoxObra(gdFecSis, Right(gsCodAge, 2), gsCodUser, gsopecod, fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), "Registro del Comprobante Nº " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, lblProveedorCod.Caption, lblAreaAgeCod.Caption, fnMoneda, fnTipo, Trim(txtObservacion.Text), Trim(Right(cboTpoComprobante.Text, 3)), txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text, CDate(txtComprobanteFecEmision.value), fnTipoPago, fsProveedorIFICod, fsProveedorIFICtaCod, DatosContratoxObra, fsCtaContCodProv, fnTpoCambio, lsMovNro)
    End Select
    Else
        MsgBox "El número de comprobante ya existe, por favor ingrese otro número ", vbInformation, "Aviso"
        Exit Sub
    End If
    Screen.MousePointer = 0
    If lnMovNro = 0 Then
        MsgBox "Ha ocurrido un error al registrar el Comprobante", vbCritical, "Aviso"
        Set olog = Nothing
        Exit Sub
    End If
    Set oDLog = New DLogGeneral
    If fnTipo = LogTipoDocOrigenComprobante.OrdenCompra Or fnTipo = LogTipoDocOrigenComprobante.OrdenServicio Then
        oDLog.RegistraActaPendxComprobante lnMovNro, lblAreaAgeCod.Caption
    Else
        Set rsAreas = oDLog.ObtenerAreaxActaConformidadxContrato(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1))
        If Not rsAreas.EOF And Not rsAreas.BOF Then
            If rsAreas!cArea1 <> "00" Then
                oDLog.RegistraActaPendxComprobante lnMovNro, rsAreas!cArea1
            End If
            If rsAreas!cArea2 <> "00" Then
                oDLog.RegistraActaPendxComprobante lnMovNro, rsAreas!cArea2
            End If
        End If
    End If
    MsgBox "Se ha registrado el Comprobante de Nro. " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text & " con éxito", vbInformation, "Aviso"
    Set olog = Nothing
    Set oDLog = Nothing
    'ARLO 20160126 ***
    Dim lsMoneda As String
    If (fnMoneda = 1) Then
    lsMoneda = "SOLES"
    gsopecod = LogPistaRegistroComprobanteMN
    Else
    lsMoneda = "DOLARES"
    gsopecod = LogPistaRegistroComprobanteME
    End If
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Registro del Comprobante Nº " & txtComprobanteSerie.Text & "-" & txtComprobanteNro.Text & " En Moneda " & lsMoneda
    Set objPista = Nothing
    '***
    If MsgBox("¿Desea registrar otro Comprobante?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbYes Then
        cmdCancelar2_Click
    Else
        Unload Me
    End If
    Exit Sub
ErrRegComprobante:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdRegistrarComprobante_Click()
    Dim bselecciona As Boolean
    Dim row As Integer
    Dim rs As ADODB.Recordset
    Select Case fnTipo '***Valida Seleccion de Registros
        Case LogTipoDocOrigenComprobante.OrdenCompra _
            , LogTipoDocOrigenComprobante.OrdenServicio
            For row = 1 To feOrden.Rows - 1
                If feOrden.TextMatrix(row, 4) = "." Then
                    bselecciona = True
                    Exit For
                End If
            Next
        Case LogTipoDocOrigenComprobante.ContratoServicio _
             , LogTipoDocOrigenComprobante.ContratoArrendamiento
            For row = 1 To feContrato.Rows - 1
                If feContrato.TextMatrix(row, 2) = "." Then
                    bselecciona = True
                    Exit For
                End If
            Next
        Case LogTipoDocOrigenComprobante.ContratoObra
            For row = 1 To feContratoObra.Rows - 1
                If feContratoObra.TextMatrix(row, 1) = "." Then
                    bselecciona = True
                    Exit For
                End If
            Next
        Case LogTipoDocOrigenComprobante.ContratoAdqBienes
            For row = 1 To feOrden.Rows - 1
                If feOrden.TextMatrix(row, 4) = "." Then
                    bselecciona = True
                    Exit For
                End If
            Next
    End Select
    If Not bselecciona Then
        MsgBox "Ud. debe seleccionar los Items a Registrar", vbOKOnly + vbExclamation, "Aviso"
        Exit Sub
    End If
    'Para Rellenar el Detalle de Comprobante ******************
        Select Case fnTipo
            Case LogTipoDocOrigenComprobante.OrdenCompra _
                , LogTipoDocOrigenComprobante.OrdenServicio
                Set rs = olog.ListaOrdenDetxComprobante(fsMatrizDatos(1, 1))
            Case LogTipoDocOrigenComprobante.ContratoServicio _
                , LogTipoDocOrigenComprobante.ContratoArrendamiento _
                , LogTipoDocOrigenComprobante.ContratoObra _
                , LogTipoDocOrigenComprobante.ContratoAdqBienes
                Set rs = olog.ListaContratoDetxComprobante(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1))
        End Select
        If Not rs.EOF Then
            EstablecerDatosDetComprobante rs!cProveedorCod, rs!cProveedorNombre, rs!cAreaAgeCod, rs!cAreaAgeDesc, IIf(rs!cIFiCod <> "", rs!cMoneda, ""), rs!cDocTpo, rs!cDocNro, rs!cIFiCod, rs!cIFiNombre, rs!cIFiCtaCod
        End If
    '***********************************************************
    Set rs = Nothing
    EstadoControles 3
    HabilitaTab True
    txtObservacion.SetFocus
End Sub
Private Sub EstablecerDatosDetComprobante(Optional ByVal psProveedorCod As String = "", _
                                            Optional ByVal psProveedorNombre As String = "", _
                                            Optional ByVal psAreaAgeCod As String = "", _
                                            Optional ByVal psAreaAgeDesc As String = "", _
                                            Optional ByVal psMoneda As String = "", _
                                            Optional ByVal psDocTpo As String = "", _
                                            Optional ByVal psDocNro As String = "", _
                                            Optional ByVal psIFICod As String = "", _
                                            Optional ByVal psIFINombre As String = "", _
                                            Optional ByVal psIFICtaCod As String = "")
    Dim Index, row As Integer
    Dim rs As ADODB.Recordset
    
    If psProveedorCod <> "" Then
        lblProveedorCod.Caption = psProveedorCod
    End If
    If psProveedorNombre <> "" Then
        lblProveedorNombre.Caption = psProveedorNombre
    End If
    If psAreaAgeCod <> "" Then
        lblAreaAgeCod.Caption = psAreaAgeCod
    End If
    If psAreaAgeDesc <> "" Then
        lblAreaAgeNombre.Caption = psAreaAgeDesc
    End If
    If psDocTpo <> "" Then
        fsDocTpo = psDocTpo
    End If
    If psDocNro <> "" Then
        fsDocNro = psDocNro
    End If
    fsProveedorIFICod = psIFICod
    fsProveedorIFICtaCod = psIFICtaCod
    
    txtComprobanteFecEmision.value = CDate(gdFecSis) 'PASI20141229
    
    'Llenamos el Grid Detalle del Comprobante ***********
    Index = 0
    row = 0
    Call LimpiaFlex(feComprobanteDet)
    feComprobanteDet.ColumnasAEditar = "X-X-X-X-X-X-X-X-X"
    Select Case fnTipo
        Case LogTipoDocOrigenComprobante.OrdenCompra
            For Index = 1 To feOrden.Rows - 1
                If feOrden.TextMatrix(Index, 4) = "." Then
                    feComprobanteDet.AdicionaFila
                    row = feComprobanteDet.row
                    feComprobanteDet.TextMatrix(row, 1) = feOrden.TextMatrix(Index, 5)
                    feComprobanteDet.TextMatrix(row, 2) = feOrden.TextMatrix(Index, 6)
                    feComprobanteDet.TextMatrix(row, 3) = feOrden.TextMatrix(Index, 7)
                    feComprobanteDet.TextMatrix(row, 4) = feOrden.TextMatrix(Index, 8)
                    feComprobanteDet.TextMatrix(row, 5) = feOrden.TextMatrix(Index, 9)
                    feComprobanteDet.TextMatrix(row, 6) = feOrden.TextMatrix(Index, 10)
                    feComprobanteDet.TextMatrix(row, 7) = feOrden.TextMatrix(Index, 11)
                    feComprobanteDet.TextMatrix(row, 8) = feOrden.TextMatrix(Index, 3)
                    feComprobanteDet.TextMatrix(row, 9) = feOrden.TextMatrix(Index, 2)
                End If
            Next
        Case LogTipoDocOrigenComprobante.OrdenServicio
            For Index = 1 To feOrden.Rows - 1
                If feOrden.TextMatrix(Index, 4) = "." Then
                    feComprobanteDet.AdicionaFila
                    row = feComprobanteDet.row
                    feComprobanteDet.TextMatrix(row, 1) = feOrden.TextMatrix(Index, 5)
                    feComprobanteDet.TextMatrix(row, 2) = feOrden.TextMatrix(Index, 6)
                    feComprobanteDet.TextMatrix(row, 3) = feOrden.TextMatrix(Index, 7)
                    feComprobanteDet.TextMatrix(row, 4) = "Und."
                    feComprobanteDet.TextMatrix(row, 5) = "1"
                    feComprobanteDet.TextMatrix(row, 6) = feOrden.TextMatrix(Index, 11)
                    feComprobanteDet.TextMatrix(row, 7) = feOrden.TextMatrix(Index, 11)
                    feComprobanteDet.TextMatrix(row, 8) = feOrden.TextMatrix(Index, 3)
                    feComprobanteDet.TextMatrix(row, 9) = feOrden.TextMatrix(Index, 2)
                End If
            Next
        Case LogTipoDocOrigenComprobante.ContratoServicio
            For Index = 1 To feContrato.Rows - 1
                If feContrato.TextMatrix(Index, 2) = "." Then
                    Set rs = olog.ListaContratoServicioxRegistrarComprobante(fsMatrizDatos(2, 1), fsMatrizDatos(1, 1), feContrato.TextMatrix(Index, 3))
                    If Not rs.EOF Then
                        If fnTipoPagoContrato = 2 Then 'Para el ingreso de los Montos cuando el contrato tiene pago variable
                            feComprobanteDet.ColumnasAEditar = "X-X-X-X-X-5-6-X-X"
                        End If
                        Do While Not rs.EOF
                            feComprobanteDet.AdicionaFila
                            row = feComprobanteDet.row
                            feComprobanteDet.TextMatrix(row, 1) = rs!cAgeDest
                            feComprobanteDet.TextMatrix(row, 2) = rs!cCtaContCod
                            feComprobanteDet.TextMatrix(row, 3) = rs!cDescripcion
                            feComprobanteDet.TextMatrix(row, 4) = "Und."
                            feComprobanteDet.TextMatrix(row, 5) = "1"
                            feComprobanteDet.TextMatrix(row, 6) = rs!nMovImporte
                            feComprobanteDet.TextMatrix(row, 7) = rs!nMovImporte
                            feComprobanteDet.TextMatrix(row, 8) = rs!cCtaContCod
                            feComprobanteDet.TextMatrix(row, 9) = rs!nMovItem
                            rs.MoveNext
                        Loop
                    End If
                    Set rs = Nothing
                End If
            Next
        Case LogTipoDocOrigenComprobante.ContratoArrendamiento
             For Index = 1 To feContrato.Rows - 1
                If feContrato.TextMatrix(Index, 2) = "." Then
                    feComprobanteDet.AdicionaFila
                    row = feComprobanteDet.row
                    feComprobanteDet.TextMatrix(row, 1) = feContrato.TextMatrix(Index, 9)
                    feComprobanteDet.TextMatrix(row, 2) = feContrato.TextMatrix(Index, 10)
                    feComprobanteDet.TextMatrix(row, 3) = "PAGO CUOTA Nº " & feContrato.TextMatrix(Index, 3)
                    feComprobanteDet.TextMatrix(row, 4) = "Und."
                    feComprobanteDet.TextMatrix(row, 5) = "1"
                    feComprobanteDet.TextMatrix(row, 6) = feContrato.TextMatrix(Index, 6)
                    feComprobanteDet.TextMatrix(row, 7) = feContrato.TextMatrix(Index, 6)
                    feComprobanteDet.TextMatrix(row, 8) = feContrato.TextMatrix(Index, 1)
                    feComprobanteDet.TextMatrix(row, 9) = feContrato.TextMatrix(Index, 3)
                End If
            Next
        Case LogTipoDocOrigenComprobante.ContratoObra
             For Index = 1 To feContratoObra.Rows - 1
                If feContratoObra.TextMatrix(Index, 1) = "." Then
                    feComprobanteDet.AdicionaFila
                    row = feComprobanteDet.row
                    feComprobanteDet.TextMatrix(row, 1) = feContratoObra.TextMatrix(Index, 5)
                    feComprobanteDet.TextMatrix(row, 2) = feContratoObra.TextMatrix(Index, 6)
                    feComprobanteDet.TextMatrix(row, 3) = feContratoObra.TextMatrix(Index, 3)
                    feComprobanteDet.TextMatrix(row, 4) = "Und."
                    feComprobanteDet.TextMatrix(row, 5) = "1"
                    feComprobanteDet.TextMatrix(row, 6) = feContratoObra.TextMatrix(Index, 2)
                    feComprobanteDet.TextMatrix(row, 7) = feContratoObra.TextMatrix(Index, 2)
                    feComprobanteDet.TextMatrix(row, 8) = feContratoObra.TextMatrix(Index, 4)
                    feComprobanteDet.TextMatrix(row, 9) = 0
                End If
            Next
        Case LogTipoDocOrigenComprobante.ContratoAdqBienes
            For Index = 1 To feOrden.Rows - 1
                If feOrden.TextMatrix(Index, 4) = "." Then
                    feComprobanteDet.AdicionaFila
                    row = feComprobanteDet.row
'                    feComprobanteDet.TextMatrix(row, 1) = feContratoAdqBienes.TextMatrix(Index, 8)
'                    feComprobanteDet.TextMatrix(row, 2) = feContratoAdqBienes.TextMatrix(Index, 3)
'                    feComprobanteDet.TextMatrix(row, 3) = feContratoAdqBienes.TextMatrix(Index, 4)
'                    feComprobanteDet.TextMatrix(row, 4) = feContratoAdqBienes.TextMatrix(Index, 5)
'                    feComprobanteDet.TextMatrix(row, 5) = feContratoAdqBienes.TextMatrix(Index, 6)
'                    feComprobanteDet.TextMatrix(row, 6) = feContratoAdqBienes.TextMatrix(Index, 7)
'                    feComprobanteDet.TextMatrix(row, 7) = (CInt(feContratoAdqBienes.TextMatrix(Index, 6)) * CDbl(feContratoAdqBienes.TextMatrix(Index, 7)))
'                    feComprobanteDet.TextMatrix(row, 8) = feContratoAdqBienes.TextMatrix(Index, 9)
'                    feComprobanteDet.TextMatrix(row, 9) = 0
                    
                    feComprobanteDet.TextMatrix(row, 1) = feOrden.TextMatrix(row, 5)
                    feComprobanteDet.TextMatrix(row, 2) = feOrden.TextMatrix(row, 6)
                    feComprobanteDet.TextMatrix(row, 3) = feOrden.TextMatrix(row, 7)
                    feComprobanteDet.TextMatrix(row, 4) = feOrden.TextMatrix(row, 8)
                    feComprobanteDet.TextMatrix(row, 5) = feOrden.TextMatrix(row, 9)
                    feComprobanteDet.TextMatrix(row, 6) = feOrden.TextMatrix(row, 10)
                    feComprobanteDet.TextMatrix(row, 7) = feOrden.TextMatrix(row, 11)
                    feComprobanteDet.TextMatrix(row, 8) = feOrden.TextMatrix(row, 3)
                    feComprobanteDet.TextMatrix(row, 9) = feOrden.TextMatrix(row, 2)
                End If
            Next
    End Select
    '****************************************************
    If Len(fsProveedorIFICod) = 0 Or Len(fsProveedorIFICtaCod) = 0 Then
        MsgBox "Ud. debe verificar que el Proveedor " & UCase(lblProveedorNombre.Caption) & Chr(10) & "se encuentre registrado en la BD de Proveedores de Logística, además tenga" & Chr(10) & "configurado una cuenta en " & UCase(psMoneda) & " para poder continuar con el proceso.", vbInformation, "Aviso"
    Else
        '*** Predeterminamos Tipo de Pago
        If fsProveedorIFICod = "1090100012521" Then 'CMACMAYNAS
            fnTipoPago = LogTipoPagoComprobante.gPagoCuentaCMAC
        Else
            If fsProveedorIFICod = "1090100824640" Then 'BCP
               fnTipoPago = LogTipoPagoComprobante.gPagoTransferencia
            Else 'OTRO BANCO
                fnTipoPago = LogTipoPagoComprobante.gPagoCheque
            End If
        End If
        '***
    End If
End Sub
Private Sub feComprobanteDet_OnCellChange(pnRow As Long, pnCol As Long)
    On Error GoTo ErrfeOrden_OnCellChange
    If feComprobanteDet.TextMatrix(1, 0) <> "" Then
            If pnCol = 5 Or pnCol = 6 Then
                feComprobanteDet.TextMatrix(pnRow, 7) = Round(Format((feComprobanteDet.TextMatrix(pnRow, 5)) * feComprobanteDet.TextMatrix(pnRow, 6), gsFormatoNumeroView), 2)
            End If
        
    End If
    Exit Sub
ErrfeOrden_OnCellChange:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub Form_Load()
    Dim oConstSist As New NConstSistemas
    fsAreaAgeCod = gsCodArea & Right(gsCodAge, 2)
    
    If gbBitTCPonderado Then
        fnTpoCambio = gnTipCambioPonderado
    Else
        fnTpoCambio = gnTipCambioC
    End If
    If fnMoneda = gMonedaNacional Then
        fsCtaContCodOC = oConstSist.LeeConstSistema(gsLogCtasContOCMN)
        fsCtaContCodOS = oConstSist.LeeConstSistema(gsLogCtasContOSMN)
    Else
        fsCtaContCodOC = oConstSist.LeeConstSistema(gsLogCtasContOCME)
        fsCtaContCodOS = oConstSist.LeeConstSistema(gsLogCtasContOSME)
    End If
    Set olog = New DLogGeneral
    Me.Caption = "Registro de Comprobantes en " & IIf(Mid(gsopecod, 3, 1) = 1, "SOLES", "DOLARES")
    fnFormTamanioIni = 4500
    fnFormTamanioActiva = 8450
    Height = fnFormTamanioIni
    EstadoInicial True
End Sub
Private Sub EstadoInicial(ByVal phEstado As Boolean)
    ReDim fsMatrizDatos(3, 1)
    ReDim fsMatrizDatObra(2, 1)
    txtBuscar.Text = ""
    lblDescripcion.Caption = ""
    txtBuscar.Enabled = True
    MuestraGrid
    If phEstado Then
        HabilitaTab False
    End If
    CargaCombos
    EstadoControles 0
End Sub
Private Sub CargaCombos()
    Dim oConst As DConstantes
    Set oConst = New DConstantes
    Dim oDoc As New DOperacion
    Dim rs As New ADODB.Recordset
    
    
    'CargaCombo oConst.GetConstante("10042"), Me.cbotpoOrigen
    cbotpoOrigen.Clear
    CargaCombo olog.ListaTpoDocOrigenComprobante, Me.cbotpoOrigen, , 1, 0
    cbotpoOrigen.ListIndex = -1
    cbotpoOrigen.Enabled = True
    
    'Set rs = odoc.CargaOpeDoc(gnAlmaComprobanteRegistroMN, OpeDocMetDigitado)
    Set rs = oDoc.CargaOpeDoc(gnAlmaComprobanteLibreRegistroMN, OpeDocMetDigitado)
    cboTpoComprobante.Clear
    Do While Not rs.EOF
        cboTpoComprobante.AddItem Format(rs!nDocTpo, "00") & " " & Mid(rs!cDocDesc & Space(100), 1, 100) & rs!nDocTpo
        rs.MoveNext
    Loop
    Set rs = Nothing
    Set oDoc = Nothing
End Sub
Private Sub EstadoControles(ByVal pnEstado As Integer)
    Select Case pnEstado
        Case 0
                cmdMostrar.Enabled = True
                fraItemComprobante.Enabled = False
                cmdRegistrarComprobante.Enabled = False
                cmdCancelar.Enabled = False
                'cmdVerContrato.Enabled = False
                cmdNuevaValorizacion.Enabled = False
        Case 1
                cmdRegistrarComprobante.Enabled = True
                fraItemComprobante.Enabled = True
                cmdCancelar.Enabled = True
                'cmdVerContrato.Enabled = True
                cmdNuevaValorizacion.Enabled = False
        Case 2
                cmdRegistrarComprobante.Enabled = False
                fraItemComprobante.Enabled = False
                cmdCancelar.Enabled = True
                'cmdVerContrato.Enabled = True
                cmdNuevaValorizacion.Enabled = True
        Case 3
                cmdMostrar.Enabled = False
                fraItemComprobante.Enabled = False
                cmdRegistrarComprobante.Enabled = False
                cmdCancelar.Enabled = False
                'cmdVerContrato.Enabled = True
                cmdNuevaValorizacion.Enabled = False
        Case 4
                cmdRegistrarComprobante.Enabled = True
                fraItemComprobante.Enabled = True
                cmdCancelar.Enabled = True
                'cmdVerContrato.Enabled = False
                cmdNuevaValorizacion.Enabled = False
    End Select
End Sub
Private Sub HabilitaTab(ByVal phEstado As Boolean)
    Me.Top = 2000
    If Not phEstado Then
        Height = fnFormTamanioIni
        sstComprobante.Visible = False
    Else
        Height = fnFormTamanioActiva
        sstComprobante.Visible = True
    End If
End Sub
Private Sub txtBuscar_Click(psCodigo As String, psDescripcion As String)
    If cbotpoOrigen.ListIndex = -1 Then
        MsgBox "No se ha seleccionado ningún tipo de Documento.", vbExclamation, "Aviso."
        cbotpoOrigen.SetFocus
        Exit Sub
    Else
        Select Case CInt(Trim(Right(cbotpoOrigen, 100)))
            Case LogTipoDocOrigenComprobante.OrdenCompra _
                , LogTipoDocOrigenComprobante.OrdenServicio
                fsMatrizDatos = frmLogExaminaOCS.Inicio(Trim(Right(cbotpoOrigen, 100)), fsAreaAgeCod, IIf(Trim(Right(cbotpoOrigen, 100)) = 1, fsCtaContCodOC, fsCtaContCodOS))
            Case LogTipoDocOrigenComprobante.ContratoServicio _
                , LogTipoDocOrigenComprobante.ContratoArrendamiento _
                , LogTipoDocOrigenComprobante.ContratoObra _
                , LogTipoDocOrigenComprobante.ContratoAdqBienes
                fsMatrizDatos = frmLogExaminaCBSO.Inicio(Trim(Right(cbotpoOrigen, 100)), fsAreaAgeCod, CInt(Mid(gsopecod, 3, 1)))
        End Select
    End If
    
    If fsMatrizDatos(2, 1) = "" Then
        EstadoInicial False
        cbotpoOrigen.SetFocus
    Else
        cbotpoOrigen.Enabled = False
        fnTipo = CInt(Trim(Right(cbotpoOrigen.Text, 100)))
        psCodigo = fsMatrizDatos(2, 1)
        psDescripcion = fsMatrizDatos(3, 1)
        lblDescripcion.Caption = psDescripcion
        cmdMostrar.SetFocus
    End If
    
End Sub
Private Sub cboTpoComprobante_LostFocus()
    If Trim(Left(cboTpoComprobante, 2)) = "05" Then
        txtComprobanteSerie = Trim(Str("3"))
    End If
End Sub '***NAGL ERS012-2017 20170710
Private Sub txtComprobanteNro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdRegistrar.SetFocus
    End If
End Sub
Private Sub txtComprobanteSerie_KeyPress(KeyAscii As Integer)
    KeyAscii = LetrasNumeros(KeyAscii)
    If KeyAscii = 13 Then
        If Trim(Left(cboTpoComprobante, 2)) = "05" Then
           txtComprobanteSerie = Trim(Str("3"))
        Else
            txtComprobanteSerie = Right(String(4, "0") & txtComprobanteSerie, 4)
        End If 'NAGL ERS012-2017 20170710
        txtComprobanteNro.SetFocus
    End If
End Sub
Private Sub txtComprobanteSerie_LostFocus()
    If Trim(Left(cboTpoComprobante, 2)) = "05" Then
        txtComprobanteSerie = Trim(Str("3"))
    Else
        txtComprobanteSerie = Right(String(4, "0") & txtComprobanteSerie, 4)
    End If
End Sub '***NAGL ERS012-2017 20170710
Private Sub txtObservacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTpoComprobante.SetFocus
    End If
End Sub
Private Sub LimpiarDetcomprobante()
    lblProveedorCod.Caption = ""
    lblProveedorNombre.Caption = ""
    lblAreaAgeCod.Caption = ""
    lblAreaAgeNombre.Caption = ""
    txtObservacion.Text = ""
    txtComprobanteSerie.Text = ""
    txtComprobanteNro.Text = ""
    txtComprobanteFecEmision.value = CDate(gdFecSis)
    fsCtaContCodProv = ""
    fsProveedorIFICod = ""
    fsProveedorIFICtaCod = ""
    fsDocTpo = ""
    fsDocNro = ""
    fnTipoPago = 0
End Sub





