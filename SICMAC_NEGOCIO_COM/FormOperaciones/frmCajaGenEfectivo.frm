VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCajaGenEfectivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A Rendir en Efectivo"
   ClientHeight    =   5355
   ClientLeft      =   750
   ClientTop       =   2220
   ClientWidth     =   9585
   ControlBox      =   0   'False
   Icon            =   "frmCajaGenEfectivo.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreCuadre 
      Caption         =   "&Pre-Cuadre"
      Height          =   375
      Left            =   5900
      TabIndex        =   30
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   4950
      Width           =   1200
   End
   Begin TabDlg.SSTab TabBilletaje 
      Height          =   4110
      Left            =   90
      TabIndex        =   13
      Top             =   750
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   7250
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Moneda Nacional"
      TabPicture(0)   =   "frmCajaGenEfectivo.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraBilletajes(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Moneda Extranjera"
      TabPicture(1)   =   "frmCajaGenEfectivo.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraBilletajes(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fraBilletajes 
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
         Height          =   3690
         Index           =   1
         Left            =   105
         TabIndex        =   21
         Top             =   330
         Width           =   9195
         Begin SICMACT.FlexEdit fgBilletes 
            Height          =   2610
            Index           =   1
            Left            =   90
            TabIndex        =   4
            Top             =   195
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   4604
            Cols0           =   6
            FixedCols       =   2
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Descripción-Cantidad-Monto-cEfectivoCod-nEfectivoValor"
            EncabezadosAnchos=   "350-2000-800-1200-0-0"
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
            ColumnasAEditar =   "X-X-2-3-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-R-R-C-C"
            FormatosEdit    =   "0-0-3-4-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   285
            CellBackColor   =   -2147483633
         End
         Begin SICMACT.FlexEdit fgMonedas 
            Height          =   2610
            Index           =   1
            Left            =   4620
            TabIndex        =   5
            Top             =   195
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   4604
            Cols0           =   6
            FixedCols       =   2
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Descripción-Cantidad-Monto-cEfectivoCod-nEfectivoValor"
            EncabezadosAnchos=   "350-2000-800-1200-0-0"
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
            ColumnasAEditar =   "X-X-2-3-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-R-R-C-C"
            FormatosEdit    =   "0-0-3-4-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   285
            CellBackColor   =   -2147483633
         End
         Begin VB.Label lbl3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "TOTAL BILLETES :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   150
            Index           =   1
            Left            =   1140
            TabIndex        =   27
            Top             =   2925
            Width           =   1410
         End
         Begin VB.Label lblTotalBilletes 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "0.00"
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
            Height          =   300
            Index           =   1
            Left            =   2580
            TabIndex        =   26
            Top             =   2850
            Width           =   1965
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   1
            Left            =   6165
            TabIndex        =   25
            Top             =   3300
            Width           =   735
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   315
            Index           =   1
            Left            =   7125
            TabIndex        =   24
            Top             =   3255
            Width           =   1965
         End
         Begin VB.Label lblTotMoneda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "0.00"
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
            Height          =   315
            Index           =   1
            Left            =   7140
            TabIndex        =   23
            Top             =   2850
            Width           =   1965
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "TOTAL MONEDAS :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   165
            Index           =   1
            Left            =   5670
            TabIndex        =   22
            Top             =   2925
            Width           =   1440
         End
         Begin VB.Shape ShapeS 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   345
            Index           =   1
            Left            =   1035
            Top             =   2835
            Width           =   3525
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   345
            Index           =   1
            Left            =   5595
            Top             =   2835
            Width           =   3525
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   345
            Index           =   1
            Left            =   5985
            Top             =   3240
            Width           =   3120
         End
      End
      Begin VB.Frame fraBilletajes 
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
         Height          =   3690
         Index           =   0
         Left            =   -74895
         TabIndex        =   14
         Top             =   330
         Width           =   9195
         Begin SICMACT.FlexEdit fgBilletes 
            Height          =   2610
            Index           =   0
            Left            =   90
            TabIndex        =   2
            Top             =   195
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   4604
            Cols0           =   6
            FixedCols       =   2
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Descripción-Cantidad-Monto-cEfectivoCod-nEfectivoValor"
            EncabezadosAnchos=   "350-2000-800-1200-0-0"
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
            ColumnasAEditar =   "X-X-2-3-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-R-R-C-C"
            FormatosEdit    =   "0-0-3-4-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   285
            CellBackColor   =   -2147483633
         End
         Begin SICMACT.FlexEdit fgMonedas 
            Height          =   2610
            Index           =   0
            Left            =   4620
            TabIndex        =   3
            Top             =   195
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   4604
            Cols0           =   6
            FixedCols       =   2
            HighLight       =   2
            AllowUserResizing=   3
            EncabezadosNombres=   "N°-Descripción-Cantidad-Monto-cEfectivoCod-nEfectivoValor"
            EncabezadosAnchos=   "350-2000-800-1200-0-0"
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
            ColumnasAEditar =   "X-X-2-3-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0-0-0-0"
            EncabezadosAlineacion=   "C-L-R-R-C-C"
            FormatosEdit    =   "0-0-3-4-0-0"
            AvanceCeldas    =   1
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   285
            CellBackColor   =   -2147483633
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "TOTAL MONEDAS :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   165
            Index           =   0
            Left            =   5670
            TabIndex        =   20
            Top             =   2925
            Width           =   1440
         End
         Begin VB.Label lblTotMoneda 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
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
            Height          =   315
            Index           =   0
            Left            =   7140
            TabIndex        =   19
            Top             =   2850
            Width           =   1965
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   315
            Index           =   0
            Left            =   7125
            TabIndex        =   18
            Top             =   3255
            Width           =   1965
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004080&
            Height          =   195
            Index           =   0
            Left            =   6165
            TabIndex        =   17
            Top             =   3300
            Width           =   735
         End
         Begin VB.Label lblTotalBilletes 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
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
            Height          =   300
            Index           =   0
            Left            =   2580
            TabIndex        =   16
            Top             =   2850
            Width           =   1965
         End
         Begin VB.Label lbl3 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "TOTAL BILLETES :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   150
            Index           =   0
            Left            =   1140
            TabIndex        =   15
            Top             =   2925
            Width           =   1410
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   345
            Index           =   0
            Left            =   5985
            Top             =   3240
            Width           =   3120
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   345
            Index           =   0
            Left            =   5595
            Top             =   2835
            Width           =   3525
         End
         Begin VB.Shape ShapeS 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H8000000C&
            Height          =   345
            Index           =   0
            Left            =   1050
            Top             =   2835
            Width           =   3525
         End
      End
   End
   Begin SICMACT.Usuario oUser 
      Left            =   5040
      Top             =   4140
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   4950
      Width           =   1200
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Cuadre"
      Height          =   375
      Left            =   7095
      TabIndex        =   6
      Top             =   4950
      Width           =   1200
   End
   Begin VB.Frame fraDatosPrinc 
      Height          =   690
      Left            =   90
      TabIndex        =   9
      Top             =   0
      Width           =   9360
      Begin SICMACT.TxtBuscar txtBuscarUser 
         Height          =   360
         Left            =   840
         TabIndex        =   0
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   8388608
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   345
         Left            =   7770
         TabIndex        =   1
         Top             =   225
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCaptionUser 
         AutoSize        =   -1  'True
         Caption         =   "Cajero :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   12
         Top             =   240
         Width           =   675
      End
      Begin VB.Label lblDescUser 
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
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2205
         TabIndex        =   11
         Top             =   195
         Width           =   4695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   195
         Left            =   7155
         TabIndex        =   10
         Top             =   255
         Width           =   540
      End
   End
   Begin VB.Label lblMontoSol 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "0.00"
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
      Height          =   345
      Left            =   1920
      TabIndex        =   29
      Top             =   4935
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "MONTO SOLICITADO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   165
      Left            =   180
      TabIndex        =   28
      Top             =   5025
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Shape ShapeMonSol 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   375
      Left            =   120
      Top             =   4920
      Visible         =   0   'False
      Width           =   3855
   End
End
Attribute VB_Name = "frmCajaGenEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vbOk As Boolean
Dim rsBillAuxMN As ADODB.Recordset
Dim rsMonAuxMN As ADODB.Recordset
Dim rsBillAuxME As ADODB.Recordset
Dim rsMonAuxME As ADODB.Recordset
Dim lnMoneda As Moneda
Dim lnMonto As Double
Dim lnArendirFase As COMDConstantes.ARendirFases
Dim lnDiferencia As Double
Dim lbDiferencia As Boolean
Dim lsCaptionArendir As String
Dim lbEnableFecha As Boolean
Dim lbMuestra As Boolean
Dim lbRegistro As Boolean
Dim lsMovNro As String
Dim lnMovNro As Long
Dim lbModifica As Boolean
Dim lbSalir As Boolean
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim lnParamDif As Double

'**DAOR 20080125 *************************************
Dim fnNumMaxRegEfecSinSolicExt As Integer
Dim fnVecesTotal As Integer
Dim fnVecesVigente As Integer
Dim fnMovNroRegEfec As Long, fnMovNroRegEfecUlt As Long
'*****************************************************
'**MADM 20101006 *************************************
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
'*****************************************************

Public Sub Muestra(ByVal psMovNro As String, Optional nMontoSol As Double = 0, _
        Optional ByVal nMon As Moneda = gMonedaNacional)
lbMuestra = True
lsMovNro = psMovNro
lnMoneda = nMon
lnMonto = nMontoSol
lblMontoSol = Format(Abs(nMontoSol), "##,###0.00")
Me.Show 1
End Sub

Public Sub Inicio(ByVal psOpeCod As String, ByVal psOpeDesc As String, _
            ByVal pnMontoSol As Double, ByVal pnMoneda As Moneda, _
            Optional ByVal pbEnableFecha As Boolean = True, Optional pbCubreDif As Boolean = False)
 
lbDiferencia = pbCubreDif
lbEnableFecha = pbEnableFecha
cmdImprimir.Visible = False
lnMoneda = pnMoneda
If lnMoneda = gMonedaNacional Then
    TabBilletaje.TabEnabled(0) = True
    TabBilletaje.TabEnabled(1) = False
    TabBilletaje.Tab = 0
Else
    TabBilletaje.TabEnabled(0) = False
    TabBilletaje.TabEnabled(1) = True
    TabBilletaje.Tab = 1
End If
lnMonto = pnMontoSol
lblMontoSol = Format(Abs(pnMontoSol), "##,###0.00")
If Val(pnMontoSol) < 0 Then
    lblMontoSol.ForeColor = &HFF&
Else
    lblMontoSol.ForeColor = &HC00000
End If
'MIOL 20120625, SEGUN RQ12093 *************************************************************
If psOpeCod = 901016 Then
    Me.cmdPreCuadre.Visible = True
Else
    Me.cmdPreCuadre.Visible = False
    Me.cmdAceptar.Caption = "&Aceptar"
End If
'END MIOL *********************************************************************************
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
'**DAOR 20080125 ********************************
Dim lRsExt As ADODB.Recordset
Dim lsOperacion As String
Dim loCaja As COMNCajaGeneral.NCOMCajaGeneral
Dim lsMovNroExt As String
'************************************************
Dim lsConSisValor As String 'FRHU 20160226
Dim lbConSisValor As Boolean 'FRHU 20160226
Dim lsOpeCod As String 'FRHU 20160226
'MIOL 20120705, SEGUN RQ12093 *******************
Set oCajero = New COMNCajaGeneral.NCOMCajero
If gsOpeCod = "901016" Then
    If oCajero.YaRealizoDevBilletaje(gsCodUser, gdFecSis, gsCodAge) Then
        MsgBox "Ud. ha realizado la operación de registro de efectivo, esta operación no esta disponible después del registro de efectivo", vbInformation, "Aviso"
        cmdAceptar.Enabled = False
        lbModifica = False
        Exit Sub
    End If
End If
'END MIOL ***************************************

'MIOL 20120727, *********************************
Set oCajero = New COMNCajaGeneral.NCOMCajero
If gsOpeCod = "901007" Then
    If oCajero.YaRealizoRegEfectivo(gsCodUser, gdFecSis, gsCodAge) Then
        MsgBox "Ud. ha realizado la operación de registro de efectivo, esta operación no esta disponible después del registro de efectivo", vbInformation, "Aviso"
        cmdAceptar.Enabled = False
        lbModifica = False
        Exit Sub
    End If
End If
'END MIOL ***************************************

'GIPO ERS051-2016
Dim rsTarj As ADODB.Recordset
Set rsTarj = oCajero.ObtenerTarjetasADevolver(gsCodUser, gdFecSis, gsCodAge)
If rsTarj.RecordCount > 0 Then
    MsgBox "No se puede realizar el Cuadre porque aún tiene pendiente devolución de " & rsTarj.RecordCount & " Tarjetas. " & _
    "Diríjase al SICMACM Tarjetas>Control Stock>Interno>Devolver Tarjetas", vbInformation, "Aviso"
    Exit Sub
End If
'END GIPO

lnDiferencia = 0
If Val(lblTotal(0)) = 0 And Val(lblTotal(1)) = 0 Then
    If MsgBox("¿Desea Registrar Monto de Billetaje = Cero", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
End If
If lbRegistro Then
    If lbModifica Then
        Dim oCont As COMNContabilidad.NCOMContFunciones  'NContFunciones
        Dim lbNuevo  As Boolean
        Set oCont = New COMNContabilidad.NCOMContFunciones
        '''If MsgBox("Desea Registrar el billetaje Ingresado por S/." & lblTotal(0) & " y $." & lblTotal(1), vbYesNo + vbQuestion, "Aviso") = vbYes Then 'marg ers044-2016
        If MsgBox("Desea Registrar el billetaje Ingresado por " & gcPEN_SIMBOLO & lblTotal(0) & " y $." & lblTotal(1), vbYesNo + vbQuestion, "Aviso") = vbYes Then 'marg ers044-2016
            vbOk = True

            Set rsBillAuxMN = fgBilletes(0).GetRsNew()
            Set rsMonAuxMN = fgMonedas(0).GetRsNew()
            Set rsBillAuxME = fgBilletes(1).GetRsNew()
            Set rsMonAuxME = fgMonedas(1).GetRsNew()
            
            'MIOL 20120705, SEGUN RQ12093 *******************
            'Set oCajero = New COMNCajaGeneral.NCOMCajero
            
            '**DAOR 20080125, Extorno automático del registro de efectivo anterior*******
            If fnVecesTotal < fnNumMaxRegEfecSinSolicExt And fnVecesVigente >= 1 Then
                If gsOpeCod = gOpeBoveAgeRegEfect Then
                    lsOperacion = COMDConstSistema.gOpeBoveAgeRegEfect
                    Set lRsExt = oCajero.ObtenerRegistrosDeEfectivo(lsOperacion, CDate(txtfecha.Text), CDate(txtfecha.Text), txtBuscarUser.Text, gsCodAge)
                    lsOperacion = COMDConstSistema.gOpeBoveAgeExtRegSobFalt
                Else
                    lsOperacion = COMDConstSistema.gOpeHabCajDevBilletaje
                    Set lRsExt = oCajero.GetDevCajero(lsOperacion, CDate(txtfecha.Text), CDate(txtfecha.Text), gsCodUser)
                    lsOperacion = COMDConstSistema.gOpeHabCajExtDevBilletaje
                End If
                If lRsExt.RecordCount > 0 Then
                    lsMovNroExt = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                    Set loCaja = New COMNCajaGeneral.NCOMCajaGeneral
                    loCaja.GrabaExtornoMov gdFecSis, lRsExt!Fecha, lsMovNroExt, lRsExt!nMovNro, lsOperacion, "Extorno Automatico (Primer Billetaje)", lRsExt!nMovImporte
                    Set loCaja = Nothing
                End If
            End If
            '****************************************************************************
            
            '**Modficado por DAOR 20080125, Todos las operaciones de efectivo deberan crear un nuevo registro ***
            'If lsMovNro <> "" Then
            '    lbNuevo = False
            'Else
            '    lbNuevo = True
            '    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            'End If

            lbNuevo = True
            lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            '*****************************************************************************************************
            
            '**DAOR 20080128, se aumentó el parámetro fnMovNroRegEfec
            If oCajero.GrabaRegistroEfectivo(gsFormatoFecha, lsMovNro, _
                            gsOpeCod, gsOpeDesc, rsBillAuxMN, rsBillAuxME, rsMonAuxMN, rsMonAuxME, txtBuscarUser.Text, lbNuevo, fnMovNroRegEfec) = 0 Then
                cmdImprimir.Enabled = True
                cmdImprimir.SetFocus
            End If
            fgBilletes(0).SetFocus
            '**DAOR 20080125 ***********************************************************************
            '**Iniciar el registro de sobrante y faltante ******************************************
            cmdAceptar.Enabled = False
            'FRHU 20160222 Riesgo Operativo
            lsOpeCod = Trim(LeeConstanteSist(gConstSistOpeTpoMostrarRegistroEfectivo))
            lsConSisValor = LeeConstanteSist(gConstSistMostrarRegistroEfectivo)
            lbConSisValor = IIf(lsConSisValor = "1", 1, 0)
            If lbConSisValor And InStr(1, lsOpeCod, gsOpeCod) > 0 Then
                Call frmCajaGenMostrarBilletaje.Inicio(fnMovNroRegEfec)
            End If
            'FIN FRHU
            If gsOpeCod = gOpeBoveAgeRegEfect Then
                frmCajeroIngEgre.Inicia True, False, , lsMovNro, fnMovNroRegEfec, fnVecesTotal + 1
            Else
                frmCajeroIngEgre.Inicia False, False, , lsMovNro, fnMovNroRegEfec, fnVecesTotal + 1
            End If
            '**************************************************************************************
        End If
        Set oCont = Nothing
    Else
        vbOk = True
        Me.Hide
    End If
Else
    Dim nMonto As Double, nMontoSol As Double
    If lnMoneda = gMonedaNacional Then
        Set rsBillAuxMN = fgBilletes(0).GetRsNew()
        Set rsMonAuxMN = fgMonedas(0).GetRsNew()
        nMonto = CDbl(lblTotal(0))
    Else
        Set rsBillAuxMN = fgBilletes(1).GetRsNew()
        Set rsMonAuxMN = fgMonedas(1).GetRsNew()
        nMonto = CDbl(lblTotal(1))
    End If
    nMontoSol = CDbl(lblMontoSol)
    If nMonto <> nMontoSol Then
        MsgBox "Monto total de descomposisión de efectivo no coincide con el monto solicitado", vbInformation, "Aviso"
        TabBilletaje.SetFocus
        Exit Sub
    End If
    vbOk = True
    Me.Hide
End If
Set oCajero = Nothing

End Sub

Private Sub cmdCancelar_Click()
vbOk = False
If lbRegistro Then
    Unload Me
    Set frmCajaGenEfectivo = Nothing
Else
    Me.Hide
End If
End Sub

Public Property Get lbOk() As Variant
    lbOk = vbOk
End Property

Public Property Let lbOk(ByVal vNewValue As Variant)
    vbOk = vNewValue
End Property

Public Property Get rsBilletes() As ADODB.Recordset
    Set rsBilletes = rsBillAuxMN
End Property

Public Property Get rsMonedas() As ADODB.Recordset
    Set rsMonedas = rsMonAuxMN
End Property

Public Property Set rsBilletes(ByVal vNewValue As ADODB.Recordset)
    Set rsBilletes = vNewValue
End Property

Public Property Set rsMonedas(ByVal vNewValue As ADODB.Recordset)
    Set rsMonedas = vNewValue
End Property

Public Function ImprimirBilletaje()
    cmdImprimir_Click
End Function

Private Sub cmdImprimir_Click()

Dim oPrevio As previo.clsprevio
Dim rsMonMN As New ADODB.Recordset, rsMonME As New ADODB.Recordset
Dim rsBillMN As New ADODB.Recordset, rsBillME As New ADODB.Recordset
Dim sUsuario As String
Dim sCadImp As String

If fgBilletes(0).GetRsNew() Is Nothing Then
    Set rsBillMN = Nothing
Else
    Set rsBillMN = fgBilletes(0).GetRsNew()
End If

If fgBilletes(1).GetRsNew() Is Nothing Then
    Set rsBillME = Nothing
Else
    Set rsBillME = fgBilletes(1).GetRsNew()
End If

If fgMonedas(0).GetRsNew() Is Nothing Then
    Set rsMonMN = Nothing
Else
    Set rsMonMN = fgMonedas(0).GetRsNew()
End If

If fgMonedas(1).GetRsNew() Is Nothing Then
    Set rsMonME = Nothing
Else
    Set rsMonME = fgMonedas(1).GetRsNew()
End If
sUsuario = Trim(lblDescUser)

'ARCV 01-05-2007
'Set oCajero = New COMNCajaGeneral.NCOMCajero

    sCadImp = ImprimeBilletaje(rsBillMN, rsBillME, rsMonMN, rsMonME, gsNomAge, txtBuscarUser.Text, sUsuario, gdFecSis)
'Set oCajero = Nothing
    
Set oPrevio = New previo.clsprevio
    oPrevio.Show sCadImp, "BILLETAJE - " & txtBuscarUser, False
Set oPrevio = Nothing
End Sub

'ARCV 01-05-2007
'------------------------------
Public Function ImprimeBilletaje(ByVal rsBillMN As ADODB.Recordset, ByVal rsBillME As ADODB.Recordset, _
        ByVal rsMonMN As ADODB.Recordset, ByVal rsMonME As ADODB.Recordset, _
        ByVal sNomAge As String, ByVal sUsuario As String, ByVal sNomUsuario As String, _
        ByVal dFecSis As Date) As String

Dim i As Integer, nCarLin As Long
Dim sCad As String
Dim sNumPag As String
Dim sTitRp1 As String, sTitRp2 As String
Dim sTotalMonto As String * 16
Dim nTotalMon As Double, nTotalBill As Double
Dim sMoneda As String, sMonedaTit As String
Dim sSimbolo As String
Dim oImpre As New COMNCaptaGenerales.NCOMCaptaImpresion
Dim nTotalMN, nTotalME As Double 'EJVG 20110726
'Dim oImp As COMFunciones.FCOMVarImpresion
'Set oImp = New COMFunciones.FCOMVarImpresion


sCad = ""

nCarLin = 85

sTitRp1 = "DESCOMPOSICION DE EFECTIVO - " & sUsuario
sTitRp2 = sNomUsuario

sCad = sCad & oImpre.CabeRepoCaptac("", "", nCarLin, "SECCION OPERACIONES", sTitRp1, sTitRp2, sMoneda, "1", Trim(sNomAge), dFecSis, Chr$(10)) & Chr$(10)

'''sMoneda = "SOLES" 'marg ers044-2016
sMoneda = StrConv(gcPEN_PLURAL, vbUpperCase) 'marg ers044-2016
sMonedaTit = "MONEDA NACIONAL"
'''sSimbolo = "S/." 'MARG ERS044-2016
sSimbolo = gcPEN_SIMBOLO & " " 'MARG ERS044-2016
sCad = sCad & sMonedaTit & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)

'MsgBox "Entra a imprime BilletajeRS"

sCad = sCad & ImprimeRsBilletaje(rsBillMN, nTotalBill)
nTotalMN = nTotalBill 'EJVG 20110726
'MsgBox "Sale de imprime BilletajeRS"

RSet sTotalMonto = Format$(nTotalBill, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & "SUB TOTAL BILLETAJE" & Space(23) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & ImprimeRsBilletaje(rsMonMN, nTotalMon)
RSet sTotalMonto = Format$(nTotalMon, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & "SUB TOTAL MONEDA" & Space(26) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
RSet sTotalMonto = sSimbolo & " " & Format$(nTotalBill + nTotalMon, "#,##0.00")
sCad = sCad & "TOTAL " & sMonedaTit & Space(21) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & Chr$(10)

nTotalMN = nTotalMN + nTotalMon

sMoneda = "DOLARES"
sMonedaTit = "MONEDA EXTRANJERA"
sSimbolo = "US$"
sCad = sCad & sMonedaTit & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & ImprimeRsBilletaje(rsBillME, nTotalBill)
RSet sTotalMonto = Format$(nTotalBill, "#,##0.00")
sCad = sCad & String(nCarLin, "-") & Chr$(10)
sCad = sCad & "SUB TOTAL BILLETAJE" & Space(23) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)
'sCad = sCad & ImprimeRsBilletaje(rsMonME, nTotalMon)
'RSet sTotalMonto = Format$(nTotalMon, "#,##0.00")
'sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'sCad = sCad & "SUB TOTAL MONEDA" & Space(26) & sTotalMonto & oImp.gPrnSaltoLinea
'sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'RSet sTotalMonto = sSimbolo & " " & Format$(nTotalBill + nTotalMon, "#,##0.00")
sCad = sCad & "TOTAL " & sMonedaTit & Space(21) & sTotalMonto & Chr$(10)
sCad = sCad & String(nCarLin, "-") & Chr$(10)

nTotalME = nTotalBill

'Set oImp = Nothing

'EJVG 20110726
'Se adicionó al Reporte de Descomposición de Dinero el Acta de Arqueo para la Operación 901016 y 901007
If gsOpeCod = CStr(gOpeHabCajRegEfect) Or gsOpeCod = CStr(gOpeBoveAgeRegEfect) Then 'EAAS 20180202
    Dim sTitulo As String
    sTitulo = "ACTA DE ARQUEO"
    'sCad = sCad & Chr$(10)
    sCad = sCad & String(Int((IIf(nCarLin <= Len(sTitulo), Len(sTitulo) + 1, nCarLin) - Len(sTitulo)) / 2), " ") & sTitulo & Chr$(10) _
    & "FECHA: " & Format$(dFecSis, "dd/mm/yyyy") & "  " & "HORA: " & Chr$(10) _
    & "DATOS DEL SISTEMA" & Space(5) & "DATOS FISICO" & Space(5) & "DIFERENCIA" & Chr$(10) _
    & "-----------------" & Space(5) & "------------" & Space(5) & "----------" & Chr$(10) _
    & gcPEN_SIMBOLO & " " & Space(18) & FillNum(Format$(nTotalMN, "#,##0.00"), 12, " ") & Space(5) & Space(10) & Chr$(10) _
    & "US$ " & Space(18) & FillNum(Format$(nTotalME, "#,##0.00"), 12, " ") & Space(5) & Space(10) & Chr$(10) _
    & String(nCarLin, "-") & Chr$(10) _
    & "OBSERVACIONES:" & Chr$(10) _
    & Space(nCarLin) & Chr$(10) _
    & String(nCarLin, "-") & Chr$(10) _
    & Space(8) & "REALIZA EL ARQUEO" & Space(15) & "|" & Space(12) & "ACEPTA EL ARQUEO" & Chr$(10) _
    & String(nCarLin, "-") & Chr$(10) _
    & Space(40) & "|" & Space(44) & Chr$(10) _
    & Space(40) & "|" & Space(44) & Chr$(10) _
    & Space(40) & "|" & Space(44) & Chr$(10) _
    & Space(40) & "|" & Space(44) & Chr$(10) _
    & "NOMBRE:" & Space(33) & "|NOMBRE:" & Mid(oUser.UserNom, 1, 44) & Chr$(10) _
    & "CARGO :" & Space(33) & "|CARGO :" & Mid(oUser.PersCargo, 1, 44) & Chr$(10) _
    & String(nCarLin, "-") & Chr$(10) 'marg ers044-2016
End If
ImprimeBilletaje = sCad
End Function

Private Function ImprimeRsBilletaje(ByVal rs As ADODB.Recordset, ByRef nTotalMonto As Double) As String
Dim sCantidad As String * 8
Dim sMonto As String * 16
Dim sDescripcion As String * 30
Dim sCad As String
sCad = ""
nTotalMonto = 0

'Dim oImp As COMFunciones.FCOMVarImpresion
'Set oImp = New COMFunciones.FCOMVarImpresion

If Not rs Is Nothing Then
    Do While Not rs.EOF
        RSet sCantidad = Trim(rs("Cantidad"))
        nTotalMonto = nTotalMonto + rs("Monto")
        RSet sMonto = Trim(rs("Monto"))
        sDescripcion = Trim(rs("Descripción"))
        sCad = sCad & sDescripcion & Space(2) & sCantidad & Space(2) & sMonto & Chr$(10)
        rs.MoveNext
    Loop
End If

'Set oImp = Nothing

ImprimeRsBilletaje = sCad
End Function

'MIOL 20120601, SEGUN RQ12093 ***************************************************************
Private Sub cmdPreCuadre_Click()
    Dim oDCredDoc As DCredDoc
    Set oDCredDoc = New DCredDoc
    
    If oDCredDoc.YaRealizoDevBilletajePreCuadre(gsCodUser, gdFecSis, gsCodAge) Then
        MsgBox "Ud. ha realizado la operación de registro de PreCuadre, esta operación no esta disponible después del registro de efectivo", vbInformation, "Aviso"
        cmdPreCuadre.Enabled = False
        Exit Sub
    Else
        frmCajeroCorte.Show 1
    End If
End Sub
'END MIOL  **********************************************************************************
'--------------------------------

Private Sub fgBilletes_OnValidate(Index As Integer, ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lnTotal As Currency
Dim lnValor As Currency
Select Case pnCol
    Case 3
        lnValor = CCur(IIf(fgBilletes(Index).TextMatrix(pnRow, pnCol) = "", "0", fgBilletes(Index).TextMatrix(pnRow, pnCol)))
        If Residuo(lnValor, CCur(fgBilletes(Index).TextMatrix(pnRow, 5))) Then
            fgBilletes(Index).TextMatrix(pnRow, 2) = Format(Round(lnValor / fgBilletes(Index).TextMatrix(pnRow, 5), 0), "#,##0")
        Else
            Cancel = False
            Exit Sub
        End If
    Case 2
        lnValor = CCur(IIf(fgBilletes(Index).TextMatrix(pnRow, pnCol) = "", "0", fgBilletes(Index).TextMatrix(pnRow, pnCol)))
        fgBilletes(Index).TextMatrix(pnRow, 3) = Format(lnValor * CCur(IIf(fgBilletes(Index).TextMatrix(pnRow, 5) = "", "0", fgBilletes(Index).TextMatrix(pnRow, 5))), "#,##0.00")
End Select
lblTotalBilletes(Index) = Format(fgBilletes(Index).SumaRow(3), "#,##0.00")
lnTotal = Format(CCur(lblTotalBilletes(Index)) + CCur(lblTotMoneda(Index)), "#,##0.00")
lblTotal(Index) = Format(lnTotal, "#,##0.00")
End Sub
    
Private Sub fgMonedas_OnValidate(Index As Integer, ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lnTotal As Currency
Dim lnValor As Currency
Select Case pnCol
    Case 3
        lnValor = CCur(IIf(fgMonedas(Index).TextMatrix(pnRow, pnCol) = "", "0", fgMonedas(Index).TextMatrix(pnRow, pnCol)))
        If Residuo(lnValor, CCur(fgMonedas(Index).TextMatrix(pnRow, 5))) Then
            fgMonedas(Index).TextMatrix(pnRow, 2) = Format(Round(lnValor / fgMonedas(Index).TextMatrix(pnRow, 5), 0), "#,##0")
        Else
            Cancel = False
            Exit Sub
        End If
    Case 2
        lnValor = CCur(IIf(fgMonedas(Index).TextMatrix(pnRow, pnCol) = "", "0", fgMonedas(Index).TextMatrix(pnRow, pnCol)))
        fgMonedas(Index).TextMatrix(pnRow, 3) = Format(lnValor * CCur(IIf(fgMonedas(Index).TextMatrix(pnRow, 5) = "", "0", fgMonedas(Index).TextMatrix(pnRow, 5))), "#,##0.00")
End Select
lblTotMoneda(Index) = Format(fgMonedas(Index).SumaRow(3), "#,##0.00")
lnTotal = Format(CCur(lblTotalBilletes(Index)) + CCur(lblTotMoneda(Index)), "#,##0.00")
lblTotal(Index) = Format(lnTotal, "#,##0.00")
End Sub

Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Set oCajero = New COMNCajaGeneral.NCOMCajero
Dim oGen As COMDConstSistema.DCOMGeneral  'DGeneral
Dim lrs As ADODB.Recordset 'DAOR 20080125

Set oGen = New COMDConstSistema.DCOMGeneral
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
lnParamDif = oGen.GetParametro(4000, 1001)
fnNumMaxRegEfecSinSolicExt = oGen.GetParametro(4000, 1002) 'DAOR 20080125

Set oGen = Nothing
lbSalir = False
CargaBilletajes gMonedaNacional, lsMovNro
CargaBilletajes gMonedaExtranjera, lsMovNro

If lbMuestra = False Then
    txtfecha = gdFecSis
    txtfecha.Enabled = lbEnableFecha
Else
    txtfecha = Mid(lsMovNro, 7, 2) & "/" & Mid(lsMovNro, 5, 2) & "/" & Left(lsMovNro, 4)
    txtfecha.Enabled = False
    fgBilletes(0).lbEditarFlex = False
    fgMonedas(0).lbEditarFlex = False
    fgBilletes(1).lbEditarFlex = False
    fgMonedas(1).lbEditarFlex = False
End If

lblCaptionUser.Visible = False
txtBuscarUser.Visible = False
lblDescUser.Visible = False

If gsOpeCod = "901003" Then
    lbRegistro = False
End If

If lbRegistro = True Then
    Dim oGeneral As COMDConstSistema.DCOMGeneral
    Set oGeneral = New COMDConstSistema.DCOMGeneral
    Dim lsOpeCod As String, sUsuario As String
    
    lblCaptionUser.Visible = True
    txtBuscarUser.Visible = True
    lblDescUser.Visible = True
    txtfecha = gdFecSis
    txtfecha.Enabled = False
    txtBuscarUser.Enabled = False
    oUser.Inicio gsCodUser
    If gsOpeCod = gOpeBoveAgeRegEfect Then
        sUsuario = gsUsuarioBOVEDA
        lsOpeCod = gOpeBoveAgeRegEfect
        lblDescUser = "BOVEDA"
        cmdPreCuadre.Visible = False
    Else
        sUsuario = gsCodUser
        lsOpeCod = gOpeHabCajRegEfect
        lblDescUser = PstaNombre(oUser.UserNom)
    End If
    txtBuscarUser = sUsuario
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    lsMovNro = oCajero.GetMovUserBilletaje(lsOpeCod, sUsuario, gdFecSis, , gsCodAge)
    lnMovNro = Val(oCajero.GetMovUserBilletaje(lsOpeCod, sUsuario, gdFecSis, True))
    If oCajero.GetMovUserBilletajeDevuelto(lsOpeCod, sUsuario, gdFecSis) <> "" Then
        MsgBox "Se ha realizado la Devolución del billetaje registrado ", vbInformation, "Aviso"
        lbSalir = True
    End If
    CargaBilletajes gMonedaNacional, lsMovNro
    CargaBilletajes gMonedaExtranjera, lsMovNro
    
    '**DAOR 20080125, Verificar número de veces de registro de efectivo (Billetaje)**
    Set lrs = oCajero.ObtenerValidacionRegistroEfectivo(sUsuario, gsCodAge, Format(gdFecSis, "yyyymmdd"), lsOpeCod)
    If Not lrs.EOF And Not lrs.BOF Then
        fnMovNroRegEfecUlt = lrs!nMovNro
        fnVecesTotal = lrs!nVecesTotal
        fnVecesVigente = lrs!nVecesVigente
        'Add By Gitu 23-10-2009
        If fnVecesTotal >= 1 Then
            CargaBilletajes gMonedaNacional, lsMovNro, True
            CargaBilletajes gMonedaExtranjera, lsMovNro, True
        End If
        'End Gitu
        'Comentado x MADM 20110923
        ''        If fnVecesTotal >= fnNumMaxRegEfecSinSolicExt And fnVecesVigente >= 1 Then
        ''            MsgBox "Se ha llegado al número máximo de registros de efectivo, para volver a realizar esta operación es necesario que realice un : Extorno de Devolución por Billetaje o Extorno de Registro de Efectivo ", vbInformation, "Aviso"
        ''            cmdAceptar.Enabled = False
        ''            lbModifica = False
        ''        End If
        'END MADM
         
        'MADM 20110923 -MADM 20101012
                'If ((fnVecesTotal = 1) Or (fnVecesTotal = 2 And fnVecesVigente = 0)) Then
                
                'Comentado Por MIOL 20120705 ***
                'If ((fnVecesTotal >= 1) And (fnVecesVigente = 0)) Then
                '    Set loVistoElectronico = New SICMACT.frmVistoElectronico
                '    lbVistoVal = loVistoElectronico.Inicio(3, gsOpeCod)
                '    If Not lbVistoVal Then
                '           CmdAceptar.Enabled = False
                '           lbModifica = False
                '    End If
                'End If
                '***
                
         'END MADM
    End If
    '********************************************************************************
    If lbModifica Then
        fgBilletes(0).lbEditarFlex = True
        fgMonedas(0).lbEditarFlex = True
        fgBilletes(1).lbEditarFlex = True
        fgMonedas(1).lbEditarFlex = True
    Else
        fgBilletes(0).lbEditarFlex = False
        fgMonedas(0).lbEditarFlex = False
        fgBilletes(1).lbEditarFlex = False
        fgMonedas(1).lbEditarFlex = False
    End If
    Set oGeneral = Nothing
    TabBilletaje.Tab = 0
End If
Set oCajero = Nothing
End Sub

Public Sub RegistroEfectivo(Optional pbModifica As Boolean = True, Optional pcOpeCod As Long = 0)
lbRegistro = True
lbModifica = pbModifica

Me.Caption = "Operaciones - Billetaje"
Me.cmdImprimir.Enabled = False 'MIOL 20130506, SEGUN RQ13207
'MADM 20110207
If pcOpeCod <> 0 Then
    gsOpeCod = CStr(pcOpeCod)
End If
'END MADM
Me.Show 1
End Sub

Private Sub CargaBilletajes(ByVal nMoneda As Moneda, Optional sMovNro As String = "", Optional bLimBille As Boolean = False)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oContFunct As COMNContabilidad.NCOMContFunciones  'NContFunciones
Dim oEfec As COMDCajaGeneral.DCOMEfectivo   'Defectivo
Dim lnFila As Long
Dim i As Integer

Set oContFunct = New COMNContabilidad.NCOMContFunciones
Set oEfec = New COMDCajaGeneral.DCOMEfectivo

Set rs = New ADODB.Recordset
If lbRegistro Then
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    If sMovNro <> "" Then
        Set rs = oCajero.GetBilletajeCajero(sMovNro, txtBuscarUser, nMoneda, "B")
    Else
        Set rs = oEfec.EmiteBilletajes(nMoneda, "B")
    End If
     Set oCajero = Nothing
Else
    If lbMuestra = False Then
        Set rs = oEfec.EmiteBilletajes(nMoneda, "B")
    Else
        Set rs = oEfec.GetBilletajesMov(sMovNro, nMoneda, "B")
    End If
End If

i = IIf(nMoneda = gMonedaNacional, 0, 1)

fgBilletes(i).FontFixed.Bold = True
fgBilletes(i).Clear
fgBilletes(i).FormaCabecera
fgBilletes(i).Rows = 2
Do While Not rs.EOF
    fgBilletes(i).AdicionaFila
    lnFila = fgBilletes(i).row
    fgBilletes(i).TextMatrix(lnFila, 1) = rs!Descripcion
    
    'Modificado por gitu 23-10-2009
    If bLimBille Then
        fgBilletes(i).TextMatrix(lnFila, 2) = Format(0, "#,##0")
        fgBilletes(i).TextMatrix(lnFila, 3) = Format(0, "#,##0.00")
    Else
        '*** PEAC 20090924
        fgBilletes(i).TextMatrix(lnFila, 2) = Format(rs!Cantidad, "#,##0")
        fgBilletes(i).TextMatrix(lnFila, 3) = Format(rs!Monto, "#,##0.00")
        '*** FIN PEAC
    End If
    'End Gitu
    
    fgBilletes(i).TextMatrix(lnFila, 4) = rs!cEfectivoCod
    fgBilletes(i).TextMatrix(lnFila, 5) = rs!nEfectivoValor
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
fgBilletes(i).Col = 2

Set rs = New ADODB.Recordset
Set oCajero = New COMNCajaGeneral.NCOMCajero
If lbRegistro Then
    If sMovNro <> "" Then
        Set rs = oCajero.GetBilletajeCajero(sMovNro, txtBuscarUser, nMoneda, "M")
    Else
        Set rs = oEfec.EmiteBilletajes(nMoneda, "M")
    End If
Else
    If lbMuestra = False Then
        Set rs = oEfec.EmiteBilletajes(nMoneda, "M")
    Else
        Set rs = oEfec.GetBilletajesMov(sMovNro, nMoneda, "M")
    End If
End If

fgMonedas(i).FontFixed.Bold = True
fgMonedas(i).Clear
fgMonedas(i).FormaCabecera
fgMonedas(i).Rows = 2
Do While Not rs.EOF
    fgMonedas(i).AdicionaFila
    lnFila = fgMonedas(i).row
    fgMonedas(i).TextMatrix(lnFila, 1) = rs!Descripcion
    
    'comentado por gitu 23-10-2009
    If bLimBille Then
        '*** PEAC 20090924
        fgMonedas(i).TextMatrix(lnFila, 2) = Format(0, "#,##0")
        fgMonedas(i).TextMatrix(lnFila, 3) = Format(0, "#,##0.00")
        '*** FIN PEAC
    Else
        fgMonedas(i).TextMatrix(lnFila, 2) = Format(rs!Cantidad, "#,##0")
        fgMonedas(i).TextMatrix(lnFila, 3) = Format(rs!Monto, "#,##0.00")
    End If
    
    fgMonedas(i).TextMatrix(lnFila, 4) = rs!cEfectivoCod
    fgMonedas(i).TextMatrix(lnFila, 5) = rs!nEfectivoValor
    rs.MoveNext
Loop

rs.Close
Set rs = Nothing
Set oContFunct = Nothing
Set oEfec = Nothing
fgMonedas(i).Col = 2
lblTotalBilletes(i) = Format(fgBilletes(i).SumaRow(3), "#,##0.00")
lblTotMoneda(i) = Format(fgMonedas(i).SumaRow(3), "#,##0.00")
lblTotal(i) = Format(CDbl(lblTotalBilletes(i)) + CDbl(lblTotMoneda(i)), "#,##0.00")
End Sub

Public Function Residuo(Dividendo As Currency, Divisor As Currency) As Boolean
Dim X As Currency
X = Round(Dividendo / Divisor, 0)
Residuo = True
X = X * Divisor
If X <> Dividendo Then
   Residuo = False
End If
End Function

Public Property Get vnDiferencia() As Double
vnDiferencia = lnDiferencia
End Property

Public Property Let vnDiferencia(ByVal vNewValue As Double)
lnDiferencia = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
Set oCajero = Nothing
End Sub

Private Sub txtBuscarUser_EmiteDatos()
lblDescUser = txtBuscarUser.psDescripcion
End Sub

Public Property Get MovNro() As String
MovNro = lsMovNro
End Property

Public Property Get nMovNro() As Long
nMovNro = lnMovNro
End Property

Public Property Get FechaMov() As Date
FechaMov = CDate(txtfecha)
End Property

Public Property Get Total() As Currency
Total = 0
End Property



