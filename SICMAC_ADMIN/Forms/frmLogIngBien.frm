VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogIngBien 
   ClientHeight    =   6840
   ClientLeft      =   990
   ClientTop       =   2205
   ClientWidth     =   10470
   Icon            =   "frmLogIngBien.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10470
   Visible         =   0   'False
   Begin VB.TextBox txtDescServ 
      Height          =   315
      Left            =   3780
      TabIndex        =   68
      Top             =   6300
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   7620
      TabIndex        =   21
      Top             =   6330
      Width           =   1125
   End
   Begin VB.PictureBox PicOk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5085
      Picture         =   "frmLogIngBien.frx":030A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   59
      Top             =   6390
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Frame FrameTipCambio 
      Caption         =   "Tipo de Cambio"
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
      Height          =   615
      Left            =   120
      TabIndex        =   51
      Top             =   6150
      Visible         =   0   'False
      Width           =   3345
      Begin VB.TextBox txtTipCambio 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   480
         TabIndex        =   53
         Top             =   210
         Width           =   960
      End
      Begin VB.TextBox txtTipCompra 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   2220
         TabIndex        =   52
         Top             =   210
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Fijo"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   270
         Width           =   435
      End
      Begin VB.Label Label5 
         Caption         =   "Compra"
         Height          =   255
         Left            =   1590
         TabIndex        =   54
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.TextBox txtProvCod 
      Height          =   315
      Left            =   3930
      MaxLength       =   20
      TabIndex        =   50
      Top             =   6360
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Caption         =   "Observaciones"
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
      Height          =   885
      Left            =   120
      TabIndex        =   47
      Top             =   1530
      Width           =   5655
      Begin VB.TextBox txtMovDesc 
         Height          =   585
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   210
         Width           =   5340
      End
   End
   Begin VB.Frame frameDestino 
      Caption         =   "Proveedor "
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
      Height          =   750
      Left            =   120
      TabIndex        =   24
      Top             =   750
      Width           =   5655
      Begin VB.CommandButton cmdExaCab 
         Caption         =   "..."
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   300
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txtExaCab 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1200
         TabIndex        =   36
         Top             =   270
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox txtProvRuc 
         Height          =   315
         Left            =   150
         MaxLength       =   20
         TabIndex        =   1
         Top             =   270
         Width           =   1335
      End
      Begin VB.TextBox txtProvNom 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   25
         Top             =   270
         Width           =   3945
      End
   End
   Begin TabDlg.SSTab TabDoc 
      CausesValidation=   0   'False
      Height          =   3645
      Left            =   120
      TabIndex        =   32
      Top             =   2490
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   6429
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   564
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Detalle de Documento      "
      TabPicture(0)   =   "frmLogIngBien.frx":064C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ShapeIGV"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblIGV"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fgDetalle"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCant"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtObj"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdExaminar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdAgregar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdEliminar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtIGV"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtTot"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdValVenta"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Asiento C&ontable         "
      TabPicture(1)   =   "frmLogIngBien.frx":0668
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtMonto"
      Tab(1).Control(1)=   "cmdServicio"
      Tab(1).Control(2)=   "txtDebe"
      Tab(1).Control(3)=   "txtHaber"
      Tab(1).Control(4)=   "fgAsiento"
      Tab(1).Control(5)=   "FrameServicio"
      Tab(1).Control(6)=   "lblTotal(0)"
      Tab(1).ControlCount=   7
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   295
         Left            =   -68700
         TabIndex        =   19
         Top             =   1350
         Visible         =   0   'False
         Width           =   1510
      End
      Begin VB.CommandButton cmdServicio 
         Caption         =   "Ser&vicio"
         Height          =   360
         Left            =   -74760
         TabIndex        =   61
         Top             =   3090
         Width           =   1095
      End
      Begin VB.CommandButton cmdValVenta 
         Caption         =   "Valor Venta"
         Height          =   360
         Left            =   2550
         TabIndex        =   60
         ToolTipText     =   "Calcula el Valor Venta de Subtotales"
         Top             =   3000
         Width           =   1080
      End
      Begin VB.TextBox txtTot 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   255
         Left            =   8580
         TabIndex        =   57
         Top             =   2940
         Width           =   1185
      End
      Begin VB.TextBox txtIGV 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   255
         Left            =   8580
         TabIndex        =   18
         Top             =   3240
         Width           =   1185
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   1380
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3000
         Width           =   1080
      End
      Begin VB.TextBox txtDebe 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   -68250
         TabIndex        =   34
         Top             =   3000
         Width           =   1510
      End
      Begin VB.TextBox txtHaber 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   315
         Left            =   -66750
         TabIndex        =   33
         Top             =   3000
         Width           =   1510
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAsiento 
         Height          =   2535
         Left            =   -74820
         TabIndex        =   20
         Top             =   450
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4471
         _Version        =   393216
         Cols            =   6
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483638
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "A&gregar"
         Height          =   360
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3000
         Width           =   1110
      End
      Begin VB.CommandButton cmdExaminar 
         Caption         =   "..."
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3090
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1020
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.TextBox txtObj 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   295
         Left            =   1800
         TabIndex        =   12
         Top             =   1020
         Visible         =   0   'False
         Width           =   1510
      End
      Begin VB.TextBox txtCant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   295
         Left            =   8700
         TabIndex        =   14
         Top             =   705
         Visible         =   0   'False
         Width           =   1110
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDetalle 
         Height          =   2475
         Left            =   180
         TabIndex        =   15
         Top             =   450
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4366
         _Version        =   393216
         Cols            =   11
         ForeColorSel    =   -2147483643
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483637
         AllowBigSelection=   0   'False
         Enabled         =   0   'False
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
         RowSizingMode   =   1
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
         _NumberOfBands  =   1
         _Band(0).Cols   =   11
      End
      Begin VB.Frame FrameServicio 
         Height          =   585
         Left            =   -74820
         TabIndex        =   62
         Top             =   2940
         Visible         =   0   'False
         Width           =   5475
         Begin VB.CommandButton cmdSeekServ 
            Caption         =   "..."
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2970
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   210
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.TextBox txtImporte 
            Height          =   315
            Left            =   4020
            TabIndex        =   65
            Top             =   180
            Width           =   1365
         End
         Begin VB.TextBox txtCodServ 
            Height          =   315
            Left            =   1890
            TabIndex        =   64
            Top             =   180
            Width           =   1365
         End
         Begin VB.Label Label15 
            Caption         =   "Importe"
            Height          =   195
            Left            =   3390
            TabIndex        =   66
            Top             =   270
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "Código"
            Height          =   195
            Left            =   1320
            TabIndex        =   63
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Sub-Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7560
         TabIndex        =   58
         Top             =   2970
         Width           =   1035
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   7350
         Top             =   2910
         Width           =   2445
      End
      Begin VB.Label lblIGV 
         BackColor       =   &H00E0E0E0&
         Caption         =   "I.G.V."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7800
         TabIndex        =   56
         Top             =   3270
         Width           =   615
      End
      Begin VB.Label lblTotal 
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   -69030
         TabIndex        =   35
         Top             =   3060
         Width           =   615
      End
      Begin VB.Shape ShapeIGV 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   7350
         Top             =   3210
         Width           =   2445
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Movimiento"
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
      Height          =   690
      Left            =   120
      TabIndex        =   27
      Top             =   30
      Width           =   5655
      Begin VB.TextBox txtOpeCod 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   750
         TabIndex        =   30
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox txtMovNro 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   1620
         TabIndex        =   28
         Top             =   240
         Width           =   2160
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   4380
         TabIndex        =   0
         Top             =   240
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label8 
         Caption         =   "Número"
         Height          =   255
         Left            =   150
         TabIndex        =   31
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   3870
         TabIndex        =   29
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8820
      TabIndex        =   22
      Top             =   6330
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtxtAsiento 
      Height          =   315
      Left            =   240
      TabIndex        =   26
      Top             =   6240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmLogIngBien.frx":0684
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdDocumento 
      Caption         =   "&Imprimir"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7620
      TabIndex        =   23
      Top             =   6330
      Visible         =   0   'False
      Width           =   1125
   End
   Begin TabDlg.SSTab TabDocRef 
      Height          =   1695
      Left            =   5850
      TabIndex        =   4
      Top             =   750
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   2990
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Orden de Com&pra"
      TabPicture(0)   =   "frmLogIngBien.frx":06FD
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(3)=   "Shape1"
      Tab(0).Control(4)=   "Shape4"
      Tab(0).Control(5)=   "txtOCNro"
      Tab(0).Control(6)=   "txtOCFecha"
      Tab(0).Control(7)=   "txtOCPlazo"
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Guía de &Remisión"
      TabPicture(1)   =   "frmLogIngBien.frx":0719
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(2)=   "Shape2"
      Tab(1).Control(3)=   "Shape3"
      Tab(1).Control(4)=   "txtGRFecha"
      Tab(1).Control(5)=   "txtGRSerie"
      Tab(1).Control(6)=   "txtGRNro"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "&Factura            "
      TabPicture(2)   =   "frmLogIngBien.frx":0735
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label9"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Shape5"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Shape6"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtFacSerie"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "txtFacNro"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cboFacDestino"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txtFacFecha"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      Begin VB.TextBox txtGRNro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74160
         MaxLength       =   8
         TabIndex        =   6
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtGRSerie 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74520
         MaxLength       =   3
         TabIndex        =   5
         Top             =   630
         Width           =   375
      End
      Begin MSMask.MaskEdBox txtFacFecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   3150
         TabIndex        =   10
         Top             =   630
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtOCPlazo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73410
         TabIndex        =   44
         Top             =   1110
         Width           =   2655
      End
      Begin VB.ComboBox cboFacDestino 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmLogIngBien.frx":0751
         Left            =   810
         List            =   "frmLogIngBien.frx":0761
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1110
         Width           =   3435
      End
      Begin VB.TextBox txtFacNro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   840
         MaxLength       =   8
         TabIndex        =   9
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtFacSerie 
         Enabled         =   0   'False
         Height          =   315
         Left            =   480
         MaxLength       =   3
         TabIndex        =   8
         Top             =   630
         Width           =   375
      End
      Begin VB.TextBox txtOCFecha 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71850
         TabIndex        =   38
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtOCNro 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74520
         MaxLength       =   8
         TabIndex        =   37
         Top             =   630
         Width           =   1155
      End
      Begin MSMask.MaskEdBox txtGRFecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   -71850
         TabIndex        =   7
         Top             =   630
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H8000000E&
         Height          =   1095
         Left            =   135
         Top             =   495
         Width           =   4245
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H8000000C&
         Height          =   1095
         Left            =   120
         Top             =   480
         Width           =   4245
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H8000000E&
         Height          =   1095
         Left            =   -74865
         Top             =   495
         Width           =   4245
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H8000000C&
         Height          =   1095
         Left            =   -74880
         Top             =   480
         Width           =   4245
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H8000000E&
         Height          =   1095
         Left            =   -74865
         Top             =   495
         Width           =   4245
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H8000000C&
         Height          =   1095
         Left            =   -74880
         Top             =   480
         Width           =   4245
      End
      Begin VB.Label Label11 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74790
         TabIndex        =   49
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   -72960
         TabIndex        =   48
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label14 
         Caption         =   "Plazo de Entrega"
         Height          =   240
         Left            =   -74790
         TabIndex        =   45
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   -72960
         TabIndex        =   43
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Nº"
         Height          =   165
         Left            =   -74790
         TabIndex        =   42
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Emisión"
         Height          =   165
         Left            =   2040
         TabIndex        =   41
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Nº"
         Height          =   165
         Left            =   210
         TabIndex        =   40
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Destino"
         Height          =   195
         Left            =   210
         TabIndex        =   39
         Top             =   1200
         Width           =   555
      End
   End
   Begin VB.Label lblDoc 
      Alignment       =   1  'Right Justify
      Caption         =   "NOTA DE INGRESO Nº 00000000"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5850
      TabIndex        =   46
      Top             =   240
      Width           =   4515
   End
   Begin VB.Menu mnuLog 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAtender 
         Caption         =   "&Atender"
      End
      Begin VB.Menu mnuNoAtender 
         Caption         =   "&No Atender"
      End
   End
   Begin VB.Menu mnuObj 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuAgregar 
         Caption         =   "&Agregar"
      End
      Begin VB.Menu mnuEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnuGravado 
         Caption         =   "&Gravado"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmLogIngBien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim sSQL As String, sSqlObj As String
'Dim lTransActiva As Boolean      ' Controla si la transaccion esta activa o no
'Dim rs As New ADODB.Recordset    'Rs temporal para lectura de datos
'Dim lMN As Boolean, sMoney As String  'Identifican el Simbolo del Tipo de Moneda
'Dim nUltimoTxt As String 'Ultimo TextBox seleccionado
'Dim lSalir As Boolean
'Dim lLlenaObj As Boolean, OK As Boolean
'Dim sObjCod As String, sObjDesc As String, sObjUnid
'Dim sCtaCod As String, sCtaDesc As String
'Dim sProvRuc As String, sDocOC As String, sDocGR As String
'Dim nTasaIGV As Currency, nVariaIGV As Currency
'
'Public Sub Inicio(LlenaObj As Boolean, Optional ProvRuc As String)
'lLlenaObj = LlenaObj
'sProvRuc = ProvRuc
'Me.Show 1
'End Sub
'
'Private Sub ControlesTabDocRef(lGR As Boolean, lFac As Boolean)
'txtGRSerie.Enabled = lGR
'txtGRNro.Enabled = lGR
'txtGRFecha.Enabled = lGR
'txtFacSerie.Enabled = lFac
'txtFacNro.Enabled = lFac
'txtFacFecha.Enabled = lFac
'cboFacDestino.Enabled = lFac
'End Sub
'
'Private Sub FormatoAsiento()
'fgAsiento.TextMatrix(0, 0) = "#"
'fgAsiento.TextMatrix(0, 1) = "Cta.Contable"
'fgAsiento.TextMatrix(0, 2) = "Descripción"
'fgAsiento.TextMatrix(0, 3) = "Debe"
'fgAsiento.TextMatrix(0, 4) = "Haber"
'fgAsiento.ColWidth(0) = 380
'fgAsiento.ColWidth(1) = 1500
'fgAsiento.ColWidth(2) = 4680
'fgAsiento.ColWidth(3) = 1500
'fgAsiento.ColWidth(4) = 1500
'fgAsiento.ColWidth(5) = 0
'
'fgAsiento.ColAlignment(0) = 4
'fgAsiento.ColAlignment(1) = 1
'fgAsiento.ColAlignment(3) = 7
'fgAsiento.ColAlignment(4) = 7
'fgAsiento.ColAlignmentFixed(0) = 4
'fgAsiento.ColAlignmentFixed(3) = 7
'fgAsiento.ColAlignmentFixed(4) = 7
'fgAsiento.Row = 1
'fgAsiento.Col = 1
'fgAsiento.RowHeight(-1) = 285
'End Sub
'
'Private Function FormaSelect(Optional sObj As String, Optional nNiv As Integer) As String
'Dim sText As String
'sText = "SELECT d.cObjetoCod, upper(d.cObjetoDesc) as cObjetoDesc, d.nObjetoNiv, c.nCtaObjNiv, b.cCtaContCod, b.cCtaContDesc, e.cUnidadAbrev " _
'      & "FROM  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b,  " & gcCentralCom & "CtaObj c,  " & gcCentralCom & "Objeto d LEFT JOIN Bienes e ON (e.cObjetoCod = d.cObjetoCod) " _
'      & "WHERE b.cCtaContCod = a.cCtaContCod AND c.cCtaContCod = b.cCtaContCod AND " _
'      & "      ((d.cObjetoCod Like c.cObjetoCod+'%') "
'sText = sText & "AND (a.cOpeCod='" & gcOpeCod & "') AND (a.cOpeCtaDH='D')) "
'If nNiv > 0 Then
'   sText = sText & "and d.nobjetoniv = " & nNiv & " "
'End If
'FormaSelect = sText & IIf(sObj <> "", "and d.cObjetoCod = '" & sObj & "' ", sObj) _
'            & "order by d.cObjetoCod"
'End Function
'Private Sub VerTxtObj()
'If txtObj.Visible Then
'   txtObj.Visible = False
'   cmdExaminar.Visible = False
'End If
'End Sub
'Private Function ValidarProvee(ProvRuc As String) As Boolean
'Dim sSqlProv As String
'Dim rsProv As New ADODB.Recordset
''ValidarProvee = False
''   If Len(Trim(ProvRuc)) = 0 Then
''      Exit Function
''   End If
''   sSqlProv = "select rtrim(a.cProvRuc) as cRuc, b.cnompers, a.cCodPers from  proveedor a, " & gcCentralPers & "Persona b where b.ccodpers = a.ccodpers and a.cProvRuc = '" & ProvRuc & "' order by cRuc"
''   Set rsProv = CargaRecord(sSqlProv)
''   If RSVacio(rsProv) Then
''      MsgBox " Proveedor no Encontrado ...! ", vbCritical, "Aviso"
''   Else
''      txtProvNom = rsProv(1)
''      txtProvCod = rsProv(2)
''      ValidarProvee = True
''   End If
''rsProv.Close
'End Function
'
'Private Sub cboFacDestino_KeyPress(KeyAscii As Integer)
'Dim N As Integer
'If KeyAscii = 13 Then
'   If cboFacDestino.ListIndex = -1 Then
'      MsgBox "Selección de destino no Valido...!", vbCritical, "Error"
'   Else
'      TabDoc.Tab = 0
'      If cboFacDestino.ListIndex <> 1 Then
'         fgDetalle.Col = 1
'         For N = 1 To fgDetalle.Rows - 1
'             If fgDetalle.TextMatrix(N, 10) = "X" Then
'                fgDetalle.TextMatrix(N, 10) = ""
'                fgDetalle.Row = N
'                Set fgDetalle.CellPicture = LoadPicture("")
'             End If
'         Next
'      End If
'      If fgDetalle.Enabled Then
'         fgDetalle.SetFocus
'      Else
'         cmdAgregar.SetFocus
'      End If
'   End If
'End If
'End Sub
'
'Private Sub cmdExaCab_Click()
''Dim sSqlProv As String
''Dim rsProv As New ADODB.Recordset
''   sSqlProv = "SELECT rtrim(a.cProvRuc) as cRuc, b.cNomPers, a.cCodPers " _
''           & "FROM  Proveedor a, " & gcCentralPers & "Persona b " _
''           & "WHERE b.cCodPers = a.cCodPers AND a.cProvEstado='1' order by a.cProvRuc"
''   frmBusqDiversa.Inicio 3, sSqlProv
''   If frmBusqDiversa.lOk Then
''      sSqlProv = "select * from  proveedor where cProvRuc = '" & frmBusqDiversa.pCod & "'"
''      Set rsProv = CargaRecord(sSqlProv)
''      txtProvRuc = Trim(rsProv!cProvRuc)
''      txtProvNom = frmBusqDiversa.pDesc
''      txtProvCod = rsProv!cCodPers
''      txtMovDesc.SetFocus
''      rsProv.Close
''   Else
''      txtProvRuc.SetFocus
''   End If
'End Sub
'
'Private Sub cmdExaCab_LostFocus()
'cmdExaCab.Visible = False
'txtExaCab.Visible = False
'End Sub
'
'Private Sub cmdSeekServ_Click()
''sSQL = "SELECT cCtaContCod as cObjetoCod, cCtaContDesc as cObjetoDesc " _
''     & ""
''frmBusqDiversa.Inicio 1, sSQL
'End Sub
'
'Private Sub cmdServicio_Click()
'FrameServicio.Visible = True
'txtCodServ.SetFocus
'End Sub
'
'Private Sub cmdValVenta_Click()
'Dim N As Integer
'For N = 1 To fgDetalle.Rows - 1
'    fgDetalle.TextMatrix(N, 7) = Format(Round(Val(Format(fgDetalle.TextMatrix(N, 7), gcFormDato)) / (1 + (nTasaIGV / 100)), 2), gcFormView)
'Next
'SumasDoc
''cmdValVenta.Enabled = False
'End Sub
'
'Private Sub fgAsiento_KeyPress(KeyAscii As Integer)
'If fgAsiento.Col = 3 Then
'   If KeyAscii = 13 Then
'      EnfocaTexto txtMonto, 0, fgAsiento
'   Else
'      If InStr("0123456789.", KeyAscii) > 0 Then
'         EnfocaTexto txtMonto, KeyAscii, fgAsiento
'      End If
'   End If
'End If
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
''If lTransActiva Then
''   dbCmact.RollbackTrans
''End If
''CierraConexion
'End Sub
'
'Private Sub mnuAgregar_Click()
'cmdAgregar_Click
'End Sub
'
'Private Sub mnuAtender_Click()
'Dim nPos As Integer, nSalto As Integer
'nSalto = IIf(fgDetalle.Row < fgDetalle.RowSel, 1, -1)
'For nPos = fgDetalle.Row To fgDetalle.RowSel Step nSalto
'    fgDetalle.TextMatrix(nPos, 5) = fgDetalle.TextMatrix(nPos, 4)
'Next
'SumasDoc
'End Sub
'
'Private Sub mnueliminar_Click()
'cmdEliminar_Click
'End Sub
'
'Private Sub mnuGravado_Click()
'If mnuGravado.Checked Then
'  Set fgDetalle.CellPicture = LoadPicture("")
'  fgDetalle.TextMatrix(fgDetalle.Row, 10) = ""
'Else
'  Set fgDetalle.CellPicture = PicOk.Picture
'  fgDetalle.CellPictureAlignment = 7
'  fgDetalle.TextMatrix(fgDetalle.Row, 10) = "X"
'End If
'End Sub
'
'Private Sub mnuNoAtender_Click()
'Dim nPos As Integer, nSalto As Integer
'nSalto = IIf(fgDetalle.Row < fgDetalle.RowSel, 1, -1)
'For nPos = fgDetalle.Row To fgDetalle.RowSel Step nSalto
'    fgDetalle.TextMatrix(nPos, 5) = ""
'Next
'SumasDoc
'End Sub
'
'Private Sub ActualizaFG(nItem As Integer)
''fgDetalle.TextMatrix(nItem, 1) = sObjCod
''fgDetalle.TextMatrix(nItem, 2) = sObjDesc
''fgDetalle.TextMatrix(nItem, 3) = IIf(IsNull(sObjUnid), "UND", sObjUnid)
''fgDetalle.Col = 5
''
''If nItem > fgAsiento.Rows - 1 Then
''   AdicionaRow fgAsiento
''   nItem = fgAsiento.Row
''End If
''fgAsiento.TextMatrix(nItem, 0) = nItem
''fgAsiento.TextMatrix(nItem, 1) = sCtaCod
''fgAsiento.TextMatrix(nItem, 2) = sCtaDesc
''fgAsiento.TextMatrix(nItem, 5) = "D"
''If Not GetAsignaProvision(sCtaCod, nItem) Then
''   lSalir = True
''   Exit Sub
''End If
'''Calculo el Costo del Bien que sale
''GetSaldoCtaObj GetCtaObjFiltro(sCtaCod, sObjCod), sObjCod, gdFecSis
''fgDetalle.TextMatrix(nItem, 6) = Format(gnCant, gcFormView)
''If gnCant > 0 Then
''   fgDetalle.TextMatrix(nItem, 7) = Format(Round(gnSaldo / gnCant, 2), gcFormView)
''End If
'End Sub
'
'Private Sub SumaAsiento()
'Dim nDebe As Currency, nHaber As Currency
'Dim I As Integer
'For I = 1 To fgAsiento.Rows - 1
'    nDebe = nDebe + Val(Format(fgAsiento.TextMatrix(I, 3), gcFormDato))
'    nHaber = nHaber + Val(Format(fgAsiento.TextMatrix(I, 4), gcFormDato))
'Next
'txtDebe = Format(nDebe, gcFormView)
'txtHaber = Format(nHaber, gcFormView)
'End Sub
'
'Private Sub ActivaExaminar(txt As TextBox)
'txt.Width = txt.Width - cmdExaminar.Width + 10
'cmdExaminar.Top = txt.Top + 15
'cmdExaminar.Left = txt.Left + txt.Width - 10
'cmdExaminar.Visible = True
'cmdExaminar.TabIndex = txt.TabIndex + 1
'End Sub
'Private Sub DesactivaExaminar(txt As TextBox)
'txt.Width = txt.Width + cmdExaminar.Width - 10
'End Sub
'
'Private Sub cmdAgregar_Click()
'If Not fgDetalle.Enabled Then
'   AdicionaRow fgDetalle
'   EnfocaTexto txtObj, 0, fgDetalle
'Else
'   If Val(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 5), gcFormDato)) > 0 And _
'      Val(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 1), gcFormDato)) > 0 And _
'      Val(Format(fgDetalle.TextMatrix(fgDetalle.Rows - 1, 7), gcFormDato)) > 0 Then
'      AdicionaRow fgDetalle
'      EnfocaTexto txtObj, 0, fgDetalle
'   Else
'      fgDetalle.SetFocus
'   End If
'End If
'End Sub
'
'Private Sub cmdCerrar_Click()
''If lTransActiva Then
''   dbCmact.RollbackTrans
''End If
''lTransActiva = False
''Unload Me
'End Sub
'
'Private Sub cmdDocumento_Click()
'Dim N As Integer
'Dim sTexto As String
'  sTexto = "N O T A   D E   I N G R E S O    Nro." & gcDocNro
'  gdFecha = txtFecha
'  rtxtAsiento.Text = ImpreCabAsiento(sTexto, False)
'  rtxtAsiento.Text = rtxtAsiento.Text & "Proveedor     : " & BON & txtProvNom & BOFF & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea
'  sTexto = ""
'  If txtOCNro <> "" Then
'     sTexto = "Orden de Compra No " & txtOCNro & "    "
'  End If
'  If txtGRNro <> "" Then
'     sTexto = sTexto & "Guía de Remisión No " & txtGRSerie & "-" & txtGRNro & IIf(sTexto = "", "    ", oImpresora.gPrnSaltoLinea  & Space(16))
'  End If
'  If txtFacNro <> "" Then
'     sTexto = sTexto & "Factura No " & txtFacSerie & "-" & txtFacNro
'  End If
'  If Trim(sTexto) <> "" Then
'     rtxtAsiento.Text = rtxtAsiento.Text & "Referencias   : " & BON & sTexto & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & BOFF
'  End If
'  gsGlosa = txtMovDesc
'  rtxtAsiento.Text = rtxtAsiento.Text & ImpreGlosa("Observaciones : ")
'
'sTexto = CON
'For N = 1 To fgDetalle.Rows - 1
'    sTexto = sTexto & Left(fgDetalle.TextMatrix(N, 1) + Space(18), 18)
'    sTexto = sTexto & "  " & Left(fgDetalle.TextMatrix(N, 2) + Space(80), 80)
'    sTexto = sTexto & "   " & Right(Space(12) & fgDetalle.TextMatrix(N, 4), 12)
'    sTexto = sTexto & "  " & Right(Space(12) & fgDetalle.TextMatrix(N, 5), 12) & oImpresora.gPrnSaltoLinea
'Next
'sTexto = sTexto & COFF
'rtxtAsiento.Text = rtxtAsiento.Text & ImpreDetLog(sTexto, "Solicitada   Atendida")
'rtxtAsiento.Text = rtxtAsiento.Text & oImpresora.gPrnSaltoLinea
'rtxtAsiento.Text = rtxtAsiento.Text & BON
'rtxtAsiento.Text = rtxtAsiento.Text & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea
'rtxtAsiento.Text = rtxtAsiento.Text & "      _____________________                        _____________________" & oImpresora.gPrnSaltoLinea
'rtxtAsiento.Text = rtxtAsiento.Text & "           RECIBIDO POR                               Vo Bo LOGISTICA   " & BOFF & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoPagina
'sTexto = ImprimeAsientoContable(gcMovNro)
'rtxtAsiento.Text = rtxtAsiento.Text & sTexto
'rtxtAsiento.Text = ImpreCarEsp(rtxtAsiento.Text)
'frmPrevio.Previo rtxtAsiento, "Documento: Nota de Ingreso", False, gnLinPage
'End Sub
'
'Private Function FormaAsiento(lMontoDig As Boolean) As String
'Dim N As Integer
'Dim nFactor As Currency
'nFactor = IIf(lMontoDig, 1, gnTipCambio)
'For N = 1 To fgAsiento.Rows - 1
'    If fgAsiento.TextMatrix(N, 5) = "D" Then
'       FormaAsiento = FormaAsiento & Left(fgAsiento.TextMatrix(N, 1) & String(18, " "), 18) & " " & Left(fgAsiento.TextMatrix(N, 2) & String(70, " "), 70) & " " & GSSIMBOLO & Right(String(15, " ") & Format(Round(Val(Format(fgAsiento.TextMatrix(N, 3), gcFormDato)) * nFactor, 2), gcFormView), 15) & oImpresora.gPrnSaltoLinea
'    Else
'       FormaAsiento = FormaAsiento & Left(fgAsiento.TextMatrix(N, 1) & String(18, " "), 18) & " " & Left(fgAsiento.TextMatrix(N, 2) & String(70, " "), 70) & Space(22) & GSSIMBOLO & Right(String(15, " ") & Format(Round(Val(Format(fgAsiento.TextMatrix(N, 4), gcFormDato)) * nFactor, 2), gcFormView), 15) & oImpresora.gPrnSaltoLinea
'    End If
'Next
'End Function
'Private Sub cmdEliminar_Click()
'Dim nPos As Integer
'nPos = fgDetalle.Row
'EliminaRow fgDetalle, nPos
'EnumeraItems fgDetalle
'EliminaRow fgAsiento, nPos
'EnumeraItems fgAsiento
'End Sub
'
'Private Sub cmdExaminar_Click()
'Dim sSqlO As String
'frmBuscaBien.Inicio 2, 6, "181"
'AbreConexion
'If frmBuscaBien.lOk Then
'   sSqlO = FormaSelect(gaObj(0, 0, UBound(gaObj, 3)))
'   Set rs = CargaRecord(sSqlO)
'   If RSVacio(rs) Then
'      MsgBox "Objeto no asignado a Operación", vbCritical, "Error"
'      txtObj.SetFocus
'      Exit Sub
'   End If
'   sObjCod = rs!cObjetoCod
'   sObjDesc = rs!cObjetoDesc
'   sObjUnid = rs!cUnidadAbrev
'   sCtaCod = rs!cCtaContCod
'   sCtaDesc = rs!cCtaContDesc
'
'   ActualizaFG fgDetalle.Row
'   fgDetalle.Enabled = True
'   txtObj.Visible = False
'   cmdExaminar.Visible = False
'   fgDetalle.SetFocus
'   rs.Close
'Else
'   If txtObj.Enabled Then
'      txtObj.SetFocus
'   End If
'End If
'End Sub
'
'Private Sub fgDetalle_DblClick()
'If fgDetalle.Col = 1 Then
'   EnfocaTexto txtObj, 0, fgDetalle
'End If
'If fgDetalle.Col = 4 Then
'   mnuAtender_Click
'End If
'If fgDetalle.Col = 5 Or fgDetalle.Col = 7 Then
'   EnfocaTexto txtCant, 0, fgDetalle
'End If
'End Sub
'
'Private Sub fgDetalle_GotFocus()
'VerTxtObj
'End Sub
'
'Private Sub fgDetalle_KeyPress(KeyAscii As Integer)
'If fgDetalle.Col = 1 Then
'   If KeyAscii = 13 Then
'      EnfocaTexto txtObj, IIf(KeyAscii = 13, 0, KeyAscii), fgDetalle
'   End If
'   If KeyAscii = 32 Then
'      If cboFacDestino.ListIndex = 1 Then
'         If fgDetalle.TextMatrix(fgDetalle.Row, 10) = "X" Then
'            mnuGravado.Checked = True
'         Else
'            mnuGravado.Checked = False
'         End If
'         mnuGravado_Click
'      End If
'   End If
'End If
'If fgDetalle.Col = 4 Then
'   mnuAtender_Click
'End If
'If fgDetalle.Col = 5 Or fgDetalle.Col = 7 Then
'   If InStr("0123456789.", Chr(KeyAscii)) > 0 Then
'      EnfocaTexto txtCant, KeyAscii, fgDetalle
'   Else
'      If KeyAscii = 13 Then EnfocaTexto txtCant, 0, fgDetalle
'   End If
'End If
'End Sub
'Private Sub fgDetalle_KeyUp(KeyCode As Integer, Shift As Integer)
'If fgDetalle.Col = 5 Then
'   KeyUp_Flex fgDetalle, KeyCode, Shift
'   SumasDoc
'End If
'End Sub
'
'Private Sub fgDetalle_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
'If Button = 2 Then
'   Select Case fgDetalle.Col
'          Case 1:
'                  If cboFacDestino.ListIndex = 1 Then
'                     mnuGravado.Visible = True
'                     If fgDetalle.TextMatrix(fgDetalle.Row, 10) = "X" Then
'                        mnuGravado.Checked = True
'                     Else
'                        mnuGravado.Checked = False
'                     End If
'                  Else
'                     mnuGravado.Visible = False
'                  End If
'                  PopupMenu mnuObj
'          Case 4: mnuNoAtender.Enabled = False
'                  mnuAtender.Enabled = True
'                  PopupMenu mnuLog
'          Case 5: mnuAtender.Enabled = False
'                  mnuNoAtender.Enabled = True
'                  PopupMenu mnuLog
'   End Select
'End If
'End Sub
'
'Private Sub Form_Activate()
'If lSalir Then
'   Unload Me
'End If
'If Not lLlenaObj Then
'   TabDocRef.TabEnabled(0) = False
'End If
'End Sub
'
'Private Sub Form_Load()
'Dim N As Integer, nSaldo As Currency, nCant As Currency
'Dim sCtaCod As String, nItem As Integer
'CentraSdi Me
'lSalir = False
'AbreConexion
'If Mid(gcOpeCod, 3, 1) = "2" Then  'Identificación de Tipo de Moneda
'   lMN = False
'   GSSIMBOLO = gcME
'   If gnTipCambio = 0 Then
'      If Not GetTipCambio(gdFecSis) Then
'         lSalir = True
'         Exit Sub
'      End If
'   End If
'   FrameTipCambio.Visible = True
'   txtTipCambio = Format(gnTipCambio, gcFormView)
'   txtTipCompra = Format(gnTipCambioC, "##,###,##0.000")
'Else
'   lMN = True
'   GSSIMBOLO = gcMN
'End If
'
'' Defino el Nro de Movimiento
'txtOpeCod = gcOpeCod
'txtMovNro = GeneraMovNro
'txtFecha = Format(gdFecSis, "dd/mm/yyyy")
'
'' Identifico al Usuario
'sSQL = "Select cCodUsu, cNomUsu, cCodAge, cCodPers from " & gcCentralCom & "Usuario where cCodUsu = '" & gsCodUser & "'"
'Set rs = CargaRecord(sSQL)
'If RSVacio(rs) Then
'   MsgBox "Usuario no registrado en el Sistema. Consultar con Sistemas!", vbCritical, "Error de Acceso"
'   lSalir = True
'   Exit Sub
'End If
'If lLlenaObj Then
'   txtProvCod = Mid(gsPerCod, 3, 10)
'   sSQL = "select cProvRuc from  Proveedor where cCodPers = '" & txtProvCod & "'"
'   Set rs = CargaRecord(sSQL)
'   If RSVacio(rs) Then
'      MsgBox "Proveedor de O/C no registrado. Por favor verificar", vbCritical, "Error"
'      lSalir = True
'      Exit Sub
'   End If
'   txtProvRuc = Trim(rs!cProvRuc)
'   txtProvNom = gsPerDes
'   txtProvRuc.Enabled = False
'   sDocOC = gaDoc(0, 0)
'   txtOCNro = gaDoc(0, 1)
'   txtOCFecha = gaDoc(0, 2)
'   For N = 0 To UBound(gaObj, 3)
'       sSQL = FormaSelect(gaObj(0, 0, N), 6)
'       Set rs = CargaRecord(sSQL)
'       If Not RSVacio(rs) Then
'          AdicionaRow fgDetalle
'          nItem = fgDetalle.Row
'          fgDetalle.TextMatrix(nItem, 0) = nItem
'          fgDetalle.TextMatrix(nItem, 1) = gaObj(0, 0, N)
'          fgDetalle.TextMatrix(nItem, 2) = gaObj(0, 1, N)
'          fgDetalle.TextMatrix(nItem, 4) = Format(gaObj(0, 2, N), gcFormView)
'          fgDetalle.TextMatrix(nItem, 7) = Format(Round(gaObj(0, 3, N) / gaObj(0, 2, N), 2), gcFormView)
'          AdicionaRow fgAsiento
'          sCtaCod = rs!cCtaContCod
'          fgAsiento.TextMatrix(nItem, 0) = nItem
'          fgAsiento.TextMatrix(nItem, 1) = sCtaCod
'          fgAsiento.TextMatrix(nItem, 2) = rs!cCtaContDesc
'             ' Cuenta Provision
'          fgAsiento.TextMatrix(nItem, 5) = "D"
'          If Not GetAsignaProvision(sCtaCod, nItem) Then
'             lSalir = True
'             Exit Sub
'          End If
'       End If
'   Next
'   cmdAgregar.Enabled = False
'   cmdEliminar.Enabled = False
'   fgDetalle.Enabled = True
'End If
'
'Me.Caption = "Operación: " & gcOpeDesc
'fgDetalle.TextMatrix(0, 0) = "#"
'fgDetalle.TextMatrix(0, 1) = "Objeto"
'fgDetalle.TextMatrix(0, 2) = "Descripción"
'fgDetalle.TextMatrix(0, 3) = "Unid."
'fgDetalle.TextMatrix(0, 4) = "Solicitado"
'fgDetalle.TextMatrix(0, 5) = "Atendido"
'fgDetalle.TextMatrix(0, 6) = "Saldo"
'fgDetalle.TextMatrix(0, 7) = "Sub Total"
'fgDetalle.ColWidth(0) = 380
'fgDetalle.ColWidth(1) = 1500
'fgDetalle.ColWidth(2) = 3700
'fgDetalle.ColWidth(3) = 700
'fgDetalle.ColWidth(4) = 1100
'fgDetalle.ColWidth(5) = 1100
'fgDetalle.ColWidth(6) = 0
'fgDetalle.ColWidth(7) = 1100
'fgDetalle.ColWidth(8) = 0
'fgDetalle.ColWidth(9) = 0
'fgDetalle.ColWidth(10) = 0
'
'fgAsiento.ColAlignment(0) = 4
'fgDetalle.ColAlignment(1) = 1
'fgDetalle.ColAlignmentFixed(0) = 4
'fgDetalle.ColAlignmentFixed(4) = 7
'fgDetalle.ColAlignmentFixed(5) = 7
'fgDetalle.ColAlignmentFixed(6) = 7
'fgDetalle.ColAlignmentFixed(7) = 7
'fgDetalle.Row = 1
'fgDetalle.Col = 1
'fgDetalle.RowHeight(-1) = 285
'
'sSQL = "select * from  " & gcCentralCom & "OpeDoc where cOpeCod = '" & gcOpeCod & "' and cOpeDocMetodo = '2' "
'Set rs = CargaRecord(sSQL)
'If RSVacio(rs) Then
'   MsgBox "No se definió tipo de Documento. Por favor verificar...!", vbCritical, "Error"
'   lSalir = True
'   Exit Sub
'Else
'   gcDocTpo = rs!cDocTpo
'   gcDocNro = GeneraDocNro(gcDocTpo)
'   lblDoc.Caption = "NOTA DE INGRESO Nº " & gcDocNro
'End If
'sSQL = "select * from  " & gcCentralCom & "OpeDoc where cOpeCod = '" & gcOpeCod & "' and cOpeDocMetodo = '3' "
'Set rs = CargaRecord(sSQL)
'If RSVacio(rs) Then
'   MsgBox "No se definió Tipo de Documento emitido por Proveedor. " & oImpresora.gPrnSaltoLinea _
'        & "Por favor verificar...!", vbCritical, "Error"
'   lSalir = True
'   Exit Sub
'End If
'sSQL = ""
'sDocGR = ""
'Do While Not rs.EOF
'   If rs!cDocTpo = gcDocTpoFac Then
'      sSQL = "SELECT b.nImpTasa FROM  " & gcCentralCom & "DocImpuesto a,  " & gcCentralCom & "Impuesto b " _
'           & "WHERE  a.cDocTpo = '" & rs!cDocTpo & "' and a.cCtaContCod = '" & gcCtaIGV & "' and " _
'           & "       a.cCtaContCod = b.cCtaContCod  "
'   Else
'      sDocGR = rs!cDocTpo
'   End If
'   rs.MoveNext
'Loop
'If sDocGR = "" Then
'   MsgBox "Falta asignar Documento Guia de Remisión a Operación", vbCritical, "Error"
'   lSalir = True
'   Exit Sub
'End If
'If sSQL = "" Then
'   MsgBox "Falta asignar Documento FACTURA a Operación", vbCritical, "Error"
'   lSalir = True
'   Exit Sub
'End If
'Set rs = CargaRecord(sSQL)
'If RSVacio(rs) Then
'   MsgBox "Falta relacionar FACTURA con Impuesto IGV", vbCritical, "Error"
'   lSalir = True
'   Exit Sub
'End If
'nTasaIGV = rs!nImpTasa
'FormatoAsiento
'rs.Close
'cboFacDestino.ListIndex = 2
'End Sub
'
'Private Sub TabDoc_Click(PreviousTab As Integer)
'Select Case TabDoc.Tab
'       Case 0
'              If Len(fgDetalle.TextMatrix(1, 0)) > 0 Then fgDetalle.Enabled = True
'              fgAsiento.Enabled = False
'              If fgDetalle.Enabled Then
'                 fgDetalle.SetFocus
'                 EliminaAsiento
'              End If
'              cmdAceptar.Enabled = False
'              cmdServicio.Enabled = False
'       Case 1
'              If Len(fgAsiento.TextMatrix(1, 0)) > 0 Then fgAsiento.Enabled = True
'              fgDetalle.Enabled = False
'              If fgAsiento.Enabled Then
'                 RefrescaAsiento
'                 AsignaImpuesto
'                 SumaAsiento
'                 fgAsiento.SetFocus
'              End If
'              cmdAceptar.Enabled = True
'              cmdServicio.Enabled = True
'End Select
'End Sub
'
'Private Sub TabDocRef_Click(PreviousTab As Integer)
'Select Case TabDocRef.Tab
'       Case 0
'            If txtOCNro.Enabled Then
'               txtOCNro.SetFocus
'            End If
'            ControlesTabDocRef False, False
'       Case 1
'            ControlesTabDocRef True, False
'            txtGRSerie.SetFocus
'       Case 2
'            ControlesTabDocRef False, True
'            txtFacSerie.SetFocus
'End Select
'End Sub
'
'Private Sub txtCodServ_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   KeyAscii = 0
'   sSQL = "SELECT * FROM " & gcCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & txtCodServ & "%'"
'   Set rs = CargaRecord(sSQL)
'   If rs.EOF Then
'      MsgBox "Cuenta no Asignada a Operación", vbInformation, "Error"
'      txtCodServ.SelStart = 0
'      txtCodServ.SelLength = Len(txtCodServ)
'      rs.Close: Set rs = Nothing
'      Exit Sub
'   End If
'   If rs.RecordCount > 1 Then
'      MsgBox "Cuenta no es de Asiento...!", vbInformation, "Aviso"
'      txtCodServ.SelStart = 0
'      txtCodServ.SelLength = Len(txtCodServ)
'      rs.Close: Set rs = Nothing
'      Exit Sub
'   End If
'   txtDescServ = rs!cCtaContDesc
'   txtImporte.SetFocus
'   rs.Close: Set rs = Nothing
'End If
'End Sub
'
'Private Sub txtFacFecha_GotFocus()
'      txtFacFecha.SelStart = 0
'      txtFacFecha.SelLength = Len(txtFacFecha)
'End Sub
'
'Private Sub txtFacFecha_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   'If Not IsDate(txtFacFecha.Text) Then
'   If ValidaFecha(txtFacFecha) <> "" Then
'      MsgBox " Fecha no Válida... ", vbCritical, "Error"
'      txtFacFecha.SelStart = 0
'      txtFacFecha.SelLength = Len(txtFacFecha)
'   Else
'      GetTipCambio CDate(txtFacFecha)
'      txtTipCompra = Format(gnTipCambioC, "###,###,##0.000")
'      gnTipCambio = Val(Format(txtTipCambio, gcFormDato))
'      cboFacDestino.SetFocus
'   End If
'End If
'End Sub
'
'Private Sub txtFacFecha_Validate(Cancel As Boolean)
'If ValidaFecha(txtFacFecha) <> "" Then
'   MsgBox " Fecha no Válida... ", vbCritical, "Error"
'   txtFacFecha.SelStart = 0
'   txtFacFecha.SelLength = Len(txtFacFecha)
'   Cancel = True
'Else
'   GetTipCambio CDate(txtFacFecha)
'   txtTipCompra = Format(gnTipCambioC, "###,###,##0.000")
'End If
'End Sub
'
'Private Sub txtFecha_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 Then
''   If ValidaFecha(txtFecha.Text) <> "" Then
''      MsgBox " Fecha no Válida... ", vbInformation, "Aviso"
''      txtFecha.SelStart = 0
''      txtFecha.SelLength = Len(txtFecha)
''   Else
''      txtMovNro = GeneraMovNro(, , txtFecha)
''      txtProvRuc.SetFocus
''   End If
''End If
'End Sub
'
'Private Sub txtFecha_Validate(Cancel As Boolean)
''If ValidaFecha(txtFecha.Text) <> "" Then
''   MsgBox " Fecha no Válida ", vbInformation, "Aviso"
''   Cancel = True
''Else
''   txtMovNro = GeneraMovNro(, , txtFecha)
''End If
'End Sub
'
'Private Sub txtGRFecha_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
''   If Not IsDate(txtGRFecha.Text) Then
'   If ValidaFecha(txtGRFecha.Text) <> "" Then
'      MsgBox " Fecha no Válida... ", vbCritical, "Error"
'      txtGRFecha.SelStart = 0
'      txtGRFecha.SelLength = Len(txtGRFecha)
'   Else
'      TabDocRef.Tab = 2
'   End If
'End If
'End Sub
'
'Private Sub txtGRNro_KeyPress(KeyAscii As Integer)
''KeyAscii = intfNumEnt(KeyAscii)
'If KeyAscii = 13 Then
'   txtGRNro = Right(String(8, "0") & txtGRNro, 8)
'   txtGRFecha.SetFocus
'End If
'End Sub
'Private Sub txtFacNro_KeyPress(KeyAscii As Integer)
''KeyAscii = intfNumEnt(KeyAscii)
'If KeyAscii = 13 Then
'   txtFacNro = Right(String(8, "0") & txtFacNro, 8)
'   txtFacFecha.SetFocus
'End If
'End Sub
'
'Private Sub txtGRSerie_KeyPress(KeyAscii As Integer)
''KeyAscii = intfNumEnt(KeyAscii)
'If KeyAscii = 13 Then
'   txtGRSerie = Right(String(3, "0") & txtGRSerie, 3)
'   txtGRNro.SetFocus
'End If
'End Sub
'Private Sub txtFacSerie_KeyPress(KeyAscii As Integer)
''KeyAscii = intfNumEnt(KeyAscii)
'If KeyAscii = 13 Then
'   txtFacSerie = Right(String(3, "0") & txtFacSerie, 3)
'   txtFacNro.SetFocus
'End If
'End Sub
'
'Private Sub txtIGV_GotFocus()
'txtIGV.SelStart = 0
'txtIGV.SelLength = Len(txtIGV)
'End Sub
'
'Private Sub txtIGV_KeyPress(KeyAscii As Integer)
'Dim nIGV As Currency
'KeyAscii = intfNumDec(txtIGV, KeyAscii, 14, 2)
'If KeyAscii = 13 Then
'   nIGV = Round(Val(Format(txtTot, gcFormDato)) * nTasaIGV / 100, 2)
'   If Val(Format(txtIGV, gcFormDato)) > nIGV + 0.01 Or _
'      Val(Format(txtIGV, gcFormDato)) < nIGV - 0.01 Then
'      txtIGV = Format(nIGV, gcFormView)
'      nVariaIGV = 0
'   Else
'      txtIGV = Format(txtIGV, gcFormView)
'      fgDetalle.SetFocus
'      nVariaIGV = Val(Format(txtIGV, gcFormDato)) - nIGV
'   End If
'End If
'End Sub
'
'Private Sub txtImporte_KeyPress(KeyAscii As Integer)
'Dim nPos As Integer
'Dim nImporte As Currency
'KeyAscii = intfNumDec(txtImporte, KeyAscii, 14, 2)
'If KeyAscii = 13 Then
'   nPos = fgAsiento.Rows - 1
'   AdicionaRow fgAsiento, nPos
'   EnumeraItems fgAsiento
'   nImporte = Val(Format(txtImporte, gcFormDato))
'   fgAsiento.TextMatrix(nPos, 1) = txtCodServ
'   fgAsiento.TextMatrix(nPos, 2) = txtDescServ
'   fgAsiento.TextMatrix(nPos, 3) = Format(nImporte, gcFormView)
'   fgAsiento.TextMatrix(nPos, 5) = "D"
'   FrameServicio.Visible = False
'   txtIGV = Format(Val(Format(txtIGV, gcFormDato)) + Round(Val(Format(txtImporte, gcFormDato)) - (Val(Format(txtImporte, gcFormDato)) * 100 / (nTasaIGV + 100)), 2), gcFormView)
'   fgAsiento.TextMatrix(nPos + 1, 4) = Format(Val(Format(fgAsiento.TextMatrix(nPos + 1, 4), gcFormDato)) + Val(Format(nImporte, gcFormDato)), gcFormView)
'   SumaAsiento
'   fgAsiento.SetFocus
'End If
'End Sub
'
'Private Sub txtMonto_KeyPress(KeyAscii As Integer)
'Dim nDif As Currency
'If KeyAscii = 13 Then
'   nDif = Val(Format(fgAsiento.Text, gcFormDato)) - Val(Format(txtMonto, gcFormDato))
'   fgAsiento.Text = Format(txtMonto, gcFormView)
'   fgAsiento.TextMatrix(fgAsiento.Rows - 1, 4) = Format(Val(Format(fgAsiento.TextMatrix(fgAsiento.Rows - 1, 4), gcFormDato)) - nDif, gcFormView)
'   SumaAsiento
'   fgAsiento.SetFocus
'End If
'End Sub
'
'Private Sub txtMonto_LostFocus()
'txtMonto.Visible = False
'End Sub
'
'Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   KeyAscii = 0
'   Select Case TabDocRef.Tab
'          Case 0:
'                  If TabDocRef.TabEnabled(0) Then
'                     txtOCNro.SetFocus
'                  Else
'                     TabDocRef.Tab = 1
'                  End If
'          Case 1: txtGRSerie.SetFocus
'          Case 2: txtFacSerie.SetFocus
'   End Select
'End If
'End Sub
'
'Private Sub txtProvRuc_GotFocus()
'txtProvCod.Width = txtProvRuc.Width - cmdExaCab.Width
'cmdExaCab.Visible = True
'txtExaCab.Visible = True
'VerTxtObj
'End Sub
'Private Sub txtProvRuc_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   If ValidarProvee(txtProvRuc) Then
'      txtMovDesc.SetFocus
'   End If
'End If
'End Sub
'Private Sub txtProvRuc_LostFocus()
'txtProvCod.Width = txtProvCod.Width + cmdExaCab.Width
'End Sub
'Private Sub txtProvRuc_Validate(Cancel As Boolean)
'If Not ValidarProvee(txtProvRuc) Then
'   Cancel = True
'End If
'End Sub
'
'Private Sub txtMovDesc_GotFocus()
'cmdExaCab.Visible = False
'txtExaCab.Visible = False
'VerTxtObj
'End Sub
'
'Private Sub txtObj_GotFocus()
'ActivaExaminar txtObj
'End Sub
'
'Private Sub txtObj_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 40 Or KeyCode = 38 Then
'   txtObj_KeyPress 13
'   SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
'End If
'End Sub
'
'Private Sub txtObj_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   If ValidaObj Then
'      ActualizaFG fgDetalle.Row
'      fgDetalle.Enabled = True
'      txtObj.Visible = False
'      cmdExaminar.Visible = False
'      fgDetalle.SetFocus
'   End If
'End If
'End Sub
'Private Function ValidaObj() As Boolean
'ValidaObj = False
'If Len(txtObj) = 0 Then
'   txtObj.Visible = False
'   cmdExaminar.Visible = False
'   EliminaRow fgDetalle, fgDetalle.Row
'   Exit Function
'End If
'sSQL = FormaSelect(txtObj, 6)
'Set rs = CargaRecord(sSQL)
'If Not RSVacio(rs) Then
'   sObjCod = rs!cObjetoCod
'   sObjDesc = rs!cObjetoDesc
'   sObjUnid = rs!cUnidadAbrev
'   sCtaCod = rs!cCtaContCod
'   sCtaDesc = rs!cCtaContDesc
'Else
'   MsgBox "Objeto no encontrado...!", vbCritical, "Error de Búsqueda"
'   Exit Function
'End If
'ValidaObj = True
'End Function
'
'Private Sub txtObj_LostFocus()
'DesactivaExaminar txtObj
'End Sub
'
'Private Sub txtCant_KeyPress(KeyAscii As Integer)
'KeyAscii = intfNumDec(txtCant, KeyAscii, 12, 2)
'If KeyAscii = 13 Then
'   fgDetalle.Text = Format(txtCant.Text, gcFormView)
'   txtCant.Visible = False
'   fgDetalle.SetFocus
'   SumasDoc
'   nVariaIGV = 0
'End If
'End Sub
'Private Sub SumasDoc()
'Dim N As Integer
'Dim nTot As Currency
'For N = 1 To fgDetalle.Rows - 1
'    If fgDetalle.TextMatrix(N, 7) <> "" And fgDetalle.TextMatrix(N, 5) <> "" Then
'       nTot = nTot + Val(Format(fgDetalle.TextMatrix(N, 7), gcFormDato))
'    End If
'Next
'If nTot > 0 Then
'   txtTot = Format(nTot, gcFormView)
'   txtIGV = Format(Round(nTot * nTasaIGV / 100, 2), gcFormView)
'Else
'   txtTot = "": txtIGV = ""
'End If
'End Sub
'
'Private Sub txtCant_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 40 Or KeyCode = 38 Then
'   txtCant_KeyPress 13
'   SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
'End If
'End Sub
'
'Private Sub txtCant_LostFocus()
'txtCant.Text = ""
'txtCant.Visible = False
'End Sub
'Private Sub txtCant_Validate(Cancel As Boolean)
'If Val(txtCant) = 0 Then
'   MsgBox "Cantidad debe ser mayor que Cero...!", vbCritical, "Aviso"
'   Cancel = True
'End If
'End Sub
'
'Private Sub cmdAceptar_Click()
''Dim N As Integer 'Contador
''Dim nItem As Integer, nCol  As Integer
''Dim sTexto As String, lOk As Boolean
''Dim sMovNro As String
''On Error GoTo ErrAceptar
''For N = 1 To fgDetalle.Rows - 1
''    lOk = True
''    If Len(fgDetalle.TextMatrix(N, 1)) = 0 Then
''       MsgBox " Item " & N & " vacio sin Objetos...! ", vbCritical, "Error de Movimiento"
''       nCol = 1
''       lOk = False
''    Else
''     If Val(Format(fgDetalle.TextMatrix(N, 5), gcFormDato)) = 0 Then
''        MsgBox " No se asignaron cantidades a un Objeto...! ", vbCritical, "Error de Movimiento"
''        nCol = 5
''        lOk = False
''     Else
''      If Val(Format(fgDetalle.TextMatrix(N, 7), gcFormDato)) = 0 Then
''         MsgBox " No se asignó Costo de un Objeto...! ", vbCritical, "Error de Movimiento"
''         nCol = 7
''         lOk = False
''      End If
''     End If
''    End If
''    If Not lOk Then
''       TabDoc.Tab = 0
''       fgDetalle.Row = N
''       fgDetalle.Col = nCol
''       fgDetalle.SetFocus
''       Exit Sub
''    End If
''Next
''If txtFacSerie = "" Or txtFacNro = "" Or Not ValidaFecha(txtFacFecha) = "" Then
''   MsgBox "Faltan datos de Factura. Por favor verificar...!", vbCritical, "Error"
''   Exit Sub
''End If
''
'''Primero Salvamos MovNro para hacer referencia cuando es atención de Orden de Compra
''sMovNro = gcMovNro
''gsGlosa = txtMovDesc
''
''If MsgBox(" ¿ Seguro de grabar Operación ? ", vbOKCancel, "Aviso de Confirmación") = vbCancel Then
''   Exit Sub
''End If
''
''' Iniciamos transaccion
''If lTransActiva Then
''   dbCmact.RollbackTrans
''   lTransActiva = False
''End If
''dbCmact.BeginTrans
''lTransActiva = True
'''gcMovNro = GeneraMovNro(txtMovNro)
''
''' Grabamos en Mov
''GrabaMov
''
''' Grabamos en MovObj y MovCant
''For N = 1 To fgAsiento.Rows - 1
''    If fgAsiento.TextMatrix(N, 5) = "D" Then
''       GrabaMovCta Format(N, "000"), fgAsiento.TextMatrix(N, 1), Val(Format(fgAsiento.TextMatrix(N, 3), gcFormDato))
''    Else
''       GrabaMovCta Format(N, "000"), fgAsiento.TextMatrix(N, 1), Val(Format(fgAsiento.TextMatrix(N, 4), gcFormDato)) * -1
''    End If
''    If N <= fgDetalle.Rows - 1 Then    'Para Bienes
''       GrabaMovObj Format(N, "000"), fgDetalle.TextMatrix(N, 1), "1"
''       'Actualizamos los Stocks
''       GrabaMovCant Format(N, "000"), Val(Format(fgDetalle.TextMatrix(N, 5), gcFormDato))
''    Else 'Asignamos el Objeto Proveedor a la cuenta de Provisión
''       If N = fgAsiento.Rows - 1 Then
''          GrabaMovObj Format(N, "000"), "00" & txtProvCod, "1"
''       Else
''       End If
''    End If
''Next
''
''gcDocNro = GeneraDocNro(gcDocTpo)
''sSql = "INSERT INTO  MovDoc (cMovNro, cDocTpo, cDocNro, dDocFecha) VALUES('" & gcMovNro & "', '" & gcDocTpo & "', '" & gcDocNro _
''     & "', '" & GetFechaMov(gcMovNro, False) & "')"
''dbCmact.Execute sSql
''
''' Actualizamos Nro de Nota de Ingreso
''lblDoc.Caption = "VALE DE INGRESO Nº " & gcDocNro
''
'''Grabamos MovRef Referencias a la Orden de Compra ...si existe
'''Primero referenciamos a la Orden de Compra
''If txtOCNro <> "" Then
''   sSql = "INSERT INTO  MovRef VALUES('" & gcMovNro & "', '" & sMovNro & "')" _
''        & ""
''   dbCmact.Execute sSql
''End If
'''Ahora la Guía de Remisión
''If txtGRNro <> "" Then
''   sSql = "INSERT INTO MovDoc (cMovNro, cDocTpo, cDocNro, dDocFecha) VALUES('" & gcMovNro & "', '" & sDocGR & "', '" _
''        & IIf(txtGRSerie <> "", txtGRSerie & "-", "") & txtGRNro & "', '" & Format(CDate(txtGRFecha), "mm/dd/yyyy") & "') "
''   dbCmact.Execute sSql
''End If
'''Por ultimo, la Factura
''If txtFacNro <> "" Then
''   sSql = "INSERT INTO MovDoc (cMovNro, cDocTpo, cDocNro, dDocFecha) VALUES('" & gcMovNro & "', '" & gcDocTpoFac & "', '" _
''        & IIf(txtFacSerie <> "", txtFacSerie & "-", "") & txtFacNro & "', '" & Format(CDate(txtFacFecha), "mm/dd/yyyy") & "') "
''   dbCmact.Execute sSql
''End If
''' Como se trata de Factura de posee IGV, los grabamos en movotros
''sSql = "INSERT INTO  MovOtrosItem VALUES('" & gcMovNro & "', '001', '" & gcCtaIGV & "', " & Val(Format(txtIGV, gcFormDato)) & ", '')"
''dbCmact.Execute sSql
''
'''Grabamos el Tipo de Cambio de Compra
''If GSSIMBOLO = gcME Then
''   GeneraMovME gcMovNro, nVal(txtTipCompra)
''End If
''
''dbCmact.CommitTrans
''   OK = True
''   lTransActiva = False
''   cmdDocumento_Click
''   If MsgBox("¿ Desea continuar ingresando Documentos?", vbQuestion + vbYesNo, "¿Confimación!") = vbNo Then
''      Exit Sub
''   End If
''   txtMovDesc = ""
''   txtProvRuc = ""
''   txtProvNom = ""
''   txtMovNro = GeneraMovNro(, , txtFecha)
''   txtFacNro = ""
''   txtFacSerie = ""
''   txtFacFecha = "  /  /    "
''   txtTot = ""
''   txtIGV = ""
''   fgDetalle.Rows = 2
''   EliminaRow fgDetalle, 1
''   fgAsiento.Rows = 2
''   EliminaRow fgAsiento, 1
''   SumaAsiento
''   txtProvRuc.SetFocus
''Exit Sub
''ErrAceptar:
''   MsgBox TextErr(Err.Description), vbCritical, "Error de Actualización"
'End Sub
'
'Public Property Get lOk() As Boolean
'lOk = OK
'End Property
'
'Public Property Let lOk(ByVal vNewValue As Boolean)
'OK = vNewValue
'End Property
'
'Private Sub AsignaImpuesto()
'Dim nIGV As Currency, nIGVn As Currency, nCant As Currency
'Dim nPos As Integer
'Dim N As Integer
' 'Buscamos Impuestos de Factura
' sSQL = "SELECT b.cCtaContCod, b.cCtaContDesc, c.cImpAbrev " _
'      & "FROM    " & gcCentralCom & "CtaCont as b,  " & gcCentralCom & "Impuesto c " _
'      & "WHERE  b.cCtaContCod = c.cCtaContCod and " _
'      & "       c.cCtaContCod = '" & gcCtaIGV & "'"
' Set rs = CargaRecord(sSQL)
' If RSVacio(rs) Then
'    MsgBox "Impuesto IGV de Factura no existe. Por favor verificar...!", vbCritical, "Error"
' Else
'    If cboFacDestino.ListIndex = 0 Then
'       nIGV = Round(Val(Format(txtTot, gcFormDato) * (nTasaIGV / 100)), 2)
'    Else
'         nIGV = 0: nIGVn = 0
'         For N = 1 To fgDetalle.Rows - 1
'             If fgDetalle.TextMatrix(N, 1) <> "" Then
'                nCant = Round(Val(Format(fgAsiento.TextMatrix(N, 3), gcFormDato)) * nTasaIGV / 100, 2)
'                If fgDetalle.TextMatrix(N, 10) = "X" And cboFacDestino.ListIndex = 1 Then
'                   nIGV = nIGV + nCant
'                Else
'                   nIGVn = nIGVn + Round(Val(Format(fgDetalle.TextMatrix(N, 7), gcFormDato)) * nTasaIGV / 100, 2)
'                   fgAsiento.TextMatrix(N, 3) = Format(Val(Format(fgAsiento.TextMatrix(N, 3), gcFormDato)) + nCant, gcFormView)
'                End If
'             End If
'         Next
'    End If
'    If nIGV > 0 Then
'       AdicionaRow fgAsiento, fgAsiento.Rows - 1
'       EnumeraItems fgAsiento
'       nPos = fgAsiento.Row
'       fgAsiento.TextMatrix(nPos, 0) = nPos
'       fgAsiento.TextMatrix(nPos, 1) = rs!cCtaContCod
'       fgAsiento.TextMatrix(nPos, 2) = rs!cCtaContDesc & " (" & rs!cimpabrev & ")"
'       fgAsiento.TextMatrix(nPos, 3) = Format(nIGV + nVariaIGV, gcFormView)
'       fgAsiento.TextMatrix(nPos, 5) = "D"
'       fgAsiento.TextMatrix(nPos + 1, 0) = nPos + 1
'       fgAsiento.TextMatrix(nPos + 1, 4) = Format(Val(Format(fgAsiento.TextMatrix(nPos + 1, 4), gcFormDato)) + Round(Val(Format(txtTot, gcFormDato) * (nTasaIGV / 100)), 2) + (nVariaIGV), gcFormView)
'    End If
'    If nIGVn > 0 Then
'       If nIGV = 0 Then
'          fgAsiento.TextMatrix(fgAsiento.Rows - 1, 4) = Format(Val(Format(fgAsiento.TextMatrix(fgAsiento.Rows - 1, 4), gcFormDato)) + nIGVn + (nVariaIGV), gcFormView)
'          fgAsiento.TextMatrix(fgAsiento.Rows - 2, 3) = Format(Val(Format(fgAsiento.TextMatrix(fgAsiento.Rows - 2, 3), gcFormDato)) + (nVariaIGV), gcFormView)
'       Else
'          fgAsiento.TextMatrix(fgAsiento.Rows - 1, 4) = Format(Val(Format(fgAsiento.TextMatrix(fgAsiento.Rows - 1, 4), gcFormDato)) + nIGVn, gcFormView)
'       End If
'    End If
' End If
'End Sub
'Private Sub RefrescaAsiento()
'Dim N As Integer, nPos As Integer
'Dim nItem As Integer
'Dim nCant As Currency, nCosto As Currency
'If fgDetalle.TextMatrix(fgDetalle.Rows - 1, 1) = "" Or fgDetalle.TextMatrix(fgDetalle.Rows - 1, 5) = "" Then
'   EliminaRow fgDetalle, fgDetalle.Rows - 1
'End If
'For N = 1 To fgDetalle.Rows - 1
'    If Len(fgDetalle.TextMatrix(N, 1)) > 0 Then
'       fgAsiento.TextMatrix(N, 3) = Format(Round(Val(Format(fgDetalle.TextMatrix(N, 7), gcFormDato)), 2), gcFormView)
'    End If
'Next
'For N = 1 To fgDetalle.Rows - 1
'    If fgDetalle.TextMatrix(N, 1) <> "" Then
'       nPos = CuentaEnAsiento(fgDetalle.TextMatrix(N, 8))
'       If nPos = 0 Then
'          AdicionaRow fgAsiento
'          nPos = fgAsiento.Row
'          fgAsiento.TextMatrix(nPos, 0) = nPos
'          fgAsiento.TextMatrix(nPos, 1) = fgDetalle.TextMatrix(N, 8)
'          fgAsiento.TextMatrix(nPos, 2) = fgDetalle.TextMatrix(N, 9)
'       End If
'       nCant = Val(Format(fgDetalle.TextMatrix(N, 5), gcFormDato))
'       nCosto = Round(Val(Format(fgDetalle.TextMatrix(N, 7), gcFormDato)), 2)
'       fgAsiento.TextMatrix(nPos, 3) = ""
'       fgAsiento.TextMatrix(nPos, 4) = Format(Val(Format(fgAsiento.TextMatrix(nPos, 4), gcFormDato)) + nCosto, gcFormView)
'       fgAsiento.TextMatrix(nPos, 5) = "H"
'    End If
'Next
'End Sub
'Private Sub EliminaAsiento()
'Dim nItems As Integer, N As Integer
'nItems = fgAsiento.Rows - 1
'For N = fgDetalle.Rows To nItems
'    EliminaRow fgAsiento, fgDetalle.Rows
'Next
'If fgDetalle.TextMatrix(fgDetalle.Rows - 1, 1) = "" Then
'   EliminaRow fgAsiento, fgDetalle.Rows - 1
'End If
'End Sub
'
'Private Function CuentaEnAsiento(sCod As String) As Integer
'Dim I As Integer
'For I = fgDetalle.Rows To fgAsiento.Rows - 1
'    If fgAsiento.TextMatrix(I, 1) = sCod Then
'       CuentaEnAsiento = I
'    End If
'Next
'End Function
'
'Private Function GetAsignaProvision(sCod As String, nItem As Integer) As Boolean
'' Cuenta de Provision
'GetAsignaProvision = False
'sSQL = "select a.cCtaContCod, b.cCtaContDesc " _
'     & "from  " & gcCentralCom & "OpeCta a,  " & gcCentralCom & "CtaCont b " _
'     & "where b.cCtaContCod = a.cCtaContCod " _
'     & "and a.cOpeCtaDH = 'H' and a.cOpeCod = '" & gcOpeCod & "'"
'Set rs = CargaRecord(sSQL)
'If RSVacio(rs) Then
'   MsgBox "Falta definir Cuenta de Provision...Por favor revisar Operaciones", vbCritical, "Error"
'   Exit Function
'End If
'fgDetalle.TextMatrix(nItem, 8) = rs!cCtaContCod
'fgDetalle.TextMatrix(nItem, 9) = rs!cCtaContDesc
'GetAsignaProvision = True
'End Function
'
'
