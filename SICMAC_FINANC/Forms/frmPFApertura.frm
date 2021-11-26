VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPFApertura 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apertura de Cuentas"
   ClientHeight    =   6690
   ClientLeft      =   1110
   ClientTop       =   1245
   ClientWidth     =   8025
   Icon            =   "frmPFApertura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   6510
      TabIndex        =   58
      Top             =   6120
      Width           =   1365
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3780
      TabIndex        =   57
      Top             =   6120
      Width           =   1365
   End
   Begin VB.CommandButton cmdCarta 
      Caption         =   "&Carta Apertura"
      Height          =   375
      Left            =   120
      TabIndex        =   56
      Top             =   6120
      Width           =   1365
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5145
      TabIndex        =   55
      Top             =   6120
      Width           =   1365
   End
   Begin VB.CommandButton cmdCalendario 
      Caption         =   "Calen&dario"
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
      Left            =   1950
      TabIndex        =   54
      Top             =   6120
      Width           =   1365
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
      Height          =   630
      Left            =   6285
      TabIndex        =   24
      Top             =   30
      Visible         =   0   'False
      Width           =   1620
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
         Left            =   495
         TabIndex        =   1
         Top             =   225
         Width           =   810
      End
      Begin VB.Label lblTipCambio 
         AutoSize        =   -1  'True
         Caption         =   "Fijo"
         Height          =   195
         Left            =   165
         TabIndex        =   25
         Top             =   285
         Width           =   240
      End
   End
   Begin TabDlg.SSTab TabDoc 
      Height          =   2760
      Left            =   75
      TabIndex        =   19
      Tag             =   "0"
      Top             =   3285
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   4868
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   617
      TabCaption(0)   =   "Cheque de &Gerencia"
      TabPicture(0)   =   "frmPFApertura.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frameOrigen"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Transferencia"
      TabPicture(1)   =   "frmPFApertura.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Efectivo"
      TabPicture(2)   =   "frmPFApertura.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Shape2"
      Tab(2).Control(1)=   "Label12"
      Tab(2).Control(2)=   "txtBilleteImporte"
      Tab(2).Control(3)=   "FlexEdit1"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&Otras..."
      TabPicture(3)   =   "frmPFApertura.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdExaminar1"
      Tab(3).Control(1)=   "txtTotalOtrasCtas"
      Tab(3).Control(2)=   "cmdAceptarObj"
      Tab(3).Control(3)=   "cmdCancelarObjeto"
      Tab(3).Control(4)=   "cmdEliminar"
      Tab(3).Control(5)=   "cmdAgregarCta"
      Tab(3).Control(6)=   "txtObj"
      Tab(3).Control(7)=   "fgObj"
      Tab(3).Control(8)=   "txtCta"
      Tab(3).Control(9)=   "txtCelda1"
      Tab(3).Control(10)=   "fgDetalle"
      Tab(3).Control(11)=   "Label18"
      Tab(3).Control(12)=   "Label2"
      Tab(3).Control(13)=   "Shape3"
      Tab(3).ControlCount=   14
      Begin Sicmact.FlexEdit FlexEdit1 
         Height          =   1785
         Left            =   -74790
         TabIndex        =   68
         Top             =   450
         Width           =   7365
         _extentx        =   11615
         _extenty        =   2672
         highlight       =   1
         allowuserresizing=   3
         font            =   "frmPFApertura.frx":037A
         font            =   "frmPFApertura.frx":03A6
         font            =   "frmPFApertura.frx":03D2
         font            =   "frmPFApertura.frx":03FE
         font            =   "frmPFApertura.frx":042A
         fontfixed       =   "frmPFApertura.frx":0456
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
      End
      Begin VB.CommandButton cmdExaminar1 
         Caption         =   "..."
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -72900
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1095
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox txtTotalOtrasCtas 
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
         Height          =   285
         Left            =   -70605
         TabIndex        =   44
         Tag             =   "3"
         Top             =   1455
         Width           =   1680
      End
      Begin VB.CommandButton cmdAceptarObj 
         Caption         =   "&OK"
         Height          =   360
         Left            =   -68775
         TabIndex        =   43
         Top             =   1815
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CommandButton cmdCancelarObjeto 
         Caption         =   "Cance&lar"
         Height          =   360
         Left            =   -68775
         TabIndex        =   42
         Top             =   2190
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   -68760
         TabIndex        =   38
         Top             =   975
         Width           =   1350
      End
      Begin VB.CommandButton cmdAgregarCta 
         Caption         =   "A&gregar"
         Height          =   360
         Left            =   -68760
         TabIndex        =   37
         Top             =   600
         Width           =   1350
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
         Height          =   300
         Left            =   -74340
         TabIndex        =   34
         Top             =   2100
         Visible         =   0   'False
         Width           =   1700
      End
      Begin VB.Frame frameOrigen 
         Caption         =   "Institucion Financiera Giradora"
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
         Height          =   2100
         Left            =   135
         TabIndex        =   28
         Top             =   420
         Width           =   7500
         Begin Sicmact.TxtBuscar txtBuscarIFChq 
            Height          =   345
            Left            =   1155
            TabIndex        =   66
            Top             =   270
            Width           =   1860
            _extentx        =   3281
            _extenty        =   609
            appearance      =   1
            appearance      =   1
            font            =   "frmPFApertura.frx":0484
            appearance      =   1
         End
         Begin VB.TextBox txtNumChq 
            Height          =   315
            Left            =   4875
            MaxLength       =   8
            TabIndex        =   64
            Top             =   630
            Width           =   1185
         End
         Begin VB.TextBox txtCtaBanco 
            Height          =   315
            Left            =   1155
            MaxLength       =   15
            TabIndex        =   63
            Top             =   630
            Width           =   2670
         End
         Begin VB.CheckBox chkPlaza 
            Caption         =   "Misma plaza"
            Height          =   255
            Left            =   4920
            TabIndex        =   5
            Top             =   1065
            Value           =   1  'Checked
            Width           =   1425
         End
         Begin VB.Frame Frame7 
            Height          =   630
            Left            =   4305
            TabIndex        =   48
            Top             =   1470
            Width           =   3195
            Begin VB.TextBox txtChequeImporte 
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
               Height          =   330
               Left            =   870
               TabIndex        =   6
               Tag             =   "2"
               Top             =   180
               Width           =   2175
            End
            Begin VB.Label lblTotal 
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               Caption         =   "Importe :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   90
               TabIndex        =   49
               Top             =   225
               Width           =   765
            End
         End
         Begin MSMask.MaskEdBox txtFechaRegChq 
            Height          =   315
            Left            =   1155
            TabIndex        =   3
            Top             =   990
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin MSMask.MaskEdBox txtFechaValorChq 
            Height          =   315
            Left            =   3330
            TabIndex        =   4
            Top             =   1005
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   " "
         End
         Begin VB.Label txtDescBcoChq 
            ForeColor       =   &H00C00000&
            Height          =   240
            Left            =   3060
            TabIndex        =   17
            Top             =   300
            Width           =   4275
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Cuenta :"
            Height          =   195
            Left            =   150
            TabIndex        =   32
            Top             =   645
            Width           =   945
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Inst. &Financ"
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   285
            Width           =   825
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Cheque"
            Height          =   195
            Left            =   3885
            TabIndex        =   31
            Top             =   675
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Registro :"
            Height          =   195
            Left            =   150
            TabIndex        =   30
            Top             =   1050
            Width           =   675
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Valorización :"
            Height          =   195
            Left            =   2355
            TabIndex        =   29
            Top             =   1050
            Width           =   945
         End
      End
      Begin VB.TextBox txtBilleteImporte 
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
         Height          =   285
         Left            =   -69075
         TabIndex        =   26
         Tag             =   "0"
         Top             =   2265
         Width           =   1680
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cuenta de Institución Financiera :"
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
         Height          =   2100
         Left            =   -74865
         TabIndex        =   23
         Top             =   420
         Width           =   7500
         Begin Sicmact.TxtBuscar txtBuscarCtaTrans 
            Height          =   345
            Left            =   1035
            TabIndex        =   67
            Top             =   270
            Width           =   2565
            _extentx        =   4524
            _extenty        =   609
            appearance      =   1
            appearance      =   1
            font            =   "frmPFApertura.frx":04B0
            appearance      =   1
         End
         Begin VB.Frame Frame5 
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
            Height          =   600
            Left            =   4320
            TabIndex        =   51
            Top             =   1500
            Width           =   3180
            Begin VB.TextBox txtBancoImporte 
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
               Height          =   345
               Left            =   960
               TabIndex        =   11
               Tag             =   "1"
               Top             =   180
               Width           =   2115
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Importe :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   765
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Documento"
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
            Height          =   630
            Left            =   105
            TabIndex        =   50
            Top             =   1260
            Width           =   3705
            Begin VB.CheckBox chkDocOrigen 
               Height          =   195
               Left            =   1755
               TabIndex        =   9
               Top             =   270
               Width           =   225
            End
            Begin VB.CommandButton cmdDoc 
               Caption         =   "Cheque"
               CausesValidation=   0   'False
               Enabled         =   0   'False
               Height          =   345
               Left            =   2025
               TabIndex        =   10
               Top             =   195
               Width           =   1605
            End
            Begin VB.OptionButton optDoc 
               Caption         =   "Carta"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   8
               Top             =   240
               Width           =   720
            End
            Begin VB.OptionButton optDoc 
               Caption         =   "Cheque"
               Height          =   255
               Index           =   0
               Left            =   75
               TabIndex        =   7
               Top             =   240
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.Label txtBancoCta 
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   1140
            TabIndex        =   62
            Top             =   705
            Width           =   6150
         End
         Begin VB.Label txtBancoDesc 
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3720
            TabIndex        =   53
            Top             =   315
            Width           =   3615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nro C&uenta :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   105
            TabIndex        =   18
            Top             =   315
            Width           =   900
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgObj 
         Height          =   855
         Left            =   -74850
         TabIndex        =   35
         Top             =   1755
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1508
         _Version        =   393216
         Cols            =   10
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483638
         AllowBigSelection=   0   'False
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
         _Band(0).Cols   =   10
      End
      Begin VB.TextBox txtCta 
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
         Left            =   -74475
         TabIndex        =   39
         Top             =   1155
         Visible         =   0   'False
         Width           =   1310
      End
      Begin VB.TextBox txtCelda1 
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
         Height          =   285
         Left            =   -70275
         TabIndex        =   40
         Top             =   660
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDetalle 
         Height          =   1020
         Left            =   -74850
         TabIndex        =   41
         Top             =   420
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   1799
         _Version        =   393216
         Cols            =   6
         ForeColorSel    =   -2147483643
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483638
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   2
         ScrollBars      =   2
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
         _Band(0).Cols   =   6
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
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
         Left            =   -71745
         TabIndex        =   45
         Top             =   1485
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Objeto :"
         Height          =   195
         Left            =   -74850
         TabIndex        =   36
         Top             =   1485
         Width           =   555
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
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
         Left            =   -70230
         TabIndex        =   27
         Top             =   2310
         Width           =   615
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   -70425
         Top             =   2250
         Width           =   3045
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H8000000C&
         Height          =   315
         Left            =   -71955
         Top             =   1440
         Width           =   3045
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "&Glosa"
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
      Height          =   1410
      Left            =   90
      TabIndex        =   22
      Top             =   1785
      Width           =   7875
      Begin VB.TextBox txtMovDesc 
         Height          =   735
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   255
         Width           =   7530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4830
         TabIndex        =   47
         Top             =   1095
         Width           =   570
      End
      Begin VB.Label txtImporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Left            =   5460
         TabIndex        =   46
         Top             =   1065
         Width           =   2160
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
      Height          =   630
      Left            =   75
      TabIndex        =   20
      Top             =   30
      Width           =   6195
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   4905
         TabIndex        =   0
         Top             =   195
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label txtMovNro 
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
         Height          =   210
         Left            =   1035
         TabIndex        =   61
         Top             =   255
         Width           =   3000
      End
      Begin VB.Label txtOpeCod 
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
         Height          =   225
         Left            =   315
         TabIndex        =   60
         Top             =   255
         Width           =   705
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha "
         Height          =   240
         Left            =   4410
         TabIndex        =   21
         Top             =   225
         Width           =   555
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   1020
      Left            =   75
      TabIndex        =   13
      Top             =   705
      Width           =   7875
      Begin Sicmact.TxtBuscar txtBuscarBanco 
         Height          =   360
         Left            =   150
         TabIndex        =   65
         Top             =   300
         Width           =   2445
         _extentx        =   3810
         _extenty        =   635
         appearance      =   1
         appearance      =   1
         font            =   "frmPFApertura.frx":04DC
         appearance      =   1
      End
      Begin VB.Label Label6 
         Caption         =   "&Nueva Cuenta"
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
         Left            =   165
         TabIndex        =   12
         Top             =   0
         Width           =   1320
      End
      Begin VB.Label txtCtaTpo 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cuenta      :"
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   2685
         TabIndex        =   14
         Top             =   630
         Width           =   5070
      End
      Begin VB.Label txtObjetoDesc 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Inst. Financiera :"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2685
         TabIndex        =   15
         Top             =   270
         Width           =   4425
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   285
      Left            =   825
      TabIndex        =   59
      Top             =   6120
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmPFApertura.frx":0508
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
End
Attribute VB_Name = "frmPFApertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim sCtaCod As String, sCtaDes As String
Dim lSalir As Boolean
Dim lNuevoBanco As Boolean
Dim aObj() As String ',  aCta(100, 3) As String

'Variables de Billetaje
Dim aMonedas() As String, aBilletes() As String  'Guarda las Monedas y Billetes con valores
Dim nMonedas As Integer, nBilletes As Integer    'Nro de Monedas y Billetes
Dim nMontoM As Currency, nMontoB As Currency            'Subtotal de Monedas y billetes
Dim lnPos As Integer
Dim lsDHUnic As String * 1
Dim lsDHTab As String * 1

''-**************************************************************************
'********************************* NUEVO MODELO ***************************
Dim oDCtasIf As DCajaCtasIF
Dim oNCtasIf As NCajaCtaIF
Dim oNContFunc  As NContFunciones
Dim vsCtaTpo As String
Dim lsDocNroVoucher  As String
Dim lsDocNroOrigen  As String
Dim ldDocFechaOrigen  As Date
Dim lsFileCartaTrans As String
Dim lsFileCartaAper As String
Dim oConec As DConecta
Private Sub Sumas()
    Dim nDebe As Currency
    Dim i As Integer
    For i = 1 To fgBillete.Rows - 1
        nDebe = nDebe + Val(Format(fgBillete.TextMatrix(i, 3), gcFormDato))
    Next
    txtBilleteImporte = Format(nDebe, gcFormView)
    fgBillete.SetFocus
End Sub
Private Sub CargaFlex(psObjCod As String)
    Dim nLong As Integer, nItem As Integer, Nivel As Integer
    Dim lsObjDesc As String, SQLObjs As String
    Dim rs As ADODB.Recordset
    
    nLong = Len(Trim(psObjCod))  ' Del Objeto Tipo
    SQLObjs = "SELECT cObjetoCod, cObjetoDesc from  " & gcCentralCom & "Objeto " _
       & "where SubString(cObjetoCod,1," & nLong & ") = '" & psObjCod _
       & "' and nobjetoNiv = 4 Order by cObjetoCod Desc"
        
        
    'Set rs = CargaRecord(SQLObjs)
    If RSVacio(rs) Then
        MsgBox "Error al asignar Objetos a la Operación. Consultar con Sistemas" & psObjCod & " no tiene Instancias ", vbCritical, "Error de Objetos"
        Unload Me
    End If
    Do While Not rs.EOF
        'AdicionaRow fgBillete
        'nItem = fgBillete.Row
        'fgBillete.TextMatrix(nItem, 0) = ""
        'fgBillete.TextMatrix(nItem, 1) = rs!cObjetoCod
        'fgBillete.TextMatrix(nItem, 2) = rs!cObjetoDesc
        'fgBillete.TextMatrix(nItem, 4) = "D"
        'fgBillete.Col = 2
        'fgBillete.CellBackColor = "&H00DBDBDB"
        'rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Sub

Private Sub TotalImporte()
    txtImporte = Format(Val(Format(txtChequeImporte, gcFormDato)) + Val(Format(txtBancoImporte, gcFormDato)) + Val(Format(txtBilleteImporte, gcFormDato)) + Val(Format(txtTotalOtrasCtas, gcFormDato)), gcFormView)
End Sub
Private Sub chkDocOrigen_Click()
    If chkDocOrigen.Value = Checked Then
        cmdDoc.Enabled = True
    Else
        cmdDoc.Enabled = False
    End If
End Sub

Private Sub chkPlaza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtChequeImporte.SetFocus
End If
End Sub

Private Sub cmdAceptar_Click()
    Dim lnImporte As Currency
    Dim lsCheque As String
    Dim lsCarta As String
    Dim lbNuevoTpo As Boolean
    
    On Error GoTo ErrorAceptar1
    'lbBilletaje = False
    If lsDHUnic = "H" Then
        'If UBound(ltCalendario) = 0 Then
        '    MsgBox "No se ha definido calendario de pagos para pagaré respectivo", vbInformation, "Aviso"
        '    Me.cmdCalendario.SetFocus
        '    Exit Sub
        'End If
    End If
    If Valida = False Then Exit Sub
    
    lnImporte = Val(Format(txtBancoImporte, gcFormDato))
    If lnImporte > 0 And chkDocOrigen.Value = Checked Then
        lsCheque = ""
        If optDoc(0).Value Then
            'lsCheque = FormaCheque(Mid(txtBancoCod, 5, 2))
        Else
            If Dir(sFileCartaOrigen) <> "" Then
                rtf.LoadFile sFileCartaOrigen, 1
                'lsCheque = ProcesaPlantilla(rtf.Text, True)
            End If
        End If
    End If
    lnImporte = Val(Format(txtImporte, gcFormDato))
    If lnImporte = 0 Then
        If lsDHUnic = "D" Then
            If MsgBox("No se indicó el monto de Apertura...!. Se va a generar una Apertura Sin Saldo." & Chr(13) & "Desea Proseguir con la Operación??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                Exit Sub
            End If
        Else
            MsgBox "No se indicó el monto de Registro...!", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

    lsCarta = ""
    If Dir(sFileCarta) <> "" And sDocNroCarta <> "" Then
        rtf.LoadFile sFileCarta, 1
        'lsCarta = ProcesaPlantilla(rtf.Text, True)
    End If

    'Verificamos Nueva Cuenta de Banco en OBJETO
    lbNuevoTpo = False
    SSQL = "SELECT cObjetoCod FROM " & gcCentralCom & "Objeto WHERE cObjetoCod = '" & Mid(txtObjetoCod, 1, 9) & "'"
    'Set rs = CargaRecord(sSql)
    If rs.EOF Then   'No Existe Tipo de Cuenta
        lbNuevoTpo = True
    End If
    rs.Close
    Set rs = Nothing

    If MsgBox(" ¿ Seguro de Grabar Operación ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
        lTransActiva = False
        Exit Sub
    End If

    lTransActiva = True
    
    'sSql = "INSERT INTO mov (cMovNro,cOpeCod,cMovDesc,cMovEstado) " & _
       " VALUES( '" & gcMovNro & "','" & gsOpeCod & "','" & Replace(gcGlosa, "'", "''") & "','0')"


'    nItem1 = 1
    'nImporte = lnImporte
    If lnImporte > 0 Then
        If gbGeneraCal Then
            If ExisteOBjetoFiltro(sCtaCod, Trim(txtObjetoCod)) = False Then
                lsFiltro = GeneraCtaBancos(sCtaCod, Trim(txtObjetoCod))
                SSQL = "INSERT INTO " & gcCentralCom & "CtaObjFiltro(cCtaContCod,cCtaObjOrden,cObjetoCod,cCtaObjMascara) " _
                   & "VALUES('" & sCtaCod & "','1','" & Trim(txtObjetoCod) & "','" & lsFiltro & "')"
     
                dbCmact.Execute SSQL
            End If
            SSQL = "Select * from Opecta where copecod ='" & gsOpeCod & "' and cOpeCtaTpo>='8' and cOpeCtaOpc='1' "
            Set rs = New ADODB.Recordset
            Set rs = CargaRecord(SSQL)
            Do While Not rs.EOF
                lsFiltro = GeneraCtaBancos(rs!cCtaContCod, Trim(txtObjetoCod))
                SSQL = "INSERT INTO " & gcCentralCom & "CtaObjFiltro(cCtaContCod,cCtaObjOrden,cObjetoCod,cCtaObjMascara) " _
                   & "VALUES('" & rs!cCtaContCod & "','1','" & Trim(txtObjetoCod) & "','" & lsFiltro & "')"
     
                dbCmact.Execute SSQL
                rs.MoveNext
            Loop
            rs.Close
            Set rs = Nothing
        End If

        SSQL = "Insert INTO MovCta (cMovNro,cMovItem,cCtaContCod,nMovImporte) " & _
           "   VALUES ('" & gcMovNro & "','" & FillNum(Trim(Str(nItem1)), 3, "0") & "','" & GetCtaObjFiltro(sCtaCod, Trim(Me.txtObjetoCod)) & "'," & IIf(lsDHUnic = "D", nImporte, nImporte * (-1)) & "  )"
 
        dbCmact.Execute (SSQL)

        SSQL = "Insert INTO MovObj(cMovNro,cMovItem,cObjetoCod,cMovObjOrden)  " & _
           " VALUES ('" & gcMovNro & "','" & FillNum(Trim(Str(nItem1)), 3, "0") & "','" & txtObjetoCod & "','1')"
        dbCmact.Execute (SSQL)


    End If

    'Cheques de Gerencia
    nImporte = Val(Format(txtChequeImporte, gcFormDato))
    If nImporte > 0 Then
        nItem1 = nItem1 + 1
        lsCtaHaber = EmiteCuenta(gsOpeCod, txtChequeImporte.Tag, lsDHTab)

        SSQL = "INSERT INTO MOVCTA(cMovNro, cMovItem, cCtaContCod, nMovImporte) " _
           & " VALUES('" & gcMovNro & "','" & FillNum(Trim(Str(nItem1)), 3, "0") & "','" _
           & GetCtaObjFiltro(lsCtaHaber, Trim(aObj(1, 6)), True) & "'," & IIf(lsDHTab = "H", CCur(nImporte) * (-1), nImporte) & ")"
     
        dbCmact.Execute SSQL
 
        SSQL = "INSERT INTO MOVOBJ(cMovNro, cMovItem, cMovObjOrden, cObjetoCod  ) " _
           & " VALUES('" & gcMovNro & "','" & FillNum(Trim(Str(nItem1)), 3, "0") & "','1','" _
           & Trim(aObj(1, 6)) & "')"
     
        dbCmact.Execute SSQL
 
        SSQL = "INSERT INTO MOVOBJ(cMovNro, cMovItem, cMovObjOrden, cObjetoCod  ) " _
           & " VALUES('" & gcMovNro & "','" & FillNum(Trim(Str(nItem1)), 3, "0") & "','2','" _
           & Trim(aObj(1, 5)) & "')"
 
        dbCmact.Execute SSQL
        
        lsCodCta = Right(gsCodAge, 2) & "001" & Mid(gsOpeCod, 3, 1) & Format(gdFecSis, "mmddyy")
        'Inserta el cheque en la tabla de Cheques como depositado por la operacion de apertura
        SSQL = "INSERT INTO ChequeCaja(cCodCta, cNumChq, dRegChq, dValorChq, nMontoTC, nMontoChq, cEstChq, cCodBco, cPlaza, cCodUsu, cCtaBco, cDepBco, cMovNro, cCodAge,cMoneda)" _
           & " VALUES('" & lsCodCta & "', '" & Mid(txtNumChq, 1, 8) & "', '" & Format(txtFechaRegChq, "mm/dd/yyyy") & "', '" _
           & Format(txtFechaValorChq, "mm/dd/yyyy") & "', " & nImporteS & ", " & nImporteS & ", 'E', '" & Mid(txtCodBcoChq, 5, 2) & "', '" _
           & cboPlaza.ListIndex + 1 & "', '" & gsCodUser & "', '" & txtCtaBanco & "', '1', '" & gcMovNro & "', '" & Trim(aObj(1, 6)) & "','" & Mid(gsOpeCod, 3, 1) & "')"
 
        dbCmact.Execute SSQL

        SSQL = "INSERT INTO MOVDOC(cMovNro, nDocTpo, cDocNro, dDocFecha) VALUES('" & gcMovNro & "','" & gnDocTpoCheque & "','" & Mid(txtCodBcoChq, 5, 2) & "-" & FillNum(Mid(txtNumChq, 1, 8), 8, "0") & " ','" & Format(txtFecha, "mm/dd/yyyy") & "')"
        dbCmact.Execute SSQL
    End If

    nImporte = Val(Format(txtBancoImporte, gcFormDato))
    If nImporte > 0 Then
        lsCtaHaber = EmiteCuenta(gsOpeCod, txtBancoImporte.Tag, lsDHTab)
        nItem1 = nItem1 + 1
        sTexto = GetCtaObjFiltro(lsCtaHaber, txtBancoCod, False)
        GrabaMovCta FillNum(Trim(Str(nItem1)), 3, "0"), GetCtaObjFiltro(lsCtaHaber, txtBancoCod, True), IIf(lsDHTab = "H", CCur(nImporte) * (-1), nImporte)
        GrabaMovObj FillNum(Trim(Str(nItem1)), 3, "0"), txtBancoCod, "1"
        If sCheque <> "" Then
            SSQL = "INSERT INTO movdoc(cMovNro, nDocTpo, cDocNro, dDocFecha) VALUES ('" & gcMovNro & "', '" & sDocTpoOrigen & "', '" & sDocNroOrigen & "', '" & Format(dDocFechaOrigen, "mm/dd/yyyy") & "')"
            dbCmact.Execute SSQL
            If optDoc(0).Value Then
                SSQL = "INSERT INTO movdoc(cMovNro, nDocTpo, cDocNro, dDocFecha) VALUES ('" & gcMovNro & "', '" & sDocTpoVoucher & "', '" & sDocNroVoucher & "', '" & Format(dDocFechaOrigen, "mm/dd/yyyy") & "')"
                dbCmact.Execute SSQL
            End If
        End If
    End If
    nImporte = Val(Format(txtBilleteImporte, gcFormDato))
    If nImporte > 0 Then
        lbBilletaje = True
        nMonedas = 0: nBilletes = 0
        nMontoM = 0: nMontoB = 0
        lsCtaHaber = EmiteCuenta(gsOpeCod, txtBilleteImporte.Tag, lsDHTab)
        With fgBillete
            For N = 1 To fgBillete.Rows - 1
                nImporte = Val(Format(.TextMatrix(N, 3), gcFormDato))
                'nImporteS = nImporte
                If nImporte <> 0 Then
                    nItem1 = nItem1 + 1
                    SSQL = "INSERT INTO  MovCta(cMovNro, cMovItem, cCtaContCod, nMovImporte) " _
                       & " VALUES ('" & gcMovNro & "', '" & FillNum(Trim(Str(nItem1)), 3, "0") & "', '" _
                       & GetCtaObjFiltro(lsCtaHaber, .TextMatrix(N, 1), True) & "', " & IIf(lsDHTab = "H", CCur(nImporte) * (-1), nImporte) & ")"
          
                    dbCmact.Execute SSQL
       
                    SSQL = "INSERT INTO MovObj(cMovNro, cMovItem, cMovObjOrden, cObjetoCod) " _
                       & " VALUES ('" & gcMovNro & "', '" & FillNum(Trim(Str(nItem1)), 3, "0") & "','1','" _
                       & .TextMatrix(N, 1) & "')"
            
                    dbCmact.Execute SSQL
                End If
            Next
        End With
    End If
    nImporte = Val(Format(Me.txtTotalOtrasCtas, gcFormDato))
    If nImporte > 0 Then
        For i = 1 To fgDetalle.Rows - 1
            nImporte = Abs(Val(Format(fgDetalle.TextMatrix(i, 3), gcFormDato)))
            nItem1 = nItem1 + 1
            'Vemos el Objeto
            lsCtaHaber = fgDetalle.TextMatrix(i, 1)
            For N = 1 To fgObj.Rows - 1
                If fgObj.TextMatrix(N, 4) = fgDetalle.TextMatrix(i, 1) And fgObj.TextMatrix(N, 0) = fgDetalle.TextMatrix(i, 0) Then
                    If Len(fgObj.TextMatrix(N, 2)) > 0 Then
                        SSQL = "INSERT INTO  Movobj(cMovNro, cMovItem, cMovObjOrden, cObjetoCod) values ('" & gcMovNro & "', '" & Format(nItem1, "000") _
                           & "', '" & fgObj.TextMatrix(N, 1) & "', '" & fgObj.TextMatrix(N, 2) & "')"
                        dbCmact.Execute SSQL
                        lsCtaHaber = GetCtaObjFiltro(fgDetalle.TextMatrix(i, 1), fgObj.TextMatrix(N, 2), True)
                    End If
    
                Else
                    If lsDHTab = "D" Then
                        SSQL = "INSERT INTO  Movobj(cMovNro, cMovItem, cMovObjOrden, cObjetoCod) values ('" & gcMovNro & "', '" & Format(nItem1, "000") _
                           & "', '1', '" & Trim(txtObjetoCod) & "')"
                        dbCmact.Execute SSQL
                    End If
                End If
            Next
            SSQL = "INSERT INTO  MovCta(cMovNro, cMovItem, cCtaContCod, nMovImporte) VALUES ('" & gcMovNro & "', '" & Format(nItem1, "000") _
               & "', '" & lsCtaHaber & "', " & IIf(lsDHTab = "H", CCur(nImporte) * (-1), nImporte) & ")"
            dbCmact.Execute SSQL
        Next
    End If

    'Grabamos el Objeto Banco
    If lNuevoTpo Then
        SSQL = "INSERT INTO " & gcCentralCom & "Objeto(cObjetoCod, cObjetoDesc,nObjetoNiv ) VALUES ('" & Mid(txtObjetoCod, 1, 9) & "', '" & txtCtaTpo & "', 4 )"
        dbCmact.Execute SSQL
    End If
    SSQL = "INSERT INTO " & gcCentralCom & "Objeto(cObjetoCod, cObjetoDesc,nObjetoNiv ) VALUES ('" & txtObjetoCod & "', 'Pendiente', 5 )"
    dbCmact.Execute SSQL

    'Grabamos en CtaBanco
    If gbGeneraCal Then
        'Grabamos en CtaBanco
        SSQL = "INSERT INTO CtaBanco (cObjetoCod, cCtaBcoDesc, dCtaBcoAper, dCtaBcoVenc, cMovNro, cCtaBcoTipo,nCtaBcoCuotas,nCtaBcoPlazo)  " _
           & " VALUES ('" & txtObjetoCod & "', '" & Trim(txtObjetoDesc & " " & txtCtaTpo) & "', '" & Format(txtFecha, "mm/dd/yyyy") & "', '" _
           & Format(txtFecha, "mm/dd/yyyy") & "','" & gcMovNro & "','A'," & frmAdeudCal.SpnCuotas.Valor & "," & frmAdeudCal.txtPlazoCuotas & ")"

        dbCmact.Execute SSQL

        'cObjetoCod           nAdeudSaldoCap        nAdeudPerGracia nAdeudComIni cAdeudPlaza dAdeudUltPago               nAdeudCuotaCap cAdeudEjecutado cAdeudMonedaPag
        SSQL = "INSERT INTO CtaBancoAdeud (cObjetoCod, nAdeudSaldoCap, nAdeudPerGracia, nAdeudComIni, cAdeudPlaza, " _
           & " dAdeudUltPago, nAdeudCuotaCap , cAdeudEjecutado , cAdeudMonedaPag ) " _
           & " VALUES ('" & txtObjetoCod & "'," & Format(frmAdeudCal.txtCapital, "#0.00") & "," & frmAdeudCal.SpnGracia.Valor & "," _
           & frmAdeudCal.txtComisionIni & ",'" & IIf(frmAdeudCal.chkInterno.Value = 1, "0", "1") & "','" _
           & Format(txtFecha, "mm/dd/yyyy") & "'," _
           & frmAdeudCal.txtCuotaPagoK & ",'E','" & IIf(frmAdeudCal.chkSoles.Value = 1, "1", "2") & "')"
 
        dbCmact.Execute SSQL
        'Grabamos en CtaBancoInteres
        SSQL = "INSERT INTO CtaBancoInteres(cObjetoCod, dCtaBcoIntReg, nCtaBcoIntPorc, nCtaBcoIntPeriodo) " _
           & "VALUES ('" & txtObjetoCod & "', '" & Format(txtFecha, "mm/dd/yyyy") & "'," & frmAdeudCal.txtInteres & "," & IIf(frmAdeudCal.optPeriodo(0).Value, "360", "30") & ")"
  
        dbCmact.Execute SSQL
    Else
        SSQL = "INSERT INTO CtaBanco (cObjetoCod, cCtaBcoDesc, dCtaBcoAper, dCtaBcoVenc, cMovNro, cCtaBcoTipo,nCtaBcoCuotas,nCtaBcoPlazo)  " _
           & "VALUES ('" & txtObjetoCod & "', '" & Trim(txtObjetoDesc & " " & txtCtaTpo) & "', '" & Format(txtFecha, "mm/dd/yyyy") & "', '" & Format(txtFecha, "mm/dd/yyyy") & "','" & gcMovNro & "','C',0,0)"

        dbCmact.Execute SSQL
    End If

    If Not lMN Then
        If Not GeneraMovME(gcMovNro) Then
            dbCmact.RollbackTrans
            lTransActiva = False
            Exit Sub
        End If
    End If
    'GRABAMOS EL CALENDARIO DE LOS ADEUDADOS
    If gbGeneraCal Then
        If UBound(ltCalendario) > 0 Then
            'GRABAMOS LOS POSIBLES PAGOS A REALIZARSE O PAGOS PROYECTADOS
            For N = 0 To UBound(ltCalendario)
                If ltCalendario(N).Cuota <> -1 Then
                    If ltCalendario(N).Cuota = 0 Then
                        'cPlanTipoCuota  A  Cuota adelantada C Cuota Normal
                        SSQL = "INSERT INTO CTABANCOCAL(cObjetoCod, cPlanTipoCuota, nNroCuota , dPlanPago, nPlanCapital, nPlanInteres) " _
                           & " VALUES('" & Trim(txtObjetoCod) & "','A','" & Format(ltCalendario(N).FechaPago, "mm/dd/yyyy") & "'," _
                           & ltCalendario(N).Capital & "," & ltCalendario(N).Interes & ")"
                     
                        dbCmact.Execute SSQL
                    Else
                        SSQL = "INSERT INTO CTABANCOCAL(cObjetoCod, cPlanTipoCuota, nNroCuota , dPlanPago, nPlanCapital, nPlanInteres) " _
                           & " VALUES('" & Trim(txtObjetoCod) & "','C'," & ltCalendario(N).Cuota & ",'" & Format(ltCalendario(N).FechaPago, "mm/dd/yyyy") & "' ," _
                           & ltCalendario(N).Capital & "," & ltCalendario(N).Interes & ")"
                     
                        dbCmact.Execute SSQL
                 
                        If ltCalendario(N).Cuota = 1 Then
                            SSQL = "UPDATE CTABANCO SET dCtaBcoVenc='" & Format(ltCalendario(N).FechaPago, "mm/dd/yyyy") & "' where cObjetocod ='" & Trim(txtObjetoCod) & "'"
                            dbCmact.Execute SSQL
                        End If
                 
                    End If
         
                End If
     
            Next N
        End If
    End If

    dbCmact.CommitTrans
    lTransActiva = False

    Dim lsCadena As String
    lsCadena = ""
    If gnImporte = 0 Then
        lsCadena = ImpreCabAsiento("APERTURAS DE CUENTAS ")
    End If

    If Mid(gsOpeCod, 3, 2) = "17" Or Mid(gsOpeCod, 3, 2) = "27" Then
        lsCadena = lsCadena & " Apertura en Cmac : " & PrnSet("B+") & txtObjetoDesc & PrnSet("B-") & Chr(10)
    Else
        lsCadena = lsCadena & " Apertura en Banco : " & PrnSet("B+") & txtObjetoDesc & PrnSet("B-") & Chr(10)
    End If
    lsCadena = lsCadena & " Propietario      : " & PrnSet("B+") & txtPropietario & PrnSet("B-") & Chr(10)
    lsCadena = lsCadena & "                    " & PrnSet("B+") & txtObjetoCod & " - " & txtCtaTpo & PrnSet("B-")

    If gnImporte > 0 Then
        If sChequeGerencia <> "" Then
            lsCadena = lsCadena & Chr(10) & " Cheque Gerencia   : " & PrnSet("B+") & sChequeGerencia & PrnSet("B-")
        End If
        If sCheque <> "" Then
            lsCadena = lsCadena & Chr(10) & " Cheque Emitido : " & PrnSet("B+") & sDocNroOrigen & PrnSet("B-") & "      Voucher Nro.: " & PrnSet("B+") & sDocNroVoucher & PrnSet("B-")
        End If
        rtf.Text = ImprimeAsientoContable(gcMovNro, "ASIENTO CONTABLE DE APERTURAS", lsCadena, lbBilletaje, "19")
    Else
        lsCadena = lsCadena & Chr(10) & String(90, "=") & Chr(10)
        lsCadena = lsCadena & "APERTURA DE CUENTA SIN SALDO" & Chr(10)
        lsCadena = lsCadena & String(90, "=") & Chr(10)
        rtf.Text = lsCadena
    End If

    frmPrevio.Previo rtf, "Asiento Contable", False, 66
    If MsgBox(" ¿ Desea registrar nueva Operación ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
        Unload Me
        Exit Sub
    End If
    cmdCancelar.Value = True
    Exit Sub
ErrorAceptar1:
    If lTransActiva Then
        dbCmact.RollbackTrans
    End If
    MsgBox "Error N°[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdAceptarObj_Click()
    Dim i As Integer
    For i = 1 To fgObj.Rows - 1
        If Len(Trim(fgObj.TextMatrix(i, 2))) = 0 Then
            MsgBox "Faltan objetos por definir ...!", vbInformation, "Error de Objetos"
            fgObj.SetFocus
            Exit Sub
        End If
    Next
    AsignaEnabled True, True
    fgDetalle.SetFocus
End Sub

Private Sub cmdAgregarCta_Click()
    Dim sCod As String
    Dim nRow As Integer
    nRow = fgDetalle.Rows - 1
    If fgDetalle.TextMatrix(nRow, 0) <> "" Then
        If fgDetalle.TextMatrix(nRow, 3) <> "" Then
            AdicionaRow fgDetalle
            EnfocaObj 0, txtCta
        End If
    Else
        AdicionaRow fgDetalle
        EnfocaObj 0, txtCta
    End If
End Sub

Private Sub cmdCalendario_Click()
    If txtObjetoCod <> "" And Val(txtImporte) > 0 Then
        ReDim ltCalendario(0)
        frmAdeudCal.txtFecha = txtFecha
        frmAdeudCal.txtFecha.Enabled = False
        frmAdeudCal.lnCapital = Format(CCur(txtImporte), "#,#0.00")
        frmAdeudCal.lblCodAdeudado = Trim(txtObjetoCod)
        frmAdeudCal.lblDescEntidad = Trim(txtObjetoDesc)
        frmAdeudCal.Inicio True
    End If
End Sub

Private Sub cmdCancelar_Click()
    Dim N As Integer
    'gbGeneraCal = False
    'txtMovNro = GeneraMovNro(, , txtFecha)
    'For n = 1 To fgBillete.Rows - 1
    '    fgBillete.TextMatrix(n, 3) = ""
    'Next
    txtChequeImporte = ""
    AsignaEnabled True, False
    Me.txtCelda1.Visible = False
    Me.txtCta.Visible = False
    Me.cmdExaminar1.Visible = False
    suma
    If lsDHUnic = "H" Then
        cmdCalendario.Visible = True
    Else
        cmdCalendario.Visible = False
    End If
    ReDim ltCalendario(0)
    txtBancoImporte = ""
    txtBilleteImporte = ""
    'sDocNroCarta = ""
    'sDocNroOrigen = ""
    txtBancoCta = ""
    txtBancoDesc = ""
    txtImporte = ""
    
    txtObjetoDesc = "Inst. Financiera :"
    txtMovDesc = ""
    txtCtaBanco = ""
    txtNumChq = ""
    txtDescBcoChq = ""
    txtChequeImporte = "0.00"
    txtNumChq.Enabled = False
    'txtObjetoCod.Enabled = True
    txtBuscarBanco.Text = ""
    txtBuscarCtaTrans.Text = ""
    txtBuscarIFChq.Text = ""
    'txtObjetoCod.SetFocus
    CabeceraGridOtrasCtasContables
    CabGridObjetosOtrasCtas

    Unload frmAdeudCal
    Set frmAdeudCal = Nothing
  
End Sub

Private Sub cmdCancelarObjeto_Click()
    AsignaEnabled True, True
    cmdEliminar_Click
End Sub

Private Sub cmdCarta_Click()
    'gcFile = sFileCarta
    'gcEntiOrig = txtBancoDesc
    'gcCtaEntiOrig = txtBancoCta
    'gcEntiDest = txtObjetoDesc
    'gcPersona = gcEntiDest
    'gcCtaEntiOrig = txtCtaTpo
    'gcGlosa = Me.txtMovDesc
    'gnImporte = Val(Format(txtImporte, gcFormDato))
    'frmDocPago.Inicio True, False, False
    'If frmDocPago.lOk Then
    '    sDocNroCarta = gcDocNro
    '    dDocFechaCarta = frmDocPago.dDocFecha
    '    cmdAceptar.SetFocus
    'End If
    Dim oNContFun As NContFunciones
    Dim lsTpoDoc As String
    Dim lsDocTpoVoucher As String
    Dim lnImporte As Currency
    Dim lsDocTpo As String
    Dim oDocPago As clsDocPago
    
    Set oDocPago = New clsDocPago
    Set oNContFun = New NContFunciones
    
    lsFileCartaAper = App.Path & "\" & gsOpeCod & ".txt"
    
    If txtBancoDesc = "" And txtBancoCta = "" Then
        MsgBox "Cuenta no válida", vbInformation, "Aviso"
        Exit Sub
    End If
    If txtObjetoDesc = "" Then
        MsgBox "Cuenta de Destino no válida", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Val(txtBancoImporte) = 0 Then
        MsgBox "Monto Ingresado no válido", vbInformation, "Aviso"
        txtBancoImporte.SetFocus
        Exit Sub
    End If
    
    lsDocTpo = gnDocTpoCarta
    lsFileCartaAper = App.Path & "\Spooler\" & gsOpeCod & "T.txt"
    
    lnImporte = txtBancoImporte
    oDocPago.DocPagoShow True, False, False, gsOpeCod, gsOpeCod, _
                         Mid(txtBuscarBanco.Text, 1, 13), gdFecSis, txtBancoDesc, _
                         txtBancoCta, txtObjetoDesc, "", lsFileCartaAper, True, lsDocTpoVoucher, lnImporte, txtMovDesc
    If oDocPago.vbOk Then
        lsDocNroVoucher = oDocPago.vsNroVoucher
        lsDocNroOrigen = oDocPago.vsNroDoc
        ldDocFechaOrigen = oDocPago.vdFechaDoc
        txtMovDesc = oDocPago.vsGlosa
        lsFileCartaTrans = oDocPago.vsFileCarta
        cmdCarta.SetFocus
    End If
    
End Sub

Private Sub cmdDoc_Click()
    Dim oNContFun As NContFunciones
    Dim lsTpoDoc As String
    Dim lsDocTpoVoucher As String
    Dim lsFileCarta As String
    Dim lnImporte As Currency
    Dim lsDocTpo As String
    Dim oDocPago As clsDocPago
    
    Set oDocPago = New clsDocPago
    Set oNContFun = New NContFunciones
     
    If txtBancoDesc = "" And txtBancoCta = "" Then
        MsgBox "Cuenta no válida", vbInformation, "Aviso"
        Exit Sub
    End If
    If Val(txtBancoImporte) = 0 Then
        MsgBox "Monto Ingresado no válido", vbInformation, "Aviso"
        txtBancoImporte.SetFocus
        Exit Sub
    End If
    
    If optDoc(0).Value Then
        lsTpoDoc = gnDocTpoCheque
        If lsDocTpoVoucher <> "" Then
            lsDocNroVoucher = oNContFun.GeneraDocNro(gnDocTpoCheque, Mid(gsOpeCod, 3, 1))
        End If
    Else
        lsDocTpo = gnDocTpoCarta
        lsFileCarta = App.Path & "\Spooler\" & gsOpeCod & "T.txt"
    End If
    lnImporte = txtBancoImporte
    oDocPago.DocPagoShow optDoc(1).Value, optDoc(0).Value, False, gsOpeCod, gsOpeCod, _
                         Mid(txtBuscarBanco, 1, 13), gdFecSis, txtBancoDesc, _
                         txtBancoCta, "", "", lsFileCarta, True, lsDocTpoVoucher, lnImporte, txtMovDesc
    
    If oDocPago.vbOk Then
        lsDocNroVoucher = oDocPago.vsNroVoucher
        lsDocNroOrigen = oDocPago.vsNroDoc
        ldDocFechaOrigen = oDocPago.vdFechaDoc
        txtMovDesc = oDocPago.vsGlosa
        lsFileCartaTrans = oDocPago.vsFileCarta
        cmdCarta.SetFocus
    End If
End Sub

Private Sub cmdEliminar_Click()
    If fgDetalle.TextMatrix(fgDetalle.Row, 0) <> "" Then
        EliminaCuenta fgDetalle.TextMatrix(fgDetalle.Row, 1), fgDetalle.TextMatrix(fgDetalle.Row, 0)
        suma

        If fgDetalle.Enabled Then
            fgDetalle.SetFocus
        End If
        txtCta.Visible = False
        cmdExaminar1.Visible = False
    End If
End Sub
Private Sub EliminaCuenta(sCod As String, nItem As Integer)
    Dim K  As Integer, m As Integer
    EliminaRow fgDetalle, fgDetalle.Row
    K = 1
inicioDo:
    Do While K < fgObj.Rows
        If Len(fgObj.TextMatrix(K, 1)) Then
            If fgObj.TextMatrix(K, 4) = sCod And Val(fgObj.TextMatrix(K, 0)) = nItem Then
                EliminaRow fgObj, K
                GoTo inicioDo
            End If
        End If
        K = K + 1
    Loop
    If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
        RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.Row, 1), fgDetalle.TextMatrix(fgDetalle.Row, 0)
        suma
    Else
        txtTotalOtrasCtas = ""
    End If
End Sub
Private Sub suma()
    Dim nTotal As Currency
    Dim i As Integer
    For i = 1 To fgDetalle.Rows - 1
        nTotal = nTotal + Val(Format(fgDetalle.TextMatrix(i, 3), gcFormDato))
        If fgDetalle.Rows > 2 Then
            fgDetalle.TextMatrix(i, 0) = i
        End If
    Next
    txtTotalOtrasCtas = Format(nTotal, gcFormView)
    TotalImporte
End Sub
Private Sub RefrescaFgObj(sCod As String, nItem As Integer)
    Dim K  As Integer
    For K = 1 To fgObj.Rows - 1
        If Len(fgObj.TextMatrix(K, 1)) Then
            If fgObj.TextMatrix(K, 0) = nItem And fgObj.TextMatrix(K, 4) = sCod Then
                fgObj.RowHeight(K) = 285
            Else
                fgObj.RowHeight(K) = 0
            End If
        End If
    Next
End Sub

'Private Sub cmdExaBanco_Click()
'    Dim rs As adodb.Recordset
'    Dim oDescObj As DescObjeto.ClassDescObjeto
'    Set oDescObj = New DescObjeto.ClassDescObjeto
'
'    Set rs = New adodb.Recordset
'
'    Set rs = New adodb.Recordset
'    Set rs = oDCtasIf.CargaCtasIF("", "01-02")
'    If rs.EOF Then
'        Exit Sub
'    End If
'    oDescObj.Show rs, txtObjetoCod, 3
'    If oDescObj.lbOk Then
'        txtBancoCod = oDescObj.gsSelecCod
'        txtBancoCta = oDescObj.gsSelecDesc
'        txtBancoDesc = oDCtasIf.NombreIF(Mid(oDescObj.gsSelecCod, 1, 13))
'
'        txtBancoImporte.SetFocus
'    Else
'        txtBancoCod = ""
'        txtBancoDesc = ""
'        txtBancoCta = ""
'        cmdExaBanco.SetFocus
'    End If
'    rs.Close: Set rs = Nothing
'End Sub

'Private Sub cmdExaBcoChq_Click()
'    Dim sql As String
'    Dim rs As New adodb.Recordset
'    Dim lsObjeto As String
'    Dim lsNivel As String
'    Dim oDescObj As DescObjeto.ClassDescObjeto
'    Set oDescObj = New DescObjeto.ClassDescObjeto
'
'    Set rs = New adodb.Recordset
'    Set rs = oDCtasIf.CargaCtasIF("01", "", 1)
'    If rs.EOF Then
'        Exit Sub
'    End If
'    oDescObj.Show rs, "", 1, "BANCOS"
'    If oDescObj.lbOk Then
'        txtCodBcoChq = oDescObj.gsSelecCod
'        txtDescBcoChq = oDescObj.gsSelecDesc
'        txtCtaBanco.SetFocus
'        txtChequeImporte.Enabled = True
'    Else
'        txtCodBcoChq = ""
'        txtDescBcoChq = ""
'        If txtCodBcoChq.Enabled Then
'            txtCodBcoChq.SetFocus
'        End If
'    End If
'    rs.Close
'    Set rs = Nothing
'    Set oDescObj = Nothing
'End Sub
'Private Sub cmdExaminar_Click()
'Dim rs As adodb.RecordSet
'Set rs = New adodb.RecordSet
'Dim oDescObj As DescObjeto.ClassDescObjeto
'Set oDescObj = New DescObjeto.ClassDescObjeto
'
'    Set rs = oDCtasIf.CargaCtasIF("01", "", 1)
'    If rs.EOF Then
'        Exit Sub
'    End If
'    oDescObj.Show rs, txtObjetoCod, 1, "BANCOS"
'    If oDescObj.lbOk Then
'        txtObjetoCod = oDescObj.gsSelecCod
'        txtObjetoDesc = "Inst. Financiera :" & oDescObj.gsSelecDesc
'        txtObjetoCod = txtObjetoCod & "." & Trim(oDCtasIf.GeneraCodCtaIF(Trim(oDescObj.gsSelecCod), Mid(gsOpeCod, 3, 1), vsCtaTpo))
'        txtMovDesc.SetFocus
'    Else
'        If txtObjetoCod.Enabled And txtObjetoCod.Visible Then
'            txtObjetoCod.SetFocus
'        End If
'    End If
'    rs.Close
'    Set rs = Nothing
'    Set oDescObj = Nothing
'End Sub
Private Sub cmdExaminar1_Click()
    Dim sSqlObj As String
    Dim rsObj As New ADODB.Recordset
    Dim nRow As Integer
    If txtObj.Visible Then
        nRow = fgObj.Row
        If Mid(fgObj.TextMatrix(nRow, 8), 1, 2) = "00" Then
            Call frmBuscaCli.Inicia(Me, True)
            If Len(Trim(CodGrid)) > 0 Then
                txtObj.Visible = False
                cmdExaminar1.Visible = False
                SendKeys "{Down}", True
            Else
                txtObj.SetFocus
            End If
        Else
            sSqlObj = gcCentralCom & "spGetTreeObj '" & fgObj.TextMatrix(nRow, 8) & "', " _
               & Val(fgObj.TextMatrix(nRow, 5)) + Val(fgObj.TextMatrix(nRow, 6)) & ", '" _
               & fgObj.TextMatrix(nRow, 7) & "'"
            Set rsObj = dbCmact.Execute(sSqlObj)
            If RSVacio(rsObj) Then Exit Sub
            frmDescObjeto.Inicio rsObj, txtObj, Val(fgObj.TextMatrix(nRow, 5)) + Val(fgObj.TextMatrix(nRow, 6))
            If frmDescObjeto.lOk Then
                fgObj.TextMatrix(nRow, 2) = gaObj(0, 0, UBound(gaObj, 3) - 1)
                fgObj.TextMatrix(nRow, 3) = gaObj(0, 1, UBound(gaObj, 3) - 1)
                txtObj.Visible = False
                cmdExaminar1.Visible = False
                SendKeys "{Down}", True
            Else
                txtObj.SetFocus
            End If
            rsObj.Close
            Set rsObj = Nothing
        End If
    Else
        SSQL = "SELECT a.cCtaContCod as cObjetoCod, a.cCtaContDesc as cObjetoDesc " _
           & "FROM  " & gcCentralCom & "CtaCont a,  " & gcCentralCom & "OpeCta b WHERE b.cOpeCod = '" & gsOpeCod & "' and " _
           & "b.cOpeCtaDH = '" & lsDHTab & "' and a.cCtaContCod = b.cCtaContCod and b.cOpeCtaOpc='2' "
 
 
        Set rs = CargaRecord(SSQL)
        SSQL = "SELECT a.cCtaContCod as cObjetoCod, a.cCtaContDesc as cObjetoDesc " _
           & "FROM  " & gcCentralCom & "CtaCont a,  " & gcCentralCom & "OpeCta b WHERE b.cOpeCod = '" & gsOpeCod & "' and " _
           & "b.cOpeCtaDH = 'F' and a.cCtaContCod = b.cCtaContCod " _
        
           Set rsObj = CargaRecord(SSQL)
        If RSVacio(rs) Then
            rsObj.Close
            Set rsObj = Nothing
            Exit Sub
        End If
        frmSeleCuentas.Inicio IIf(lsDHTab = "D", rs, rsObj), IIf(lsDHTab = "D", rsObj, rs)
        If frmSeleCuentas.lOk Then
            AdicionaSeleccion frmSeleCuentas.psCtaCod, frmSeleCuentas.psCtaDesc, frmSeleCuentas.psDH
            RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.Row, 1), fgDetalle.TextMatrix(fgDetalle.Row, 0)
            txtCta.Visible = False
            cmdExaminar1.Visible = False
        End If
        rsObj.Close
        Set rsObj = Nothing
    End If
  
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub
Private Sub fgDetalle_DblClick()
    If Me.fgDetalle.TextMatrix(fgDetalle.Row, 0) <> "" Then
        If fgDetalle.Col = 1 Then
            EnfocaObj 0, txtCta
        Else
            EnfocaTexto1 0
        End If
    End If
End Sub
Private Sub fgDetalle_GotFocus()
    txtObj.Visible = False
    cmdExaminar1.Visible = False
End Sub

Private Sub fgDetalle_KeyPress(KeyAscii As Integer)
    If Me.fgDetalle.TextMatrix(Me.fgDetalle.Row, 0) <> "" Then
        If InStr(".0123456789", Chr(KeyAscii)) > 0 Then
            If fgDetalle.Col = 1 Then
                EnfocaObj KeyAscii, txtCta
            Else
                EnfocaTexto1 KeyAscii
            End If
        Else
            If KeyAscii = 13 Then
                If fgDetalle.Col = 1 Then
                    EnfocaObj 0, txtCta
                Else
                    EnfocaTexto1 0
                End If
            End If
        End If
    End If
End Sub

Private Sub fgDetalle_KeyUp(KeyCode As Integer, Shift As Integer)
    If (fgDetalle.TextMatrix(fgDetalle.Row, 5) = "D" And fgDetalle.Col = 3) Or _
       (fgDetalle.TextMatrix(fgDetalle.Row, 5) = "H" And fgDetalle.Col = 4) Then
        'KeyUp_Flex fgDetalle, KeyCode, Shift
        Sumas
    End If
End Sub

Private Sub fgDetalle_RowColChange()
    If Len(fgDetalle.TextMatrix(1, 1)) > 0 Then
        RefrescaFgObj fgDetalle.TextMatrix(fgDetalle.Row, 1), fgDetalle.TextMatrix(fgDetalle.Row, 0)
    End If
End Sub

Private Sub fgObj_DblClick()
    If Len(fgObj.TextMatrix(1, 1)) > 0 Then
        EnfocaObj 0, txtObj
        
        
    End If
End Sub

Private Sub fgObj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaObj 0, txtObj
    Else
        If KeyAscii <> 8 And KeyAscii <> 27 Then
            EnfocaObj KeyAscii, txtObj
        End If
    End If
End Sub

Private Sub Form_Activate()
    If lSalir Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Dim N As Integer
    Dim SSQL As String
    Dim rs As ADODB.Recordset
    
    
    Set rs = New ADODB.Recordset
    
    Me.Caption = gsOpeCod
    lSalir = False
    
    Set oConec = New DConecta
    If oConec.AbreConexion() = False Then
        lSalir = True
        Exit Sub
    End If
    Set oDCtasIf = New DCajaCtasIF
    Set oNCtasIf = New NCajaCtaIF
    
    oNCtasIf.Inicio gsFormatoFecha
    gsOpeCod = "401111"
    txtBuscarBanco.rs = oDCtasIf.CargaCtasIF(Mid(gsOpeCod, 3, 1), gTpoIFBanco, , 1, "1")
    txtBuscarBanco.psRaiz = "BANCOS"

    
    txtBuscarIFChq.rs = oDCtasIf.CargaCtasIF(Mid(gsOpeCod, 3, 1), gTpoIFBanco, , 1)
    txtBuscarIFChq.psRaiz = "BANCOS"
    
    txtBuscarCtaTrans.rs = oDCtasIf.CargaCtasIF(Mid(gsOpeCod, 3, 1), , gTpoCtaIFCtaCte & "-" & gTpoCtaIFCtaAho)
     
    Set oNContFunc = New NContFunciones
    
    TabDoc.Tab = 0
    txtCtaTpo = "Tipo Cuenta      :" & oNCtasIf.EmiteTipoCuentaIF(vsCtaTpo)
    
    txtOpeCod = gsOpeCod & "-"
    txtMovNro = oNContFunc.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    txtFecha.Text = gdFecSis

'    gbGeneraCal = False

    txtChequeImporte = "0.00"
    txtFechaRegChq = gdFecSis
    txtFechaValorChq = gdFecSis
    If Mid(gsOpeCod, 3, 1) = "2" Then  'Identificación de Tipo de Moneda
        If gnTipCambio = 0 Then
            If Not GetTipCambio(gdFecSis) Then
                lSalir = True
                Exit Sub
            End If
        End If
        FrameTipCambio.Visible = True
        txtTipCambio = Format(gnTipCambio, gcFormView)
    Else
        gsSimbolo = gcMN
    End If
    'Verificamos la asignación de Tipo Doc Voucher a Operación
    SSQL = "SELECT nDocTpo FROM OpeDoc WHERE cOpeCod = '" & gsOpeCod & "' and cOpeDocMetodo = '2' "
    Set rs = oConec.CargaRecordSet(SSQL)
    If Not rs.EOF Then
        'sDocTpoVoucher = rs!nDocTpo
    End If
    rs.Close
    Set rs = Nothing
    SSQL = "select Count(cCtacontcod) as TotalCtas,cOpeCtaDh from Opecta where cOpeCod ='" & gsOpeCod & "' " _
       & " group by cOpeCtaDh having count(*)=1"
    Set rs = oConec.CargaRecordSet(SSQL)
    If Not RSVacio(rs) Then
        If Not IsNull(rs!TotalCtas) Then
            lsDHUnic = Trim(rs!cOpeCtaDH)
            lsDHTab = IIf(lsDHUnic = "D", "H", "D")
        Else
            MsgBox " Pantalla no esta preparado para Nro de Cuentas contables en uno de sus rubros", vbInformation, "Aviso"
            lSalir = True
            Exit Sub
        End If
    Else
        MsgBox " Pantalla no esta preparado para Nro de Cuentas contables en uno de sus rubros", vbInformation, "Aviso"
        'lSalir = True
        Exit Sub
    End If
    If lsDHUnic = "H" Then
        cmdCalendario.Visible = True
    Else
        cmdCalendario.Visible = False
    End If

    'Cuenta unica ya sea en el debe  o en el haber
    'SSQL = "SELECT a.cCtaContCod, b.cCtaContDesc FROM OpeCta a JOIN " & gcCentralCom & "Ctacont b ON (b.cCtaContCod = a.cCtaContCod) WHERE cOpeCod = '" & gsOpeCod & "' and a.cOpeCtaDH = '" & lsDHUnic & "'"
    'Set rs = oConec.CargaRecordSet(SSQL)
    'If rs.EOF Then
    '    MsgBox "No se definió Cuenta del Debe en Operación", vbCritical, "Error"
    '    lSalir = True
    '    Exit Sub
    'End If
    'sCtaCod = rs!cCtaContCod
    'sCtaDes = rs!cCtaContDesc


    'SSQL = "SELECT * FROM " & gcCentralCom & " OpeObj WHERE cOpeCod ='" & gsOpeCod & "' order by cOpeObjOrden  "
    'Set rs = oConec.CargaRecordSet(SSQL)
    'N = 0
    'ReDim Preserve aObj(4, rs.RecordCount)
    'If Not RSVacio(rs) Then
    '    Do While Not rs.EOF
    '        N = N + 1
    '        aObj(1, N) = rs!cObjetoCod
    '        aObj(2, N) = rs!nOpeObjNiv
    '        aObj(3, N) = rs!cOpeObjOrden
    '        aObj(4, N) = rs!cOpeObjFiltro
     '       rs.MoveNext
      '  Loop
    'Else
    '    MsgBox "No se definieron los Objetos Necesarios para Operación ", vbInformation, "Aviso"
    '    lSalir = True
    '    rs.Close
    '    Set rs = Nothing
    '    Exit Sub
    'End If

    rs.Close
    Set rs = Nothing
   ' CargaFlex aObj(1, 2)

    CabeceraGridOtrasCtasContables
    CabGridObjetosOtrasCtas
    oConec.CierraConexion
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set oDCtasIf = Nothing
    
End Sub
Private Sub optDoc_Click(Index As Integer)
    Select Case Index
        Case 0
            cmdDoc.Caption = "Che&que"
        Case 1
            cmdDoc.Caption = "Carta Trans&ferencia"
    End Select
End Sub

Private Sub optDoc_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.cmdDoc.Enabled And Me.cmdDoc.Visible Then
            cmdDoc.SetFocus
        Else
            Me.cmdAceptar.SetFocus
        End If
    End If
End Sub

'Private Sub txtBancoCod_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    cmdExaBanco.SetFocus
'End If
'End Sub

Private Sub txtBancoImporte_Change()
    TotalImporte
End Sub
Private Sub txtBancoImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtBancoImporte, KeyAscii, 14, 2)
    If KeyAscii = 13 Then
        TotalImporte
        txtBancoImporte = Format(txtBancoImporte, gcFormView)
        If optDoc(0).Value Then
            optDoc(0).SetFocus
        Else
            optDoc(1).SetFocus
        End If
    End If
End Sub
Private Sub txtBancoImporte_LostFocus()
    If Val(txtBancoImporte) > 0 Then
        TotalImporte
        txtBancoImporte = Format(txtBancoImporte, gcFormView)
        If optDoc(0).Value Then
            optDoc(0).SetFocus
        Else
            optDoc(1).SetFocus
        End If
    End If
End Sub

Private Sub txtBuscarBanco_EmiteDatos()
If txtBuscarBanco.Text = "" Then
    Exit Sub
End If
'txtBuscarBanco.Text = txtBuscarBanco.Text & "." & Trim(oDCtasIf.GE(Trim(txtBuscarBanco.Text), Mid(gsOpeCod, 3, 1), vsCtaTpo))
txtObjetoDesc = "Inst. Financiera :" & txtBuscarBanco.psDescripcion

End Sub

Private Sub txtBuscarCtaTrans_EmiteDatos()
txtBancoCta = Trim(oNCtasIf.EmiteTipoCuentaIF(Mid(txtBuscarCtaTrans.Text, 15, 2))) & " " & txtBuscarCtaTrans.psDescripcion
txtBancoDesc = oDCtasIf.NombreIF(Mid(txtBuscarCtaTrans.Text, 1, 13))
End Sub
Private Sub txtBuscarIFChq_EmiteDatos()
txtDescBcoChq = Me.txtBuscarIFChq.psDescripcion
End Sub
Private Sub txtCelda_KeyPress(KeyAscii As Integer)
    Dim nMonto As Currency
    KeyAscii = NumerosDecimales(txtCelda, KeyAscii, 10, 2)
    If KeyAscii = 13 Then
        nMonto = Val(Right(fgBillete.TextMatrix(fgBillete.Row, 1), 6))
        If Residuo(Val(txtCelda), nMonto) Then
            fgBillete.Text = Format(txtCelda.Text, gcFormView)
            If fgBillete.Row < fgBillete.Rows - 1 Then
                fgBillete.Row = fgBillete.Row + 1
                If fgBillete.Row - fgBillete.TopRow >= 4 Then
                    fgBillete.TopRow = fgBillete.TopRow + 1
                End If
            End If
        End If
        Sumas
        txtCelda.Visible = False
        txtCelda = ""
    End If
End Sub
Private Sub txtCelda_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then
        txtCelda_KeyPress 13
        SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
    End If
End Sub
Private Sub txtCelda_LostFocus()
    txtCelda.Text = ""
    txtCelda.Visible = False
End Sub
Private Sub txtCelda1_KeyPress(KeyAscii As Integer)
    Dim nMonto As Currency
    KeyAscii = intfNumDec(txtCelda1, KeyAscii, 15, 2)
    If KeyAscii = 13 Then
        fgDetalle.Text = Format(txtCelda1.Text, gcFormView)
        suma
        txtCelda1.Visible = False
    End If
End Sub
Private Sub txtCelda1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then
        txtCelda1_KeyPress 13
        SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
    End If
End Sub
Private Sub txtCelda1_LostFocus()
    txtCelda1.Text = ""
    txtCelda1.Visible = False
End Sub

Private Sub txtChequeImporte_Change()
    TotalImporte
End Sub
Private Sub txtChequeImporte_GotFocus()
    fEnfoque txtChequeImporte
End Sub
Private Sub txtChequeImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtChequeImporte, KeyAscii, 14, 2)
    If KeyAscii = 13 Then
        If Val(txtChequeImporte) > 0 Then
            TotalImporte
            cmdAceptar.Enabled = True
            cmdAceptar.SetFocus
        End If
    End If
End Sub
Private Sub txtChequeImporte_LostFocus()
    If txtChequeImporte = "" Then txtChequeImporte = 0
    txtChequeImporte = Format(txtChequeImporte, "#####,####0.00")
    If Val(Me.txtChequeImporte) > 0 Then
        TotalImporte
        cmdAceptar.Enabled = True
        cmdAceptar.SetFocus
    End If
End Sub

'Private Sub txtCodBcoChq_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        txtCtaBanco.SetFocus
'    End If
'End Sub

Private Sub txtCta_GotFocus()
    ActivaExaminar txtCta
  
End Sub

Private Sub txtCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCta = ""
        cmdExaminar1.SetFocus
        'If Me.TxtCta = "" Then
        '   TxtCta.SetFocus
        'Else
        '   TxtCta.Visible = False
        '   cmdExaminar1.Visible = False
        '   SendKeys "{Down}", True
        'End If
    End If
End Sub

Private Sub txtCta_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then
        txtCta_KeyPress 13
        SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
    End If
End Sub

Private Sub txtCta_LostFocus()
    DesactivaExaminar txtCta
End Sub

Private Sub txtCtaBanco_GotFocus()
    fEnfoque txtCtaBanco
End Sub

Private Sub txtCtaBanco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNumChq.SetFocus
    End If
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValFecha(txtFecha) = False Then Exit Sub
        txtMovNro = oMov.GeneraMovNro(CDate(txtFecha))
        txtFechaRegChq = txtFecha
        txtFechaValorChq = txtFecha
        'If txtObjetoCod.Enabled Then
        '    txtObjetoCod.SetFocus
        'Else
            txtMovDesc.SetFocus
        'End If
    End If
End Sub
Private Sub txtFecha_LostFocus()
    'txtMovNro = GeneraMovNro(, , Me.txtFecha)
    txtFechaRegChq = txtFecha
    txtFechaValorChq = txtFecha
End Sub

Private Sub txtFechaRegChq_GotFocus()
    fEnfoque txtFechaRegChq
End Sub
Private Sub txtFechaRegChq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValFecha(Me.txtFechaRegChq) Then
            Me.txtFechaValorChq.SetFocus
        End If
    End If
End Sub

Private Sub txtFechaValorChq_GotFocus()
    fEnfoque txtFechaValorChq
End Sub

Private Sub txtFechaValorChq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ValFecha(Me.txtFechaValorChq) = False Then Exit Sub
        chkPlaza.SetFocus
    End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Select Case TabDoc.Tab
            Case 0:
                If cmdExaBcoChq.Enabled Then
                    cmdExaBcoChq.SetFocus
                End If
            Case 1
                If Me.txtBancoCod.Enabled Then
                    txtBancoCod.SetFocus
                Else
                    txtBancoImporte.SetFocus
                End If
            Case 2
                If Me.fgBillete.Enabled Then
                    fgBillete.SetFocus
                End If
            Case 3
                If Me.cmdAgregarCta.Enabled Then
                    cmdAgregarCta.SetFocus
                End If
        End Select
    End If
End Sub

Private Sub txtNumChq_GotFocus()
    fEnfoque txtNumChq
End Sub

Private Sub txtNumChq_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaRegChq.SetFocus
    End If
End Sub

Private Sub txtObj_GotFocus()
    ActivaExaminar txtObj
End Sub
Private Sub txtObj_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 40 Or KeyCode = 38 Then
        txtObj_KeyPress 13
        SendKeys IIf(KeyCode = 38, "{Up}", "{Down}"), True
    End If
End Sub
Private Sub txtObj_LostFocus()
    DesactivaExaminar txtObj
End Sub
Private Sub txtObj_Validate(Cancel As Boolean)
    If Not ValidaObj(txtObj) Then
        Cancel = True
    End If
End Sub
Private Sub txtObj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidaObj(txtObj) Then
            txtObj.SetFocus
        Else
            txtObj.Visible = False
            cmdExaminar1.Visible = False
            If fgObj.Row < fgObj.Rows - 1 Then
                fgObj.Row = fgObj.Row + 1
            End If
        End If
    End If
End Sub

Private Sub txtObjetoCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMovDesc.SetFocus
    End If
End Sub

Private Sub txtObjetoCod_Validate(Cancel As Boolean)
    'If Not ValidaBanco(txtObjetoCod, 0) Then
    '   Cancel = True
    'End If
End Sub

Private Sub txtPropietario_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        txtMovDesc.SetFocus
    End If
End Sub
Private Sub ActualizaTot(ByVal Item As MSComctlLib.ListItem)
'    If Item.Checked Then
'        txtChequeImporte = Format(Val(Format(txtChequeImporte, gcFormDato)) + Val(Format(Item.SubItems(4), gcFormDato)), gcFormView)
'    Else
'        txtChequeImporte = Format(Val(Format(txtChequeImporte, gcFormDato)) - Val(Format(Item.SubItems(4), gcFormDato)), gcFormView)
'    End If
'    TotalImporte
End Sub
Private Sub ActivaExaminar(txt As TextBox)
    txt.Width = txt.Width - cmdExaminar1.Width + 25
    cmdExaminar1.Top = txt.Top + 10
    cmdExaminar1.Left = txt.Left + txt.Width - 15
    cmdExaminar1.Visible = True
    cmdExaminar1.TabIndex = txt.TabIndex + 1
End Sub
Private Sub DesactivaExaminar(txt As TextBox)
    txt.Width = txt.Width + cmdExaminar1.Width - 25
End Sub
Private Sub AdicionaSeleccion(sCod As String, sDes As String, sDH As String)
    Dim nRow As Integer
    SSQL = "Select a.cObjetoCod, b.cObjetoDesc, a.cCtaObjOrden, a.nCtaObjNiv, a.cCtaObjFiltro, b.nObjetoNiv, a.cCtaObjImpre " _
       & "from  " & gcCentralCom & "CtaObj as a,  " & gcCentralCom & "Objeto as b where a.cCtaContCod = '" & sCod & "' " _
       & "and a.cObjetoCod = b.cObjetoCod"

    If rs.State = adStateOpen Then rs.Close: Set rs = Nothing
    rs.Open SSQL, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        AdicionaCuenta sCod, sDes, sDH
        nRow = fgObj.Row
        AdicionaObj sCod, fgDetalle.Row, rs
        RefrescaFgObj sCod, fgDetalle.Rows - 1
        fgObj.Row = IIf(fgObj.Rows = 2, 1, nRow + 1)
        fgObj.Col = 2
        AsignaEnabled False, True
        fgObj.SetFocus
    Else
        AdicionaCuenta sCod, sDes, sDH
        fgDetalle.TopRow = fgDetalle.Row
        If fgDetalle.Enabled Then
            fgDetalle.Row = fgDetalle.Row
            fgDetalle.Col = 3
            fgDetalle.SetFocus
        End If
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub AdicionaCuenta(sCod As String, sDes As String, sDH As String)
    Dim nItem As Integer
    nItem = fgDetalle.Row
    fgDetalle.TextMatrix(nItem, 1) = sCod
    fgDetalle.TextMatrix(nItem, 2) = sDes
    fgDetalle.TextMatrix(nItem, 5) = sDH
    fgDetalle.Col = 3
    EliminaFila nItem
End Sub

Private Sub AdicionaObj(sCodCta As String, nFila As Integer, rs As ADODB.Recordset)
    Dim nItem As Integer
    Do While Not rs.EOF
        AdicionaRow fgObj
        nItem = fgObj.Row
        fgObj.TextMatrix(nItem, 0) = nFila
        fgObj.TextMatrix(nItem, 1) = rs!cCtaObjOrden
        ' Fila 2 y 3 para Instancia de Objeto
        fgObj.TextMatrix(nItem, 4) = sCodCta
        fgObj.TextMatrix(nItem, 5) = rs!nCtaObjNiv
        fgObj.TextMatrix(nItem, 6) = rs!nObjetoNiv
        fgObj.TextMatrix(nItem, 7) = rs!cCtaObjFiltro
        fgObj.TextMatrix(nItem, 8) = rs!cObjetoCod
        fgObj.TextMatrix(nItem, 9) = rs!cCtaObjImpre
        rs.MoveNext
    Loop
End Sub

Private Sub AsignaEnabled(lOp As Boolean, lisobj As Boolean)
    fgDetalle.Enabled = lOp
    txtMovDesc.Enabled = lOp
    cmdAgregarCta.Enabled = lOp
    cmdEliminar.Enabled = lOp
    If lisobj Then
        Me.fgObj.Enabled = Not lOp
        cmdAceptarObj.Visible = Not lOp
        cmdCancelarObjeto.Visible = Not lOp
    End If
End Sub
Private Sub EnfocaTexto1(KeyAscii As Integer)
    Dim nx, ny As Integer, N As Integer
    txtCelda1 = ""
    If fgDetalle.Col = 3 Then
        nx = fgDetalle.Left + fgDetalle.ColPos(fgDetalle.Col)
        ny = fgDetalle.Top + fgDetalle.RowPos(fgDetalle.Row)
        If KeyAscii = 0 Then
            txtCelda1 = fgDetalle.Text
            txtCelda1.SelStart = 0
            txtCelda1.SelLength = Len(txtCelda)
        Else
            txtCelda1 = Chr(KeyAscii)
            txtCelda1.SelStart = 1
        End If
        txtCelda1.Left = nx + 15
        txtCelda1.Top = ny + 15
        txtCelda1.Visible = True
        txtCelda1.SetFocus
    End If
End Sub
Private Sub EnfocaObj(KeyAscii As Integer, txt As TextBox)
    Dim nx, ny As Integer
    Dim w As Integer
    Dim sText As String
    If (txt.Name = "txtObj" And fgObj.Col = 2) Or _
       (txt.Name = "txtCta" And fgDetalle.Col = 1) Then
        txt = ""
        If txt.Name = "txtObj" Then
            nx = fgObj.Left + fgObj.ColPos(fgObj.Col)
            ny = fgObj.Top + fgObj.RowPos(fgObj.Row)
            w = fgObj.CellWidth
            sText = fgObj.Text
        Else
            nx = fgDetalle.Left + fgDetalle.ColPos(fgDetalle.Col)
            ny = fgDetalle.Top + fgDetalle.RowPos(fgDetalle.Row)
            w = fgDetalle.CellWidth
            sText = fgDetalle.Text
        End If
        If KeyAscii = 0 Then
            txt = sText
            txt.SelStart = 0
            txt.SelLength = Len(txt)
        Else
            txt = Chr(KeyAscii)
            txt.SelStart = 1
        End If
        txt.Left = nx + 15
        txt.Top = ny + 15
        txt.Width = w
        txt.Visible = True
        txt.SetFocus
    End If
End Sub
Private Function ValidaObj(sCod As String) As Boolean
    Dim rsObj As New ADODB.Recordset
    Dim sSqlObj As String
    Dim nRow As Integer
    ValidaObj = False
    nRow = fgObj.Row
    If Len(Trim(txtObj)) = 0 Then
        Exit Function
    End If
    sSqlObj = gcCentralCom & "spExistObj '" & txtObj & "', " _
       & Val(fgObj.TextMatrix(nRow, 5)) + Val(fgObj.TextMatrix(nRow, 6)) & ", '" _
       & fgObj.TextMatrix(nRow, 7) & "'"
    
    Set rsObj = dbCmact.Execute(sSqlObj)
    If Not RSVacio(rsObj) Then
        GetTreeObjSele rsObj!cObjetoCod, fgObj.TextMatrix(1, 6)
        fgObj.TextMatrix(nRow, 2) = gaObj(0, 0, UBound(gaObj, 3) - 1)
        fgObj.TextMatrix(nRow, 3) = gaObj(0, 1, UBound(gaObj, 3) - 1)
        ValidaObj = True
    Else
        MsgBox " Dato no Encontrado...! ", vbCritical, "Aviso"
    End If
    rsObj.Close
End Function

Private Sub CabeceraGridOtrasCtasContables()
 '   frmOpeGeneral.Caption = "Operación: " & gsOpeCod
    fgDetalle.Clear
    fgDetalle.Rows = 2

    fgDetalle.TextMatrix(0, 0) = " #"
    fgDetalle.TextMatrix(0, 1) = "Cuenta"
    fgDetalle.TextMatrix(0, 2) = "Descripción"
    fgDetalle.TextMatrix(0, 3) = "Monto"
    fgDetalle.ColWidth(0) = 350
    fgDetalle.ColWidth(1) = 1100
    fgDetalle.ColWidth(2) = 2800
    fgDetalle.ColWidth(3) = 1200
    fgDetalle.ColWidth(4) = 0
    fgDetalle.ColWidth(5) = 0
    fgDetalle.ColAlignment(1) = 1
    fgDetalle.ColAlignmentFixed(1) = 1
    fgDetalle.ColAlignment(3) = 7
    fgDetalle.ColAlignmentFixed(3) = 7
    fgDetalle.ColAlignment(4) = 7
    fgDetalle.ColAlignmentFixed(4) = 7
End Sub

Private Sub CabGridObjetosOtrasCtas()
    fgObj.Clear
    fgObj.Rows = 2
    fgObj.TextMatrix(0, 0) = " #"
    fgObj.TextMatrix(0, 1) = "Ord"
    fgObj.TextMatrix(0, 2) = "Código"
    fgObj.TextMatrix(0, 3) = "Descripción"
    fgObj.ColWidth(0) = 350
    fgObj.ColWidth(1) = 420
    fgObj.ColWidth(2) = 1700
    fgObj.ColWidth(3) = 3500
    fgObj.ColWidth(4) = 0
    fgObj.ColWidth(5) = 0
    fgObj.ColWidth(6) = 0
    fgObj.ColWidth(7) = 0
    fgObj.ColWidth(8) = 0
    fgObj.ColWidth(9) = 0
    fgObj.ColAlignment(1) = 7
    fgObj.ColAlignment(2) = 1
  
End Sub
Private Sub EliminaFila(lnItem As Integer)
    Dim i As Integer
    For i = 0 To Me.fgObj.Rows - 1
        If Me.fgObj.TextMatrix(i, 0) = Trim(Str(lnItem)) Then
            EliminaRow fgObj, i
            Exit For
        End If
    Next
End Sub
Public Sub Inicio(psDHUnic As String, psDHTab As String, psTpoCtaIf As String)
    lsDHUnic = psDHUnic
    lsDHTab = psDHTab
    vsCtaTpo = psTpoCtaIf
    Me.Show 1
End Sub
Private Function Valida() As Boolean
Valida = True
    If txtObjetoCod = "" Then
        Valida = False
        MsgBox "No se indicó entidad financiera donde se apertura Cuenta ", vbInformation, "Error"
        txtObjetoCod.SetFocus
        Exit Function
    End If
    If Len(Trim(txtMovDesc)) = 0 Then
        Valida = False
        MsgBox "Es necesario ingresar Glosa de Operación", vbInformation, "Error"
        txtMovDesc.SetFocus
        Exit Function
    End If
    If Val(txtChequeImporte) > 0 Then
        If txtCodBcoChq = "" Then
            Valida = False
            MsgBox "Objeto Cheque no Válido", vbInformation, "Aviso"
            txtCodBcoChq.SetFocus
            Exit Function
        End If
        If txtCtaBanco = "" Then
            Valida = False
            MsgBox "Cuenta de Banco no Válida", vbInformation, "Aviso"
            txtCtaBanco.SetFocus
            Exit Function
        End If
        If txtNumChq = "" Then
            Valida = False
            MsgBox "Ingrese Numero de Cheque", vbInformation, "Aviso"
            txtNumChq.SetFocus
            Exit Function
        End If
        If ValFecha(Me.txtFechaRegChq) = False Then
            Valida = False
            Exit Function
        End If
        If ValFecha(txtFechaValorChq) = False Then
            Valida = False
            Exit Function
        End If
    End If
    If Val(txtTotalOtrasCtas) > 0 Then
        For i = 1 To Me.fgDetalle.Rows - 1
            If Val(fgDetalle.TextMatrix(i, 3)) = 0 Then
                Valida = False
                MsgBox "Falta Ingresar Monto en Cuenta contable", vbInformation, "Aviso"
                Me.TabDoc.Tab = 3
                Exit Function
            End If
        Next
    Else
        If Me.fgDetalle.TextMatrix(1, 0) <> "" Then
            Valida = False
            MsgBox "Ingrese Monto de Cuenta contable", vbInformation, "Aviso"
            Me.TabDoc.Tab = 3
            Exit Function
        End If
    End If
End Function
