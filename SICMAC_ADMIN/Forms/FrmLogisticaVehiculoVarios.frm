VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmLogisticavehiculoVarios 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8970
   Icon            =   "FrmLogisticaVehiculoVarios.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Height          =   735
      Left            =   60
      TabIndex        =   40
      Top             =   6150
      Width           =   8820
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   405
         Left            =   7440
         TabIndex        =   41
         Top             =   210
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   8805
      Begin VB.PictureBox Picture1 
         Height          =   525
         Left            =   7605
         Picture         =   "FrmLogisticaVehiculoVarios.frx":08CA
         ScaleHeight     =   465
         ScaleWidth      =   480
         TabIndex        =   55
         Top             =   390
         Width           =   540
      End
      Begin VB.CommandButton CmdBuscarPlaca 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   6060
         TabIndex        =   3
         Top             =   210
         Width           =   1035
      End
      Begin VB.TextBox TxtPlaca 
         Height          =   315
         Left            =   4635
         TabIndex        =   2
         Top             =   255
         Width           =   1245
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   7425
         X2              =   8400
         Y1              =   1500
         Y2              =   1500
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   7425
         X2              =   8400
         Y1              =   1005
         Y2              =   1005
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         Caption         =   "Control de Vehiculos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   7425
         TabIndex        =   56
         Top             =   1050
         Width           =   1005
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000B&
         BorderColor     =   &H000000C0&
         BorderWidth     =   4
         FillColor       =   &H000080FF&
         Height          =   1425
         Left            =   7245
         Top             =   255
         Width           =   1380
      End
      Begin VB.Label LblEstado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   4635
         TabIndex        =   29
         Top             =   1005
         Width           =   2500
      End
      Begin VB.Label LblBSSerie 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4635
         TabIndex        =   28
         Top             =   1350
         Width           =   2500
      End
      Begin VB.Label LblBSCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   990
         TabIndex        =   27
         Top             =   1350
         Width           =   2595
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "BSCod"
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
         Left            =   180
         TabIndex        =   26
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Marca"
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
         Left            =   180
         TabIndex        =   25
         Top             =   1035
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "BSSerie"
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
         Left            =   3870
         TabIndex        =   12
         Top             =   1425
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Modelo"
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
         Left            =   180
         TabIndex        =   11
         Top             =   660
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Placa"
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
         Left            =   3870
         TabIndex        =   10
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo "
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
         Left            =   180
         TabIndex        =   9
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Left            =   3900
         TabIndex        =   8
         Top             =   1035
         Width           =   600
      End
      Begin VB.Label LblMarca 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1005
         TabIndex        =   7
         Top             =   1005
         Width           =   2595
      End
      Begin VB.Label LblModelo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1005
         TabIndex        =   6
         Top             =   645
         Width           =   2595
      End
      Begin VB.Label LblTipo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1005
         TabIndex        =   5
         Top             =   270
         Width           =   2595
      End
      Begin VB.Label LblPlaca 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000001&
         Height          =   300
         Left            =   4635
         TabIndex        =   4
         Top             =   645
         Width           =   2500
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   4200
      Left            =   120
      TabIndex        =   0
      Top             =   1905
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7408
      _Version        =   393216
      TabOrientation  =   1
      Tabs            =   6
      Tab             =   4
      TabsPerRow      =   6
      TabHeight       =   476
      TabCaption(0)   =   "SOAT"
      TabPicture(0)   =   "FrmLogisticaVehiculoVarios.frx":1194
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(2)=   "fraFecha"
      Tab(0).Control(3)=   "MSHSOAT"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "KILOMETRAJE"
      TabPicture(1)   =   "FrmLogisticaVehiculoVarios.frx":11B0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSHKM"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame6"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "CARGA"
      TabPicture(2)   =   "FrmLogisticaVehiculoVarios.frx":11CC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame10"
      Tab(2).Control(1)=   "Frame9"
      Tab(2).Control(2)=   "Frame8"
      Tab(2).Control(3)=   "MSHCarga"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "PAPELETA"
      TabPicture(3)   =   "FrmLogisticaVehiculoVarios.frx":11E8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label20"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "MSHPapeleta"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame11"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame12"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "TxtDescripcionPapeleta"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Frame13"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "INCIDENCIA"
      TabPicture(4)   =   "FrmLogisticaVehiculoVarios.frx":1204
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label21"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label24"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "MSHIncidencia"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Frame14"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "Frame15"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "Frame16"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Frame17"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "TxtDescripcionIncidencia"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "TxtLugarInsidencia"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).ControlCount=   9
      TabCaption(5)   =   "ASIGNACION"
      TabPicture(5)   =   "FrmLogisticaVehiculoVarios.frx":1220
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "MSHAsigna"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "FraAsigna"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "FraDesasigna"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).ControlCount=   3
      Begin VB.Frame FraDesasigna 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   -70800
         TabIndex        =   97
         Top             =   2040
         Width           =   4455
         Begin VB.Frame Frame22 
            Appearance      =   0  'Flat
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
            Height          =   660
            Left            =   120
            TabIndex        =   101
            Top             =   240
            Width           =   4185
            Begin MSMask.MaskEdBox TxtFechaFinAsignacion 
               Height          =   300
               Left            =   3000
               TabIndex        =   102
               Top             =   240
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   529
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Fin"
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
               TabIndex        =   103
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Frame Frame21 
            Appearance      =   0  'Flat
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
            Height          =   615
            Left            =   1320
            TabIndex        =   98
            Top             =   960
            Width           =   2985
            Begin VB.CommandButton CmdFechaFin 
               Caption         =   "&Cancelar"
               Height          =   360
               Left            =   210
               TabIndex        =   100
               Top             =   180
               Width           =   1140
            End
            Begin VB.CommandButton CmdGrabaFin 
               Caption         =   "&Grabar"
               Height          =   360
               Left            =   1605
               TabIndex        =   99
               Top             =   180
               Width           =   1140
            End
         End
      End
      Begin VB.Frame FraAsigna 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   1935
         Left            =   -70800
         TabIndex        =   88
         Top             =   50
         Width           =   4455
         Begin VB.Frame Frame18 
            Appearance      =   0  'Flat
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
            Height          =   615
            Left            =   1440
            TabIndex        =   94
            Top             =   1200
            Width           =   2985
            Begin VB.CommandButton cmdGrabarAsigna 
               Caption         =   "&Grabar"
               Height          =   360
               Left            =   1605
               TabIndex        =   96
               Top             =   180
               Width           =   1140
            End
            Begin VB.CommandButton cmdCancelarAsigna 
               Caption         =   "&Cancelar"
               Height          =   360
               Left            =   210
               TabIndex        =   95
               Top             =   180
               Width           =   1140
            End
         End
         Begin VB.ComboBox CboConductor 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   840
            Width           =   4305
         End
         Begin VB.Frame Frame19 
            Appearance      =   0  'Flat
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
            Height          =   660
            Left            =   1200
            TabIndex        =   90
            Top             =   120
            Width           =   3225
            Begin MSMask.MaskEdBox TxtFechaAsigna 
               Height          =   300
               Left            =   2040
               TabIndex        =   91
               Top             =   240
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   529
               _Version        =   393216
               Appearance      =   0
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "dd/mm/yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Fecha de Asignacion"
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
               TabIndex        =   92
               Top             =   240
               Width           =   1800
            End
         End
         Begin VB.CheckBox chkAsigna 
            Caption         =   "Asignar Vehiculo"
            Height          =   495
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtLugarInsidencia 
         Height          =   420
         Left            =   5670
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   87
         Top             =   1680
         Width           =   2910
      End
      Begin VB.TextBox TxtDescripcionIncidencia 
         Height          =   420
         Left            =   5670
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   86
         Top             =   2400
         Width           =   2910
      End
      Begin VB.Frame Frame17 
         Appearance      =   0  'Flat
         Caption         =   "Tipo de Incidencia"
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
         Height          =   960
         Left            =   5670
         TabIndex        =   80
         Top             =   45
         Width           =   3045
         Begin VB.CommandButton CmdNuevo 
            Caption         =   "&Nuevo"
            Height          =   345
            Left            =   1965
            TabIndex        =   82
            Top             =   540
            Width           =   990
         End
         Begin VB.ComboBox CboIncidencia 
            Height          =   315
            Left            =   45
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   225
            Width           =   2985
         End
      End
      Begin VB.Frame Frame16 
         Appearance      =   0  'Flat
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
         Height          =   510
         Left            =   5670
         TabIndex        =   77
         Top             =   2775
         Width           =   2985
         Begin VB.TextBox TxtMontoIncidencia 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1890
            TabIndex        =   78
            Top             =   150
            Width           =   1005
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Monto S/."
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
            Left            =   105
            TabIndex        =   79
            Top             =   210
            Width           =   870
         End
      End
      Begin VB.Frame Frame15 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   5670
         TabIndex        =   74
         Top             =   3240
         Width           =   2985
         Begin VB.CommandButton CmdGrabarIncidencia 
            Caption         =   "&Grabar"
            Height          =   360
            Left            =   1605
            TabIndex        =   76
            Top             =   180
            Width           =   1140
         End
         Begin VB.CommandButton CmdCancelarIncidencia 
            Caption         =   "&Cancelar"
            Height          =   360
            Left            =   210
            TabIndex        =   75
            Top             =   180
            Width           =   1140
         End
      End
      Begin VB.Frame Frame14 
         Appearance      =   0  'Flat
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
         Height          =   540
         Left            =   5670
         TabIndex        =   71
         Top             =   945
         Width           =   2985
         Begin MSMask.MaskEdBox TxtFechaIncidencia 
            Height          =   300
            Left            =   1785
            TabIndex        =   72
            Top             =   165
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
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
            TabIndex        =   73
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.Frame Frame13 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   -69330
         TabIndex        =   66
         Top             =   3240
         Width           =   2985
         Begin VB.CommandButton CmdCancelarPapeleta 
            Caption         =   "&Cancelar"
            Height          =   360
            Left            =   210
            TabIndex        =   68
            Top             =   180
            Width           =   1140
         End
         Begin VB.CommandButton CmdGrabarPapeleta 
            Caption         =   "&Grabar"
            Height          =   360
            Left            =   1605
            TabIndex        =   67
            Top             =   180
            Width           =   1140
         End
      End
      Begin VB.TextBox TxtDescripcionPapeleta 
         Height          =   540
         Left            =   -69330
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   2205
         Width           =   2910
      End
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
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
         Height          =   510
         Left            =   -69330
         TabIndex        =   61
         Top             =   2730
         Width           =   2985
         Begin VB.TextBox TxtMontoPapeleta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1890
            TabIndex        =   62
            Top             =   150
            Width           =   1005
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Monto Papeleta S/."
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
            Left            =   105
            TabIndex        =   63
            Top             =   210
            Width           =   1680
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
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
         Height          =   540
         Left            =   -69330
         TabIndex        =   58
         Top             =   1275
         Width           =   2985
         Begin MSMask.MaskEdBox TxtFechaPapeleta 
            Height          =   300
            Left            =   1785
            TabIndex        =   59
            Top             =   165
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
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
            TabIndex        =   60
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
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
         Height          =   2550
         Left            =   -69330
         TabIndex        =   49
         Top             =   735
         Width           =   2985
         Begin VB.TextBox TxtDestinoDesc 
            Height          =   540
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Top             =   1920
            Width           =   2910
         End
         Begin VB.TextBox TxtDescripcion 
            Height          =   540
            Left            =   30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   51
            Top             =   1080
            Width           =   2910
         End
         Begin VB.ComboBox CboAgencia 
            Height          =   315
            Left            =   45
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   375
            Width           =   2880
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Agencia"
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
            TabIndex        =   54
            Top             =   150
            Width           =   705
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion de la Carga"
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
            Left            =   75
            TabIndex        =   53
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Descripcion Destino"
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
            Left            =   75
            TabIndex        =   52
            Top             =   1695
            Width           =   1725
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   -69330
         TabIndex        =   46
         Top             =   3240
         Width           =   2985
         Begin VB.CommandButton CmdGrabaCarga 
            Caption         =   "&Grabar"
            Height          =   360
            Left            =   1605
            TabIndex        =   48
            Top             =   180
            Width           =   1140
         End
         Begin VB.CommandButton CmdCancelarCarga 
            Caption         =   "&Cancelar"
            Height          =   360
            Left            =   210
            TabIndex        =   47
            Top             =   180
            Width           =   1140
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
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
         Height          =   540
         Left            =   -69330
         TabIndex        =   43
         Top             =   195
         Width           =   2985
         Begin MSMask.MaskEdBox TxtFechaCarga 
            Height          =   300
            Left            =   1785
            TabIndex        =   44
            Top             =   165
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
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
            TabIndex        =   45
            Top             =   225
            Width           =   600
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
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
         Height          =   645
         Left            =   -69330
         TabIndex        =   37
         Top             =   810
         Width           =   2985
         Begin MSMask.MaskEdBox txtFEchaKM 
            Height          =   300
            Left            =   1785
            TabIndex        =   38
            Top             =   240
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Inicio"
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
            TabIndex        =   39
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
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
         Height          =   975
         Left            =   -69330
         TabIndex        =   34
         Top             =   1440
         Width           =   2985
         Begin VB.TextBox TxtKMFin 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1500
            TabIndex        =   104
            Top             =   600
            Width           =   1380
         End
         Begin VB.TextBox TxtKM 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1500
            TabIndex        =   35
            Top             =   225
            Width           =   1380
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Tacometro Fin"
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
            Left            =   15
            TabIndex        =   105
            Top             =   600
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tacometro Ini"
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
            Left            =   15
            TabIndex        =   36
            Top             =   285
            Width           =   1185
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   -69330
         TabIndex        =   31
         Top             =   3240
         Width           =   2985
         Begin VB.CommandButton CmdCancelarKM 
            Caption         =   "&Cancelar"
            Height          =   360
            Left            =   210
            TabIndex        =   33
            Top             =   180
            Width           =   1140
         End
         Begin VB.CommandButton CmdGrabarKm 
            Caption         =   "&Grabar"
            Height          =   360
            Left            =   1605
            TabIndex        =   32
            Top             =   180
            Width           =   1140
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   -69330
         TabIndex        =   22
         Top             =   3240
         Width           =   2985
         Begin VB.CommandButton CmdGrabarSoat 
            Caption         =   "&Grabar"
            Height          =   360
            Left            =   1605
            TabIndex        =   24
            Top             =   180
            Width           =   1140
         End
         Begin VB.CommandButton CmdCancelarSoat 
            Caption         =   "&Cancelar"
            Height          =   360
            Left            =   210
            TabIndex        =   23
            Top             =   180
            Width           =   1140
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
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
         Height          =   615
         Left            =   -69330
         TabIndex        =   19
         Top             =   2640
         Width           =   2985
         Begin VB.TextBox TxtSoat 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1875
            TabIndex        =   20
            Top             =   225
            Width           =   1005
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Costo SOAT  $"
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
            TabIndex        =   21
            Top             =   285
            Width           =   1275
         End
      End
      Begin VB.Frame fraFecha 
         Appearance      =   0  'Flat
         Caption         =   "Fechas"
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
         Height          =   1080
         Left            =   -69330
         TabIndex        =   14
         Top             =   1560
         Width           =   2985
         Begin MSMask.MaskEdBox txtFechaF 
            Height          =   300
            Left            =   1875
            TabIndex        =   15
            Top             =   690
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   300
            Left            =   1875
            TabIndex        =   16
            Top             =   300
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Vencimiento"
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
            Left            =   45
            TabIndex        =   18
            Top             =   735
            Width           =   1905
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Inicio"
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
            Left            =   60
            TabIndex        =   17
            Top             =   345
            Width           =   1335
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHSOAT 
         Height          =   3750
         Left            =   -74925
         TabIndex        =   13
         Top             =   90
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6615
         _Version        =   393216
         Cols            =   4
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHKM 
         Height          =   3750
         Left            =   -74925
         TabIndex        =   30
         Top             =   90
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6615
         _Version        =   393216
         Cols            =   3
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHCarga 
         Height          =   3750
         Left            =   -74925
         TabIndex        =   42
         Top             =   90
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6615
         _Version        =   393216
         Cols            =   5
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHPapeleta 
         Height          =   3750
         Left            =   -74925
         TabIndex        =   57
         Top             =   90
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6615
         _Version        =   393216
         Cols            =   4
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHIncidencia 
         Height          =   3750
         Left            =   60
         TabIndex        =   69
         Top             =   90
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   6615
         _Version        =   393216
         Cols            =   5
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHAsigna 
         Height          =   3750
         Left            =   -74940
         TabIndex        =   84
         Top             =   90
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   6615
         _Version        =   393216
         Cols            =   3
         SelectionMode   =   1
         AllowUserResizing=   2
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Lugar de Insidencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   5670
         TabIndex        =   83
         Top             =   1500
         Width           =   1695
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion Incidencia"
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
         Left            =   5670
         TabIndex        =   70
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion Papeleta"
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
         Left            =   -69330
         TabIndex        =   65
         Top             =   1950
         Width           =   1830
      End
   End
End
Attribute VB_Name = "FrmLogisticavehiculoVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CargaComboConductorLibre()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset
Set LV = New DLogvehiculo
Set Rs = LV.GetConductorLibre
Me.CboConductor.Clear
While Not Rs.EOF
    Me.CboConductor.AddItem Rs!cPersNombre & Space(100) & Rs!cPersCod
    Rs.MoveNext
Wend
Set Rs = Nothing
Set LV = Nothing
End Sub

Sub MarcoAsigna()
With Me.MSHAsigna
    .Rows = 2
    .Clear
    .ColWidth(0) = 300
    .ColWidth(1) = 1200
    .ColWidth(2) = 3000
    .TextMatrix(0, 1) = "Fecha"
    .TextMatrix(0, 2) = "Consductor"
End With
End Sub
Sub MarcoPapeleta()
With Me.MSHPapeleta
    .Rows = 2
    .Clear
    .ColWidth(0) = 300
    .ColWidth(1) = 1200
    .ColWidth(2) = 1200
    .ColWidth(3) = 3000
    .TextMatrix(0, 1) = "Fecha"
    .TextMatrix(0, 2) = "Monto"
    .TextMatrix(0, 3) = "Descripcion"
End With
End Sub
Sub MarcoIncidencia()
With Me.MSHIncidencia
    .Rows = 2
    .Clear
    .ColWidth(0) = 300
    .ColWidth(1) = 1200
    .ColWidth(2) = 2000
    .ColWidth(3) = 2000
    .ColWidth(4) = 3000
    .TextMatrix(0, 1) = "Fecha"
    .TextMatrix(0, 2) = "Tipo de Incidencia"
    .TextMatrix(0, 3) = "Lugar"
    .TextMatrix(0, 4) = "Descripcion"
End With
End Sub
Sub MarcoCarga()
With Me.MSHCarga
    .Rows = 2
    .Clear
    .ColWidth(0) = 300
    .ColWidth(1) = 1000
    .ColWidth(2) = 4500
    .ColWidth(3) = 1500
    .ColWidth(4) = 4500
    .TextMatrix(0, 1) = "Fecha"
    .TextMatrix(0, 2) = "Descripcion"
    .TextMatrix(0, 3) = "Agencia Dest"
    .TextMatrix(0, 4) = "Descripcion Dest"
End With

End Sub
Sub MarcoSoat()
With Me.MSHSOAT
    .Rows = 2
    .Clear
    .ColWidth(0) = 300
    .ColWidth(1) = 1300
    .ColWidth(2) = 1300
    .ColWidth(3) = 1300
    .TextMatrix(0, 1) = "Fec Ini"
    .TextMatrix(0, 2) = "Fec Venc"
    .TextMatrix(0, 3) = "Monto ($)"
End With
End Sub

Sub MarcoKm()
With Me.MSHKM
    .Rows = 2
    .Clear
    .ColWidth(0) = 300
    .ColWidth(1) = 1300
    .ColWidth(2) = 2000
    .TextMatrix(0, 1) = "Fecha"
    .TextMatrix(0, 2) = "Tacometro (KM)"
End With
End Sub



Sub InicioLimpia()
    MarcoSoat
    MarcoKm
    MarcoCarga
    MarcoPapeleta
    MarcoIncidencia
    MarcoAsigna
    LblBSCod = ""
    LblBSSerie = ""
    lblEstado = ""
    LblMarca = ""
    LblModelo = ""
    LblPlaca = ""
    LblTipo = ""
    CmdCancelarCarga_Click
    CmdCancelarKM_Click
    CmdCancelarSoat_Click
    CmdCancelarPapeleta_Click
End Sub



Private Sub CmdBuscarPlaca_Click()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset
Set LV = New DLogvehiculo
Set Rs = LV.GetDatoVehiculo(Trim(Me.TxtPlaca))
InicioLimpia
If Not (Rs.EOF And Rs.BOF) Then
    LblBSCod = Rs!cBSCod
    LblBSSerie = Rs!cBSSerie
    
    Select Case Rs!nEstado
        Case "1"
            lblEstado = "LIBRE" & Space(100) & Rs!nEstado
        Case "2"
            lblEstado = "ASIGNADO" & Space(100) & Rs!nEstado
        Case "3"
            lblEstado = "REPARACION" & Space(100) & Rs!nEstado
    End Select
    
    LblMarca = Rs!Marcar
    LblModelo = Rs!cModelo
    LblPlaca = Rs!cPlaca
    LblTipo = Rs!TipoV
    CargaSoatVehiculo
    CargaKm
    CargaCargaVehiculo
    CargaPapeleta
    CargaInsidencia
    CargaAsignacion
    
Else
    MsgBox "No existe Vehiculo", vbInformation, "AVISO"
    Set Rs = Nothing
    Set LV = Nothing
    Exit Sub
End If
Set Rs = Nothing
Set LV = Nothing
End Sub

Public Sub fEnfoque(ctrControl As Control)
    ctrControl.SelStart = 0
    ctrControl.SelLength = Len(ctrControl.Text)
End Sub


Private Sub cmdCancelarAsigna_Click()
Me.TxtFechaAsigna = "__/__/____"
Me.CboConductor.ListIndex = -1
Me.chkAsigna.value = IIf(Right(Me.lblEstado, 1) = 2, 1, 0)
End Sub

Private Sub CmdCancelarCarga_Click()
Me.TxtFechaCarga = "__/__/____"
Me.TxtDescripcion = ""
Me.TxtDestinoDesc = ""
Me.CboAgencia.ListIndex = -1
End Sub

Private Sub CmdCancelarIncidencia_Click()
Me.TxtFechaIncidencia = "__/__/____"
Me.TxtMontoIncidencia = 0
Me.TxtDescripcionIncidencia = ""
Me.TxtLugarInsidencia = ""
Me.CboIncidencia.ListIndex = -1
End Sub

Private Sub CmdCancelarKM_Click()
Me.TxtKM = ""
Me.txtFEchaKM = "__/__/____"
End Sub

Private Sub CmdCancelarPapeleta_Click()
Me.TxtFechaPapeleta = "__/__/____"
Me.TxtDescripcionPapeleta = ""
Me.TxtMontoPapeleta = ""
End Sub

Private Sub CmdCancelarSoat_Click()
Me.txtFecha = "__/__/____"
Me.txtFechaF = "__/__/____"
Me.TxtSoat = 0
End Sub

Private Sub CmdFechaFin_Click()
Me.TxtFechaFinAsignacion = "__/__/____"
End Sub

Private Sub CmdGrabaCarga_Click()
Dim opt As Integer
Dim LV As DLogvehiculo

If Me.CboAgencia.ListIndex = -1 Then
    MsgBox "Elija una Agencia", vbInformation, "AVISO"
    Exit Sub
End If
If ValFecha(Me.TxtFechaCarga) = False Then
    Exit Sub
End If

If Trim(Me.TxtDescripcion) = "" Then
    MsgBox "Ingrese la Descripcion de la Carga", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.TxtDestinoDesc) = "" Then
    MsgBox "Ingrese la Descripcion del Destino", vbInformation, "AVISO"
    Exit Sub
End If
opt = MsgBox("Esta Seguaro de Grabar", vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub

Set LV = New DLogvehiculo
Call LV.InsertRegVehiculoCarga(Me.TxtFechaCarga, Me.TxtDescripcion, Right(Me.CboAgencia.Text, 2), Me.TxtDestinoDesc, Me.LblBSCod, Me.LblBSSerie, gdFecSis, Right(gsCodAge, 2), gsCodUser)
CmdCancelarCarga_Click
CargaCargaVehiculo
Set LV = Nothing
End Sub

Private Sub CmdGrabaFin_Click()
Dim LV As DLogvehiculo
Dim opt As Integer

If Right(Me.lblEstado, 1) = 1 Then
    MsgBox "El vehiculo se encuentra libre", vbInformation, "AVISO"
    Exit Sub
End If

If ValFecha(Me.TxtFechaFinAsignacion) = False Then
    Exit Sub
End If

opt = MsgBox("Esta Seguaro de Grabar", vbQuestion + vbYesNo, "AVISO")
Set LV = New DLogvehiculo
If opt = vbNo Then Exit Sub

Call LV.LiberaAsignacionVehiculo(Me.LblBSCod, Me.LblBSSerie, Me.TxtFechaFinAsignacion)
lblEstado = "LIBRE" & Space(100) & "1"
CargaAsignacion
CargaComboConductorLibre

Set LV = Nothing
End Sub

Private Sub cmdGrabarAsigna_Click()
Dim LV As DLogvehiculo
Dim opt As Integer

If Right(Me.lblEstado, 1) = 2 Then
    MsgBox "El vehiculo se encuentra asignado", vbInformation, "AVISO"
    Exit Sub
End If

If Me.chkAsigna.value = 1 Then
    If ValFecha(Me.TxtFechaAsigna) = False Then
        Exit Sub
    End If
    If Me.CboConductor.ListIndex = -1 Then Exit Sub
End If

If Me.chkAsigna.value = 0 Then
    MsgBox "Debe marcar el check de Asignacion", vbInformation, "AVISO"
    Exit Sub
End If

opt = MsgBox("Esta Seguaro de Grabar", vbQuestion + vbYesNo, "AVISO")
Set LV = New DLogvehiculo
If opt = vbNo Then Exit Sub

Call LV.AsiganacionVehiculo(Trim(Right(Me.CboConductor.Text, 15)), Me.LblBSCod, Me.LblBSSerie, Me.TxtFechaAsigna, Me.chkAsigna.value)
lblEstado = "ASIGNADO" & Space(100) & "2"
CargaAsignacion
CargaComboConductorLibre

Set LV = Nothing
End Sub

Private Sub CmdGrabarIncidencia_Click()
Dim opt As Integer
Dim LV As DLogvehiculo

If ValFecha(Me.TxtFechaIncidencia) = False Then
    Exit Sub
End If

If Trim(Me.TxtLugarInsidencia) = "" Then
    MsgBox "Ingrese el lugar de Incidencia", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.TxtDescripcionIncidencia) = "" Then
    MsgBox "Ingrese el Tipo de Incidencia", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.TxtMontoIncidencia) = "" Then
    MsgBox "Ingrese el Monto de la Incidencia", vbInformation, "AVISO"
    Exit Sub
End If

If Me.CboIncidencia.ListIndex = -1 Then
    MsgBox "Elija el Tipo de Incidencia", vbInformation, "AVISO"
    Exit Sub
End If

opt = MsgBox("Esta Seguaro de Grabar", vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
Set LV = New DLogvehiculo
Call LV.InsertRegVehiculoIncidencia(Me.TxtFechaIncidencia, Me.TxtDescripcionIncidencia, CInt(Trim(Right(Me.CboIncidencia, 3))), Me.TxtLugarInsidencia, Me.LblBSCod, Me.LblBSSerie, gdFecSis, Right(gsCodAge, 2), gsCodUser)
CmdCancelarIncidencia_Click
CargaInsidencia
Set LV = Nothing
End Sub

Private Sub CmdGrabarKm_Click()
Dim opt As Integer
Dim LV As DLogvehiculo
If ValFecha(Me.txtFEchaKM) = False Then
    Exit Sub
End If

If Me.TxtKM = 0 Then
    MsgBox "Dato Incorrecto", vbInformation, "AVISO"
    Me.TxtKM.SetFocus
    Exit Sub
End If

If Me.TxtKMFin = 0 Then
    MsgBox "Dato Incorrecto", vbInformation, "AVISO"
    Me.TxtKMFin.SetFocus
    Exit Sub
End If

If Trim(Me.TxtKM) = "" Then
    MsgBox "Dato Incorrecto", vbInformation, "AVISO"
    Me.TxtKM.SetFocus
    Exit Sub
End If

If Trim(Me.TxtKMFin) = "" Then
    MsgBox "Dato Incorrecto", vbInformation, "AVISO"
    Me.TxtKMFin.SetFocus
    Exit Sub
End If

If TxtKMFin > TxtKM Then
    MsgBox "Rango de Kilomentraje Incorrecto", vbInformation, "AVISO"
    Me.TxtKMFin.SetFocus
    Exit Sub
End If



opt = MsgBox("Esta Seguaro de Grabar", vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
Set LV = New DLogvehiculo
MarcoKm
Call LV.InsertRegVehiculoKm(Me.txtFEchaKM, Me.TxtKM, Me.TxtKMFin, Me.LblBSCod, Me.LblBSSerie, gdFecSis, Right(gsCodAge, 2), gsCodUser)
CmdCancelarKM_Click
CargaKm
Set LV = Nothing
End Sub

Private Sub CmdGrabarPapeleta_Click()
Dim opt As Integer
Dim LV As DLogvehiculo

If ValFecha(Me.TxtFechaPapeleta) = False Then
    Exit Sub
End If

If Trim(Me.TxtDescripcionPapeleta) = "" Then
    MsgBox "Ingrese la Descripcion del Destino", vbInformation, "AVISO"
    Exit Sub
End If

If Trim(Me.TxtMontoPapeleta) = "" Or Me.TxtMontoPapeleta = 0 Then
    MsgBox "Ingrese el Monto de la Papeleta", vbInformation, "AVISO"
    Exit Sub
End If

opt = MsgBox("Esta Seguaro de Grabar", vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
Set LV = New DLogvehiculo
Call LV.InsertRegVehiculoPapeleta(Me.TxtFechaPapeleta, Me.TxtDescripcionPapeleta, Me.TxtMontoPapeleta, Me.LblBSCod, Me.LblBSSerie, gdFecSis, Right(gsCodAge, 2), gsCodUser)
CmdCancelarPapeleta_Click
CargaPapeleta
Set LV = Nothing
End Sub

Private Sub CmdGrabarSoat_Click()
Dim opt As Integer
Dim LV As DLogvehiculo

If ValFecha(Me.txtFecha) = False Then
    Exit Sub
End If

If ValFecha(Me.txtFechaF) = False Then
    Exit Sub
End If

If Me.txtFechaF < Me.txtFecha Then
    MsgBox "Rango de Fechas Incorrectas", vbInformation, "AVISO"
    Exit Sub
End If

opt = MsgBox("Esta Seguaro de Grabar", vbQuestion + vbYesNo, "AVISO")
If opt = vbNo Then Exit Sub
Set LV = New DLogvehiculo
Call LV.InsertRegVehiculoSoat(Me.txtFecha, Me.txtFechaF, Me.LblBSCod, Me.LblBSSerie, Me.TxtSoat, gdFecSis, Right(gsCodAge, 2), gsCodUser)
CmdCancelarSoat_Click
CargaSoatVehiculo
Set LV = Nothing
End Sub

Sub CargaCargaVehiculo()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset

Set LV = New DLogvehiculo
Set Rs = LV.GetVehiculoCarga(Me.LblBSCod, Me.LblBSSerie)
i = 1
While Not Rs.EOF
    Me.MSHCarga.TextMatrix(i, 1) = Format(Rs!dFecha, "DD/MM/YYYY")
    Me.MSHCarga.TextMatrix(i, 2) = Rs!cDescripcion
    Me.MSHCarga.TextMatrix(i, 3) = Rs!NombreAgencia
    Me.MSHCarga.TextMatrix(i, 4) = Rs!cDestinoDesc
    i = i + 1
    MSHCarga.Rows = MSHCarga.Rows + 1
    Rs.MoveNext
Wend
If Not (Rs.EOF And Rs.BOF) Then Me.MSHCarga.Rows = MSHCarga.Rows - 1
Set LV = Nothing
Set Rs = Nothing
End Sub

Sub CargaPapeleta()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset

Set LV = New DLogvehiculo
Set Rs = LV.GetVehiculoPapeleta(Me.LblBSCod, Me.LblBSSerie)
i = 1
While Not Rs.EOF
    Me.MSHPapeleta.TextMatrix(i, 1) = Format(Rs!dFecha, "DD/MM/YYYY")
    Me.MSHPapeleta.TextMatrix(i, 2) = Format(Rs!nMonto, "#0.00")
    Me.MSHPapeleta.TextMatrix(i, 3) = Rs!cDescripcion
    i = i + 1
    MSHPapeleta.Rows = MSHPapeleta.Rows + 1
    Rs.MoveNext
Wend
If Not (Rs.EOF And Rs.BOF) Then Me.MSHPapeleta.Rows = MSHPapeleta.Rows - 1
Set LV = Nothing
Set Rs = Nothing
End Sub
Sub CargaInsidencia()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset

Set LV = New DLogvehiculo
Set Rs = LV.GetVehiculoIncidencia(Me.LblBSCod, Me.LblBSSerie)
i = 1
While Not Rs.EOF
    Me.MSHIncidencia.TextMatrix(i, 1) = Format(Rs!dFecha, "DD/MM/YYYY")
    Me.MSHIncidencia.TextMatrix(i, 2) = Rs!TipoIncidencia
    Me.MSHIncidencia.TextMatrix(i, 3) = Rs!cDescripcion
    Me.MSHIncidencia.TextMatrix(i, 4) = Rs!cLugar
    i = i + 1
    MSHIncidencia.Rows = MSHIncidencia.Rows + 1
    Rs.MoveNext
Wend
If Not (Rs.EOF And Rs.BOF) Then Me.MSHIncidencia.Rows = MSHIncidencia.Rows - 1
Set LV = Nothing
Set Rs = Nothing
End Sub

Sub CargaKm()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset

Set LV = New DLogvehiculo
Set Rs = LV.GetVehiculoKm(Me.LblBSCod, Me.LblBSSerie)
i = 1
While Not Rs.EOF
    Me.MSHKM.TextMatrix(i, 1) = Format(Rs!dFecha, "DD/MM/YYYY")
    Me.MSHKM.TextMatrix(i, 2) = Rs!nkm
    i = i + 1
    MSHKM.Rows = MSHKM.Rows + 1
    Rs.MoveNext
Wend
If Not (Rs.EOF And Rs.BOF) Then Me.MSHKM.Rows = MSHKM.Rows - 1
Set LV = Nothing
Set Rs = Nothing
End Sub

Sub CargaAsignacion()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset

Me.TxtFechaAsigna = gdFecSis
Me.TxtFechaFinAsignacion = gdFecSis
If Right(lblEstado, 1) = 1 Then
    Me.FraAsigna.Visible = True
    Me.FraDesasigna.Visible = False
Else
    Me.FraDesasigna.Visible = True
    Me.FraAsigna.Visible = False
End If


Set LV = New DLogvehiculo
Set Rs = LV.GetvehiculoAsignacion(Me.LblBSCod, Me.LblBSSerie)
i = 1
MarcoAsigna
While Not Rs.EOF
    Me.MSHAsigna.TextMatrix(i, 1) = Format(Rs!dFecha, "DD/MM/YYYY")
    Me.MSHAsigna.TextMatrix(i, 2) = Rs!cPersNombre
    i = i + 1
    MSHAsigna.Rows = MSHAsigna.Rows + 1
    Rs.MoveNext
Wend
If Not (Rs.EOF And Rs.BOF) Then Me.MSHAsigna.Rows = MSHAsigna.Rows - 1
Set LV = Nothing
Set Rs = Nothing
End Sub
Sub CargaSoatVehiculo()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset

Set LV = New DLogvehiculo
Set Rs = LV.GetVehiculoSoat(Me.LblBSCod, Me.LblBSSerie)
i = 1
While Not Rs.EOF
    Me.MSHSOAT.TextMatrix(i, 1) = Format(Rs!dInicio, "DD/MM/YYYY")
    Me.MSHSOAT.TextMatrix(i, 2) = Format(Rs!dVencimiento, "DD/MM/YYYY")
    Me.MSHSOAT.TextMatrix(i, 3) = Format(Rs!nMonto, "#0.00")
    i = i + 1
    MSHSOAT.Rows = MSHSOAT.Rows + 1
    Rs.MoveNext
Wend
If Not (Rs.EOF And Rs.BOF) Then Me.MSHSOAT.Rows = MSHSOAT.Rows - 1
Set LV = Nothing
Set Rs = Nothing
End Sub

Private Sub cmdNuevo_Click()
FrmLogisticaVehiculoTipoIncidencia.Show 1
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub



Private Sub Form_Load()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset
Set LV = New DLogvehiculo
Set Rs = LV.GetAgencias
While Not Rs.EOF
    Me.CboAgencia.AddItem Rs!cAgeDescripcion & Space(100) & Rs!cAgeCod
    Rs.MoveNext
Wend
cmdCancelarAsigna_Click
CargaComboTipoIncidencia
CargaComboConductorLibre

Set Rs = Nothing
Set LV = Nothing
End Sub
Sub CargaComboTipoIncidencia()
Dim LV As DLogvehiculo
Dim Rs As ADODB.Recordset
Set LV = New DLogvehiculo
Set Rs = LV.GetTipoIncidencia
Me.CboIncidencia.Clear
While Not Rs.EOF
    Me.CboIncidencia.AddItem Rs!cDescripcion & Space(100) & Rs!ntipoIncidencia
    Rs.MoveNext
Wend
Set Rs = Nothing
Set LV = Nothing
End Sub


Private Sub TxtDescripcion_GotFocus()
fEnfoque TxtDescripcion
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.TxtDestinoDesc.SetFocus
End If
End Sub





Private Sub TxtDescripcionIncidencia_GotFocus()
fEnfoque TxtDescripcionIncidencia
End Sub

Private Sub TxtDescripcionIncidencia_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.TxtMontoIncidencia.SetFocus
End If
End Sub

Private Sub TxtDescripcionPapeleta_GotFocus()
fEnfoque TxtDescripcionPapeleta
End Sub

Private Sub TxtDescripcionPapeleta_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.TxtMontoPapeleta.SetFocus
End If
End Sub

Private Sub TxtDestinoDesc_GotFocus()
fEnfoque TxtDestinoDesc
End Sub

Private Sub TxtDestinoDesc_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.CmdGrabaCarga.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtFechaF.SetFocus
End If
End Sub



Private Sub TxtFechaCarga_GotFocus()
fEnfoque TxtFechaCarga
End Sub

Private Sub TxtFechaCarga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtDescripcion.SetFocus
End If
End Sub

Private Sub txtFechaF_GotFocus()
fEnfoque txtFechaF
End Sub

Private Sub txtFechaF_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtSoat.SetFocus
End If
End Sub




Private Sub TxtFechaIncidencia_GotFocus()
fEnfoque TxtFechaIncidencia
End Sub

Private Sub TxtFechaIncidencia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtLugarInsidencia.SetFocus
End If
End Sub

Private Sub txtFEchaKM_GotFocus()
fEnfoque txtFEchaKM
End Sub

Private Sub txtFEchaKM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtKM.SetFocus
End If
End Sub




Private Sub TxtFechaPapeleta_GotFocus()
fEnfoque TxtFechaPapeleta
End Sub

Private Sub TxtFechaPapeleta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtDescripcionPapeleta.SetFocus
End If
End Sub

Private Sub TxtKM_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(TxtKM, KeyAscii)
If KeyAscii = 13 Then
    Me.TxtKM.SetFocus
End If
End Sub



Private Sub TxtKMFin_GotFocus()
fEnfoque TxtKMFin
End Sub

Private Sub TxtKMFin_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(TxtKMFin, KeyAscii)
If KeyAscii = 13 Then
    Me.CmdGrabarKm.SetFocus
End If
End Sub


Private Sub TxtLugarInsidencia_GotFocus()
fEnfoque TxtLugarInsidencia
End Sub


Private Sub TxtLugarInsidencia_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    Me.TxtDescripcionIncidencia.SetFocus
End If
End Sub

Private Sub TxtMontoIncidencia_GotFocus()
fEnfoque TxtMontoIncidencia
End Sub

Private Sub TxtMontoIncidencia_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(TxtSoat, KeyAscii)
If KeyAscii = 13 Then
    Me.CmdGrabarIncidencia.SetFocus
End If
End Sub

Private Sub TxtMontoPapeleta_GotFocus()
fEnfoque TxtMontoPapeleta
End Sub

Private Sub TxtMontoPapeleta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(TxtSoat, KeyAscii)
If KeyAscii = 13 Then
    Me.CmdGrabarPapeleta.SetFocus
End If
End Sub

Private Sub TxtPlaca_KeyPress(KeyAscii As Integer)
KeyAscii = intfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    CmdBuscarPlaca_Click
End If
End Sub

Private Sub TxtSoat_GotFocus()
fEnfoque TxtSoat
End Sub

Private Sub TxtSoat_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(TxtSoat, KeyAscii)
If KeyAscii = 13 Then
    Me.CmdGrabarSoat.SetFocus
End If
End Sub
Public Function intfMayusculas(intTecla As Integer) As Integer
 If Chr(intTecla) >= "a" And Chr(intTecla) <= "z" Then
    intTecla = intTecla - 32
 End If
 If intTecla = 39 Then
    intTecla = 0
 End If
 If intTecla = 209 Or intTecla = 241 Or intTecla = 8 Or intTecla = 32 Then
    intfMayusculas = Asc(UCase(Chr(intTecla)))
     Exit Function
 End If
 intfMayusculas = intTecla
End Function
