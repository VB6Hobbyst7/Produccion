VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPMetasXAgencias 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Metas Por Agencias"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12075
   Icon            =   "frmCredBPPMetasXAgencias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Parametros de Cumplimiento"
      TabPicture(0)   =   "frmCredBPPMetasXAgencias.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "SSTab2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   7223
         _Version        =   393216
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Nivel I"
         TabPicture(0)   =   "frmCredBPPMetasXAgencias.frx":0326
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label10"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label11"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label12"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label13"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "feMetasN1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdGuardarN1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdCancelarN1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdMetasGM1"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Frame2"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Nivel II"
         TabPicture(1)   =   "frmCredBPPMetasXAgencias.frx":0342
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmdMetasGM2"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdCancelarN2"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdGuardarN2"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "feMetasN2"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Label8"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).Control(6)=   "Label7"
         Tab(1).Control(6).Enabled=   0   'False
         Tab(1).Control(7)=   "Label6"
         Tab(1).Control(7).Enabled=   0   'False
         Tab(1).Control(8)=   "Label5"
         Tab(1).Control(8).Enabled=   0   'False
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Nivel III"
         TabPicture(2)   =   "frmCredBPPMetasXAgencias.frx":035E
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraCategoria"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "cmdMetasGM3"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "cmdCancelarN3"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).Control(3)=   "cmdGuardarN3"
         Tab(2).Control(3).Enabled=   0   'False
         Tab(2).Control(4)=   "feMetasN3"
         Tab(2).Control(4).Enabled=   0   'False
         Tab(2).Control(5)=   "Label3"
         Tab(2).Control(5).Enabled=   0   'False
         Tab(2).Control(6)=   "Label2"
         Tab(2).Control(6).Enabled=   0   'False
         Tab(2).Control(7)=   "Label1"
         Tab(2).Control(7).Enabled=   0   'False
         Tab(2).Control(8)=   "lblSaldoCartera"
         Tab(2).Control(8).Enabled=   0   'False
         Tab(2).ControlCount=   9
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   240
            TabIndex        =   26
            Top             =   600
            Width           =   1030
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Analista"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   200
               TabIndex        =   27
               Top             =   315
               Width           =   690
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   -74760
            TabIndex        =   19
            Top             =   600
            Width           =   1030
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Analista"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   200
               TabIndex        =   20
               Top             =   315
               Width           =   690
            End
         End
         Begin VB.Frame fraCategoria 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   730
            Left            =   -74760
            TabIndex        =   12
            Top             =   555
            Width           =   1030
            Begin VB.Label lblCategoria 
               AutoSize        =   -1  'True
               Caption         =   "Analista"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   200
               TabIndex        =   13
               Top             =   315
               Width           =   690
            End
         End
         Begin VB.CommandButton cmdMetasGM3 
            Caption         =   "Metas GM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   11
            Top             =   3600
            Width           =   1170
         End
         Begin VB.CommandButton cmdMetasGM2 
            Caption         =   "Metas GM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -74760
            TabIndex        =   10
            Top             =   3600
            Width           =   1170
         End
         Begin VB.CommandButton cmdMetasGM1 
            Caption         =   "Metas GM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   3600
            Width           =   1170
         End
         Begin VB.CommandButton cmdCancelarN3 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64920
            TabIndex        =   7
            Top             =   3600
            Width           =   1170
         End
         Begin VB.CommandButton cmdGuardarN3 
            Caption         =   "Guardar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -66240
            TabIndex        =   6
            Top             =   3600
            Width           =   1170
         End
         Begin VB.CommandButton cmdCancelarN2 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -64920
            TabIndex        =   5
            Top             =   3600
            Width           =   1170
         End
         Begin VB.CommandButton cmdGuardarN2 
            Caption         =   "Guardar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -66240
            TabIndex        =   4
            Top             =   3600
            Width           =   1170
         End
         Begin VB.CommandButton cmdCancelarN1 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10080
            TabIndex        =   3
            Top             =   3600
            Width           =   1170
         End
         Begin VB.CommandButton cmdGuardarN1 
            Caption         =   "Guardar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8760
            TabIndex        =   2
            Top             =   3600
            Width           =   1170
         End
         Begin SICMACT.FlexEdit feMetasN3 
            Height          =   2535
            Left            =   -74760
            TabIndex        =   8
            Top             =   960
            Width           =   11040
            _extentx        =   19473
            _extenty        =   4471
            cols0           =   11
            highlight       =   1
            encabezadosnombres=   "#-Analista-GM-AG.-GM-AG.-GM-AG.-GM-AG.-Aux"
            encabezadosanchos=   "0-1000-1200-1200-1200-1200-1200-1200-1200-1200-0"
            font            =   "frmCredBPPMetasXAgencias.frx":037A
            font            =   "frmCredBPPMetasXAgencias.frx":03A2
            font            =   "frmCredBPPMetasXAgencias.frx":03CA
            font            =   "frmCredBPPMetasXAgencias.frx":03F2
            font            =   "frmCredBPPMetasXAgencias.frx":041A
            fontfixed       =   "frmCredBPPMetasXAgencias.frx":0442
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-3-X-5-X-7-X-9-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-R-R-R-R-R-R-R-R-C"
            formatosedit    =   "0-0-2-2-3-3-3-3-2-2-0"
            cantentero      =   15
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin SICMACT.FlexEdit feMetasN2 
            Height          =   2535
            Left            =   -74760
            TabIndex        =   18
            Top             =   1005
            Width           =   11040
            _extentx        =   19473
            _extenty        =   4471
            cols0           =   11
            highlight       =   1
            encabezadosnombres=   "#-Analista-GM-AG.-GM-AG.-GM-AG.-GM-AG.-Aux"
            encabezadosanchos=   "0-1000-1200-1200-1200-1200-1200-1200-1200-1200-0"
            font            =   "frmCredBPPMetasXAgencias.frx":0468
            font            =   "frmCredBPPMetasXAgencias.frx":0490
            font            =   "frmCredBPPMetasXAgencias.frx":04B8
            font            =   "frmCredBPPMetasXAgencias.frx":04E0
            font            =   "frmCredBPPMetasXAgencias.frx":0508
            fontfixed       =   "frmCredBPPMetasXAgencias.frx":0530
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-3-X-5-X-7-X-9-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-R-R-R-R-R-R-R-R-C"
            formatosedit    =   "0-0-2-2-3-3-3-3-2-2-0"
            cantentero      =   15
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin SICMACT.FlexEdit feMetasN1 
            Height          =   2535
            Left            =   240
            TabIndex        =   25
            Top             =   1000
            Width           =   11040
            _extentx        =   19473
            _extenty        =   4471
            cols0           =   11
            highlight       =   1
            encabezadosnombres=   "#-Analista-GM-AG.-GM-AG.-GM-AG.-GM-AG.-Aux"
            encabezadosanchos=   "0-1000-1200-1200-1200-1200-1200-1200-1200-1200-0"
            font            =   "frmCredBPPMetasXAgencias.frx":0556
            font            =   "frmCredBPPMetasXAgencias.frx":057E
            font            =   "frmCredBPPMetasXAgencias.frx":05A6
            font            =   "frmCredBPPMetasXAgencias.frx":05CE
            font            =   "frmCredBPPMetasXAgencias.frx":05F6
            fontfixed       =   "frmCredBPPMetasXAgencias.frx":061E
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            lbultimainstancia=   -1
            tipobusqueda    =   3
            columnasaeditar =   "X-X-X-3-X-5-X-7-X-9-X"
            listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-C-R-R-R-R-R-R-R-R-C"
            formatosedit    =   "0-0-2-2-3-3-3-3-2-2-0"
            cantentero      =   15
            textarray0      =   "#"
            lbeditarflex    =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mora(8-30)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   8460
            TabIndex        =   31
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nº Operaciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   6060
            TabIndex        =   30
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nº Clientes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3660
            TabIndex        =   29
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   1260
            TabIndex        =   28
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mora(8-30)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -66540
            TabIndex        =   24
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nº Operaciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -68940
            TabIndex        =   23
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nº Clientes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -71340
            TabIndex        =   22
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -73740
            TabIndex        =   21
            Top             =   690
            Width           =   2415
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Mora(8-30)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -66540
            TabIndex        =   17
            Top             =   645
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nº Operaciones"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -68940
            TabIndex        =   16
            Top             =   645
            Width           =   2415
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nº Clientes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -71340
            TabIndex        =   15
            Top             =   645
            Width           =   2415
         End
         Begin VB.Label lblSaldoCartera 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Saldo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   -73740
            TabIndex        =   14
            Top             =   645
            Width           =   2415
         End
      End
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Mes a generar:"
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label lblMes 
      AutoSize        =   -1  'True
      Caption         =   "@Mes"
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
      Left            =   1320
      TabIndex        =   32
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmCredBPPMetasXAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private fgFecActual As Date
'Private i As Integer
'
'Private Sub cmdGuardarN1_Click()
'On Error GoTo Error
'If ValidaDatos(1) Then
'    If MsgBox("Estas seguro de guradar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    For i = 1 To feMetasN1.Rows - 1
'        Call oBPP.OpeMetasAgencia(Trim(feMetasN1.TextMatrix(i, 10)), CDbl(feMetasN1.TextMatrix(i, 3)), CDbl(feMetasN1.TextMatrix(i, 5)), _
'                        CDbl(feMetasN1.TextMatrix(i, 7)), CDbl(feMetasN1.TextMatrix(i, 9)), Month(fgFecActual), Year(fgFecActual), gsCodUser, fgFecActual, 1)
'    Next i
'
'    MsgBox "Datos registrados satisfactoriamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarN2_Click()
'On Error GoTo Error
'If ValidaDatos(2) Then
'    If MsgBox("Estas seguro de guradar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    For i = 1 To feMetasN2.Rows - 1
'        Call oBPP.OpeMetasAgencia(Trim(feMetasN2.TextMatrix(i, 10)), CDbl(feMetasN2.TextMatrix(i, 3)), CDbl(feMetasN2.TextMatrix(i, 5)), _
'                        CDbl(feMetasN2.TextMatrix(i, 7)), CDbl(feMetasN2.TextMatrix(i, 9)), Month(fgFecActual), Year(fgFecActual), gsCodUser, fgFecActual, 2)
'    Next i
'
'    MsgBox "Datos registrados satisfactoriamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Sub cmdGuardarN3_Click()
'On Error GoTo Error
'If ValidaDatos(3) Then
'    If MsgBox("Estas seguro de guradar los datos?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Dim oBPP As COMNCredito.NCOMBPPR
'    Set oBPP = New COMNCredito.NCOMBPPR
'
'    For i = 1 To feMetasN3.Rows - 1
'        Call oBPP.OpeMetasAgencia(Trim(feMetasN3.TextMatrix(i, 10)), CDbl(feMetasN3.TextMatrix(i, 3)), CDbl(feMetasN3.TextMatrix(i, 5)), _
'                        CDbl(feMetasN3.TextMatrix(i, 7)), CDbl(feMetasN3.TextMatrix(i, 9)), Month(fgFecActual), Year(fgFecActual), gsCodUser, fgFecActual, 3)
'    Next i
'
'    MsgBox "Datos registrados satisfactoriamente", vbInformation, "Aviso"
'End If
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'
'Private Function ValidaDatos(ByVal pnNivel As Integer) As Boolean
'Select Case pnNivel
'    Case 1:
'            For i = 1 To feMetasN1.Rows - 1
'                If CDbl(feMetasN1.TextMatrix(i, 3)) < CDbl(feMetasN1.TextMatrix(i, 2)) Then
'                    MsgBox "Saldo Cartera de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN1.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                If CDbl(feMetasN1.TextMatrix(i, 5)) < CDbl(feMetasN1.TextMatrix(i, 4)) Then
'                    MsgBox "Número de Clientes de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN1.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                If CDbl(feMetasN1.TextMatrix(i, 7)) < CDbl(feMetasN1.TextMatrix(i, 6)) Then
'                    MsgBox "Número de Operaciones de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN1.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                If CDbl(feMetasN1.TextMatrix(i, 9)) < CDbl(feMetasN1.TextMatrix(i, 8)) Then
'                    MsgBox "Mora(3-30) de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN1.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            Next i
'    Case 2:
'            For i = 1 To feMetasN2.Rows - 1
'                If CDbl(feMetasN2.TextMatrix(i, 3)) < CDbl(feMetasN2.TextMatrix(i, 2)) Then
'                    MsgBox "Saldo Cartera de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN2.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                If CDbl(feMetasN2.TextMatrix(i, 5)) < CDbl(feMetasN2.TextMatrix(i, 4)) Then
'                    MsgBox "Número de Clientes de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN2.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                If CDbl(feMetasN2.TextMatrix(i, 7)) < CDbl(feMetasN2.TextMatrix(i, 6)) Then
'                    MsgBox "Número de Operaciones de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN2.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                If CDbl(feMetasN2.TextMatrix(i, 9)) < CDbl(feMetasN2.TextMatrix(i, 8)) Then
'                    MsgBox "Mora(3-30) de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN2.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            Next i
'    Case 3:
'            For i = 1 To feMetasN3.Rows - 1
'                If CDbl(feMetasN3.TextMatrix(i, 3)) < CDbl(feMetasN3.TextMatrix(i, 2)) Then
'                    MsgBox "Saldo Cartera de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN3.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                If CDbl(feMetasN3.TextMatrix(i, 5)) < CDbl(feMetasN3.TextMatrix(i, 4)) Then
'                    MsgBox "Número de Clientes de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN3.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                If CDbl(feMetasN3.TextMatrix(i, 7)) < CDbl(feMetasN3.TextMatrix(i, 6)) Then
'                    MsgBox "Número de Operaciones de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN3.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'                If CDbl(feMetasN3.TextMatrix(i, 9)) < CDbl(feMetasN3.TextMatrix(i, 8)) Then
'                    MsgBox "Mora(3-30) de Agencia no puede ser menor que el parametro de Gerencia del Analista " & Trim(feMetasN3.TextMatrix(i, 1)), vbInformation, "Aviso"
'                    ValidaDatos = False
'                    Exit Function
'                End If
'            Next i
'End Select
'
'
'ValidaDatos = True
'End Function
'
'Private Sub cmdMetasGM1_Click()
'MetasAG 1
'End Sub
'
'Private Sub cmdMetasGM2_Click()
'MetasAG 2
'End Sub
'
'Private Sub cmdMetasGM3_Click()
'MetasAG 3
'End Sub
'
'Private Sub Form_Load()
'    MesActual
'    lblMes.Caption = MesAnio(fgFecActual)
'    CargaGridNivel 1
'    CargaGridNivel 2
'    CargaGridNivel 3
'End Sub
'
'Private Sub MetasAG(ByVal pnNivel As Integer)
'Select Case pnNivel
'    Case 1:
'            For i = 1 To feMetasN1.Rows - 1
'                If Trim(feMetasN1.TextMatrix(i, 3)) = "" Or Trim(feMetasN1.TextMatrix(i, 3)) = "0.00" Then
'                    feMetasN1.TextMatrix(i, 3) = feMetasN1.TextMatrix(i, 2)
'                End If
'
'                If Trim(feMetasN1.TextMatrix(i, 5)) = "" Or Trim(feMetasN1.TextMatrix(i, 5)) = "0" Then
'                    feMetasN1.TextMatrix(i, 5) = feMetasN1.TextMatrix(i, 4)
'                End If
'
'                If Trim(feMetasN1.TextMatrix(i, 7)) = "" Or Trim(feMetasN1.TextMatrix(i, 7)) = "0" Then
'                    feMetasN1.TextMatrix(i, 7) = feMetasN1.TextMatrix(i, 6)
'                End If
'
'                If Trim(feMetasN1.TextMatrix(i, 9)) = "" Or Trim(feMetasN1.TextMatrix(i, 9)) = "0.00" Then
'                    feMetasN1.TextMatrix(i, 9) = feMetasN1.TextMatrix(i, 8)
'                End If
'            Next i
'            feMetasN1.TopRow = 1
'    Case 2:
'            For i = 1 To feMetasN2.Rows - 1
'                If Trim(feMetasN2.TextMatrix(i, 3)) = "" Or Trim(feMetasN2.TextMatrix(i, 3)) = "0.00" Then
'                    feMetasN2.TextMatrix(i, 3) = feMetasN2.TextMatrix(i, 2)
'                End If
'
'                If Trim(feMetasN2.TextMatrix(i, 5)) = "" Or Trim(feMetasN2.TextMatrix(i, 5)) = "0" Then
'                    feMetasN2.TextMatrix(i, 5) = feMetasN2.TextMatrix(i, 4)
'                End If
'
'                If Trim(feMetasN2.TextMatrix(i, 7)) = "" Or Trim(feMetasN2.TextMatrix(i, 7)) = "0" Then
'                    feMetasN2.TextMatrix(i, 7) = feMetasN2.TextMatrix(i, 6)
'                End If
'
'                If Trim(feMetasN2.TextMatrix(i, 9)) = "" Or Trim(feMetasN2.TextMatrix(i, 9)) = "0.00" Then
'                    feMetasN2.TextMatrix(i, 9) = feMetasN2.TextMatrix(i, 8)
'                End If
'            Next i
'            feMetasN2.TopRow = 1
'    Case 3:
'            For i = 1 To feMetasN3.Rows - 1
'                If Trim(feMetasN3.TextMatrix(i, 3)) = "" Or Trim(feMetasN3.TextMatrix(i, 3)) = "0.00" Then
'                    feMetasN3.TextMatrix(i, 3) = feMetasN3.TextMatrix(i, 2)
'                End If
'
'                If Trim(feMetasN3.TextMatrix(i, 5)) = "" Or Trim(feMetasN3.TextMatrix(i, 5)) = "0" Then
'                    feMetasN3.TextMatrix(i, 5) = feMetasN3.TextMatrix(i, 4)
'                End If
'
'                If Trim(feMetasN3.TextMatrix(i, 7)) = "" Or Trim(feMetasN3.TextMatrix(i, 7)) = "0" Then
'                    feMetasN3.TextMatrix(i, 7) = feMetasN3.TextMatrix(i, 6)
'                End If
'
'                If Trim(feMetasN3.TextMatrix(i, 9)) = "" Or Trim(feMetasN3.TextMatrix(i, 9)) = "0.00" Then
'                    feMetasN3.TextMatrix(i, 9) = feMetasN3.TextMatrix(i, 8)
'                End If
'            Next i
'            feMetasN3.TopRow = 1
'End Select
'End Sub
'
'Private Sub CargaGridNivel(ByVal pnNivel As Integer)
'On Error GoTo Error
'
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'
'Set oBPP = New COMNCredito.NCOMBPPR
'
'Set rsBPP = oBPP.ObtenerMetasAgencia(fgFecActual, gsCodAge, pnNivel)
'
'Select Case pnNivel
'    Case 1: LimpiaFlex feMetasN1
'    Case 2: LimpiaFlex feMetasN2
'    Case 3: LimpiaFlex feMetasN3
'End Select
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    For i = 1 To rsBPP.RecordCount
'        Select Case pnNivel
'            Case 1:
'                    feMetasN1.AdicionaFila
'                    feMetasN1.TextMatrix(i, 1) = rsBPP!Usuario
'                    feMetasN1.TextMatrix(i, 10) = rsBPP!cPersCod
'                    feMetasN1.TextMatrix(i, 2) = Format(rsBPP!SG, "###," & String(15, "#") & "#0.00")
'                    feMetasN1.TextMatrix(i, 3) = Format(rsBPP!nSaldoAg, "###," & String(15, "#") & "#0.00")
'                    feMetasN1.TextMatrix(i, 4) = rsBPP!CG
'                    feMetasN1.TextMatrix(i, 5) = rsBPP!nNCliAg
'                    feMetasN1.TextMatrix(i, 6) = rsBPP!OG
'                    feMetasN1.TextMatrix(i, 7) = rsBPP!nNOpeAg
'                    feMetasN1.TextMatrix(i, 8) = Format(rsBPP!MG, "###," & String(15, "#") & "#0.00")
'                    feMetasN1.TextMatrix(i, 9) = Format(rsBPP!nMoraAg, "###," & String(15, "#") & "#0.00")
'            Case 2:
'                    feMetasN2.AdicionaFila
'                    feMetasN2.TextMatrix(i, 1) = rsBPP!Usuario
'                    feMetasN2.TextMatrix(i, 10) = rsBPP!cPersCod
'                    feMetasN2.TextMatrix(i, 2) = Format(rsBPP!SG, "###," & String(15, "#") & "#0.00")
'                    feMetasN2.TextMatrix(i, 3) = Format(rsBPP!nSaldoAg, "###," & String(15, "#") & "#0.00")
'                    feMetasN2.TextMatrix(i, 4) = rsBPP!CG
'                    feMetasN2.TextMatrix(i, 5) = rsBPP!nNCliAg
'                    feMetasN2.TextMatrix(i, 6) = rsBPP!OG
'                    feMetasN2.TextMatrix(i, 7) = rsBPP!nNOpeAg
'                    feMetasN2.TextMatrix(i, 8) = Format(rsBPP!MG, "###," & String(15, "#") & "#0.00")
'                    feMetasN2.TextMatrix(i, 9) = Format(rsBPP!nMoraAg, "###," & String(15, "#") & "#0.00")
'            Case 3:
'                    feMetasN3.AdicionaFila
'                    feMetasN3.TextMatrix(i, 1) = rsBPP!Usuario
'                    feMetasN3.TextMatrix(i, 10) = rsBPP!cPersCod
'                    feMetasN3.TextMatrix(i, 2) = Format(rsBPP!SG, "###," & String(15, "#") & "#0.00")
'                    feMetasN3.TextMatrix(i, 3) = Format(rsBPP!nSaldoAg, "###," & String(15, "#") & "#0.00")
'                    feMetasN3.TextMatrix(i, 4) = rsBPP!CG
'                    feMetasN3.TextMatrix(i, 5) = rsBPP!nNCliAg
'                    feMetasN3.TextMatrix(i, 6) = rsBPP!OG
'                    feMetasN3.TextMatrix(i, 7) = rsBPP!nNOpeAg
'                    feMetasN3.TextMatrix(i, 8) = Format(rsBPP!MG, "###," & String(15, "#") & "#0.00")
'                    feMetasN3.TextMatrix(i, 9) = Format(rsBPP!nMoraAg, "###," & String(15, "#") & "#0.00")
'        End Select
'        rsBPP.MoveNext
'    Next i
'
'    Select Case pnNivel
'        Case 1: feMetasN1.TopRow = 1
'        Case 2: feMetasN2.TopRow = 1
'        Case 3: feMetasN3.TopRow = 1
'    End Select
'
'End If
'
'Exit Sub
'Error:
'     MsgBox err.Description, vbCritical, "Error"
'End Sub
'Private Sub MesActual()
'Dim oConsSist As COMDConstSistema.NCOMConstSistema
'Set oConsSist = New COMDConstSistema.NCOMConstSistema
'fgFecActual = oConsSist.LeeConstSistema(gConstSistFechaBPP)
'Set oConsSist = Nothing
'End Sub
'
'Private Function MesAnio(ByVal dFecha As Date) As String
'Dim sFechaDesc As String
'sFechaDesc = ""
'
'Select Case Month(dFecha)
'    Case 1: sFechaDesc = "Enero"
'    Case 2: sFechaDesc = "Febrero"
'    Case 3: sFechaDesc = "Marzo"
'    Case 4: sFechaDesc = "Abril"
'    Case 5: sFechaDesc = "Mayo"
'    Case 6: sFechaDesc = "Junio"
'    Case 7: sFechaDesc = "Julio"
'    Case 8: sFechaDesc = "Agosto"
'    Case 9: sFechaDesc = "Septiembre"
'    Case 10: sFechaDesc = "Octubre"
'    Case 11: sFechaDesc = "Noviembre"
'    Case 12: sFechaDesc = "Diciembre"
'End Select
'
'sFechaDesc = sFechaDesc & " " & CStr(Year(dFecha))
'MesAnio = UCase(sFechaDesc)
'End Function
'
