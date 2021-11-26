VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContingInformeTecReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contingencias: Registro Informe Técnico"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10590
   Icon            =   "frmContingInformeTecReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVistaPrevia 
      Caption         =   "Vista Previa"
      Height          =   345
      Left            =   120
      TabIndex        =   36
      Top             =   6900
      Width           =   1170
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Height          =   345
      Left            =   1440
      TabIndex        =   1
      Top             =   6900
      Width           =   1050
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   2640
      TabIndex        =   0
      Top             =   6900
      Width           =   1050
   End
   Begin TabDlg.SSTab SSTabConting 
      Height          =   6645
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11721
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Informe Técnico"
      TabPicture(0)   =   "frmContingInformeTecReg.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraActivo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame0"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraPasivo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame FraPasivo 
         Caption         =   " Datos del Informe Técnico "
         Height          =   4815
         Left            =   120
         TabIndex        =   38
         Top             =   1680
         Width           =   10095
         Begin VB.ComboBox cboMonedaDemP 
            Height          =   315
            ItemData        =   "frmContingInformeTecReg.frx":0326
            Left            =   7920
            List            =   "frmContingInformeTecReg.frx":0330
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   360
            Width           =   680
         End
         Begin VB.ComboBox cboTipoP 
            Height          =   315
            ItemData        =   "frmContingInformeTecReg.frx":0359
            Left            =   5040
            List            =   "frmContingInformeTecReg.frx":0363
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   1520
            Width           =   1560
         End
         Begin VB.ComboBox cboCalifP 
            Height          =   315
            Left            =   7920
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   1520
            Width           =   2055
         End
         Begin VB.TextBox txtNroInformeP 
            Height          =   285
            Left            =   1080
            TabIndex        =   50
            Top             =   360
            Width           =   1935
         End
         Begin VB.ComboBox cboMonedaP 
            Height          =   315
            ItemData        =   "frmContingInformeTecReg.frx":038C
            Left            =   7920
            List            =   "frmContingInformeTecReg.frx":0396
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   725
            Width           =   680
         End
         Begin VB.TextBox txtDocOrigenP 
            Height          =   285
            Left            =   1080
            TabIndex        =   48
            Top             =   1120
            Width           =   2415
         End
         Begin VB.TextBox txtEntidadP 
            Height          =   285
            Left            =   6840
            TabIndex        =   47
            Top             =   1125
            Width           =   3075
         End
         Begin VB.TextBox txtNombreRefP 
            Height          =   285
            Left            =   1080
            TabIndex        =   46
            Top             =   1520
            Width           =   3015
         End
         Begin VB.Frame Frame5 
            Caption         =   " Cuentas Contables Ejercicio Actual "
            Height          =   1875
            Left            =   720
            TabIndex        =   41
            Top             =   2040
            Width           =   8535
            Begin Sicmact.TxtBuscar TxtBuscaCtaCodDeudorP 
               Height          =   330
               Left            =   1920
               TabIndex        =   70
               Top             =   270
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   582
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Enabled         =   0   'False
               TipoBusqueda    =   5
               EnabledText     =   0   'False
            End
            Begin Sicmact.TxtBuscar TxtBuscaCtaCodAcreedorP 
               Height          =   330
               Left            =   1920
               TabIndex        =   71
               Top             =   640
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   582
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Enabled         =   0   'False
               TipoBusqueda    =   5
               EnabledText     =   0   'False
            End
            Begin Sicmact.TxtBuscar TxtBuscaCtaCodGastosP 
               Height          =   330
               Left            =   1920
               TabIndex        =   72
               Top             =   1020
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   582
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Enabled         =   0   'False
               TipoBusqueda    =   5
               EnabledText     =   0   'False
            End
            Begin Sicmact.TxtBuscar TxtBuscaCtaCodProvP 
               Height          =   330
               Left            =   1920
               TabIndex        =   73
               Top             =   1400
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   582
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Enabled         =   0   'False
               TipoBusqueda    =   5
               EnabledText     =   0   'False
            End
            Begin VB.Label txtCtaDescProvP 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4320
               TabIndex        =   79
               Top             =   1400
               Width           =   4035
            End
            Begin VB.Label txtCtaDescGastosP 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4320
               TabIndex        =   78
               Top             =   1020
               Width           =   4035
            End
            Begin VB.Label txtCtaDescAcreedorP 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4320
               TabIndex        =   77
               Top             =   640
               Width           =   4035
            End
            Begin VB.Label txtCtaDescDeudorP 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4320
               TabIndex        =   76
               Top             =   270
               Width           =   4035
            End
            Begin VB.Label Label39 
               Caption         =   "Provisión :"
               Height          =   255
               Left            =   240
               TabIndex        =   45
               Top             =   1350
               Width           =   855
            End
            Begin VB.Label Label37 
               Caption         =   "Gastos :"
               Height          =   255
               Left            =   240
               TabIndex        =   44
               Top             =   990
               Width           =   855
            End
            Begin VB.Label Label24 
               Caption         =   "Deudor de control :"
               Height          =   255
               Left            =   240
               TabIndex        =   43
               Top             =   345
               Width           =   1455
            End
            Begin VB.Label Label23 
               Caption         =   "Acreedor de control :"
               Height          =   255
               Left            =   240
               TabIndex        =   42
               Top             =   675
               Width           =   1695
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   " Cuentas Contables Ejercicio Anterior "
            Height          =   735
            Left            =   720
            TabIndex        =   39
            Top             =   3960
            Width           =   8535
            Begin Sicmact.TxtBuscar TxtBuscaCtaCodAcreedorAntP 
               Height          =   330
               Left            =   1920
               TabIndex        =   74
               Top             =   240
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   582
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Enabled         =   0   'False
               TipoBusqueda    =   5
               EnabledText     =   0   'False
            End
            Begin VB.Label txtCtaDescAcreedorAntP 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   4320
               TabIndex        =   80
               Top             =   240
               Width           =   4035
            End
            Begin VB.Label Label20 
               Caption         =   "Acreedor Ejerc. Ant. :"
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   345
               Width           =   1575
            End
         End
         Begin Sicmact.EditMoney txtMontoP 
            Height          =   315
            Left            =   8610
            TabIndex        =   53
            Top             =   720
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
         Begin Sicmact.TxtBuscar txtBuscaPersP 
            Height          =   330
            Left            =   1080
            TabIndex        =   54
            Top             =   725
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   582
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
            Enabled         =   0   'False
            Enabled         =   0   'False
            TipoBusqueda    =   7
            TipoBusPers     =   1
            EnabledText     =   0   'False
         End
         Begin MSMask.MaskEdBox txtFecInformeP 
            Height          =   285
            Left            =   4440
            TabIndex        =   55
            Top             =   360
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFecOrigenP 
            Height          =   300
            Left            =   4680
            TabIndex        =   56
            Top             =   1125
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Sicmact.EditMoney txtDemandaP 
            Height          =   315
            Left            =   8610
            TabIndex        =   57
            Top             =   360
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "0"
         End
         Begin VB.Label Label35 
            Caption         =   "Monto Demanda (control) :"
            Height          =   375
            Left            =   5760
            TabIndex        =   69
            Top             =   390
            Width           =   1935
         End
         Begin VB.Label lblLabelP 
            Caption         =   "Tipo :"
            Height          =   255
            Left            =   4320
            TabIndex        =   68
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label Label34 
            Caption         =   "Calificación :"
            Height          =   255
            Left            =   6840
            TabIndex        =   67
            Top             =   1605
            Width           =   975
         End
         Begin VB.Label Label33 
            Caption         =   "Nº Informe :"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label32 
            Caption         =   "Personal a Cargo : "
            Height          =   480
            Left            =   120
            TabIndex        =   65
            Top             =   690
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "Fecha Informe :"
            Height          =   255
            Left            =   3240
            TabIndex        =   64
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label lblNombrePersP 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2760
            TabIndex        =   63
            Top             =   725
            Width           =   2840
         End
         Begin VB.Label Label29 
            Caption         =   "Monto Real Pérdida (gastos) :"
            Height          =   255
            Left            =   5760
            TabIndex        =   62
            Top             =   795
            Width           =   2175
         End
         Begin VB.Label Label28 
            Caption         =   "Fecha Doc. :"
            Height          =   255
            Left            =   3720
            TabIndex        =   61
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "Doc. Origen : "
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label26 
            Caption         =   "Entidad : "
            Height          =   255
            Left            =   6120
            TabIndex        =   59
            Top             =   1170
            Width           =   735
         End
         Begin VB.Label Label25 
            Caption         =   "Nombre Referencial : "
            Height          =   375
            Left            =   120
            TabIndex        =   58
            Top             =   1440
            Width           =   975
         End
      End
      Begin VB.Frame Frame0 
         Caption         =   " Datos Generales "
         Height          =   1215
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   10095
         Begin VB.Label lblOrigen 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1080
            TabIndex        =   24
            Top             =   360
            Width           =   2595
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Registro :"
            Height          =   255
            Left            =   4200
            TabIndex        =   23
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Monto : "
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Origen :"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   395
            Width           =   615
         End
         Begin VB.Label Label5 
            Caption         =   "Usuario : "
            Height          =   255
            Left            =   4200
            TabIndex        =   20
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lblMontoReg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1520
            TabIndex        =   19
            Top             =   680
            Width           =   1155
         End
         Begin VB.Label lblMoneda 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1080
            TabIndex        =   18
            Top             =   680
            Width           =   390
         End
         Begin VB.Label lblFechaReg 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   5520
            TabIndex        =   17
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label lblUsuario 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   5520
            TabIndex        =   16
            Top             =   675
            Width           =   1155
         End
         Begin VB.Label txtDesc 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Left            =   6840
            TabIndex        =   15
            Top             =   360
            Width           =   2955
         End
      End
      Begin VB.Frame FraActivo 
         Caption         =   " Datos del Informe Técnico "
         Height          =   4215
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   10095
         Begin VB.Frame Frame2 
            Caption         =   " Cuentas Contables Ejercicio Actual "
            Height          =   1095
            Left            =   720
            TabIndex        =   81
            Top             =   2040
            Width           =   8535
            Begin Sicmact.TxtBuscar txtBuscaCtaCodAcreedorA 
               Height          =   330
               Left            =   1200
               TabIndex        =   82
               Top             =   600
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   582
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Enabled         =   0   'False
               TipoBusqueda    =   5
               EnabledText     =   0   'False
            End
            Begin Sicmact.TxtBuscar txtBuscaCtaCodDeudorA 
               Height          =   330
               Left            =   1200
               TabIndex        =   83
               Top             =   240
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   582
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Enabled         =   0   'False
               TipoBusqueda    =   5
               EnabledText     =   0   'False
            End
            Begin VB.Label txtCtaDescAcreedorA 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3600
               TabIndex        =   87
               Top             =   600
               Width           =   4755
            End
            Begin VB.Label txtCtaDescDeudorA 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3600
               TabIndex        =   86
               Top             =   240
               Width           =   4755
            End
            Begin VB.Label Label8 
               Caption         =   "Deudora :"
               Height          =   255
               Left            =   240
               TabIndex        =   85
               Top             =   300
               Width           =   855
            End
            Begin VB.Label Label9 
               Caption         =   "Acreedora :"
               Height          =   255
               Left            =   240
               TabIndex        =   84
               Top             =   680
               Width           =   855
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   " Cuentas Contables Ejercicio Anterior "
            Height          =   735
            Left            =   720
            TabIndex        =   35
            Top             =   3240
            Width           =   8535
            Begin Sicmact.TxtBuscar txtBuscaCtaCodAcreedoraAntA 
               Height          =   330
               Left            =   1200
               TabIndex        =   75
               Top             =   240
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   582
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   0   'False
               Enabled         =   0   'False
               TipoBusqueda    =   5
               EnabledText     =   0   'False
            End
            Begin VB.Label txtCtaDescDeudorAntA 
               BackColor       =   &H8000000E&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   3600
               TabIndex        =   88
               Top             =   240
               Width           =   4755
            End
            Begin VB.Label Label10 
               Caption         =   "Acreedora :"
               Height          =   255
               Left            =   240
               TabIndex        =   37
               Top             =   350
               Width           =   855
            End
         End
         Begin VB.TextBox txtNombreRefA 
            Height          =   285
            Left            =   1080
            TabIndex        =   33
            Top             =   1500
            Width           =   5295
         End
         Begin VB.TextBox txtEntidadA 
            Height          =   285
            Left            =   7680
            TabIndex        =   31
            Top             =   1125
            Width           =   2115
         End
         Begin VB.TextBox txtDocOrigenA 
            Height          =   285
            Left            =   1080
            TabIndex        =   27
            Top             =   1120
            Width           =   2775
         End
         Begin VB.ComboBox cboMonedaA 
            Height          =   315
            ItemData        =   "frmContingInformeTecReg.frx":03BF
            Left            =   7680
            List            =   "frmContingInformeTecReg.frx":03C9
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   725
            Width           =   735
         End
         Begin VB.TextBox txtNroInformeA 
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            Top             =   360
            Width           =   2415
         End
         Begin VB.ComboBox cboCalifA 
            Height          =   315
            Left            =   7680
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   360
            Width           =   2175
         End
         Begin Sicmact.EditMoney txtProvisionA 
            Height          =   315
            Left            =   7750
            TabIndex        =   6
            Top             =   725
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
         Begin Sicmact.TxtBuscar TxtBuscarPersA 
            Height          =   330
            Left            =   1080
            TabIndex        =   7
            Top             =   725
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   582
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
            Enabled         =   0   'False
            Enabled         =   0   'False
            TipoBusqueda    =   7
            TipoBusPers     =   1
            EnabledText     =   0   'False
         End
         Begin MSMask.MaskEdBox txtFecInformeA 
            Height          =   285
            Left            =   5160
            TabIndex        =   26
            Top             =   360
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFecOrigenA 
            Height          =   300
            Left            =   5160
            TabIndex        =   28
            Top             =   1125
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "Nombre Referencial : "
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Entidad : "
            Height          =   255
            Left            =   6720
            TabIndex        =   32
            Top             =   1170
            Width           =   735
         End
         Begin VB.Label Label6 
            Caption         =   "Doc. Origen : "
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Doc. :"
            Height          =   255
            Left            =   4200
            TabIndex        =   29
            Top             =   1170
            Width           =   1095
         End
         Begin VB.Label lblLabelA 
            Caption         =   "Monto :"
            Height          =   255
            Left            =   6720
            TabIndex        =   13
            Top             =   795
            Width           =   615
         End
         Begin VB.Label lblNombrePersA 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2760
            TabIndex        =   12
            Top             =   720
            Width           =   3555
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha Informe :"
            Height          =   255
            Left            =   3960
            TabIndex        =   11
            Top             =   390
            Width           =   1215
         End
         Begin VB.Label Label14 
            Caption         =   "Personal a Cargo : "
            Height          =   480
            Left            =   120
            TabIndex        =   10
            Top             =   690
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Nº Informe :"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   390
            Width           =   855
         End
         Begin VB.Label Label12 
            Caption         =   "Calificación :"
            Height          =   255
            Left            =   6720
            TabIndex        =   8
            Top             =   405
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmContingInformeTecReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmContingInformeTecReg
'** Descripción : Registro de Informes Tecnicos para Contingencias creado segun RFC056-2012
'** Creación : JUEZ, 20120621 09:00:00 AM
'********************************************************************

Option Explicit
Dim rs As ADODB.Recordset
Dim oConting As DContingencia
Dim oGen As DGeneral
Dim sNumRegistro As String
Dim psOpeCod As String
Dim nTipoConting As Integer

Public Sub RegistroActivo(ByVal psNumRegistro As String)
    sNumRegistro = psNumRegistro
    If CargaDatos(True, False) Then
        Dim rsCtaCont As ADODB.Recordset
        Set oGen = New DGeneral
        Set oConting = New DContingencia
        Set rs = oGen.GetConstante(5091)
        Call CargaCombo(cboCalifA, rs)
        nTipoConting = gActivoContingente
        TxtBuscarPersA.Enabled = True
        Call CargaControlesCtaCont
        
        If oConting.ObtieneCantInformesTecxConting(sNumRegistro) > 0 Then
            Set rs = oConting.ObtieneUltimoInformeTecConting(sNumRegistro)
            Set rsCtaCont = oConting.BuscaCtaContxInformeTecnico(rs!nIdInfTec, sNumRegistro)
            txtBuscaCtaCodDeudorA.Text = rsCtaCont!cCtaContDeudorA
            txtBuscaCtaCodDeudorA_EmiteDatos
            txtBuscaCtaCodAcreedorA.Text = rsCtaCont!cCtaContAcreedorA
            txtBuscaCtaCodAcreedorA_EmiteDatos
            txtBuscaCtaCodAcreedoraAntA.Text = rsCtaCont!cCtaContDeudorAntA
            txtBuscaCtaCodAcreedoraAntA_EmiteDatos
        End If
        Me.Show 1
    End If
End Sub

Public Sub RegistroPasivo(ByVal psNumRegistro As String)
    sNumRegistro = psNumRegistro
    If CargaDatos(False, True) Then
        Dim rsCtaCont As ADODB.Recordset
        Set oGen = New DGeneral
        Set rs = oGen.GetConstante(5092)
        Call CargaCombo(cboCalifP, rs)
        Set rs = oGen.GetConstante(5093)
        Call CargaCombo(cboTipoP, rs)
        nTipoConting = gPasivoContingente
        txtBuscaPersP.Enabled = True
        txtDemandaP.Enabled = True
        Call CargaControlesCtaCont
        
        If oConting.ObtieneCantInformesTecxConting(sNumRegistro) > 0 Then
            Set rs = oConting.ObtieneUltimoInformeTecConting(sNumRegistro)
            Set rsCtaCont = oConting.BuscaCtaContxInformeTecnico(rs!nIdInfTec, sNumRegistro)
            TxtBuscaCtaCodDeudorP.Text = rsCtaCont!cCtaContDeudorP
            TxtBuscaCtaCodDeudorP_EmiteDatos
            TxtBuscaCtaCodAcreedorP.Text = rsCtaCont!cCtaContAcreedorP
            TxtBuscaCtaCodAcreedorP_EmiteDatos
            TxtBuscaCtaCodGastosP.Text = rsCtaCont!cCtaContGastosP
            TxtBuscaCtaCodGastosP_EmiteDatos
            TxtBuscaCtaCodProvP.Text = rsCtaCont!cCtaContProvP
            TxtBuscaCtaCodProvP_EmiteDatos
            TxtBuscaCtaCodAcreedorAntP.Text = rsCtaCont!cCtaContAcreedorAntP
            TxtBuscaCtaCodAcreedorAntP_EmiteDatos
        End If
        Me.Show 1
    End If
End Sub

Private Sub cboCalif_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If nTipoConting = gActivoContingente Then
            cboMonedaA.SetFocus
        Else
            cboTipoP.SetFocus
        End If
    End If
End Sub

Private Sub cboCalifA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtBuscarPersA.SetFocus
    End If
End Sub

Private Sub cboCalifP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtBuscaCtaCodDeudorP.SetFocus
    End If
End Sub

Private Sub cboMonedaA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtProvisionA.SetFocus
    End If
End Sub

'JUEZ 20131028 ********************************************
Private Sub cboMonedaDemP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDemandaP.SetFocus
    End If
End Sub
'END JUEZ *************************************************

Private Sub cboMonedaP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoP.SetFocus
    End If
End Sub

Private Sub cboTipoP_Click()
    'If Trim(Right(cboTipoP.Text, 2)) = 1 Then
        'txtDemandaP.Enabled = True
        'txtDemandaP.value = 0
        'txtDemandaP.SetFocus
    'Else
        'txtDemandaP.Enabled = False
        'txtDemandaP.value = Format(Me.txtMontoP.value, "#,##0.00")
    'End If
End Sub

Private Sub cboTipoP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'If Trim(Right(cboTipoP.Text, 2)) = 1 Then
            'txtDemandaP.Enabled = True
            'txtDemandaP.value = 0
            'txtDemandaP.SetFocus
        'Else
            'txtDemandaP.Enabled = False
            'txtDemandaP.value = Format(Me.txtMontoP.value, "#,##0.00")
            'txtDemandaP_KeyPress (13)
            
        'End If
        cboCalifP.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdRegistrar_Click()
    If ValidaDatos Then
        If MsgBox("Está seguro de registrar el informe técnico de la Contingencia? ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
               
        Dim sMensaje As String
        Dim pbGeneraAsiento As Boolean
        Dim cUser As String
        Dim oMov As DMov
        Dim oImpr As NContImprimir
        Dim oPrevio As clsPrevioFinan
        Set oPrevio = New clsPrevioFinan
        Set oConting = New DContingencia
        Set oMov = New DMov
        Set oImpr = New NContImprimir
        
        gdFecha = gdFecSis
        gsMovNro = oMov.GeneraMovNro(gdFecha, gsCodAge, gsCodUser)
        
        If nTipoConting = gActivoContingente Then
            cUser = oConting.ObtenerUserxCodPersEncargado(TxtBuscarPersA.Text)
        Else
            cUser = oConting.ObtenerUserxCodPersEncargado(txtBuscaPersP.Text)
        End If
        If cUser <> "" Then
            If nTipoConting = gActivoContingente Then
                Call oConting.RegistraInformeTecActivo(sNumRegistro, Trim(txtNroInformeA.Text), txtFecInformeA.Text, cUser, Trim(Right(cboCalifA.Text, 3)), _
                                                        Trim(Right(cboMonedaA.Text, 3)), txtProvisionA.value, Trim(txtEntidadA.Text), _
                                                        Trim(txtDocOrigenA.Text), txtFecOrigenA.Text, Trim(txtNombreRefA.Text), _
                                                        txtBuscaCtaCodDeudorA.Text, txtBuscaCtaCodAcreedorA.Text, txtBuscaCtaCodAcreedoraAntA.Text, _
                                                        psOpeCod, gsMovNro, sMensaje, pbGeneraAsiento)
            Else
                Call oConting.RegistraInformeTecPasivo(sNumRegistro, Trim(txtNroInformeP.Text), txtFecInformeP.Text, cUser, Trim(Right(cboCalifP.Text, 3)), Trim(Right(cboTipoP.Text, 3)), _
                                                        Trim(Right(cboMonedaP.Text, 3)), txtMontoP.value, Trim(txtEntidadP.Text), Trim(Right(cboMonedaDemP.Text, 3)), txtDemandaP.value, _
                                                         Trim(txtDocOrigenP.Text), txtFecOrigenP.Text, Trim(txtNombreRefP.Text), _
                                                        TxtBuscaCtaCodDeudorP.Text, TxtBuscaCtaCodAcreedorP.Text, TxtBuscaCtaCodGastosP.Text, TxtBuscaCtaCodProvP.Text, TxtBuscaCtaCodAcreedorAntP.Text, _
                                                        psOpeCod, gsMovNro, sMensaje, pbGeneraAsiento)
                'JUEZ 20131028 Se agregó Trim(Right(cboMonedaDemP.Text, 3))
            End If
            If sMensaje = "" Then
                MsgBox "Se ha registrado con exito el IT de la Contingencia", vbInformation, "Aviso"
                If pbGeneraAsiento Then
                    oPrevio.Show oImpr.ImprimeAsientoContable(gsMovNro, 66, 79), gsOpeDesc, False, 66, gImpresora
                Else
                    MsgBox "El registro no generó Asiento Contable", vbInformation, "Aviso"
                End If
                Unload Me
            Else
                MsgBox sMensaje, vbInformation, "Aviso!"
            End If
        Else
            MsgBox "La persona elegida no puede ser la encargada", vbInformation, "Aviso!"
        End If
    End If
End Sub

Private Sub cmdVistaPrevia_Click()
    If ValidaDatos Then
        Dim sMensaje As String
        Dim pbGeneraAsiento As Boolean
        Dim cUser As String
        Dim oMov As DMov
        Set oConting = New DContingencia
        Set oMov = New DMov
        Dim oImpr As NContImprimir
        Dim oPrevio As clsPrevioFinan
        Set oPrevio = New clsPrevioFinan
        Set oImpr = New NContImprimir
        
        Dim DCtaCont As String
        Dim HCtaCont As String
        Dim DCtaContDem As String
        Dim HCtaContDem As String
        Dim nMovImporte As Currency
        Dim nMovImporteDem As Currency
              
        gdFecha = gdFecSis
        gsMovNro = oMov.GeneraMovNro(gdFecha, gsCodAge, gsCodUser)
        
        If nTipoConting = gActivoContingente Then
            cUser = oConting.ObtenerUserxCodPersEncargado(TxtBuscarPersA.Text)
        Else
            cUser = oConting.ObtenerUserxCodPersEncargado(txtBuscaPersP.Text)
        End If
        If cUser <> "" Then
            If nTipoConting = gActivoContingente Then
                Call oConting.RegistraInformeTecActivo(sNumRegistro, Trim(txtNroInformeA.Text), txtFecInformeA.Text, cUser, Trim(Right(cboCalifA.Text, 3)), _
                                                        Trim(Right(cboMonedaA.Text, 3)), txtProvisionA.value, Trim(txtEntidadA.Text), _
                                                        Trim(txtDocOrigenA.Text), txtFecOrigenA.Text, Trim(txtNombreRefA.Text), _
                                                        txtBuscaCtaCodDeudorA.Text, txtBuscaCtaCodAcreedorA.Text, txtBuscaCtaCodAcreedoraAntA.Text, _
                                                        psOpeCod, gsMovNro, sMensaje, pbGeneraAsiento, True, DCtaCont, HCtaCont, nMovImporte, DCtaContDem, HCtaContDem, nMovImporteDem)
            Else
                Call oConting.RegistraInformeTecPasivo(sNumRegistro, Trim(txtNroInformeP.Text), txtFecInformeP.Text, cUser, Trim(Right(cboCalifP.Text, 3)), Trim(Right(cboTipoP.Text, 3)), _
                                                        Trim(Right(cboMonedaP.Text, 3)), txtMontoP.value, Trim(txtEntidadP.Text), Trim(Right(cboMonedaDemP.Text, 3)), txtDemandaP.value, _
                                                         Trim(txtDocOrigenP.Text), txtFecOrigenP.Text, Trim(txtNombreRefP.Text), _
                                                        TxtBuscaCtaCodDeudorP.Text, TxtBuscaCtaCodAcreedorP.Text, TxtBuscaCtaCodGastosP.Text, TxtBuscaCtaCodProvP.Text, TxtBuscaCtaCodAcreedorAntP.Text, _
                                                        psOpeCod, gsMovNro, sMensaje, pbGeneraAsiento, True, DCtaCont, HCtaCont, nMovImporte, DCtaContDem, HCtaContDem, nMovImporteDem)
            End If
            If pbGeneraAsiento Then
                    oPrevio.Show oImpr.VistaPreviaAsientoContable(gsMovNro, 66, 79, , , , , , , psOpeCod, sNumRegistro, IIf(nTipoConting = gActivoContingente, Trim(Right(cboMonedaA.Text, 3)), Trim(Right(cboMonedaP.Text, 3))), DCtaCont, HCtaCont, nMovImporte, DCtaContDem, HCtaContDem, nMovImporteDem), gsOpeDesc, False, 66, gImpresora
                Else
                    MsgBox "El registro no generará Asiento Contable", vbInformation, "Aviso"
                End If
        Else
            MsgBox "La persona elegida no puede ser la encargada", vbInformation, "Aviso!"
        End If
    End If
End Sub



Private Sub txtBuscaPersP_EmiteDatos()
    If txtBuscaPersP.Text = "" Then
        lblNombrePersP.Caption = "" 'JUEZ 20131028
        Exit Sub
    End If
    
    Dim rsDatos As New ADODB.Recordset
    Set oGen = New DGeneral
    Set rsDatos = oGen.GetDataUser(txtBuscaPersP.Text, True)
    If rsDatos.RecordCount = 0 Then
        MsgBox "Persona no Válida o no se Encuentra Registrada como Trabajador en la Institución", vbInformation, "Aviso"
        txtBuscaPersP.Text = ""
        lblNombrePersP.Caption = "" 'JUEZ 20131028
    Else
        lblNombrePersP.Caption = rsDatos!cPersNombre
        'txtDemandaP.SetFocus
        cboMonedaP.SetFocus 'JUEZ 20131028
    End If
End Sub

Private Sub txtBuscarPersA_EmiteDatos()
    If TxtBuscarPersA.Text = "" Then
        lblNombrePersA.Caption = "" 'JUEZ 20131028
        Exit Sub
    End If
    
    Dim rsDatos As New ADODB.Recordset
    Set oGen = New DGeneral
    Set rsDatos = oGen.GetDataUser(TxtBuscarPersA.Text, True)
    If rsDatos.RecordCount = 0 Then
        MsgBox "Persona no Válida o no se Encuentra Registrada como Trabajador en la Institución", vbInformation, "Aviso"
        TxtBuscarPersA.Text = ""
        lblNombrePersA.Caption = "" 'JUEZ 20131028
    Else
        lblNombrePersA.Caption = rsDatos!cPersNombre
        cboMonedaA.SetFocus
    End If
End Sub

Public Function CargaDatos(ByVal pbActivo As Boolean, ByVal pbPasivo As Boolean) As Boolean
    psOpeCod = gRegistroInfTec
    
    Set oConting = New DContingencia
    Set rs = oConting.BuscaContigenciaSeleccionada(sNumRegistro)
    If rs.RecordCount = 0 Then
        CargaDatos = False
        Exit Function
    End If
    lblOrigen.Caption = rs!cOrigen
    lblFechaReg.Caption = Format(rs!dFechaReg, "dd/mm/yyyy")
    lblMoneda.Caption = rs!cmoneda
    lblUsuario.Caption = rs!cUserReg
    lblMontoReg.Caption = Format(rs!nMonto, "#,##0.00")
    txtDesc.Caption = rs!cContigDesc
    
    'lblFecInforme.Caption = gdFecSis
    
    lblLabelA.Visible = pbActivo
    cboMonedaA.Visible = pbActivo
    txtProvisionA.Visible = pbActivo
    lblLabelP.Visible = pbPasivo
    cboTipoP.Visible = pbPasivo
    FraActivo.Visible = pbActivo
    FraPasivo.Visible = pbPasivo
    CargaDatos = True
End Function

Private Sub txtBuscaCtaCodDeudorA_EmiteDatos()
    txtCtaDescDeudorA = txtBuscaCtaCodDeudorA.psDescripcion
'If txtCtaDescDeudorA <> "" Then
'   txtBuscaCtaCodAcreedorA.SetFocus
'End If
End Sub

Private Sub txtBuscaCtaCodAcreedorA_EmiteDatos()
txtCtaDescAcreedorA = txtBuscaCtaCodAcreedorA.psDescripcion
'If txtCtaDescAcreedorA <> "" Then
'   txtBuscaCtaCodAcreedoraAntA.SetFocus
'End If
End Sub

Private Sub txtBuscaCtaCodAcreedoraAntA_EmiteDatos()
txtCtaDescDeudorAntA = txtBuscaCtaCodAcreedoraAntA.psDescripcion
'If txtCtaDescDeudorAntA <> "" Then
'   cmdVistaPrevia.SetFocus
'End If
End Sub
'************************************************************
Private Sub TxtBuscaCtaCodDeudorP_EmiteDatos()
txtCtaDescDeudorP = TxtBuscaCtaCodDeudorP.psDescripcion
'If txtCtaDescDeudorP <> "" Then
'   TxtBuscaCtaCodAcreedorP.SetFocus
'End If
End Sub

Private Sub TxtBuscaCtaCodAcreedorP_EmiteDatos()
txtCtaDescAcreedorP = TxtBuscaCtaCodAcreedorP.psDescripcion
'If txtCtaDescAcreedorP <> "" Then
'   TxtBuscaCtaCodGastosP.SetFocus
'End If
End Sub

Private Sub TxtBuscaCtaCodGastosP_EmiteDatos()
txtCtaDescGastosP = TxtBuscaCtaCodGastosP.psDescripcion
'If txtCtaDescGastosP <> "" Then
'   TxtBuscaCtaCodProvP.SetFocus
'End If
End Sub

Private Sub TxtBuscaCtaCodProvP_EmiteDatos()
txtCtaDescProvP = TxtBuscaCtaCodProvP.psDescripcion
'If txtCtaDescProvP <> "" Then
'   TxtBuscaCtaCodAcreedorAntP.SetFocus
'End If
End Sub

Private Sub TxtBuscaCtaCodAcreedorAntP_EmiteDatos()
txtCtaDescAcreedorAntP = TxtBuscaCtaCodAcreedorAntP.psDescripcion
'If txtCtaDescAcreedorAntP <> "" Then
'   cmdVistaPrevia.SetFocus
'End If
End Sub

Private Sub txtDemandaP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'cboCalifP.SetFocus
        'txtDocOrigenP.SetFocus
        txtBuscaPersP.SetFocus 'JUEZ 20131028
    End If
End Sub

Private Sub txtDocOrigenA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecOrigenA.SetFocus
    End If
End Sub

Private Sub txtDocOrigenP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecOrigenP.SetFocus
    End If
End Sub

Private Sub txtEntidadA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNombreRefA.SetFocus
    End If
End Sub

Private Sub txtEntidadP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNombreRefP.SetFocus
    End If
End Sub

Private Sub txtFecInformeA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboCalifA.SetFocus
    End If
End Sub

Private Sub txtFecInformeA_LostFocus()
    Dim sMsj As String
    sMsj = ValidaFecha(txtFecInformeA.Text)
    If Not Trim(sMsj) = "" Then
        MsgBox sMsj, vbInformation, "Aviso"
        If txtFecInformeA.Enabled Then
            txtFecInformeA.SetFocus
            Exit Sub
        End If
    End If
    If CDate(txtFecInformeA.Text) >= gdFecSis Then
        MsgBox "Fecha del informe No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtFecInformeA.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtFecInformeP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'cboMonedaP.SetFocus
        cboMonedaDemP.SetFocus 'JUEZ 20131028
    End If
End Sub

Private Sub txtFecInformeP_LostFocus()
    Dim sMsj As String
    sMsj = ValidaFecha(txtFecInformeP.Text)
    If Not Trim(sMsj) = "" Then
        MsgBox sMsj, vbInformation, "Aviso"
        If txtFecInformeP.Enabled Then
            txtFecInformeP.SetFocus
            Exit Sub
        End If
    End If
    If CDate(txtFecInformeP.Text) >= gdFecSis Then
        MsgBox "Fecha del informe No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtFecInformeP.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtFecOrigenA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntidadA.SetFocus
    End If
End Sub

Private Sub txtFecOrigenA_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFecOrigenA.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        If txtFecOrigenA.Enabled Then
            txtFecOrigenA.SetFocus
            Exit Sub
        End If
    End If
    If CDate(txtFecOrigenA.Text) >= gdFecSis Then
        MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtFecOrigenA.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtFecOrigenP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntidadP.SetFocus
    End If
End Sub

Private Sub txtFecOrigenP_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFecOrigenP.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        If txtFecOrigenP.Enabled Then txtFecOrigenP.SetFocus
        Exit Sub
    End If
    If CDate(txtFecOrigenP.Text) >= gdFecSis Then
        MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
        txtFecOrigenP.SetFocus
        Exit Sub
    End If
End Sub

'Private Sub txtMontoDemandaP_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        txtEntidadP.SetFocus
'    End If
'End Sub

Private Sub txtMontoP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
'        If cboTipoP.ListIndex > -1 Then
'            If Trim(Right(cboTipoP.Text, 2)) = 1 Then
'                txtDemandaP.Enabled = True
'                txtDemandaP.value = 0
'                txtDemandaP.SetFocus
'            Else
'                txtDemandaP.Enabled = False
'                txtDemandaP.value = Format(txtMontoP.value, "#,##0.00")
'                txtBuscaPersP.SetFocus
'            End If
'        End If
        'txtBuscaPersP.SetFocus
        txtDocOrigenP.SetFocus 'JUEZ 20131028
    End If
End Sub

Private Sub txtNombreRefA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBuscaCtaCodDeudorA.SetFocus
    End If
End Sub

Private Sub txtNombreRefP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTipoP.SetFocus
    End If
End Sub

Private Sub txtNroInformeA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecInformeA.SetFocus
    End If
End Sub

Private Sub txtNroInformeP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecInformeP.SetFocus
    End If
End Sub

Private Sub txtProvisionA_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDocOrigenA.SetFocus
    End If
End Sub

'Private Sub txtProvisionP_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If Trim(Right(Me.cboTipoP.Text, 3)) = "1" Then
'            txtMontoDemandaP.SetFocus
'        Else
'            txtEntidadP.SetFocus
'        End If
'    End If
'End Sub

Public Function ValidaDatos() As Boolean
    If nTipoConting = gActivoContingente Then
        If Trim(txtNroInformeA.Text) = "" Then
            MsgBox "Falta ingresar el Nro. de informe", vbInformation, "Aviso"
            txtNroInformeA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFecInformeA.Text) = "__/__/____" Then
            MsgBox "Falta ingresar la fecha del informe", vbInformation, "Aviso"
            txtFecInformeA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If cboCalifA.ListIndex = -1 Then
            MsgBox "Falta seleccionar la calificacion", vbInformation, "Aviso"
            cboCalifA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(TxtBuscarPersA.Text) = "" Then
            MsgBox "Falta ingresar a la persona", vbInformation, "Aviso"
            TxtBuscarPersA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Len(Trim(TxtBuscarPersA.Text)) <> 13 Then
            MsgBox "El codigo de la persona no es correcta", vbInformation, "Aviso"
            TxtBuscarPersA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If cboMonedaA.ListIndex = -1 Then
            MsgBox "Falta seleccionar la moneda", vbInformation, "Aviso"
            cboMonedaA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtProvisionA.value = 0 Then
            MsgBox "Falta ingresar el monto", vbInformation, "Aviso"
            txtProvisionA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtDocOrigenA.Text) = "" Then
            MsgBox "Falta ingresar el origen del documento", vbInformation, "Aviso"
            txtDocOrigenA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFecOrigenA.Text) = "__/__/____" Then
            MsgBox "Falta ingresar la fecha", vbInformation, "Aviso"
            txtFecOrigenA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtEntidadA.Text) = "" Then
            MsgBox "Falta ingresar la entidad", vbInformation, "Aviso"
            txtEntidadA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtNombreRefA.Text) = "" Then
            MsgBox "Falta ingresar el nombre de referencia", vbInformation, "Aviso"
            txtNombreRefA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtBuscaCtaCodDeudorA.Text) = "" Then
            MsgBox "Falta ingresar la Cta.Contable Deudora del Ejercicio Actual", vbInformation, "Aviso"
            txtBuscaCtaCodDeudorA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtBuscaCtaCodAcreedorA.Text) = "" Then
            MsgBox "Falta ingresar la Cta.Contable Acreedora del Ejercicio Actual", vbInformation, "Aviso"
            txtBuscaCtaCodAcreedorA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtBuscaCtaCodAcreedoraAntA.Text) = "" Then
            MsgBox "Falta ingresar la Cta.Contable Deudora del Ejercicio Anterior", vbInformation, "Aviso"
            txtBuscaCtaCodAcreedoraAntA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If oConting.VerificaSiExisteNroInformeTec(sNumRegistro, Trim(txtNroInformeA.Text)) = True Then
            MsgBox "El Informe Tecnico para esta contingencia ya existe!!", vbInformation, "Aviso!"
            txtNroInformeA.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    Else
        If Trim(txtNroInformeP.Text) = "" Then
            MsgBox "Falta ingresar el Nro. de informe", vbInformation, "Aviso"
            txtNroInformeP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFecInformeP.Text) = "__/__/____" Then
            MsgBox "Falta ingresar la fecha del informe", vbInformation, "Aviso"
            txtFecInformeP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If cboMonedaP.ListIndex = -1 Then
            MsgBox "Falta seleccionar la moneda", vbInformation, "Aviso"
            cboMonedaP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtMontoP.value = 0 Then
            MsgBox "Falta ingresar el monto", vbInformation, "Aviso"
            txtMontoP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtBuscaPersP.Text) = "" Then
            MsgBox "Falta ingresar a la persona", vbInformation, "Aviso"
            txtBuscaPersP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Len(Trim(txtBuscaPersP.Text)) <> 13 Then
            MsgBox "El codigo de la persona no es correcta", vbInformation, "Aviso"
            txtBuscaPersP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtDocOrigenP.Text) = "" Then
            MsgBox "Falta ingresar el origen del documento", vbInformation, "Aviso"
            txtDocOrigenP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFecOrigenP.Text) = "__/__/____" Then
            MsgBox "Falta ingresar la fecha", vbInformation, "Aviso"
            txtFecOrigenP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtEntidadP.Text) = "" Then
            MsgBox "Falta ingresar la entidad", vbInformation, "Aviso"
            txtEntidadP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtNombreRefP.Text) = "" Then
            MsgBox "Falta ingresar el nombre de referencia", vbInformation, "Aviso"
            txtNombreRefP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If cboCalifP.ListIndex = -1 Then
            MsgBox "Falta seleccionar la calificacion", vbInformation, "Aviso"
            cboCalifP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If cboTipoP.ListIndex = -1 Then
            MsgBox "Falta seleccionar el tipo", vbInformation, "Aviso"
            cboTipoP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        'If Trim(Right(cboTipoP.Text, 3)) = "1" Then
        If txtDemandaP.value = 0 Then
            MsgBox "Falta ingresar el monto de la demanda", vbInformation, "Aviso"
            txtDemandaP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        'End If
        If Trim(TxtBuscaCtaCodDeudorP.Text) = "" Then
            MsgBox "Falta ingresar la Cta.Contable Deudora de control del Ejercicio Actual", vbInformation, "Aviso"
            TxtBuscaCtaCodDeudorP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(TxtBuscaCtaCodAcreedorP.Text) = "" Then
            MsgBox "Falta ingresar la Cta.Contable Acreedora de control del Ejercicio Actual", vbInformation, "Aviso"
            TxtBuscaCtaCodAcreedorP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(TxtBuscaCtaCodGastosP.Text) = "" Then
            MsgBox "Falta ingresar la Cta.Contable para Gastos del Ejercicio Actual", vbInformation, "Aviso"
            TxtBuscaCtaCodGastosP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(TxtBuscaCtaCodProvP.Text) = "" Then
            MsgBox "Falta ingresar la Cta.Contable para provisiones del Ejercicio Actual", vbInformation, "Aviso"
            TxtBuscaCtaCodProvP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(TxtBuscaCtaCodAcreedorAntP.Text) = "" Then
            MsgBox "Falta ingresar la Cta.Contable Acreedora del Ejercicio Anterior", vbInformation, "Aviso"
            TxtBuscaCtaCodAcreedorAntP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If oConting.VerificaSiExisteNroInformeTec(sNumRegistro, Trim(txtNroInformeP.Text)) = True Then
            MsgBox "El Informe Tecnico para esta contingencia ya existe!!", vbInformation, "Aviso!"
            txtNroInformeP.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    End If
    ValidaDatos = True
End Function

Private Sub CargaControlesCtaCont()
Dim oDCtaCont As DCtaCont
Set oDCtaCont = New DCtaCont
If nTipoConting = gActivoContingente Then
    txtBuscaCtaCodDeudorA.psRaiz = "Cuenta Contable Deudor"
    txtBuscaCtaCodDeudorA.rs = oDCtaCont.CargaCtaCont("", "CtaCont")
    txtBuscaCtaCodDeudorA.TipoBusqueda = BuscaGrid
    txtBuscaCtaCodDeudorA.Enabled = True
    txtBuscaCtaCodDeudorA.lbUltimaInstancia = False
    
    txtBuscaCtaCodAcreedorA.rs = oDCtaCont.CargaCtaCont("", "CtaCont")
    txtBuscaCtaCodAcreedorA.TipoBusqueda = BuscaGrid
    txtBuscaCtaCodAcreedorA.Enabled = True
    txtBuscaCtaCodAcreedorA.lbUltimaInstancia = False
    
    txtBuscaCtaCodAcreedoraAntA.rs = oDCtaCont.CargaCtaCont("", "CtaCont")
    txtBuscaCtaCodAcreedoraAntA.TipoBusqueda = BuscaGrid
    txtBuscaCtaCodAcreedoraAntA.Enabled = True
    txtBuscaCtaCodAcreedoraAntA.lbUltimaInstancia = False
Else
    TxtBuscaCtaCodDeudorP.rs = oDCtaCont.CargaCtaCont("", "CtaCont")
    TxtBuscaCtaCodDeudorP.TipoBusqueda = BuscaGrid
    TxtBuscaCtaCodDeudorP.Enabled = True
    TxtBuscaCtaCodDeudorP.lbUltimaInstancia = False
    
    TxtBuscaCtaCodAcreedorP.rs = oDCtaCont.CargaCtaCont("", "CtaCont")
    TxtBuscaCtaCodAcreedorP.TipoBusqueda = BuscaGrid
    TxtBuscaCtaCodAcreedorP.Enabled = True
    TxtBuscaCtaCodAcreedorP.lbUltimaInstancia = False
    
    TxtBuscaCtaCodGastosP.rs = oDCtaCont.CargaCtaCont("", "CtaCont")
    TxtBuscaCtaCodGastosP.TipoBusqueda = BuscaGrid
    TxtBuscaCtaCodGastosP.Enabled = True
    TxtBuscaCtaCodGastosP.lbUltimaInstancia = False
    
    TxtBuscaCtaCodProvP.rs = oDCtaCont.CargaCtaCont("", "CtaCont")
    TxtBuscaCtaCodProvP.TipoBusqueda = BuscaGrid
    TxtBuscaCtaCodProvP.Enabled = True
    TxtBuscaCtaCodProvP.lbUltimaInstancia = False
    
    TxtBuscaCtaCodAcreedorAntP.rs = oDCtaCont.CargaCtaCont("", "CtaCont")
    TxtBuscaCtaCodAcreedorAntP.TipoBusqueda = BuscaGrid
    TxtBuscaCtaCodAcreedorAntP.Enabled = True
    TxtBuscaCtaCodAcreedorAntP.lbUltimaInstancia = False
End If
Set oDCtaCont = Nothing
End Sub
