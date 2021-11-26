VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCredReporte 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Creditos"
   ClientHeight    =   7680
   ClientLeft      =   1515
   ClientTop       =   1515
   ClientWidth     =   9795
   Icon            =   "frmCredReporte.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdImprimirA02 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6210
      TabIndex        =   90
      Top             =   7080
      Width           =   1380
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   225
      Left            =   3120
      TabIndex        =   50
      Top             =   8280
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   397
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCredReporte.frx":030A
   End
   Begin VB.Frame FraA02 
      Height          =   6885
      Index           =   0
      Left            =   5730
      TabIndex        =   2
      Top             =   0
      Width           =   3960
      Begin VB.Frame fraReporte 
         Caption         =   "Tipo de Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1020
         Left            =   90
         TabIndex        =   94
         Top             =   5760
         Visible         =   0   'False
         Width           =   1590
         Begin VB.OptionButton optReporte 
            Caption         =   "Por Credito"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   96
            Top             =   315
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Por Analista"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   97
            Top             =   600
            Width           =   1230
         End
      End
      Begin VB.CommandButton CmdSelecAge 
         Caption         =   "&Agencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2055
         TabIndex        =   92
         Top             =   5985
         Width           =   1380
      End
      Begin VB.Frame FraA02 
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   780
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   4890
         Visible         =   0   'False
         Width           =   1665
         Begin VB.CheckBox ChkMonA02 
            Caption         =   "Soles"
            Height          =   210
            Index           =   0
            Left            =   330
            TabIndex        =   5
            Top             =   240
            Width           =   915
         End
         Begin VB.CheckBox ChkMonA02 
            Caption         =   "Dolares"
            Height          =   210
            Index           =   1
            Left            =   330
            TabIndex        =   4
            Top             =   480
            Width           =   990
         End
      End
      Begin VB.Frame fraProductos 
         Height          =   4890
         Left            =   1860
         TabIndex        =   88
         Top             =   1065
         Visible         =   0   'False
         Width           =   1950
         Begin VB.CheckBox chkProducto 
            Caption         =   "Comercial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   75
            TabIndex        =   89
            Top             =   120
            Width           =   1080
         End
         Begin VB.CheckBox chkComercial 
            Caption         =   "Empresarial"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   91
            Tag             =   "101"
            Top             =   369
            Width           =   1200
         End
         Begin VB.CheckBox chkComercial 
            Caption         =   "Pesquero"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   93
            Tag             =   "102"
            Top             =   618
            Width           =   1080
         End
         Begin VB.CheckBox chkComercial 
            Caption         =   "Agropecuario"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   95
            Tag             =   "103"
            Top             =   867
            Width           =   1380
         End
         Begin VB.CheckBox chkComercial 
            Caption         =   "Carta Fianza"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   98
            Tag             =   "121"
            Top             =   1116
            Width           =   1455
         End
         Begin VB.CheckBox chkProducto 
            Caption         =   "Micro Empresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   1
            Left            =   75
            TabIndex        =   99
            Top             =   1365
            Width           =   1455
         End
         Begin VB.CheckBox chkMicroEmpresa 
            Caption         =   "PYME Empresarial"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   101
            Tag             =   "201"
            Top             =   1614
            Width           =   1755
         End
         Begin VB.CheckBox chkMicroEmpresa 
            Caption         =   "PYME Pesquero"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   102
            Tag             =   "202"
            Top             =   1863
            Width           =   1740
         End
         Begin VB.CheckBox chkMicroEmpresa 
            Caption         =   "PYME Agropecuario"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   103
            Tag             =   "203"
            Top             =   2112
            Width           =   1740
         End
         Begin VB.CheckBox chkMicroEmpresa 
            Caption         =   "PYME Carta Fianza"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   104
            Tag             =   "221"
            Top             =   2361
            Width           =   1755
         End
         Begin VB.CheckBox chkProducto 
            Caption         =   "Consumo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   270
            Index           =   2
            Left            =   75
            TabIndex        =   105
            Top             =   2610
            Width           =   1080
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Descuento x Planilla"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   106
            Tag             =   "301"
            Top             =   2874
            Width           =   1755
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Garantia Plazo Fijo"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   107
            Tag             =   "302"
            Top             =   3123
            Width           =   1650
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Garantia CTS"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   108
            Tag             =   "303"
            Top             =   3372
            Width           =   1590
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Usos Diversos"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   109
            Tag             =   "304"
            Top             =   3621
            Width           =   1470
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Prestamos Admin."
            Height          =   255
            Index           =   4
            Left            =   150
            TabIndex        =   110
            Tag             =   "320"
            Top             =   3870
            Width           =   1605
         End
         Begin VB.CheckBox chkProducto 
            Caption         =   "Hipotecario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   3
            Left            =   75
            TabIndex        =   111
            Top             =   4119
            Width           =   1515
         End
         Begin VB.CheckBox chkHipotecario 
            Caption         =   "Hipotecaja"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   112
            Tag             =   "401"
            Top             =   4368
            Width           =   1080
         End
         Begin VB.CheckBox chkHipotecario 
            Caption         =   "Mi Vivienda"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   113
            Tag             =   "423"
            Top             =   4620
            Width           =   1350
         End
      End
      Begin VB.Frame FraPagCheque 
         Caption         =   "Pago Con Cheque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1230
         Left            =   60
         TabIndex        =   83
         Top             =   1545
         Visible         =   0   'False
         Width           =   1785
         Begin VB.TextBox TxtNroCheque 
            Enabled         =   0   'False
            Height          =   300
            Left            =   360
            TabIndex        =   86
            Top             =   825
            Width           =   1230
         End
         Begin VB.OptionButton OptPagCheque 
            Caption         =   "Nro Cheque"
            Height          =   210
            Index           =   1
            Left            =   105
            TabIndex        =   85
            Top             =   570
            Width           =   1215
         End
         Begin VB.OptionButton OptPagCheque 
            Caption         =   "General"
            Height          =   210
            Index           =   0
            Left            =   105
            TabIndex        =   84
            Top             =   300
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.Frame FraA02 
         Height          =   3825
         Index           =   2
         Left            =   1860
         TabIndex        =   10
         Top             =   1065
         Visible         =   0   'False
         Width           =   1845
         Begin VB.CheckBox chkCred 
            Caption         =   "CTS"
            Height          =   210
            Index           =   8
            Left            =   360
            TabIndex        =   24
            Tag             =   "303"
            Top             =   2970
            Width           =   1245
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Plazo Fijo"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   23
            Tag             =   "302"
            Top             =   2760
            Width           =   1230
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Dscto x Planilla"
            Height          =   210
            Index           =   6
            Left            =   360
            TabIndex        =   22
            Tag             =   "301"
            Top             =   2535
            Width           =   1425
         End
         Begin VB.CheckBox chkCredConsumo 
            Caption         =   "Consumo"
            Height          =   195
            Left            =   90
            TabIndex        =   21
            Top             =   2265
            Width           =   1530
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Usos Diversos"
            Height          =   210
            Index           =   9
            Left            =   360
            TabIndex        =   20
            Tag             =   "304"
            Top             =   3210
            Width           =   1365
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Administ. Trab."
            Height          =   210
            Index           =   10
            Left            =   360
            TabIndex        =   19
            Tag             =   "320"
            Top             =   3435
            Width           =   1365
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Pesquero"
            Height          =   225
            Index           =   5
            Left            =   345
            TabIndex        =   18
            Tag             =   "103"
            Top             =   1935
            Width           =   1260
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Agropecuario"
            Height          =   225
            Index           =   4
            Left            =   345
            TabIndex        =   17
            Tag             =   "102"
            Top             =   1710
            Width           =   1260
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Empresarial"
            Height          =   225
            Index           =   3
            Left            =   345
            TabIndex        =   16
            Tag             =   "101"
            Top             =   1485
            Width           =   1260
         End
         Begin VB.CheckBox chkCredComercial 
            Caption         =   "Comercial"
            Height          =   285
            Left            =   60
            TabIndex        =   15
            Top             =   1215
            Width           =   1245
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Pesquero"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   14
            Tag             =   "203"
            Top             =   930
            Width           =   1290
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Agropecuario"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   13
            Tag             =   "202"
            Top             =   705
            Width           =   1290
         End
         Begin VB.CheckBox chkCred 
            Caption         =   "Empresarial"
            Height          =   195
            Index           =   0
            Left            =   360
            TabIndex        =   12
            Tag             =   "201"
            Top             =   465
            Width           =   1290
         End
         Begin VB.CheckBox chkCredMES 
            Caption         =   "MicroEmpresa "
            Height          =   255
            Left            =   75
            TabIndex        =   11
            Top             =   165
            Width           =   1485
         End
      End
      Begin VB.Frame FraDiasAtrConsumo 
         Caption         =   "Dias Atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1290
         Left            =   75
         TabIndex        =   72
         Top             =   1125
         Visible         =   0   'False
         Width           =   1650
         Begin VB.TextBox TxtDiasAtrCons3Ini 
            Height          =   285
            Left            =   1125
            TabIndex        =   77
            Text            =   "30"
            Top             =   915
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons2Fin 
            Height          =   285
            Left            =   1110
            TabIndex        =   76
            Text            =   "15"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons2Ini 
            Height          =   285
            Left            =   465
            TabIndex        =   75
            Text            =   "8"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons1Fin 
            Height          =   285
            Left            =   1110
            TabIndex        =   74
            Text            =   "7"
            Top             =   255
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons1Ini 
            Height          =   285
            Left            =   465
            TabIndex        =   73
            Text            =   "1"
            Top             =   255
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Mayor De :"
            Height          =   195
            Left            =   195
            TabIndex        =   82
            Top             =   960
            Width           =   780
         End
         Begin VB.Label Label16 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   81
            Top             =   630
            Width           =   150
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   80
            Top             =   630
            Width           =   210
         End
         Begin VB.Label Label18 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   79
            Top             =   285
            Width           =   150
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   78
            Top             =   285
            Width           =   210
         End
      End
      Begin VB.CommandButton CmdInstitucion 
         Caption         =   "&Instituciones"
         Height          =   450
         Left            =   960
         TabIndex        =   71
         Top             =   3210
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Frame FraIncluirMora 
         Height          =   570
         Left            =   720
         TabIndex        =   69
         Top             =   2565
         Visible         =   0   'False
         Width           =   2205
         Begin VB.CheckBox ChkIncluirMora 
            Caption         =   "Incluir Mora"
            Height          =   255
            Left            =   150
            TabIndex        =   70
            Top             =   195
            Width           =   1890
         End
      End
      Begin VB.Frame fraCredxInstOrden 
         Caption         =   "Ordenar Por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1260
         Left            =   735
         TabIndex        =   65
         Top             =   1305
         Visible         =   0   'False
         Width           =   2205
         Begin VB.OptionButton OptOrdenPagare 
            Caption         =   "&Pagare"
            Height          =   210
            Left            =   240
            TabIndex        =   68
            Top             =   915
            Width           =   1665
         End
         Begin VB.OptionButton OptOrdenAlfabetico 
            Caption         =   "Orden &Alfabetico"
            Height          =   210
            Left            =   240
            TabIndex        =   67
            Top             =   600
            Width           =   1665
         End
         Begin VB.OptionButton OptOrdenCodMod 
            Caption         =   "Codigo &Modular"
            Height          =   210
            Left            =   240
            TabIndex        =   66
            Top             =   300
            Value           =   -1  'True
            Width           =   1665
         End
      End
      Begin VB.Frame FraTipCambio 
         Caption         =   "Tipo Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   720
         Left            =   90
         TabIndex        =   63
         Top             =   1980
         Visible         =   0   'False
         Width           =   1665
         Begin VB.TextBox TxtTipCambio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   270
            TabIndex        =   64
            Text            =   "0.00"
            Top             =   270
            Width           =   1005
         End
      End
      Begin VB.CommandButton CmdUbicacion 
         Caption         =   "&Ubic. Geografica"
         Height          =   420
         Left            =   120
         TabIndex        =   62
         Top             =   4335
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame fraDatosNota 
         Height          =   945
         Left            =   150
         TabIndex        =   54
         Top             =   165
         Visible         =   0   'False
         Width           =   3480
         Begin VB.TextBox TxtNotaFin 
            Height          =   315
            Left            =   1485
            TabIndex        =   61
            Top             =   555
            Width           =   435
         End
         Begin VB.TextBox TxtNotaIni 
            Height          =   315
            Left            =   750
            TabIndex        =   59
            Top             =   555
            Width           =   435
         End
         Begin VB.CheckBox ChkPorc 
            Alignment       =   1  'Right Justify
            Caption         =   "Por Porcentaje"
            Height          =   210
            Left            =   1980
            TabIndex        =   57
            Top             =   270
            Width           =   1365
         End
         Begin VB.TextBox TxtCuotasPend 
            Height          =   315
            Left            =   1485
            TabIndex        =   56
            Top             =   225
            Width           =   435
         End
         Begin VB.Label Label12 
            Caption         =   "Al"
            Height          =   255
            Left            =   1230
            TabIndex        =   60
            Top             =   600
            Width           =   210
         End
         Begin VB.Label Label11 
            Caption         =   "Notas :"
            Height          =   240
            Left            =   120
            TabIndex        =   58
            Top             =   585
            Width           =   525
         End
         Begin VB.Label Label10 
            Caption         =   "Cuotas Pendiente :"
            Height          =   240
            Left            =   90
            TabIndex        =   55
            Top             =   255
            Width           =   1350
         End
      End
      Begin VB.CommandButton CmdAnalistas 
         Caption         =   "&Analistas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2055
         TabIndex        =   49
         Top             =   6390
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Frame FraMoraAnt 
         Height          =   660
         Left            =   90
         TabIndex        =   47
         Top             =   4185
         Visible         =   0   'False
         Width           =   1650
         Begin VB.CheckBox ChkMoraAnt 
            Caption         =   "Mora Anterior"
            Height          =   195
            Left            =   90
            TabIndex        =   48
            Top             =   270
            Width           =   1425
         End
      End
      Begin VB.Frame FraCondicion 
         Caption         =   "Condicion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   90
         TabIndex        =   43
         Top             =   2970
         Visible         =   0   'False
         Width           =   1665
         Begin VB.CheckBox ChkCond 
            Caption         =   "Refinanciado"
            Height          =   210
            Index           =   2
            Left            =   150
            TabIndex        =   46
            Tag             =   "2"
            Top             =   870
            Width           =   1320
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Paralelo"
            Height          =   210
            Index           =   1
            Left            =   150
            TabIndex        =   45
            Tag             =   "3"
            Top             =   600
            Width           =   1320
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Normal"
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   44
            Tag             =   "1"
            Top             =   315
            Width           =   1320
         End
      End
      Begin VB.Frame FraDiasAtr 
         Caption         =   "Dias Atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1680
         Left            =   75
         TabIndex        =   28
         Top             =   1050
         Visible         =   0   'False
         Width           =   1650
         Begin VB.TextBox TxtCar4I 
            Height          =   285
            Left            =   1125
            TabIndex        =   42
            Text            =   "30"
            Top             =   1260
            Width           =   330
         End
         Begin VB.TextBox TxtCar3F 
            Height          =   285
            Left            =   1125
            TabIndex        =   40
            Text            =   "30"
            Top             =   945
            Width           =   330
         End
         Begin VB.TextBox TxtCar3I 
            Height          =   285
            Left            =   480
            TabIndex        =   38
            Text            =   "16"
            Top             =   945
            Width           =   330
         End
         Begin VB.TextBox TxtCar2F 
            Height          =   285
            Left            =   1110
            TabIndex        =   36
            Text            =   "15"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtCar2I 
            Height          =   285
            Left            =   465
            TabIndex        =   34
            Text            =   "8"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtCar1F 
            Height          =   285
            Left            =   1110
            TabIndex        =   32
            Text            =   "7"
            Top             =   255
            Width           =   330
         End
         Begin VB.TextBox TxtCar1I 
            Height          =   285
            Left            =   465
            TabIndex        =   30
            Text            =   "1"
            Top             =   255
            Width           =   330
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Mayor De :"
            Height          =   195
            Left            =   195
            TabIndex        =   41
            Top             =   1305
            Width           =   780
         End
         Begin VB.Label Label8 
            Caption         =   "A"
            Height          =   255
            Left            =   900
            TabIndex        =   39
            Top             =   975
            Width           =   150
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   165
            TabIndex        =   37
            Top             =   975
            Width           =   210
         End
         Begin VB.Label Label6 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   35
            Top             =   630
            Width           =   150
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   630
            Width           =   210
         End
         Begin VB.Label Label4 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   31
            Top             =   285
            Width           =   150
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   29
            Top             =   285
            Width           =   210
         End
      End
      Begin VB.Frame FraA02 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Index           =   3
         Left            =   660
         TabIndex        =   25
         Top             =   2055
         Visible         =   0   'False
         Width           =   2370
         Begin VB.OptionButton OptSaldo 
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   27
            Top             =   210
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.OptionButton OptSaldo 
            Caption         =   "Con Saldos"
            Height          =   255
            Index           =   1
            Left            =   330
            TabIndex        =   26
            Top             =   495
            Width           =   1755
         End
      End
      Begin MSMask.MaskEdBox TxtFecIniA02 
         Height          =   300
         Left            =   1305
         TabIndex        =   6
         Top             =   315
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFecFinA02 
         Height          =   300
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Frame fraDiasAtr2 
         Caption         =   "Dias Atraso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   900
         Left            =   75
         TabIndex        =   51
         Top             =   1080
         Visible         =   0   'False
         Width           =   1650
         Begin VB.TextBox TxtDiaAtrIni 
            Height          =   315
            Left            =   240
            TabIndex        =   53
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox TxtDiasAtrFin 
            Height          =   315
            Left            =   930
            TabIndex        =   52
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final :"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   735
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial :"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   330
         Visible         =   0   'False
         Width           =   990
      End
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7785
      TabIndex        =   1
      Top             =   7065
      Width           =   1380
   End
   Begin VB.Frame Frame1 
      Height          =   7560
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5565
      Begin MSComctlLib.ImageList imglstFiguras 
         Left            =   300
         Top             =   5820
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReporte.frx":038C
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReporte.frx":06DE
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReporte.frx":0A30
               Key             =   "Hijito"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReporte.frx":0D82
               Key             =   "Bebe"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TVRep 
         Height          =   7185
         Left            =   75
         TabIndex        =   100
         Top             =   195
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   12674
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         ImageList       =   "imglstFiguras"
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label14 
      Caption         =   "Label14"
      Height          =   465
      Left            =   5775
      TabIndex        =   114
      Top             =   6975
      Width           =   285
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      SizeMode        =   1  'Stretch
      TabIndex        =   87
      Top             =   -15
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmCredReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MatAgencias() As String
Private MatProductos() As String
Private MatCondicion() As String
Private matAnalista() As String
Private MatInstitucion() As String
Private sUbicacionGeo As String

Dim sOpePadre As String
Dim sOpeHijo As String
Dim sOpeHijito As String

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Public Sub Inicia(ByVal sCaption As String)
 
    Me.Caption = sCaption
    LlenaArbol
    Me.Show 1
End Sub

Private Sub LlenaArbol()
Dim clsGen As DGeneral
Dim rsUsu As Recordset
Dim sOperacion As String, sOpeCod As String
Dim nodOpe As Node
Dim lsTipREP As String

    lsTipREP = "108"
    
    Set clsGen = New DGeneral
    Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
    Set clsGen = Nothing
      
    Do While Not rsUsu.EOF
        sOpeCod = rsUsu("cOpeCod")
        sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
        Select Case rsUsu("nOpeNiv")
            Case "1"
                sOpePadre = "P" & sOpeCod
                Set nodOpe = TVRep.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
                nodOpe.Tag = sOpeCod
            Case "2"
                sOpeHijo = "H" & sOpeCod
                Set nodOpe = TVRep.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
                nodOpe.Tag = sOpeCod
            Case "3"
                sOpeHijito = "J" & sOpeCod
                Set nodOpe = TVRep.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
                nodOpe.Tag = sOpeCod
            Case "4"
                Set nodOpe = TVRep.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
                nodOpe.Tag = sOpeCod
        End Select
        rsUsu.MoveNext
    Loop
    rsUsu.Close
    Set rsUsu = Nothing
End Sub

Private Sub HabilitaControleFrame(ByVal pbTxtFecIni As Boolean, ByVal pbTxtFecFin As Boolean, _
        pbFraProd As Boolean, pbFraMoneda As Boolean, ByVal pbFraSaldos As Boolean, _
        Optional ByVal pbFraDiasAtraso As Boolean = False, Optional pbFraCondicion As Boolean = False, _
        Optional ByVal pbFraMoraAnt As Boolean = False, Optional pbAnalistas As Boolean = False, _
        Optional pbFraDiasAtr2 As Boolean = False, Optional ByVal pbFraDatosNota As Boolean = False, _
        Optional ByVal pbCmdUbicacion As Boolean = False, Optional ByVal pbTipCambio As Boolean = False, _
        Optional ByVal pbfraCredxInstOrden As Boolean = False, Optional ByVal pbFraIncluirMora As Boolean = False, _
        Optional ByVal pbCmdInstitucion As Boolean = False, Optional ByVal pbfradiasatrconsumo As Boolean = False, _
        Optional ByVal pbSoloPrdConsumo As Boolean = False, Optional pbFraPagCheque As Boolean = False, Optional pbFraProductos As Boolean = False, _
        Optional pbFraReporte As Boolean = False, Optional pbcmdAge As Boolean = True)
        
        CmdSelecAge.Visible = pbcmdAge
        FraA02(3).Visible = pbFraSaldos
        TxtFecFinA02.Visible = pbTxtFecFin
        Label3.Visible = pbTxtFecFin
        TxtFecIniA02.Visible = pbTxtFecIni
        Label2.Visible = pbTxtFecIni
        FraA02(2).Visible = pbFraProd
        FraA02(1).Visible = pbFraMoneda
        FraDiasAtr.Visible = pbFraDiasAtraso
        FraCondicion.Visible = pbFraCondicion
        FraMoraAnt.Visible = pbFraMoraAnt
        CmdAnalistas.Visible = pbAnalistas
        fraDiasAtr2.Visible = pbFraDiasAtr2
        fraDatosNota.Visible = pbFraDatosNota
        CmdUbicacion.Visible = pbCmdUbicacion
        FraTipCambio.Visible = pbTipCambio
        fraCredxInstOrden.Visible = pbfraCredxInstOrden
        FraIncluirMora.Visible = pbFraIncluirMora
        CmdInstitucion.Visible = pbCmdInstitucion
        FraDiasAtrConsumo.Visible = pbfradiasatrconsumo
        chkCredMES.Enabled = Not pbSoloPrdConsumo
        chkCred(0).Enabled = Not pbSoloPrdConsumo
        chkCred(1).Enabled = Not pbSoloPrdConsumo
        chkCred(2).Enabled = Not pbSoloPrdConsumo
        chkCredComercial.Enabled = Not pbSoloPrdConsumo
        chkCred(3).Enabled = Not pbSoloPrdConsumo
        chkCred(4).Enabled = Not pbSoloPrdConsumo
        chkCred(5).Enabled = Not pbSoloPrdConsumo
        FraPagCheque.Visible = pbFraPagCheque
        fraProductos.Visible = pbFraProductos
        fraReporte.Visible = pbFraReporte
End Sub
 
Private Sub chkCredComercial_Click()
    If chkCredComercial.Value = 1 Then
        chkCred(3).Value = 1
        chkCred(4).Value = 1
        chkCred(5).Value = 1
    Else
        chkCred(3).Value = 0
        chkCred(4).Value = 0
        chkCred(5).Value = 0
    End If
End Sub

Private Sub chkCredConsumo_Click()
    If chkCredConsumo.Value = 1 Then
        chkCred(6).Value = 1
        chkCred(7).Value = 1
        chkCred(8).Value = 1
        chkCred(9).Value = 1
        chkCred(10).Value = 1
    Else
        chkCred(6).Value = 0
        chkCred(7).Value = 0
        chkCred(8).Value = 0
        chkCred(9).Value = 0
        chkCred(10).Value = 0
    End If
End Sub

Private Sub chkCredMES_Click()
    If chkCredMES.Value = 1 Then
        chkCred(0).Value = 1
        chkCred(1).Value = 1
        chkCred(2).Value = 1
    Else
        chkCred(0).Value = 0
        chkCred(1).Value = 0
        chkCred(2).Value = 0
    End If
End Sub


Private Sub chkProducto_Click(Index As Integer)
Dim i As Integer
    If Index = 0 Then
        For i = 0 To 3
            chkComercial(i).Value = chkProducto(Index).Value
        Next
    ElseIf Index = 1 Then
        For i = 0 To 3
            chkMicroEmpresa(i).Value = chkProducto(Index).Value
        Next
    ElseIf Index = 2 Then
        For i = 0 To 4
            chkConsumo(i).Value = chkProducto(Index).Value
        Next
    ElseIf Index = 3 Then
        For i = 0 To 1
            chkHipotecario(i).Value = chkProducto(Index).Value
        Next
    End If
End Sub

Private Sub CmdAnalistas_Click()
    frmSelectAnalistas.SeleccionaAnalistas
End Sub

Private Sub CmdImprimirA02_Click()
Dim i As Integer
Dim nContAge As Integer
Dim P As Previo.clsPrevio
Dim oNCredDoc As NCredDoc
Dim sCadImp As String
Dim nValTmp As Integer
Dim loRep As NCredReporte
Dim dUltimoDia As DCredReporte
Dim nUltimoDia As Integer
Dim sProductos As String
Dim sMoneda As String
Dim sCondicion As String
Dim sTempo As Integer
Dim lsArchivoN As String
Dim lbLibroOpen As Boolean
Dim sAnalistas As String
Dim nContAna As Integer
Dim nContAgencias As Integer
Dim sAgencias As String
Dim nTempoParam As Byte

Dim oTipCambio As nTipoCambio
                 
    If CmdUbicacion.Visible Then
        If Trim(sUbicacionGeo) = "" Then
            MsgBox "Seleccione una Ubicacion Geografica"
            Exit Sub
        End If
    End If
                                
    If TxtFecIniA02.Visible = True Then
        If IsDate(TxtFecIniA02.Text) = False Then
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            TxtFecIniA02.SetFocus
            Exit Sub
        End If
    End If
    
    If TxtFecFinA02.Visible = True Then
        If IsDate(TxtFecFinA02.Text) = False Then
            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
            TxtFecFinA02.SetFocus
            Exit Sub
        End If
    End If
    
    If CmdAnalistas.Visible Then
        ReDim matAnalista(0)
        nContAge = 0
        nContAna = 0
        For i = 0 To frmSelectAnalistas.LstAnalista.ListCount - 1
            If frmSelectAnalistas.LstAnalista.Selected(i) = True Then
                nContAge = nContAge + 1
                nContAna = nContAna + 1
                ReDim Preserve matAnalista(nContAge)
                matAnalista(nContAge - 1) = Trim(Right(frmSelectAnalistas.LstAnalista.List(i), 20))
            End If
        Next i
        If UBound(matAnalista) = 0 Then
            MsgBox "Debe Seleccionar por lo Menos un Analista"
            Exit Sub
        End If
    End If
    
    If FraA02(1).Visible = True Then
        If ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0 Then
            MsgBox "Seleccione una moneda", vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    If FraCondicion.Visible = True Then
        If ChkCond(0).Value = 0 And ChkCond(1).Value = 0 And ChkCond(2).Value = 0 Then
            MsgBox "Seleccione al menos una condición", vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    sTempo = 0
    
    If fraProductos.Visible = True Then
        For i = 0 To 3
            If chkComercial(i).Value = 1 Then
                sTempo = sTempo + 1
            End If
        Next
        For i = 0 To 3
            If chkMicroEmpresa(i).Value = 1 Then
                sTempo = sTempo + 1
            End If
        Next
        For i = 0 To 4
            If chkConsumo(i).Value = 1 Then
                sTempo = sTempo + 1
            End If
        Next
        For i = 0 To 1
            If chkHipotecario(i).Value = 1 Then
                sTempo = sTempo + 1
            End If
        Next
        If sTempo = 0 Then
            MsgBox "Ud. debe seleccionar al menos un producto", vbExclamation, "Aviso"
            chkProducto(0).SetFocus
            Exit Sub
        End If
    End If
    
    ReDim MatAgencias(0)
    nContAge = 0
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            nContAge = nContAge + 1
            nContAgencias = nContAgencias + 1
            ReDim Preserve MatAgencias(nContAge)
            MatAgencias(nContAge - 1) = Mid(frmSelectAgencias.List1.List(i), 1, 2)
        End If
    Next i
    If nContAge = 0 Then
        ReDim MatAgencias(1)
        nContAgencias = 1
        MatAgencias(0) = gsCodAge
    End If
 
    
    If CmdInstitucion.Visible Then
        ReDim MatInstitucion(0)
        nContAge = 0
        If frmSelectAnalistas.LstAnalista.ListCount > 0 Then
            For i = 0 To frmSelectAnalistas.LstAnalista.ListCount - 1
                If frmSelectAnalistas.LstAnalista.Selected(i) = True Then
                    nContAge = nContAge + 1
                    ReDim Preserve MatInstitucion(nContAge)
                    MatInstitucion(nContAge - 1) = Trim(Right(frmSelectAnalistas.LstAnalista.List(i), 20))
                End If
            Next i
        End If
        If UBound(MatInstitucion) = 0 Then
            MsgBox "Seleccione una Institucion"
            Exit Sub
        End If
    End If
    
    ReDim MatProductos(0)
    nContAge = 0
    For i = 0 To chkCred.Count - 1
        If chkCred(i).Value = 1 Then
            nContAge = nContAge + 1
            ReDim Preserve MatProductos(nContAge)
            MatProductos(nContAge - 1) = Trim(chkCred(i).Tag)
        End If
    Next i
        
    ReDim MatCondicion(0)
    nContAge = 0
    For i = 0 To ChkCond.Count - 1
        If ChkCond(0).Value = 1 Then
            nContAge = nContAge + 1
            ReDim Preserve MatCondicion(nContAge)
            MatCondicion(nContAge - 1) = Trim(ChkCond(i).Tag)
        End If
    Next i
    
    If FraTipCambio.Visible = True Then
        If Val(TxtTipCambio.Text) = 0 Then
            MsgBox "Ingrese un tipo de cambio", vbExclamation, "Aviso"
            TxtTipCambio.SetFocus
            Exit Sub
        End If
    End If
    
    Set oNCredDoc = New NCredDoc
    Set P = New Previo.clsPrevio
    
    Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
    Case gColCredRepIngxPagoCred
        sCadImp = oNCredDoc.ImprimePagodeCreditos(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
    Case gColCredRepDesemEfect
        
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            Else
                sCadImp = oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            End If
        End If
        
    Case gColCredRepSalCarVig
        
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            Else
                sCadImp = oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            End If
        End If
    
    Case gColCredRepCredCancel 'Creditos Cancelados
        HabilitaControleFrame True, True, False, True, True
        If OptSaldo(0).Value Then
            If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
                sCadImp = oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                sCadImp = sCadImp & Chr$(12)
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            Else
                If ChkMonA02(0).Value = 1 Then
                    sCadImp = oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                Else
                    sCadImp = oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
                End If
            End If
        Else
            If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
                sCadImp = oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                sCadImp = sCadImp & Chr$(12)
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            Else
                If ChkMonA02(0).Value = 1 Then
                    sCadImp = oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                Else
                    sCadImp = oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
                End If
            End If
        End If
    Case gColCredRepResSalCarxAna
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text))
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text))
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text))
            End If
        End If
    
    Case gColCredRepMoraInst
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.Value = 1, True, False))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.Value = 1, True, False))
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.Value = 1, True, False))
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.Value = 1, True, False))
            End If
        End If
    Case gColCredRepMoraxAna
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeMoraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeMoraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeMoraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeMoraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista)
            End If
        End If
    Case gColCredRepCredProtes
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            Else
                sCadImp = oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            End If
        End If
    Case gColCredRepCredRetir
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
            End If
        End If
    Case gColCredRepCartaCobMoro1
        rtfCartas.FileName = App.Path & "\FormatoCarta\CartaAMoroso1.txt"
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeCartaMorosos(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), rtfCartas.Text, matAnalista)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCartaMorosos(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), rtfCartas.Text, matAnalista)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeCartaMorosos(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), rtfCartas.Text, matAnalista)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeCartaMorosos(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), rtfCartas.Text, matAnalista)
            End If
        End If
    Case gColCredRepCartaInvCredAlt
        rtfCartas.FileName = App.Path & "\FormatoCarta\CartaInvitacionCreditoParalelo.txt"
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeCartaInvitacionCreditoParalelo(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CInt(TxtNotaIni.Text), CInt(TxtNotaFin.Text), rtfCartas.Text, matAnalista, CInt(TxtCuotasPend.Text), ChkPorc.Value)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCartaInvitacionCreditoParalelo(MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CInt(TxtNotaIni.Text), CInt(TxtNotaFin.Text), rtfCartas.Text, matAnalista, CInt(TxtCuotasPend.Text), ChkPorc.Value)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeCartaInvitacionCreditoParalelo(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CInt(TxtNotaIni.Text), CInt(TxtNotaFin.Text), rtfCartas.Text, matAnalista, CInt(TxtCuotasPend.Text), ChkPorc.Value)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeCartaInvitacionCreditoParalelo(MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CInt(TxtNotaIni.Text), CInt(TxtNotaFin.Text), rtfCartas.Text, matAnalista, CInt(TxtCuotasPend.Text), ChkPorc.Value)
            End If
        End If
    
    Case gColCredRepCredxUbiGeo
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
            End If
        End If
    Case gColCredRepCredVig
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosVigentes(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosVigentes(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text))
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosVigentes(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text))
            Else
                sCadImp = oNCredDoc.ImprimeCreditosVigentes(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text))
            End If
        End If
    Case gColCredRepCredxInst
        If OptOrdenAlfabetico.Value Then
            nValTmp = 1
        End If
        If OptOrdenCodMod.Value Then
            nValTmp = 0
        End If
        If OptOrdenPagare.Value Then
            nValTmp = 2
        End If
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.Value)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.Value)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.Value)
            Else
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.Value)
            End If
        End If
    Case gColCredRepMoraxInst
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
            End If
        End If
    Case gColCredRepResSalCartxAna
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXAnalistaConsumo(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtDiasAtrCons1Ini.Text), _
                                        CInt(TxtDiasAtrCons1Fin.Text), CInt(TxtDiasAtrCons2Ini.Text), CInt(TxtDiasAtrCons2Fin.Text), CInt(TxtDiasAtrCons3Ini.Text), TxtFecIniA02)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeResumenSaldosCarteraXAnalistaConsumo(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtDiasAtrCons1Ini.Text), _
                                        CInt(TxtDiasAtrCons1Fin.Text), CInt(TxtDiasAtrCons2Ini.Text), CInt(TxtDiasAtrCons2Fin.Text), CInt(TxtDiasAtrCons3Ini.Text), TxtFecIniA02)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXAnalistaConsumo(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtDiasAtrCons1Ini.Text), _
                                        CInt(TxtDiasAtrCons1Fin.Text), CInt(TxtDiasAtrCons2Ini.Text), CInt(TxtDiasAtrCons2Fin.Text), CInt(TxtDiasAtrCons3Ini.Text), TxtFecIniA02)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeResumenSaldosCarteraXAnalistaConsumo(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtDiasAtrCons1Ini.Text), _
                                        CInt(TxtDiasAtrCons1Fin.Text), CInt(TxtDiasAtrCons2Ini.Text), CInt(TxtDiasAtrCons2Fin.Text), CInt(TxtDiasAtrCons3Ini.Text), TxtFecIniA02)
            End If
        End If
    
    Case gColCredRepResSaldeCartxInst
        Set oNCredDoc = New NCredDoc
        sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXInstitucionConsumo(MatAgencias, gsCodUser, gdFecSis, gsNomAge, MatProductos)
        
    Case gColCredRepLisDesctoPlanilla
        'Falta el Calculo dela Cuota debe incluir los intereses a la fecha
        'ya que estos creditos son cuota libre
        'Para ello se penso realizar una funcion en sql server para calculo de interes a la fecha
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.Value)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.Value)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.Value)
            Else
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.Value)
            End If
        End If
    
    Case gColCredRepPagosconCheque
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).Value, 0, 1), Trim(TxtNroCheque.Text))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).Value, 0, 1), Trim(TxtNroCheque.Text))
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).Value, 0, 1), Trim(TxtNroCheque.Text))
            Else
                sCadImp = oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).Value, 0, 1), Trim(TxtNroCheque.Text))
            End If
        End If
        
    Case gColCredRepPagosdeOtrasAgen
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
            End If
        End If
    
    Case gColCredRepPagosEnOtrasAgen
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
            End If
        End If
        
    Case gColCredRepIntEnSusp
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).Value = 0 And ChkMonA02(1).Value = 0) Or (ChkMonA02(0).Value = 1 And ChkMonA02(1).Value = 1) Then
            sCadImp = oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
        Else
            If ChkMonA02(0).Value = 1 Then
                sCadImp = oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            Else
                sCadImp = oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            End If
        End If
    Case gColCredRepProgPagosxCuota, gColCredRepDatosReqMora, gColCredRepConsCartAltoRiesgoxAna, gColCredRepConsColocxAnalista, gColCredRepConsColocxAgencia, gColCredRepConsMetAlcanzxAna, gColCredRepConsColocxFteFinan, gColCredRepConsColocyMoraxAna, gColCredRepConsResCartSuper
        
     
            For i = 0 To 3
                If chkComercial(i).Value = 1 Then
                    If Len(Trim(sProductos)) = 0 Then
                        sProductos = "'" & chkComercial(i).Tag & "'"
                    Else
                        sProductos = sProductos & ", '" & chkComercial(i).Tag & "'"
                    End If
                End If
            Next
            For i = 0 To 3
                If chkMicroEmpresa(i).Value = 1 Then
                    If Len(Trim(sProductos)) = 0 Then
                        sProductos = "'" & chkMicroEmpresa(i).Tag & "'"
                    Else
                        sProductos = sProductos & ", '" & chkMicroEmpresa(i).Tag & "'"
                    End If
                End If
            Next
            For i = 0 To 4
                If chkConsumo(i).Value = 1 Then
                    If Len(Trim(sProductos)) = 0 Then
                        sProductos = "'" & chkConsumo(i).Tag & "'"
                    Else
                        sProductos = sProductos & ", '" & chkConsumo(i).Tag & "'"
                    End If
                End If
            Next
            For i = 0 To 1
                If chkHipotecario(i).Value = 1 Then
                    If Len(Trim(sProductos)) = 0 Then
                        sProductos = "'" & chkHipotecario(i).Tag & "'"
                    Else
                        sProductos = sProductos & ", '" & chkHipotecario(i).Tag & "'"
                    End If
                End If
            Next
        'End If
        
        If ChkMonA02(0).Value = 1 Then
            If ChkMonA02(1).Value = 1 Then
                sMoneda = "'" & gMonedaNacional & "', '" & gMonedaExtranjera & "'"
            Else
                sMoneda = "'" & gMonedaNacional & "'"
            End If
        Else
            sMoneda = "'" & gMonedaExtranjera & "'"
        End If
        
        For i = 0 To nContAna - 1
            If i = 0 Then
                sAnalistas = "'" & matAnalista(i) & "'"
            Else
                sAnalistas = sAnalistas & ", '" & matAnalista(i) & "'"
            End If
        Next
        
        If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepProgPagosxCuota Then
            If ChkCond(0).Value = 1 Then
                sCondicion = gColocCredCondNormal
            End If
            If ChkCond(1).Value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = gColocCredCondParalelo
                Else
                    sCondicion = sCondicion & ", " & gColocCredCondParalelo
                End If
            End If
            If ChkCond(2).Value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = gColocCredCondRecurrente
                Else
                    sCondicion = sCondicion & ", " & gColocCredCondRecurrente
                End If
            End If
            
            
            Set loRep = New NCredReporte
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            sCadImp = loRep.nRepo108301_ListadoProgramacionPagosCuota(IIf(optReporte(0).Value = True, 1, 2), TxtFecIniA02.Text, Me.TxtFecFinA02.Text, sMoneda, sProductos, sCondicion, sAnalistas)
        
        ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepDatosReqMora Then
         
            sCadImp = Genera_Reporte(sMoneda, sProductos, sAnalistas)
        ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsCartAltoRiesgoxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAnalista Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAgencia Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsMetAlcanzxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxFteFinan Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocyMoraxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsResCartSuper Then
            For i = 0 To nContAgencias - 1
                If i = 0 Then
                    sAgencias = "'" & MatAgencias(i) & "'"
                Else
                    sAgencias = sAgencias & ", '" & MatAgencias(i) & "'"
                End If
            Next
             
            If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsCartAltoRiesgoxAna Then
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108604_CarteraAltoRiesgoxAnalista(Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAnalista Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAgencia Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsMetAlcanzxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocyMoraxAna Then
                If DateDiff("d", Format(gdFecSis, "dd/MM/YYYY"), Format(Me.TxtFecFinA02.Text, "dd/MM/YYYY")) = 0 Then
                    'la fecha que se busca es la fecha actual
                    nTempoParam = 1
                Else
  
                    Set dUltimoDia = New DCredReporte
                    nUltimoDia = dUltimoDia.RecuperaUltimoDiaMes(Me.TxtFecFinA02.Text)
                    If nUltimoDia = Val(Mid(TxtFecFinA02.Text, 1, 2)) Then
                        'El dia es el ultimo del mes que se especifica
                        If Val(Mid(Format(gdFecSis, "dd/MM/YYYY"), 4, 2)) = Val(Mid(Format(Me.TxtFecFinA02, "dd/MM/YYYY"), 4, 2)) And Val(Mid(Format(gdFecSis, "dd/MM/YYYY"), 7, 4)) = Val(Mid(Format(Me.TxtFecFinA02, "dd/MM/YYYY"), 7, 4)) Then
                            'Es el mismo mes
                            MsgBox "Ud. no puede colocar esta fecha pues en el mes actual solo vale la fecha del sistema", vbExclamation, "Aviso"
                            Me.TxtFecFinA02.SetFocus
                            Exit Sub
                        Else
                            'Es el ultimo dia del mes pasado
                            nTempoParam = 2
                        End If
                    Else
                        MsgBox "La fecha que ud está ingresando no corresponde al último dia de ese mes", vbExclamation, "Aviso"
                        Me.TxtFecFinA02.SetFocus
                        Exit Sub
                    End If
                End If
                 
                'Recalculo el tipo de cambio para la fecha especificada
                Set oTipCambio = New nTipoCambio
                TxtTipCambio.Text = Format(oTipCambio.EmiteTipoCambio(Format(Me.TxtFecFinA02.Text, "dd/MM/YYYY"), TCFijoMes), "0.00")
                Set oTipCambio = Nothing
                  
                  
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                
                If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAnalista Then
                    sCadImp = loRep.nRepo108602_ConsolidadoColocacionesxAnalista(nTempoParam, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAgencia Then
                    sCadImp = loRep.nRepo108601_ConsolidadoColocacionesxAgencia(nTempoParam, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsMetAlcanzxAna Then
                    sCadImp = loRep.nRepo108603_CuadroMetasAlcanzadasxAnalista(nTempoParam, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocyMoraxAna Then
                    sCadImp = loRep.nRepo108606_ConsolidadoColocacionesxMoraxAnalista(nTempoParam, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
                End If
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxFteFinan Then
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108605_ConsolidadoColocxFteFinanciamiento(Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsResCartSuper Then
                sCadImp = Genera_Reporte108607(Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
            End If
        End If
          
    End Select
        If Len(Trim(sCadImp)) <= 1 Then
            MsgBox "No existen datos para el reporte", vbExclamation, "Aviso"
        Else
            P.Show sCadImp, "Reportes de Creditos", True
        End If
    Set P = Nothing
    Set oNCredDoc = Nothing
End Sub

Private Sub CmdInstitucion_Click()
    frmSelectAnalistas.SeleccionaInstituciones
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdSelecAge_Click()
Dim i As Integer
Dim nContAge As Integer

    frmSelectAgencias.Show 1
    ReDim MatAgencias(0)
    nContAge = 0
    For i = 0 To frmSelectAgencias.List1.ListCount - 1
        If frmSelectAgencias.List1.Selected(i) = True Then
            nContAge = nContAge + 1
            ReDim Preserve MatAgencias(nContAge)
            MatAgencias(nContAge - 1) = Mid(frmSelectAgencias.List1.List(i), 1, 2)
        End If
    Next i
End Sub


Private Sub CmdUbicacion_Click()
    sUbicacionGeo = frmUbicacionGeo.Inicio
End Sub

Private Sub Form_Load()

    ReDim MatAgencias(0)
    ReDim MatProductos(0)
    ReDim matAnalista(0)
    ReDim MatInstitucion(0)

    Dim oTipCambio As nTipoCambio
    
    Set oTipCambio = New nTipoCambio
    TxtTipCambio.Text = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "0.00")
    Set oTipCambio = Nothing
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmSelectAgencias
    Unload frmSelectAnalistas
    Unload frmUbicacionGeo
    Set frmCredReportes = Nothing
End Sub

Private Sub OptPagCheque_Click(Index As Integer)
    If Index = 0 Then
        TxtNroCheque.Enabled = False
    Else
        TxtNroCheque.Enabled = True
        TxtNroCheque.Text = ""
    End If
    
End Sub

Private Sub TVRep_Click()
Dim m As Control
Dim sTipo As String
    Limpia
    
    Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
        Case gColCredRepIngxPagoCred
            Call HabilitaControleFrame(True, False, False, True, False)
            CmdSelecAge.Visible = True
        Case gColCredRepDesemEfect
            Call HabilitaControleFrame(True, True, True, True, False)
            CmdSelecAge.Visible = True
        Case gColCredRepSalCarVig
            Call HabilitaControleFrame(True, False, False, True, False)
        Case gColCredRepCredCancel
            Call HabilitaControleFrame(True, True, False, True, True)
        Case gColCredRepResSalCarxAna
            Call HabilitaControleFrame(True, False, True, True, False, True, True)
        Case gColCredRepMoraInst
            Call HabilitaControleFrame(True, False, True, True, False, True, True, True)
        Case gColCredRepMoraxAna
            Call HabilitaControleFrame(True, False, True, True, False, False, True, False, True)
        Case gColCredRepCredProtes
            Call HabilitaControleFrame(True, False, False, True, False, False, False, False, False)
        Case gColCredRepCredRetir
            Call HabilitaControleFrame(True, True, True, True, False, False, True, False, False)
        Case gColCredRepCartaCobMoro1
            Call HabilitaControleFrame(True, False, True, True, False, False, True, False, True, True)
        Case gColCredRepCartaInvCredAlt
            Call HabilitaControleFrame(False, False, True, True, False, False, True, False, True, False, True)
        Case gColCredRepCredxUbiGeo
            Call HabilitaControleFrame(False, False, True, True, False, False, True, False, False, False, False, True)
        Case gColCredRepCredVig
            Call HabilitaControleFrame(True, False, True, True, False, False, True, False, False, True, False, False, True)
        Case gColCredRepCredxInst
            Call HabilitaControleFrame(True, False, False, False, False, False, False, False, False, False, False, False, False, True, True, True)
        Case gColCredRepMoraxInst
            Call HabilitaControleFrame(True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True)
        Case gColCredRepResSalCartxAna
            Call HabilitaControleFrame(True, False, True, False, False, False, True, False, False, False, False, False, False, False, False, False, True, True)
        Case gColCredRepResSaldeCartxInst
            Call HabilitaControleFrame(False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False, True)
            'Ubica el Boton para Seleccionar la Institucion en la posicion Inferior
            CmdInstitucion.Left = 1965
            CmdInstitucion.Top = 5535
        Case gColCredRepLisDesctoPlanilla
            Call HabilitaControleFrame(True, False, False, False, False, False, False, False, False, False, False, False, False, True, True, True)
        Case gColCredRepPagosconCheque
            Call HabilitaControleFrame(True, True, True, True, False, False, True, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False)
            CmdSelecAge.Visible = True
        Case gColCredRepPagosdeOtrasAgen
            Call HabilitaControleFrame(True, True, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepPagosEnOtrasAgen
            Call HabilitaControleFrame(True, True, True, True, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepIntEnSusp
            Call HabilitaControleFrame(True, False, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepProgPagosxCuota
            Call HabilitaControleFrame(True, True, False, True, False, False, True, False, True, False, False, False, False, False, False, False, False, False, False, True, True, False)
        Case gColCredRepDatosReqMora
            Call HabilitaControleFrame(False, False, False, True, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, True, False, False)
        Case gColCredRepConsCartAltoRiesgoxAna, gColCredRepConsColocxAnalista, gColCredRepConsMetAlcanzxAna, gColCredRepConsColocyMoraxAna, gColCredRepConsResCartSuper
            Call HabilitaControleFrame(False, True, False, True, False, False, False, False, True, False, False, False, True, False, False, False, False, False, False, True, False)
        Case gColCredRepConsColocxAgencia, gColCredRepConsColocxFteFinan
            Call HabilitaControleFrame(False, True, False, True, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, True, False)

    End Select
End Sub

Private Sub Limpia()
Dim i As Integer
    
    Call HabilitaControleFrame(False, False, False, False, False)
        
    For i = 0 To 3
        chkProducto(i).Value = 0
        If i < 3 Then
            ChkCond(i).Value = 0
            If i < 2 Then
                ChkMonA02(i).Value = 0
            End If
        End If
    Next
    chkCredMES.Value = 0
    chkCredComercial.Value = 0
    chkCredConsumo.Value = 0
    TxtFecIniA02.Text = Format(gdFecSis, "dd/MM/YYYY")
    TxtFecFinA02.Text = Format(gdFecSis, "dd/MM/YYYY")
    
 End Sub

Private Sub TVRep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TVRep_Click
End If
End Sub

Private Sub TVRep_NodeClick(ByVal Node As MSComctlLib.Node)
    TVRep_Click
End Sub
 
Private Function Genera_Reporte(ByVal psMoneda As String, ByVal psProducto As String, ByVal psAnalistas As String) As String
  
    Dim matFilas() As Long
    Dim matCont As Long
    
    Dim nFila As Long
    Dim i As Long
    
    Dim sTempoAnalista As String
    Dim sTempoMoneda As String
    
    Dim loExc As DCredReporte
    Dim reg As New ADODB.Recordset
    
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    Set loExc = New DCredReporte
    Set reg = loExc.RecuperaListadoMorosos(psMoneda, psProducto, psAnalistas)
    If reg.BOF Then
        Genera_Reporte = ""
        Exit Function
    Else
          
        lsArchivoN = App.Path & "\Spooler\SeguimientoMora" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
        OleExcel.Class = "ExcelWorkSheet"
        lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
        If lbLibroOpen Then
            'Abro...
            Set xlHoja1 = xlLibro.Worksheets(1)
            ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
            
            sTempoAnalista = reg!Analista
            sTempoMoneda = reg!nMoneda
            matCont = 0
            nFila = 1
            
            xlHoja1.Cells(nFila, 1) = gsNomCmac
            nFila = 2
            xlHoja1.Cells(nFila, 1) = gsNomAge
             
            nFila = 3
            xlHoja1.Cells(nFila, 1) = "REPORTE DE SEGUIMIENTO DE MORA"
             
            nFila = 5
            xlHoja1.Cells(nFila, 1) = "Credito"
            xlHoja1.Cells(nFila, 2) = "Cliente"
            xlHoja1.Cells(nFila, 3) = "Direccion"
            xlHoja1.Cells(nFila, 4) = "Zona"
            xlHoja1.Cells(nFila, 5) = "Telefono"
            xlHoja1.Cells(nFila, 6) = "Saldo Capital"
            xlHoja1.Cells(nFila, 7) = "Dias Atraso"
            xlHoja1.Cells(nFila, 8) = "Direc. Fuente Ingreso"
            xlHoja1.Cells(nFila, 9) = "Zona Fuente Ingreso"
            xlHoja1.Cells(nFila, 10) = "Telef. Fuente Ingreso"
            xlHoja1.Cells(nFila, 11) = "Nombre Garante"
            xlHoja1.Cells(nFila, 12) = "Direc. Garante"
            xlHoja1.Cells(nFila, 13) = "Zona Garante"
            
            nFila = nFila + 2
            ReDim Preserve matFilas(0) As Long
            matFilas(0) = nFila
            xlHoja1.Cells(nFila, 1) = "ANALISTA"
            xlHoja1.Cells(nFila, 2) = PstaNombre(reg!cNomAnalista)
            xlHoja1.Cells(nFila, 3) = reg!Analista
            nFila = nFila + 2
            
            xlHoja1.Cells(nFila, 1) = "MONEDA"
            xlHoja1.Cells(nFila, 2) = reg!cMoneda
            xlHoja1.Range("A" & Trim(Str(nFila)) & ":C" & Trim(Str(nFila))).Font.Bold = True
            With reg
                Do While Not reg.EOF
                    If sTempoAnalista <> !Analista Then
                        nFila = nFila + 2
                        matCont = matCont + 1
                        ReDim Preserve matFilas(matCont) As Long
                        matFilas(matCont) = nFila
                        xlHoja1.Cells(nFila, 1) = "ANALISTA"
                        xlHoja1.Cells(nFila, 2) = PstaNombre(reg!cNomAnalista)
                        xlHoja1.Cells(nFila, 3) = reg!Analista
                        nFila = nFila + 2
                        xlHoja1.Cells(nFila, 1) = "MONEDA"
                        xlHoja1.Cells(nFila, 2) = !cMoneda
                        xlHoja1.Range("A" & Trim(Str(nFila)) & ":C" & Trim(Str(nFila))).Font.Bold = True
                        sTempoAnalista = !Analista
                        sTempoMoneda = !nMoneda
                    ElseIf sTempoMoneda <> !nMoneda Then
                        nFila = nFila + 2
                        xlHoja1.Cells(nFila, 1) = "MONEDA"
                        xlHoja1.Cells(nFila, 2) = !cMoneda
                        xlHoja1.Range("A" & Trim(Str(nFila)) & ":C" & Trim(Str(nFila))).Font.Bold = True
                        sTempoMoneda = !nMoneda
                    End If
                    nFila = nFila + 1
                    xlHoja1.Cells(nFila, 1) = !cCtaCod
                    xlHoja1.Cells(nFila, 2) = PstaNombre(!cPersNombre, False)
                    xlHoja1.Cells(nFila, 3) = !cPersDireccDomicilio
                    xlHoja1.Cells(nFila, 4) = !cUbiGeoDescripcion
                    xlHoja1.Cells(nFila, 5) = Str(!cPersTelefono)
                    xlHoja1.Cells(nFila, 6) = Format(!nSaldo, "#,##0.00")
                    xlHoja1.Cells(nFila, 7) = !nDiasAtraso
                    xlHoja1.Cells(nFila, 8) = !cDirFteIngreso
                    xlHoja1.Cells(nFila, 9) = !cZonaFteIngreso
                    xlHoja1.Cells(nFila, 10) = !cFonoFteIngreso
                    xlHoja1.Cells(nFila, 11) = PstaNombre("" & !cNomGarante, False)
                    xlHoja1.Cells(nFila, 12) = !cDirGarante
                    xlHoja1.Cells(nFila, 13) = !cZonaGarante
                    
                    .MoveNext
                Loop
            End With
            reg.Close
            Set reg = Nothing
        
            xlHoja1.Range("A1:B1").MergeCells = True
            xlHoja1.Range("A3:M3").MergeCells = True
             
            xlHoja1.Range("A1:B3").Font.Bold = True
            
            xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
            xlHoja1.Range("A5:M5").HorizontalAlignment = xlCenter
            
            For i = 0 To matCont ' - 1
                 xlHoja1.Range("A" & Trim(Str(matFilas(i))) & ":C" & Trim(Str(matFilas(i)))).Font.Bold = True
                 xlHoja1.Range("A" & Trim(Str(matFilas(i))) & ":C" & Trim(Str(matFilas(i)))).Interior.ColorIndex = 24
            Next
            
            With xlHoja1.Range("A5:M5")
                .Font.Bold = True
                .Borders.LineStyle = xlContinuous
                .Borders.Weight = xlThin
                .Borders.ColorIndex = 0
                .Interior.ColorIndex = 19
            End With
             
            xlHoja1.Cells.Select
            xlHoja1.Cells.Font.Name = "Arial"
            xlHoja1.Cells.Font.Size = 9
            xlHoja1.Cells.EntireColumn.AutoFit
        
            'Cierro...
            OleExcel.Class = "ExcelWorkSheet"
            ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
            OleExcel.SourceDoc = lsArchivoN
            OleExcel.Verb = 1
            OleExcel.Action = 1
            OleExcel.DoVerb -1
        End If
    End If
End Function

Private Function Genera_Reporte108607(ByVal pnTipoCambio_ As Currency, ByVal pdFechaFin_ As String, ByVal psMoneda_ As String, ByVal psProductos_ As String, ByVal psAgencias_ As String, ByVal psAnalistas_ As String) As String
  
'    Dim matFilas() As Long
'    Dim matCont As Long
'
'    Dim nFila As Long
'    Dim i As Long
'
'    Dim sTempoAnalista As String
'    Dim sTempoMoneda As String
    
    Dim loExc As DCredReporte
    Dim reg As New ADODB.Recordset
    
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean

    Dim cMatAgencia() As String
    Dim matAgencia() As String
    Dim cMatAnalista() As String
    Dim matCarNor1() As Long
    Dim matCarNor2() As Currency
    Dim matCarVen1() As Long
    Dim matCarVen2() As Currency
    Dim matCarRef1() As Long
    Dim matCarRef2() As Currency
    Dim matCobJud1() As Long
    Dim matCobJud2() As Currency
    Dim matDesemNue1() As Long
    Dim matDesemNue2() As Currency
    Dim matDesemRepre1() As Long
    Dim matDesemRepre2() As Currency
    Dim matOpeRef1() As Long
    Dim matOpeRef2() As Currency
    Dim nContador As Long
    Dim i As Long
    
    Dim sFilaSubTotal() As Long 'En este arreglo grabare el numero de filas en donde se deberan llenar con subtotales
    Dim sFilaTotal As Long 'Fila del total general
    Dim nContadorSubFila As Long 'cuantos arreglos de subfilas hay
    Dim nFila As Long 'La fila en la cual me estoy moviendo
    
    Dim sTempoAgencia As String
    Dim sTempoAnalista As String
    
    Set loExc = New DCredReporte
    Set reg = loExc.Recupera_ConsolidadoCarteraxAnalista(pnTipoCambio_, pdFechaFin_, psMoneda_, psProductos_, psAgencias_, psAnalistas_)
    If reg.BOF Then
        Genera_Reporte108607 = ""
        Exit Function
    Else
     
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Lleno  Arreglos
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    nContador = 0
    sTempoAgencia = reg!cAgencia
    sTempoAnalista = reg!cAnalista
    
    ReDim Preserve cMatAgencia(nContador) As String
    ReDim Preserve matAgencia(nContador) As String
    ReDim Preserve cMatAnalista(nContador) As String
    ReDim Preserve matCarNor1(nContador) As Long
    ReDim Preserve matCarNor2(nContador) As Currency
    ReDim Preserve matCarVen1(nContador) As Long
    ReDim Preserve matCarVen2(nContador) As Currency
    ReDim Preserve matCarRef1(nContador) As Long
    ReDim Preserve matCarRef2(nContador) As Currency
    ReDim Preserve matCobJud1(nContador) As Long
    ReDim Preserve matCobJud2(nContador) As Currency
    ReDim Preserve matDesemNue1(nContador) As Long
    ReDim Preserve matDesemNue2(nContador) As Currency
    ReDim Preserve matDesemRepre1(nContador) As Long
    ReDim Preserve matDesemRepre2(nContador) As Currency
    ReDim Preserve matOpeRef1(nContador) As Long
    ReDim Preserve matOpeRef2(nContador) As Currency
    
    cMatAnalista(0) = reg!cAnalista
    cMatAgencia(0) = reg!cAgencia
    matAgencia(0) = reg!cDesAgencia
    matAnalista(0) = reg!cNomAnalista
    
    Do While Not reg.EOF
        If sTempoAgencia <> reg!cAgencia Or sTempoAnalista <> reg!cAnalista Then
            nContador = nContador + 1
            ReDim Preserve cMatAgencia(nContador) As String
            ReDim Preserve matAgencia(nContador) As String
            ReDim Preserve cMatAnalista(nContador) As String
            ReDim Preserve matCarNor1(nContador) As Long
            ReDim Preserve matCarNor2(nContador) As Currency
            ReDim Preserve matCarVen1(nContador) As Long
            ReDim Preserve matCarVen2(nContador) As Currency
            ReDim Preserve matCarRef1(nContador) As Long
            ReDim Preserve matCarRef2(nContador) As Currency
            ReDim Preserve matCobJud1(nContador) As Long
            ReDim Preserve matCobJud2(nContador) As Currency
            ReDim Preserve matDesemNue1(nContador) As Long
            ReDim Preserve matDesemNue2(nContador) As Currency
            ReDim Preserve matDesemRepre1(nContador) As Long
            ReDim Preserve matDesemRepre2(nContador) As Currency
            ReDim Preserve matOpeRef1(nContador) As Long
            ReDim Preserve matOpeRef2(nContador) As Currency
    
            
            cMatAnalista(nContador) = reg!cAnalista
            cMatAgencia(nContador) = reg!cCodAgencia
            matAgencia(nContador) = reg!cDesAgencia
            matAnalista(nContador) = reg!cNomAnalista
            
            sTempoAgencia = reg!cAgencia
            sTempoAnalista = reg!cAnalista
        End If
        
        If reg!Lugar = 1 Then
            'Saldo de Cartera Normal
            matCarNor1(nContador) = reg!Cantidad
            matCarNor2(nContador) = reg!Total
        ElseIf reg!Lugar = 2 Then
            'Saldo de Cartera Vencida
            matCarVen1(nContador) = reg!Cantidad
            matCarVen2(nContador) = reg!Total
        ElseIf reg!Lugar = 3 Then
            'Saldo de Cartera Refinanciada
            matCarRef1(nContador) = reg!Cantidad
            matCarRef2(nContador) = reg!Total
        ElseIf reg!Lugar = 4 Then
            'Cobranza Judicial
            matCobJud1(nContador) = reg!Cantidad
            matCobJud2(nContador) = reg!Total
        ElseIf reg!Lugar = 5 Then
            'Desembolsos Nuevos
            matDesemNue1(nContador) = reg!Cantidad
            matDesemNue2(nContador) = reg!Total
        ElseIf reg!Lugar = 6 Then
            'Desembolsos Represtados
            matDesemRepre1(nContador) = reg!Cantidad
            matDesemRepre2(nContador) = reg!Total
        ElseIf reg!Lugar = 7 Then
            'Operaciones Refinanciadas
            matOpeRef1(nContador) = reg!Cantidad
            matOpeRef2(nContador) = reg!Total
        End If
        reg.MoveNext
    Loop
    reg.Close
    Set reg = Nothing

    
    lsArchivoN = App.Path & "\Spooler\ConsolCarteraxAnalista" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
        'Abro...
        Set xlHoja1 = xlLibro.Worksheets(1)
        ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1

        xlHoja1.Cells.Select
        xlHoja1.Cells.Font.Name = "Arial"
        xlHoja1.Cells.Font.Size = 9
        xlHoja1.Cells.EntireColumn.AutoFit

        nFila = 1

        xlHoja1.Cells(nFila, 1) = gsNomCmac
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = gsNomAge

        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = "RESUMEN CONSOLIDADO DE CARTERA POR ANALISTA"

        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = "Al " & pdFechaFin_
        
        nFila = nFila + 1
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = "Funcionario"
        xlHoja1.Cells(nFila, 2) = "Saldo de"
        xlHoja1.Cells(nFila, 4) = "Saldo de"
        xlHoja1.Cells(nFila, 7) = "Saldo de"
        xlHoja1.Cells(nFila, 10) = "Cobranza"
        xlHoja1.Cells(nFila, 12) = "Resultados Mensuales"
        xlHoja1.Cells(nFila, 17) = "Desembolsos"
        xlHoja1.Cells(nFila, 21) = "Total"
        xlHoja1.Cells(nFila, 23) = "Operaciones"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = "de"
        xlHoja1.Cells(nFila, 2) = "Cartera Normal"
        xlHoja1.Cells(nFila, 4) = "Cartera Vencida"
        xlHoja1.Cells(nFila, 7) = "Cartera Refinanciada"
        xlHoja1.Cells(nFila, 10) = "Judicial"
        xlHoja1.Cells(nFila, 12) = "Saldo de"
        xlHoja1.Cells(nFila, 13) = "Saldo de"
        xlHoja1.Cells(nFila, 14) = "Indice"
        xlHoja1.Cells(nFila, 15) = "Saldo"
        xlHoja1.Cells(nFila, 16) = "Indice"
        xlHoja1.Cells(nFila, 17) = "Nuevos"
        xlHoja1.Cells(nFila, 19) = "Represtamos"
        xlHoja1.Cells(nFila, 21) = "Desembolso"
        xlHoja1.Cells(nFila, 23) = "Refinanciadas"
        
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = "Credito"
        xlHoja1.Cells(nFila, 2) = "Nro"
        xlHoja1.Cells(nFila, 3) = "Saldo"
        xlHoja1.Cells(nFila, 4) = "Nro"
        xlHoja1.Cells(nFila, 5) = "Saldo"
        xlHoja1.Cells(nFila, 6) = "%"
        xlHoja1.Cells(nFila, 7) = "Nro"
        xlHoja1.Cells(nFila, 8) = "Saldo"
        xlHoja1.Cells(nFila, 9) = "%"
        xlHoja1.Cells(nFila, 10) = "Nro"
        xlHoja1.Cells(nFila, 11) = "Saldo"
        xlHoja1.Cells(nFila, 12) = "Cartera"
        xlHoja1.Cells(nFila, 13) = "Mora"
        xlHoja1.Cells(nFila, 14) = "Mora"
        xlHoja1.Cells(nFila, 15) = "C.A.R."
        xlHoja1.Cells(nFila, 16) = "C.A.R."
        xlHoja1.Cells(nFila, 17) = "Nro"
        xlHoja1.Cells(nFila, 18) = "Saldo"
        xlHoja1.Cells(nFila, 19) = "Nro"
        xlHoja1.Cells(nFila, 20) = "Saldo"
        xlHoja1.Cells(nFila, 21) = "Nro"
        xlHoja1.Cells(nFila, 22) = "Saldo"
        xlHoja1.Cells(nFila, 23) = "Nro"
        xlHoja1.Cells(nFila, 24) = "Saldo"
         
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = cMatAgencia(0) & " " & matAgencia(0)
        
        sTempoAgencia = cMatAgencia(0)
        nContadorSubFila = 0
        
        Dim sRef As Byte
        sRef = 1
             
        ReDim Preserve sFilaSubTotal(nContadorSubFila)
        sFilaSubTotal(nContadorSubFila) = nFila
         
        For i = 0 To nContador
            If sTempoAgencia <> cMatAgencia(i) Then
                If sRef = 1 Then
                    sRef = 2
                Else
                    nContadorSubFila = nContadorSubFila + 1
                    ReDim Preserve sFilaSubTotal(nContadorSubFila)
                End If
                'Agregar fila en blanco que diga subtotal
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 1) = "SubTotal"
                 
                sFilaSubTotal(nContadorSubFila) = nFila
                
                sTempoAgencia = cMatAgencia(i)
                
                'Agrego fila en blanco
                nFila = nFila + 1
                
                'Agrego fila que diga el nombre de la agencia
                nFila = nFila + 1
                xlHoja1.Cells(nFila, 1) = cMatAgencia(i) & " " & matAgencia(i)
            End If
    
            'Imprimo la fila con los valores normales
    
            nFila = nFila + 1
            xlHoja1.Cells(nFila, 1) = cMatAnalista(i) & " " & matAnalista(i)
            xlHoja1.Cells(nFila, 2) = Format(matCarNor1(i), "#,##0")
            xlHoja1.Cells(nFila, 3) = Format(matCarNor2(i), "#,##0.00")
            xlHoja1.Cells(nFila, 4) = Format(matCarVen1(i), "#,##0")
            xlHoja1.Cells(nFila, 5) = Format(matCarVen2(i), "#,##0.00")
            
            '(6)=(5)/(3) F=E/C
            xlHoja1.Range("F" & Trim(Str(nFila))).Formula = "=$E$" & Trim(Str(nFila)) & "/$C$" & Trim(Str(nFila))
        
            xlHoja1.Cells(nFila, 7) = Format(matCarRef1(i), "#,##0")
            xlHoja1.Cells(nFila, 8) = Format(matCarRef2(i), "#,##0.00")
            
            '(9)=(8)/(3) I=H/C
            xlHoja1.Range("I" & Trim(Str(nFila))).Formula = "=$H$" & Trim(Str(nFila)) & "/$C$" & Trim(Str(nFila))
            
            xlHoja1.Cells(nFila, 10) = Format(matCobJud1(i), "#,##0")
            xlHoja1.Cells(nFila, 11) = Format(matCobJud1(i), "#,##0.00")
            
            '(12)=(3)+(11) L=C+K
            xlHoja1.Range("L" & Trim(Str(nFila))).Formula = "=$C$" & Trim(Str(nFila)) & "+$K$" & Trim(Str(nFila))
            
            '(13)=(5)+(11) M=E+K
            xlHoja1.Range("M" & Trim(Str(nFila))).Formula = "=$E$" & Trim(Str(nFila)) & "+$K$" & Trim(Str(nFila))
            
            '(14)=(13)/(12) N=M/L
            xlHoja1.Range("N" & Trim(Str(nFila))).Formula = "=$M$" & Trim(Str(nFila)) & "/$L$" & Trim(Str(nFila))
             
            '(15)=(5)+(8)+(11) O=E+H+K
            xlHoja1.Range("O" & Trim(Str(nFila))).Formula = "=$E$" & Trim(Str(nFila)) & "+$H$" & Trim(Str(nFila)) & "+$K$" & Trim(Str(nFila))
            
            '(16)=(15)/(12) P=O/L
            xlHoja1.Range("P" & Trim(Str(nFila))).Formula = "=$O$" & Trim(Str(nFila)) & "/$L$" & Trim(Str(nFila))
            
            
            xlHoja1.Cells(nFila, 17) = Format(matDesemNue1(i), "#,##0")
            xlHoja1.Cells(nFila, 18) = Format(matDesemNue2(i), "#,##0.00")
            xlHoja1.Cells(nFila, 19) = Format(matDesemRepre1(i), "#,##0")
            xlHoja1.Cells(nFila, 20) = Format(matDesemRepre2(i), "#,##0.00")
            
            '(21)=(17)+(19) U=Q+S
            xlHoja1.Range("U" & Trim(Str(nFila))).Formula = "=$Q$" & Trim(Str(nFila)) & "+$S$" & Trim(Str(nFila))
            
            '(22)=(18)+(20) V=R+T
            xlHoja1.Range("V" & Trim(Str(nFila))).Formula = "=$R$" & Trim(Str(nFila)) & "+$T$" & Trim(Str(nFila))
            
            xlHoja1.Cells(nFila, 23) = Format(matOpeRef1(i), "#,##0")
            xlHoja1.Cells(nFila, 24) = Format(matOpeRef2(i), "#,##0.00")
             
        Next
         
        'Imprimo fila del ultimo subtotal
        
        nContadorSubFila = nContadorSubFila + 1
        ReDim Preserve sFilaSubTotal(nContadorSubFila)
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = "SubTotal"
        sFilaSubTotal(nContadorSubFila) = nFila
        
        'imprimo fila en blanco
        nFila = nFila + 1
        
        'imprimo fila que diga total
        nFila = nFila + 1
        xlHoja1.Cells(nFila, 1) = "Total"
        sFilaTotal = nFila
        
        xlHoja1.Range("A1:D1").MergeCells = True
        xlHoja1.Range("A2:B2").MergeCells = True
        xlHoja1.Range("A3:X3").MergeCells = True
        xlHoja1.Range("A4:X4").MergeCells = True
        xlHoja1.Range("B6:C6").MergeCells = True
        xlHoja1.Range("D6:F6").MergeCells = True
        xlHoja1.Range("G6:I6").MergeCells = True
        xlHoja1.Range("J6:K6").MergeCells = True
        xlHoja1.Range("L6:P6").MergeCells = True
        xlHoja1.Range("Q6:T6").MergeCells = True
        xlHoja1.Range("U6:V6").MergeCells = True
        xlHoja1.Range("W6:X6").MergeCells = True
        xlHoja1.Range("B7:C7").MergeCells = True
        xlHoja1.Range("D7:F7").MergeCells = True
        xlHoja1.Range("G7:I7").MergeCells = True
        xlHoja1.Range("J7:K7").MergeCells = True
        xlHoja1.Range("Q7:R7").MergeCells = True
        xlHoja1.Range("S7:T7").MergeCells = True
        xlHoja1.Range("U7:V7").MergeCells = True
        xlHoja1.Range("W7:X7").MergeCells = True
 
        xlHoja1.Range("A1:X8").Font.Bold = True

        xlHoja1.Range("A3:X8").HorizontalAlignment = xlCenter
        
        'De los subtotales
        For i = 0 To nContadorSubFila
             xlHoja1.Range("A" & Trim(Str(sFilaSubTotal(i))) & ":X" & Trim(Str(sFilaSubTotal(i)))).Font.Bold = True
             
             
             xlHoja1.Range("B" & Trim(Str(sFilaSubTotal(i)))).Formula = "=SUMA()"
             
        Next
        
        'Del total
        xlHoja1.Range("A" & Trim(Str(sFilaTotal)) & ":X" & Trim(Str(sFilaTotal))).Font.Bold = True
        

''''
''''        With xlHoja1.Range("A5:M5")
''''            .Font.Bold = True
''''            .Borders.LineStyle = xlContinuous
''''            .Borders.Weight = xlThin
''''            .Borders.ColorIndex = 0
''''            .Interior.ColorIndex = 19
''''        End With
''''

        
        

        'Cierro...
        OleExcel.Class = "ExcelWorkSheet"
        ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
        OleExcel.SourceDoc = lsArchivoN
        OleExcel.Verb = 1
        OleExcel.Action = 1
        OleExcel.DoVerb -1
    End If
    End If
End Function

'***********************************************************
' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
'***********************************************************
Private Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox Err.Description, vbInformation, "Aviso"
  ExcelBegin = False
End Function

'***********************************************************
' Final de Trabajo con EXCEL, graba Libro
'***********************************************************
Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub


'********************************
' Adiciona Hoja a LibroExcel
'********************************
Private Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet)
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = psHojName Then
       xlHoja1.Delete
       Exit For
    End If
Next
Set xlHoja1 = xlLibro.Worksheets.Add
xlHoja1.Name = psHojName
End Sub




