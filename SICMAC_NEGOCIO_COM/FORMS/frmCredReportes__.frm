VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes de Creditos"
   ClientHeight    =   7665
   ClientLeft      =   1515
   ClientTop       =   1515
   ClientWidth     =   9795
   Icon            =   "frmCredReportes__.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Left            =   6240
      TabIndex        =   75
      Top             =   7080
      Width           =   1380
   End
   Begin RichTextLib.RichTextBox rtfCartas 
      Height          =   225
      Left            =   8325
      TabIndex        =   35
      Top             =   7125
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   397
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmCredReportes__.frx":030A
   End
   Begin VB.Frame FraA02 
      Height          =   6885
      Index           =   0
      Left            =   5730
      TabIndex        =   2
      Top             =   0
      Width           =   3960
      Begin VB.Frame fraMontoMayor 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   750
         Left            =   90
         TabIndex        =   108
         Top             =   5790
         Visible         =   0   'False
         Width           =   1350
         Begin VB.TextBox txtMontoMayor 
            Height          =   330
            Left            =   90
            TabIndex        =   109
            Top             =   285
            Width           =   1140
         End
      End
      Begin VB.Frame fraEstadistica 
         Caption         =   "Est. Mensual por "
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
         Height          =   1560
         Left            =   120
         TabIndex        =   103
         Top             =   3120
         Visible         =   0   'False
         Width           =   1740
         Begin VB.TextBox txtLineaCredito 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   105
            TabIndex        =   107
            Top             =   1125
            Width           =   1500
         End
         Begin VB.OptionButton optEstadistica 
            Caption         =   "Periodo"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   106
            Top             =   285
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton optEstadistica 
            Caption         =   "Líneas de Crédito"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   105
            Top             =   547
            Width           =   1590
         End
         Begin VB.OptionButton optEstadistica 
            Caption         =   "L.C. Específica"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   104
            Top             =   825
            Width           =   1530
         End
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
         TabIndex        =   36
         Top             =   1080
         Visible         =   0   'False
         Width           =   1650
         Begin VB.TextBox TxtDiaAtrIni 
            Height          =   315
            Left            =   255
            TabIndex        =   38
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox TxtDiasAtrFin 
            Height          =   315
            Left            =   930
            TabIndex        =   37
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraCredVig 
         Caption         =   "Crédito Vigente"
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
         Height          =   1140
         Left            =   90
         TabIndex        =   99
         Top             =   1740
         Visible         =   0   'False
         Width           =   1710
         Begin VB.OptionButton optCredVig 
            Caption         =   "Analista Esp."
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   102
            Top             =   825
            Width           =   1470
         End
         Begin VB.OptionButton optCredVig 
            Caption         =   "Agrup. Analista"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   101
            Top             =   562
            Width           =   1470
         End
         Begin VB.OptionButton optCredVig 
            Caption         =   "Ord. por Cod."
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   100
            Top             =   300
            Value           =   -1  'True
            Width           =   1470
         End
      End
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
         TabIndex        =   79
         Top             =   5760
         Visible         =   0   'False
         Width           =   1590
         Begin VB.OptionButton optReporte 
            Caption         =   "Por Credito"
            Height          =   240
            Index           =   0
            Left            =   240
            TabIndex        =   81
            Top             =   315
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VB.OptionButton optReporte 
            Caption         =   "Por Analista"
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   82
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
         TabIndex        =   77
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
         TabIndex        =   73
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
            TabIndex        =   74
            Top             =   120
            Width           =   1080
         End
         Begin VB.CheckBox chkComercial 
            Caption         =   "Empresarial"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   76
            Tag             =   "101"
            Top             =   369
            Width           =   1200
         End
         Begin VB.CheckBox chkComercial 
            Caption         =   "Pesquero"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   78
            Tag             =   "102"
            Top             =   618
            Width           =   1080
         End
         Begin VB.CheckBox chkComercial 
            Caption         =   "Agropecuario"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   80
            Tag             =   "103"
            Top             =   867
            Width           =   1380
         End
         Begin VB.CheckBox chkComercial 
            Caption         =   "Carta Fianza"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   83
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
            TabIndex        =   84
            Top             =   1365
            Width           =   1455
         End
         Begin VB.CheckBox chkMicroEmpresa 
            Caption         =   "PYME Empresarial"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   86
            Tag             =   "201"
            Top             =   1614
            Width           =   1755
         End
         Begin VB.CheckBox chkMicroEmpresa 
            Caption         =   "PYME Pesquero"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   87
            Tag             =   "202"
            Top             =   1863
            Width           =   1740
         End
         Begin VB.CheckBox chkMicroEmpresa 
            Caption         =   "PYME Agropecuario"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   88
            Tag             =   "203"
            Top             =   2112
            Width           =   1740
         End
         Begin VB.CheckBox chkMicroEmpresa 
            Caption         =   "PYME Carta Fianza"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   89
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
            TabIndex        =   90
            Top             =   2610
            Width           =   1080
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Descuento x Planilla"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   91
            Tag             =   "301"
            Top             =   2874
            Width           =   1755
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Garantia Plazo Fijo"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   92
            Tag             =   "302"
            Top             =   3123
            Width           =   1650
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Garantia CTS"
            Height          =   255
            Index           =   2
            Left            =   150
            TabIndex        =   93
            Tag             =   "303"
            Top             =   3372
            Width           =   1590
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Usos Diversos"
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   94
            Tag             =   "304"
            Top             =   3621
            Width           =   1470
         End
         Begin VB.CheckBox chkConsumo 
            Caption         =   "Prestamos Admin."
            Height          =   255
            Index           =   4
            Left            =   150
            TabIndex        =   95
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
            TabIndex        =   96
            Top             =   4119
            Width           =   1515
         End
         Begin VB.CheckBox chkHipotecario 
            Caption         =   "Hipotecaja"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   97
            Tag             =   "401"
            Top             =   4368
            Width           =   1080
         End
         Begin VB.CheckBox chkHipotecario 
            Caption         =   "Mi Vivienda"
            Height          =   255
            Index           =   1
            Left            =   150
            TabIndex        =   98
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
         TabIndex        =   68
         Top             =   1545
         Visible         =   0   'False
         Width           =   1785
         Begin VB.TextBox TxtNroCheque 
            Enabled         =   0   'False
            Height          =   300
            Left            =   360
            TabIndex        =   71
            Top             =   825
            Width           =   1230
         End
         Begin VB.OptionButton OptPagCheque 
            Caption         =   "Nro Cheque"
            Height          =   210
            Index           =   1
            Left            =   105
            TabIndex        =   70
            Top             =   570
            Width           =   1215
         End
         Begin VB.OptionButton OptPagCheque 
            Caption         =   "General"
            Height          =   210
            Index           =   0
            Left            =   105
            TabIndex        =   69
            Top             =   300
            Value           =   -1  'True
            Width           =   1545
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
         TabIndex        =   57
         Top             =   1110
         Visible         =   0   'False
         Width           =   1650
         Begin VB.TextBox TxtDiasAtrCons3Ini 
            Height          =   285
            Left            =   1125
            TabIndex        =   62
            Text            =   "30"
            Top             =   915
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons2Fin 
            Height          =   285
            Left            =   1110
            TabIndex        =   61
            Text            =   "15"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons2Ini 
            Height          =   285
            Left            =   465
            TabIndex        =   60
            Text            =   "8"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons1Fin 
            Height          =   285
            Left            =   1110
            TabIndex        =   59
            Text            =   "7"
            Top             =   255
            Width           =   330
         End
         Begin VB.TextBox TxtDiasAtrCons1Ini 
            Height          =   285
            Left            =   465
            TabIndex        =   58
            Text            =   "1"
            Top             =   255
            Width           =   330
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Mayor De :"
            Height          =   195
            Left            =   195
            TabIndex        =   67
            Top             =   960
            Width           =   780
         End
         Begin VB.Label Label16 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   66
            Top             =   630
            Width           =   150
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   65
            Top             =   630
            Width           =   210
         End
         Begin VB.Label Label18 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   64
            Top             =   285
            Width           =   150
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   63
            Top             =   285
            Width           =   210
         End
      End
      Begin VB.CommandButton CmdInstitucion 
         Caption         =   "&Instituciones"
         Height          =   450
         Left            =   960
         TabIndex        =   56
         Top             =   3210
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Frame FraIncluirMora 
         Height          =   570
         Left            =   720
         TabIndex        =   54
         Top             =   2565
         Visible         =   0   'False
         Width           =   2205
         Begin VB.CheckBox ChkIncluirMora 
            Caption         =   "Incluir Mora"
            Height          =   255
            Left            =   150
            TabIndex        =   55
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
         TabIndex        =   50
         Top             =   1305
         Visible         =   0   'False
         Width           =   2205
         Begin VB.OptionButton OptOrdenPagare 
            Caption         =   "&Pagare"
            Height          =   210
            Left            =   240
            TabIndex        =   53
            Top             =   915
            Width           =   1665
         End
         Begin VB.OptionButton OptOrdenAlfabetico 
            Caption         =   "Orden &Alfabetico"
            Height          =   210
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Width           =   1665
         End
         Begin VB.OptionButton OptOrdenCodMod 
            Caption         =   "Codigo &Modular"
            Height          =   210
            Left            =   240
            TabIndex        =   51
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
         TabIndex        =   48
         Top             =   1980
         Visible         =   0   'False
         Width           =   1665
         Begin VB.TextBox TxtTipCambio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   270
            TabIndex        =   49
            Text            =   "0.00"
            Top             =   270
            Width           =   1005
         End
      End
      Begin VB.CommandButton CmdUbicacion 
         Caption         =   "&Ubic. Geografica"
         Height          =   420
         Left            =   120
         TabIndex        =   47
         Top             =   4335
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Frame fraDatosNota 
         Height          =   945
         Left            =   150
         TabIndex        =   39
         Top             =   165
         Visible         =   0   'False
         Width           =   3480
         Begin VB.TextBox TxtNotaFin 
            Height          =   315
            Left            =   1485
            TabIndex        =   46
            Top             =   555
            Width           =   435
         End
         Begin VB.TextBox TxtNotaIni 
            Height          =   315
            Left            =   750
            TabIndex        =   44
            Top             =   540
            Width           =   435
         End
         Begin VB.CheckBox ChkPorc 
            Alignment       =   1  'Right Justify
            Caption         =   "Por Porcentaje"
            Height          =   210
            Left            =   1980
            TabIndex        =   42
            Top             =   270
            Width           =   1365
         End
         Begin VB.TextBox TxtCuotasPend 
            Height          =   315
            Left            =   1485
            TabIndex        =   41
            Top             =   225
            Width           =   435
         End
         Begin VB.Label Label12 
            Caption         =   "Al"
            Height          =   255
            Left            =   1230
            TabIndex        =   45
            Top             =   600
            Width           =   210
         End
         Begin VB.Label Label11 
            Caption         =   "Notas :"
            Height          =   240
            Left            =   120
            TabIndex        =   43
            Top             =   585
            Width           =   525
         End
         Begin VB.Label Label10 
            Caption         =   "Cuotas Pendiente :"
            Height          =   240
            Left            =   90
            TabIndex        =   40
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
         TabIndex        =   34
         Top             =   6390
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Frame FraMoraAnt 
         Height          =   660
         Left            =   90
         TabIndex        =   32
         Top             =   4185
         Visible         =   0   'False
         Width           =   1650
         Begin VB.CheckBox ChkMoraAnt 
            Caption         =   "Mora Anterior"
            Height          =   195
            Left            =   90
            TabIndex        =   33
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
         TabIndex        =   28
         Top             =   2970
         Visible         =   0   'False
         Width           =   1665
         Begin VB.CheckBox ChkCond 
            Caption         =   "Refinanciado"
            Height          =   210
            Index           =   2
            Left            =   150
            TabIndex        =   31
            Tag             =   "2"
            Top             =   870
            Width           =   1320
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Paralelo"
            Height          =   210
            Index           =   1
            Left            =   150
            TabIndex        =   30
            Tag             =   "3"
            Top             =   600
            Width           =   1320
         End
         Begin VB.CheckBox ChkCond 
            Caption         =   "Normal"
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   29
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
         TabIndex        =   13
         Top             =   1050
         Visible         =   0   'False
         Width           =   1650
         Begin VB.TextBox TxtCar4I 
            Height          =   285
            Left            =   1125
            TabIndex        =   27
            Text            =   "30"
            Top             =   1260
            Width           =   330
         End
         Begin VB.TextBox TxtCar3F 
            Height          =   285
            Left            =   1125
            TabIndex        =   25
            Text            =   "30"
            Top             =   945
            Width           =   330
         End
         Begin VB.TextBox TxtCar3I 
            Height          =   285
            Left            =   480
            TabIndex        =   23
            Text            =   "16"
            Top             =   945
            Width           =   330
         End
         Begin VB.TextBox TxtCar2F 
            Height          =   285
            Left            =   1110
            TabIndex        =   21
            Text            =   "15"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtCar2I 
            Height          =   285
            Left            =   465
            TabIndex        =   19
            Text            =   "8"
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox TxtCar1F 
            Height          =   285
            Left            =   1110
            TabIndex        =   17
            Text            =   "7"
            Top             =   255
            Width           =   330
         End
         Begin VB.TextBox TxtCar1I 
            Height          =   285
            Left            =   465
            TabIndex        =   15
            Text            =   "1"
            Top             =   255
            Width           =   330
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Mayor De :"
            Height          =   195
            Left            =   195
            TabIndex        =   26
            Top             =   1305
            Width           =   780
         End
         Begin VB.Label Label8 
            Caption         =   "A"
            Height          =   255
            Left            =   900
            TabIndex        =   24
            Top             =   975
            Width           =   150
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   165
            TabIndex        =   22
            Top             =   975
            Width           =   210
         End
         Begin VB.Label Label6 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   20
            Top             =   630
            Width           =   150
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   18
            Top             =   630
            Width           =   210
         End
         Begin VB.Label Label4 
            Caption         =   "A"
            Height          =   255
            Left            =   885
            TabIndex        =   16
            Top             =   285
            Width           =   150
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "De"
            Height          =   195
            Left            =   150
            TabIndex        =   14
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
         TabIndex        =   10
         Top             =   2055
         Visible         =   0   'False
         Width           =   2370
         Begin VB.OptionButton OptSaldo 
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   300
            TabIndex        =   12
            Top             =   210
            Value           =   -1  'True
            Width           =   1755
         End
         Begin VB.OptionButton OptSaldo 
            Caption         =   "Con Saldos"
            Height          =   255
            Index           =   1
            Left            =   330
            TabIndex        =   11
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
               Picture         =   "frmCredReportes__.frx":038C
               Key             =   "Padre"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReportes__.frx":06DE
               Key             =   "Hijo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReportes__.frx":0A30
               Key             =   "Hijito"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCredReportes__.frx":0D82
               Key             =   "Bebe"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView TVRep 
         Height          =   7185
         Left            =   120
         TabIndex        =   85
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
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      SizeMode        =   1  'Stretch
      TabIndex        =   72
      Top             =   -15
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "frmCredReportes"
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
  
Dim WithEvents loRep As NCredReporte
Attribute loRep.VB_VarHelpID = -1
Dim WithEvents lsRep As nCredRepoFinMes
Attribute lsRep.VB_VarHelpID = -1

Dim loRepFM As nCredRepoFinMes
Attribute loRepFM.VB_VarHelpID = -1

Dim Progreso As clsProgressBar
Dim Progress As clsProgressBar

Public Function DescProdConsumoSeleccionado() As String
Dim lsProductos As String
Dim i As Integer
lsProductos = "PRODUCTOS : "
  For i = 0 To Me.chkConsumo.Count - 1
    If chkConsumo(i).value Then
        lsProductos = lsProductos & "/CON-" & Mid(chkConsumo(i).Caption, 1, 5)
    End If
  Next i
DescProdConsumoSeleccionado = lsProductos

End Function
Public Function ValorProdConsumo() As String
Dim i As Integer
Dim lsCad As String

    lsCad = ""
    For i = 0 To Me.chkConsumo.Count - 1
        If chkConsumo(i).value Then
            lsCad = lsCad & "'" & chkConsumo(i).Tag & "',"
        End If
    Next i
    
    If Len(lsCad) > 0 Then
        lsCad = Mid(lsCad, 1, (Len(lsCad) - 1))
        ValorProdConsumo = " AND substring(Credito.cCtaCod,6,3) IN (" & lsCad & ") "
    Else
        ValorProdConsumo = "AND substring(Credito.cCtaCod,6,1) = '3' "
    End If
End Function
Private Function DescProdSeleccionado() As String
Dim lsProductos As String
Dim i As Integer
lsProductos = "PRODUCTOS : "
  
  For i = 0 To chkComercial.Count - 1
    If chkComercial(i).value Then
        lsProductos = lsProductos & "/MES-" & Mid(chkComercial(i).Caption, 1, 3)
    End If
  Next i
  For i = 0 To chkMicroEmpresa.Count - 1
    If chkMicroEmpresa(i).value Then
        lsProductos = lsProductos & "/MES-" & Mid(chkMicroEmpresa(i).Caption, 1, 3)
    End If
  Next i
  For i = 0 To chkConsumo.Count - 1
    If chkConsumo(i).value Then
        lsProductos = lsProductos & "/MES-" & Mid(chkConsumo(i).Caption, 1, 3)
    End If
  Next i
  For i = 0 To chkHipotecario.Count - 1
    If chkHipotecario(i).value Then
        lsProductos = lsProductos & "/MES-" & Mid(chkHipotecario(i).Caption, 1, 4)
    End If
  Next i
DescProdSeleccionado = lsProductos
End Function


Private Function ValorNorRefPar() As String
    ValorNorRefPar = ""
    '**************** CREDITOS NORMALES
    If Me.ChkCond(0).value = 1 And Me.ChkCond(1).value = 0 And Me.ChkCond(2).value = 0 Then
        ValorNorRefPar = " AND (Credito.cRefinan = 'N' and nCondCre in (1,3,2)) "
    End If
    '**************** CREDITOS PARALELOS
    If Me.ChkCond(0).value = 0 And Me.ChkCond(1).value = 1 And Me.ChkCond(2).value = 0 Then
        ValorNorRefPar = " AND (Credito.cRefinan = 'N' and nCondCre = 1) "
    End If
    '*************** CREDITOS REFINANCIADOS
    If Me.ChkCond(0).value = 0 And Me.ChkCond(1).value = 0 And Me.ChkCond(2).value = 1 Then
        ValorNorRefPar = " AND Credito.cRefinan = 'R' "
    End If
    '*************** CREDITOS NORMALES Y PARALELOS
    If Me.ChkCond(0).value = 1 And Me.ChkCond(1).value = 1 And Me.ChkCond(2).value = 0 Then
        ValorNorRefPar = " AND Credito.cRefinan = 'N' "
    End If
    '*************** CREDITOS NORMALES Y REFINANCIADOS
    If Me.ChkCond(0).value = 1 And Me.ChkCond(1).value = 0 And Me.ChkCond(2).value = 1 Then
        ValorNorRefPar = " AND Credito.nCondCre <> 3 "
    End If
    '*************** CREDITOS PARALELOS Y REFINANCIADOS
    If Me.ChkCond(0).value = 0 And Me.ChkCond(1).value = 1 And Me.ChkCond(2).value = 1 Then
        ValorNorRefPar = " AND Credito.nCondCre = 3 OR Credito.cRefinan = 'R' "
    End If
    '*************** TODOS LOS CREDITOS
    If Me.ChkCond(0).value = 1 And Me.ChkCond(1).value = 1 And Me.ChkCond(2).value = 1 Then
        ValorNorRefPar = " "
    End If
End Function
Public Function DescCondSeleccionado() As String
Dim lsCondic As String
    lsCondic = " %%% CONDICION : "
    If Me.ChkCond(0).value = 1 Then  'Normal
        lsCondic = lsCondic & "Norm"
    End If
    If Me.ChkCond(1).value = 1 Then 'Paralelo
        lsCondic = lsCondic & "/Paral"
    End If
    If Me.ChkCond(2).value = 1 Then  'Refinanciado
        lsCondic = lsCondic & "/Refin"
    End If
    If Me.ChkCond(0).value = 0 And Me.ChkCond(1).value = 0 And Me.ChkCond(2).value = 0 Then
        lsCondic = lsCondic & "Norm/Paral/Refin"
    End If
DescCondSeleccionado = lsCondic
End Function

Private Function ValorMoneda() As String

ValorMoneda = ""
If ChkMonA02(0).value = 1 And ChkMonA02(1).value = 0 Then
    ValorMoneda = " AND SUBSTRING(Credito.cCtaCod, 9,1) = '1' "
End If
If ChkMonA02(0).value = 0 And ChkMonA02(1).value = 1 Then
    ValorMoneda = " AND SUBSTRING(Credito.cCtaCod, 9,1) = '2' "
End If
If ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1 Then
    ValorMoneda = " AND SUBSTRING(Credito.cCtaCod, 9,1) in ('1', '2') "
End If

End Function

Private Sub loRep_CloseProgress()
    Progress.CloseForm Me
End Sub


Private Sub loRep_ShowProgress()
    Progress.ShowForm Me
End Sub







Public Sub Inicia(ByVal sCaption As String)
 
    Me.Caption = sCaption
    LlenaArbol
    Me.Show 0
    
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

Private Sub HabilitaControleFrame1(ByVal pbTxtFecIni As Boolean, ByVal pbTxtFecFin As Boolean, _
        pbFraMoneda As Boolean, ByVal pbFraSaldos As Boolean, _
        Optional ByVal pbFraDiasAtraso As Boolean = False, Optional pbFraCondicion As Boolean = False, _
        Optional ByVal pbFraMoraAnt As Boolean = False, Optional pbAnalistas As Boolean = False, _
        Optional pbFraDiasAtr2 As Boolean = False, Optional ByVal pbFraDatosNota As Boolean = False, _
        Optional ByVal pbCmdUbicacion As Boolean = False, Optional ByVal pbTipCambio As Boolean = False, _
        Optional ByVal pbfraCredxInstOrden As Boolean = False, Optional ByVal pbFraIncluirMora As Boolean = False, _
        Optional ByVal pbCmdInstitucion As Boolean = False, Optional ByVal pbfradiasatrconsumo As Boolean = False, _
        Optional ByVal pbSoloPrdConsumo As Boolean = False, Optional pbFraPagCheque As Boolean = False, Optional pbfraProductos As Boolean = False, _
        Optional pbFraReporte As Boolean = False, Optional pbcmdAge As Boolean = True, _
        Optional pbFraCredVig As Boolean = False, Optional pbFraEstadistica As Boolean = False, Optional pbFraMontoMayor As Boolean = False)
         
        CmdSelecAge.Visible = pbcmdAge
        FraA02(3).Visible = pbFraSaldos
        TxtFecFinA02.Visible = pbTxtFecFin
        Label3.Visible = pbTxtFecFin
        TxtFecIniA02.Visible = pbTxtFecIni
        Label2.Visible = pbTxtFecIni
        'FraA02(2).Visible = pbFraProd
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
        fraEstadistica.Visible = pbFraEstadistica
        
        fraMontoMayor.Visible = pbFraMontoMayor
        
        chkProducto(0).Enabled = Not pbSoloPrdConsumo
        chkComercial(0).Enabled = Not pbSoloPrdConsumo
        chkComercial(1).Enabled = Not pbSoloPrdConsumo
        chkComercial(2).Enabled = Not pbSoloPrdConsumo
        chkComercial(3).Enabled = Not pbSoloPrdConsumo
        
        chkProducto(1).Enabled = Not pbSoloPrdConsumo
        chkMicroEmpresa(0).Enabled = Not pbSoloPrdConsumo
        chkMicroEmpresa(1).Enabled = Not pbSoloPrdConsumo
        chkMicroEmpresa(2).Enabled = Not pbSoloPrdConsumo
        chkMicroEmpresa(3).Enabled = Not pbSoloPrdConsumo
        
        chkProducto(3).Enabled = Not pbSoloPrdConsumo
        chkHipotecario(0).Enabled = Not pbSoloPrdConsumo
        chkHipotecario(1).Enabled = Not pbSoloPrdConsumo
        
        FraPagCheque.Visible = pbFraPagCheque
        fraProductos.Visible = pbfraProductos
        fraReporte.Visible = pbFraReporte
        fraCredVig.Visible = pbFraCredVig
End Sub
 
Private Sub chkProducto_Click(Index As Integer)
Dim i As Integer
    If Index = 0 Then
        For i = 0 To 3
            chkComercial(i).value = chkProducto(Index).value
        Next
    ElseIf Index = 1 Then
        For i = 0 To 3
            chkMicroEmpresa(i).value = chkProducto(Index).value
        Next
    ElseIf Index = 2 Then
        For i = 0 To 4
            chkConsumo(i).value = chkProducto(Index).value
        Next
    ElseIf Index = 3 Then
        For i = 0 To 1
            chkHipotecario(i).value = chkProducto(Index).value
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
'--------------------
Dim lsCadenaPar As String
Dim CredRepoMEs As nCredRepoFinMes
Set CredRepoMEs = New nCredRepoFinMes
Dim FMes As Date
Dim Fechaini As String
Dim lsCadenaDesPar As String

Set lsRep = New nCredRepoFinMes

Dim TipoCambio As Currency
Dim fnRepoSelec As Long



Dim oTipCambio As nTipoCambio
                 
    If CmdUbicacion.Visible Then
        If Trim(sUbicacionGeo) = "" Then
            MsgBox "Seleccione una Ubicacion Geografica", vbExclamation, "Aviso"
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
     
    If FraCondicion.Visible = True Then
        If ChkCond(0).value = 0 And ChkCond(1).value = 0 And ChkCond(2).value = 0 Then
            MsgBox "Seleccione al menos una condición", vbExclamation, "Aviso"
            Exit Sub
        End If
    End If
    
    sTempo = 0
        
    If fraDiasAtr2.Visible = True Then
        If IsNumeric(TxtDiaAtrIni.Text) Then
            If IsNumeric(TxtDiasAtrFin.Text) Then
                If Val(TxtDiasAtrFin.Text) < Val(TxtDiaAtrIni.Text) Then
                    MsgBox "El nro. de dias final no puede ser menor al nro. de dias inicial", vbExclamation, "Aviso"
                    TxtDiasAtrFin.SetFocus
                    Exit Sub
                End If
            Else
                MsgBox "Ingrese un nro. de dias válido", vbExclamation, "Aviso"
                TxtDiasAtrCons1Fin.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Ingrese un nro. de dias válido", vbExclamation, "Aviso"
            TxtDiaAtrIni.SetFocus
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
    
    
    For i = 0 To 3
        If chkComercial(i).value = 1 Then
            nContAge = nContAge + 1
            ReDim Preserve MatProductos(nContAge)
            MatProductos(nContAge - 1) = Trim(chkComercial(i).Tag)
        End If
    Next
    
    For i = 0 To 3
        If chkMicroEmpresa(i).value = 1 Then
            nContAge = nContAge + 1
            ReDim Preserve MatProductos(nContAge)
            MatProductos(nContAge - 1) = Trim(chkMicroEmpresa(i).Tag)
        End If
    Next
    
    For i = 0 To 4
        If chkConsumo(i).value = 1 Then
            nContAge = nContAge + 1
            ReDim Preserve MatProductos(nContAge)
            MatProductos(nContAge - 1) = Trim(chkConsumo(i).Tag)
        End If
    Next
    
    For i = 0 To 1
        If chkHipotecario(i).value = 1 Then
            nContAge = nContAge + 1
            ReDim Preserve MatProductos(nContAge)
            MatProductos(nContAge - 1) = Trim(chkHipotecario(i).Tag)
        End If
    Next
 
    ReDim MatCondicion(0)
    nContAge = 0
    For i = 0 To ChkCond.Count - 1
        If ChkCond(0).value = 1 Then
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
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            Else
                sCadImp = oNCredDoc.ImprimeDesembolsosEfectuados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            End If
        End If
        
    Case gColCredRepSalCarVig
        
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            Else
                sCadImp = oNCredDoc.ImprimeSaldoCarteraVigente(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            End If
        End If
    
    Case gColCredRepCredCancel 'Creditos Cancelados
        'HabilitaControleFrame1 True, True, True, True
        If OptSaldo(0).value Then
            If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
                sCadImp = oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                sCadImp = sCadImp & Chr$(12)
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            Else
                If ChkMonA02(0).value = 1 Then
                    sCadImp = oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                Else
                    sCadImp = oNCredDoc.ImprimeCreditosCancelados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
                End If
            End If
        Else
            If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
                sCadImp = oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                sCadImp = sCadImp & Chr$(12)
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            Else
                If ChkMonA02(0).value = 1 Then
                    sCadImp = oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
                Else
                    sCadImp = oNCredDoc.ImprimeCreditosCanceladosConSaldo(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
                End If
            End If
        End If
        
    Case gColCredRepResSalCarxAna
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text))
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text))
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeResumenSaldosCarteraXAnalista(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text))
            End If
        End If
    
    Case gColCredRepMoraInst
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.value = 1, True, False))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.value = 1, True, False))
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.value = 1, True, False))
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeMoraInstitucional(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtCar1I.Text), _
                                        CInt(TxtCar1F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar3I.Text), CInt(TxtCar3F.Text), CInt(TxtCar4I.Text), IIf(ChkMoraAnt.value = 1, True, False))
            End If
        End If
    
    'Modificado Se agrego una segunda opcion
    
    Case gColCredRepMoraxAna, gColCredRepAtraPagoCuotaLib
        
        Set oNCredDoc = New NCredDoc
         
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeMoraXAnalista_AtrasoPagoCuotaLibre(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepMoraxAna, False, True))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeMoraXAnalista_AtrasoPagoCuotaLibre(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepMoraxAna, False, True))
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeMoraXAnalista_AtrasoPagoCuotaLibre(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepMoraxAna, False, True))
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeMoraXAnalista_AtrasoPagoCuotaLibre(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, matAnalista, IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepMoraxAna, False, True))
            End If
        End If
    ''''''''''''''
         
    Case gColCredRepCredProtes
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge)
            Else
                sCadImp = oNCredDoc.ImprimeCreditosProtestados(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge)
            End If
        End If
    Case gColCredRepCredRetir
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosRetirados(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos)
            End If
        End If
    Case gColCredRepCredxUbiGeo
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXUbicacionGeo(MatAgencias, gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, sUbicacionGeo)
            End If
        End If
    
    'Modificado Se agrego una segunda opcion
    Case gColCredRepCredVig, gColCredRepCredVigconCuoLibre
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCredVig_CredVigCuotLib(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, False, True))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCredVig_CredVigCuotLib(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, False, True))
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCredVig_CredVigCuotLib(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, False, True))
            Else
                sCadImp = oNCredDoc.ImprimeCredVig_CredVigCuotLib(MatAgencias, CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatCondicion, MatProductos, CDbl(TxtTipCambio.Text), CInt(TxtDiaAtrIni.Text), CInt(TxtDiasAtrFin.Text), IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, False, True))
            End If
        End If
    ''''''''''''''''''''''''''''''''''''
    
    Case gColCredRepCredxInst
        If OptOrdenAlfabetico.value Then
            nValTmp = 1
        End If
        If OptOrdenCodMod.value Then
            nValTmp = 0
        End If
        If OptOrdenPagare.value Then
            nValTmp = 2
        End If
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
            Else
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
            End If
        End If
    Case gColCredRepMoraxInst
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimeMoraXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatInstitucion)
            End If
        End If
    'Case gColCredRepResSalCartxAna
        'Set oNCredDoc = New NCredDoc
        'If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
        '    sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXAnalistaConsumo(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtDiasAtrCons1Ini.Text), _
        '                                CInt(TxtDiasAtrCons1Fin.Text), CInt(TxtDiasAtrCons2Ini.Text), CInt(TxtDiasAtrCons2Fin.Text), CInt(TxtDiasAtrCons3Ini.Text), TxtFecIniA02)
        '    sCadImp = sCadImp & Chr$(12)
        '    sCadImp = sCadImp & oNCredDoc.ImprimeResumenSaldosCarteraXAnalistaConsumo(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtDiasAtrCons1Ini.Text), _
        '                                CInt(TxtDiasAtrCons1Fin.Text), CInt(TxtDiasAtrCons2Ini.Text), CInt(TxtDiasAtrCons2Fin.Text), CInt(TxtDiasAtrCons3Ini.Text), TxtFecIniA02)
        'Else
        '    If ChkMonA02(0).value = 1 Then
        '        sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXAnalistaConsumo(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtDiasAtrCons1Ini.Text), _
        '                                CInt(TxtDiasAtrCons1Fin.Text), CInt(TxtDiasAtrCons2Ini.Text), CInt(TxtDiasAtrCons2Fin.Text), CInt(TxtDiasAtrCons3Ini.Text), TxtFecIniA02)
        '    Else
        '        sCadImp = sCadImp & oNCredDoc.ImprimeResumenSaldosCarteraXAnalistaConsumo(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, CInt(TxtDiasAtrCons1Ini.Text), _
        '                                CInt(TxtDiasAtrCons1Fin.Text), CInt(TxtDiasAtrCons2Ini.Text), CInt(TxtDiasAtrCons2Fin.Text), CInt(TxtDiasAtrCons3Ini.Text), TxtFecIniA02)
        '    End If
        'End If
    
    Case gColCredRepResSaldeCartxInst
        Set oNCredDoc = New NCredDoc
        sCadImp = oNCredDoc.ImprimeResumenSaldosCarteraXInstitucionConsumo(MatAgencias, gsCodUser, gdFecSis, gsNomAge, MatProductos)
        
    Case gColCredRepLisDesctoPlanilla
        'Falta el Calculo dela Cuota debe incluir los intereses a la fecha
        'ya que estos creditos son cuota libre
        'Para ello se penso realizar una funcion en sql server para calculo de interes a la fecha
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
            Else
                sCadImp = oNCredDoc.ImprimeCreditosXInstitucion(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, nValTmp, ChkIncluirMora.value)
            End If
        End If
    
    Case gColCredRepPagosconCheque
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).value, 0, 1), Trim(TxtNroCheque.Text))
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).value, 0, 1), Trim(TxtNroCheque.Text))
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).value, 0, 1), Trim(TxtNroCheque.Text))
            Else
                sCadImp = oNCredDoc.ImprimePagosConCheque(MatAgencias, CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos, MatCondicion, IIf(OptPagCheque(0).value, 0, 1), Trim(TxtNroCheque.Text))
            End If
        End If
        
    Case gColCredRepPagosdeOtrasAgen
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimePagosDeOtraAgencia(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC)
            End If
        End If
    
    Case gColCredRepPagosEnOtrasAgen
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
            Else
                sCadImp = sCadImp & oNCredDoc.ImprimePagosENOtrasAgencias(CDate(TxtFecIniA02.Text), CDate(TxtFecFinA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, gsCodAge, gsCodCMAC, MatProductos, MatCondicion)
            End If
        End If
        
    Case gColCredRepIntEnSusp
        Set oNCredDoc = New NCredDoc
        If (ChkMonA02(0).value = 0 And ChkMonA02(1).value = 0) Or (ChkMonA02(0).value = 1 And ChkMonA02(1).value = 1) Then
            sCadImp = oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            sCadImp = sCadImp & Chr$(12)
            sCadImp = sCadImp & oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
        Else
            If ChkMonA02(0).value = 1 Then
                sCadImp = oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaNacional, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            Else
                sCadImp = oNCredDoc.ImprimeInteresesSuspenso(MatAgencias, CDate(TxtFecIniA02.Text), gMonedaExtranjera, gsCodUser, gdFecSis, gsNomAge, MatProductos)
            End If
        End If
    Case gColCredRepProgPagosxCuota, gColCredRepDatosReqMora, gColCredRepConsCartAltoRiesgoxAna, gColCredRepConsColocxAnalista, gColCredRepConsColocxAgencia, gColCredRepConsMetAlcanzxAna, gColCredRepConsColocxFteFinan, gColCredRepConsColocyMoraxAna, gColCredRepConsResCartSuper, _
        gColCredRepCartaCobMoro1, gColCredRepCartaCobMoro2, gColCredRepCartaCobMoro3, gColCredRepCartaCobMoro4, gColCredRepCartaCobMoro5, _
        gColCredRepCartaInvCredAlt, gColCredRepCartaRecup, gColCredRepCredVigArqueo, _
        gColCredRepVisitaCobroCuotas, gColCredRepClientesNCuotasPend, gColCredRepIngresosxGasto, gColCredRepCredVigIntDeven, _
        gColCredRepEstMensual, gColCredRepCredDesmMayores, gColCredRepResSalCartxAna
        
        Dim cMensaje1 As String
        Dim cMensaje2 As String
        Dim cMensaje As String

        Dim nBandera As Boolean
            
        cMensaje1 = ""
        cMensaje2 = ""
        nBandera = False
        For i = 0 To 3
            If chkComercial(i).value = 1 Then
                If Len(Trim(sProductos)) = 0 Then
                    sProductos = "'" & chkComercial(i).Tag & "'"
                    cMensaje1 = Mid(chkComercial(i).Caption, 1, 3)
                Else
                    sProductos = sProductos & ", '" & chkComercial(i).Tag & "'"
                    cMensaje1 = cMensaje1 & "/" & Mid(chkComercial(i).Caption, 1, 3)
                End If
                nBandera = True
            Else
                If Len(Trim(cMensaje2)) = 0 Then
                    cMensaje2 = Mid(chkComercial(i).Caption, 1, 3)
                Else
                    cMensaje2 = cMensaje2 & "/" & Mid(chkComercial(i).Caption, 1, 3)
                End If
            End If
        Next
        For i = 0 To 3
            If chkMicroEmpresa(i).value = 1 Then
                If Len(Trim(sProductos)) = 0 Then
                    sProductos = "'" & chkMicroEmpresa(i).Tag & "'"
                    cMensaje1 = Mid(chkMicroEmpresa(i).Caption, 1, 3)
                Else
                    sProductos = sProductos & ", '" & chkMicroEmpresa(i).Tag & "'"
                    cMensaje1 = cMensaje1 & "/" & Mid(chkMicroEmpresa(i).Caption, 1, 3)
                End If
                nBandera = True
            Else
                If Len(Trim(cMensaje2)) = 0 Then
                    cMensaje2 = Mid(chkMicroEmpresa(i).Caption, 1, 3)
                Else
                    cMensaje2 = cMensaje2 & "/" & Mid(chkMicroEmpresa(i).Caption, 1, 3)
                End If
            End If
        Next
        For i = 0 To 4
            If chkConsumo(i).value = 1 Then
                If Len(Trim(sProductos)) = 0 Then
                    sProductos = "'" & chkConsumo(i).Tag & "'"
                    cMensaje1 = Mid(chkConsumo(i).Caption, 1, 3)
                Else
                    sProductos = sProductos & ", '" & chkConsumo(i).Tag & "'"
                    cMensaje1 = cMensaje1 & "/" & Mid(chkConsumo(i).Caption, 1, 3)
                End If
                nBandera = True
            Else
                If Len(Trim(cMensaje2)) = 0 Then
                    cMensaje2 = Mid(chkConsumo(i).Caption, 1, 3)
                Else
                    cMensaje2 = cMensaje2 & "/" & Mid(chkConsumo(i).Caption, 1, 3)
                End If
            End If
        Next
        For i = 0 To 1
            If chkHipotecario(i).value = 1 Then
                If Len(Trim(sProductos)) = 0 Then
                    sProductos = "'" & chkHipotecario(i).Tag & "'"
                    cMensaje1 = Mid(chkHipotecario(i).Caption, 1, 3)
                Else
                    sProductos = sProductos & ", '" & chkHipotecario(i).Tag & "'"
                    cMensaje1 = cMensaje1 & "/" & Mid(chkHipotecario(i).Caption, 1, 3)
                End If
                nBandera = True
            Else
                If Len(Trim(cMensaje2)) = 0 Then
                    cMensaje2 = Mid(chkHipotecario(i).Caption, 1, 3)
                Else
                    cMensaje2 = cMensaje2 & "/" & Mid(chkHipotecario(i).Caption, 1, 3)
                End If
            End If
        Next
        
        If nBandera = True Then
            cMensaje = "PRODUCTOS: " & cMensaje1
        Else
            If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepResSalCartxAna Then
                cMensaje = cMensaje & " PRODUCTOS: " & Mid(chkConsumo(0).Caption, 1, 3) & "/" & Mid(chkConsumo(1).Caption, 1, 3) & "/" & Mid(chkConsumo(2).Caption, 1, 3) & "/" & Mid(chkConsumo(3).Caption, 1, 3) & "/" & Mid(chkConsumo(4).Caption, 1, 3)
                sProductos = "'" & chkConsumo(0).Tag & "', '" & chkConsumo(1).Tag & "', '" & chkConsumo(2).Tag & "', '" & chkConsumo(3).Tag & "', '" & chkConsumo(4).Tag & "'"
            Else
                cMensaje = "PRODUCTOS: " & cMensaje2
            End If
        End If
        
        cMensaje1 = ""
        cMensaje2 = ""
        nBandera = False
        
        If ChkMonA02(0).value = 1 Then
            If ChkMonA02(1).value = 1 Then
                sMoneda = "'" & gMonedaNacional & "', '" & gMonedaExtranjera & "'"
                cMensaje1 = "Nac./Ext."
            Else
                sMoneda = "'" & gMonedaNacional & "'"
                cMensaje1 = "Nac."
            End If
            nBandera = True
        Else
            If ChkMonA02(1).value = 1 Then
                sMoneda = "'" & gMonedaExtranjera & "'"
                cMensaje1 = "Ext."
                nBandera = True
            Else
                sMoneda = ""
                cMensaje2 = "Nac./Ext."
            End If
        End If
        
        If nBandera = True Then
            cMensaje = cMensaje & " MONEDA: " & cMensaje1
        Else
            cMensaje = cMensaje & " MONEDA: " & cMensaje2
        End If
          
        For i = 0 To nContAna - 1
            If i = 0 Then
                sAnalistas = "'" & matAnalista(i) & "'"
                'cMensaje1 = matAnalista(i)
            Else
                sAnalistas = sAnalistas & ", '" & matAnalista(i) & "'"
                'cMensaje1 = cMensaje1 & "/" & matAnalista(i)
            End If
        Next
         
        'cMensaje = cMensaje & " AGENCIAS: " & cMensaje1
        
        If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepProgPagosxCuota Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro1 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro2 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro3 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro4 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro5 Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaInvCredAlt Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaRecup Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVigArqueo Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepVisitaCobroCuotas Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepClientesNCuotasPend Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepIngresosxGasto Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVigIntDeven Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredDesmMayores Or _
           Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepResSalCartxAna Then
            
            cMensaje1 = ""
            cMensaje2 = ""
            nBandera = False
                
            If ChkCond(0).value = 1 Then
                sCondicion = gColocCredCondNormal
                cMensaje1 = "Norm."
                nBandera = True
            Else
                cMensaje2 = "Norm."
            End If
            If ChkCond(1).value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = gColocCredCondParalelo
                    cMensaje1 = "Par."
                Else
                    sCondicion = sCondicion & ", " & gColocCredCondParalelo
                    cMensaje1 = cMensaje1 & "/Par."
                End If
                nBandera = True
            Else
                cMensaje2 = cMensaje2 & "/Par."
            End If
            If ChkCond(2).value = 1 Then
                If Len(Trim(sCondicion)) = 0 Then
                    sCondicion = gColocCredCondRecurrente
                    cMensaje1 = "Rec."
                Else
                    sCondicion = sCondicion & ", " & gColocCredCondRecurrente
                    cMensaje1 = cMensaje1 & "/Rec."
                End If
                nBandera = True
            Else
                cMensaje2 = cMensaje2 & "/Rec."
            End If
            If nBandera = True Then
                cMensaje = cMensaje & " CONDICION: " & cMensaje1
            Else
                cMensaje = cMensaje & " CONDICION: " & cMensaje2
            End If
            
            If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepProgPagosxCuota Then
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108301_ListadoProgramacionPagosCuota(IIf(optReporte(0).value = True, 1, 2), cMensaje, TxtFecIniA02.Text, Me.TxtFecFinA02.Text, sMoneda, sProductos, sCondicion, sAnalistas)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro1 Then
                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro1, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro2 Then
                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro2, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro3 Then
                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro3, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro4 Then
                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro4, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaCobMoro5 Then
                sCadImp = Genera_ReporteWORD(gColCredRepCartaCobMoro5, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0)
            
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaRecup Then
                sCadImp = Genera_ReporteWORD(gColCredRepCartaRecup, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), 0, 0, 0, 0)
            
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaInvCredAlt Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepClientesNCuotasPend Then
                If Val(TxtCuotasPend.Text) < 0 Or IsNumeric(TxtCuotasPend.Text) = False Then
                    MsgBox "Ingrese un número de cuotas pendientes válido", vbExclamation, "Aviso"
                    TxtCuotasPend.SetFocus
                    Exit Sub
                Else
                    If Val(TxtNotaIni.Text) < 0 Or IsNumeric(TxtNotaIni.Text) = False Then
                        MsgBox "Ingrese una nota válida", vbExclamation, "Aviso"
                        TxtNotaIni.SetFocus
                        Exit Sub
                    Else
                        If Val(TxtNotaFin.Text) < 0 Or IsNumeric(TxtNotaFin.Text) = False Then
                            MsgBox "Ingrese una nota válida", vbExclamation, "Aviso"
                            TxtNotaFin.SetFocus
                            Exit Sub
                        Else
                            If Val(TxtNotaIni.Text) > Val(TxtNotaFin.Text) Then
                                MsgBox "La nota inicial no puede ser mayor que la nota final", vbExclamation, "Aviso"
                                TxtNotaIni.SetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCartaInvCredAlt Then
                    sCadImp = Genera_ReporteWORD(gColCredRepCartaInvCredAlt, cMensaje, sCondicion, sMoneda, sProductos, sAnalistas, Val(TxtDiaAtrIni.Text), Val(TxtDiasAtrFin.Text), Val(TxtNotaIni.Text), Val(TxtNotaFin.Text), ChkPorc.value, Val(TxtCuotasPend.Text))
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepClientesNCuotasPend Then
                    Set loRep = New NCredReporte
                    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                    sCadImp = loRep.nRepo108303_ClientesCuotasPend(cMensaje, sMoneda, sProductos, sCondicion, sAnalistas, Val(TxtNotaIni.Text), Val(TxtNotaFin.Text), ChkPorc.value, Val(TxtCuotasPend.Text))
                     
                End If
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVigArqueo Then
                
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108203_CreditosVigentes_Arqueo(IIf(optCredVig(0).value = True, 1, IIf(optCredVig(1).value = True, 2, 3)), cMensaje, Me.TxtFecFinA02.Text, sMoneda, sProductos, sCondicion, sAnalistas)
                 
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepVisitaCobroCuotas Then
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108302_VisitaCobroCuotas(cMensaje, gdFecSis, sMoneda, sProductos, sCondicion, sAnalistas)
            
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepIngresosxGasto Then
                
                For i = 0 To nContAgencias - 1
                    If i = 0 Then
                        sAgencias = "'" & MatAgencias(i) & "'"
                     Else
                        sAgencias = sAgencias & ", '" & MatAgencias(i) & "'"
                     End If
                Next
                
                
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108304_IngresosxGastos(Me.TxtFecIniA02.Text, Me.TxtFecFinA02.Text, cMensaje, sMoneda, sProductos, sCondicion, sAnalistas, sAgencias)
                 
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVigIntDeven Then
                
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108201_CreditosVigentes_DiasAtraso(cMensaje, gdFecSis, Val(TxtDiaAtrIni), Val(TxtDiasAtrFin), sMoneda, sProductos, sCondicion, sAnalistas)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredDesmMayores Then
                If Val(txtMontoMayor.Text) > 0 Then
                    Set loRep = New NCredReporte
                    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                    sCadImp = loRep.nRepo108204_CreditosDesembolsadosVigentes(cMensaje, Me.TxtFecIniA02.Text, Me.TxtFecFinA02.Text, Val(txtMontoMayor.Text), sMoneda, sProductos, sCondicion, sAgencias)
                Else
                    MsgBox "Ud. debe ingresar un monto válido", vbExclamation, "Aviso"
                    txtMontoMayor.SetFocus
                    Exit Sub
                End If
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepResSalCartxAna Then
                If Val(TxtDiasAtrCons1Ini.Text) > Val(TxtDiasAtrCons1Fin.Text) Then
                    MsgBox "El rango final debe ser mayor o igual al rango inicial", vbExclamation, "Aviso"
                    TxtDiasAtrCons1Fin.SetFocus
                    Exit Sub
                Else
                    If Val(TxtDiasAtrCons2Ini.Text) > Val(TxtDiasAtrCons2Fin.Text) Then
                        MsgBox "El rango final debe ser mayor o igual al rango inicial", vbExclamation, "Aviso"
                        TxtDiasAtrCons2Fin.SetFocus
                        Exit Sub
                    Else
                        Set loRep = New NCredReporte
                        loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                        sCadImp = loRep.nRepo108404_SaldosCarteraxAnalista(cMensaje, gdFecSis, Val(TxtDiasAtrCons1Ini.Text), Val(TxtDiasAtrCons1Fin.Text), Val(TxtDiasAtrCons2Ini.Text), Val(TxtDiasAtrCons2Fin.Text), Val(TxtDiasAtrCons3Ini.Text), sMoneda, sProductos, sCondicion, sAgencias, sAnalistas)
                    End If
                End If
            End If
            
        ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepDatosReqMora Then
         
            
            sCadImp = Genera_Reporte108306(cMensaje, sMoneda, sProductos, sAnalistas)
        ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsCartAltoRiesgoxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAnalista Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAgencia Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsMetAlcanzxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxFteFinan Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocyMoraxAna Or Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsResCartSuper Then
            
            cMensaje1 = ""
            cMensaje2 = ""
            nBandera = False
                
            For i = 0 To nContAgencias - 1
                If i = 0 Then
                    sAgencias = "'" & MatAgencias(i) & "'"
                    cMensaje1 = MatAgencias(i)
                Else
                    sAgencias = sAgencias & ", '" & MatAgencias(i) & "'"
                    cMensaje1 = cMensaje1 & "/" & MatAgencias(i)
                End If
            Next
            
            cMensaje = cMensaje & " AGENCIAS: " & cMensaje1
             
            If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsCartAltoRiesgoxAna Then
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108604_CarteraAltoRiesgoxAnalista(cMensaje, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
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
                 
                'Recalculo el tipo de cambio fijo del mes para la fecha especificada
                Set oTipCambio = New nTipoCambio
                TxtTipCambio.Text = Format(oTipCambio.EmiteTipoCambio(Format(Me.TxtFecFinA02.Text, "dd/MM/YYYY"), TCFijoMes), "0.00")
                Set oTipCambio = Nothing
                  
                  
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                
                If Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAnalista Then
                    sCadImp = loRep.nRepo108602_ConsolidadoColocacionesxAnalista(nTempoParam, cMensaje, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxAgencia Then
                    sCadImp = loRep.nRepo108601_ConsolidadoColocacionesxAgencia(nTempoParam, cMensaje, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsMetAlcanzxAna Then
                    sCadImp = loRep.nRepo108603_CuadroMetasAlcanzadasxAnalista(nTempoParam, cMensaje, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
                ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocyMoraxAna Then
                    sCadImp = loRep.nRepo108606_ConsolidadoColocacionesxMoraxAnalista(nTempoParam, cMensaje, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
                End If
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsColocxFteFinan Then
                Set loRep = New NCredReporte
                loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
                sCadImp = loRep.nRepo108605_ConsolidadoColocxFteFinanciamiento(cMensaje, Val(TxtTipCambio.Text), gdFecSis, sMoneda, sProductos, sAgencias)
            ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepConsResCartSuper Then
                sCadImp = Genera_Reporte108607(cMensaje, Val(TxtTipCambio.Text), Me.TxtFecFinA02.Text, sMoneda, sProductos, sAgencias, sAnalistas)
            End If
        ElseIf Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepEstMensual Then
            If optEstadistica(2).value = True Then
                If Len(Trim(txtLineaCredito.Text)) = 0 Then
                    MsgBox "Ingrese una linea de crédito", vbExclamation, "Aviso"
                    txtLineaCredito.SetFocus
                    Exit Sub
                End If
            End If
            
            Set loRep = New NCredReporte
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            sCadImp = loRep.nRepo108202_EstadisticaMensualCreditos(IIf(optEstadistica(0).value = True, 1, IIf(optEstadistica(1).value = True, 2, 3)), cMensaje, TxtFecIniA02, TxtFecFinA02, sMoneda, sProductos, Trim(txtLineaCredito.Text))
             
        End If
          
    '-******   Reportes de Fin de Mes Para Constabilidad y Planeamiento  *******
    '
    '
    '
    '
    '----------------------------------------------------------------------------
    '"1"
    Case 108701:
         sCadImp = ""
         sCadImp = lsRep.nRepo108701_CarteraColocacionesxMoneda("1", "dbCmactConsolidada..")
         
    Case 108702:
        sCadImp = ""
        If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
        sCadImp = lsRep.nRepo108702_ImpCarteraCredConsolidada("1", TxtTipCambio, "dbCmactConsolidada..")
        End If
    Case 108703:
        sCadImp = ""
        If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
        sCadImp = lsRep.nRepo108703_ImpRepCarteraProd_Venc("1", Val(TxtTipCambio), "dbCmactConsolidada..")
        End If
        
    Case 108704: '  Reporte por  Producto  Y Agencia (A-2.3)
         sCadImp = ""
         If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
         sCadImp = lsRep.nRepo108704_ImpRepCarteraAgencia_Prod("1", Val(TxtTipCambio), "C", "dbCmactConsolidada..")
        End If

    Case 108705: 'Reporte para Reclasificacion de Cartera (A-4)
        sCadImp = ""
        sCadImp = lsRep.nRepo108705_ImpCarteraReclasificacion("1", "dbCmactConsolidada..")
         

    Case 108706: ' Reporte de Intereses Devengados Vigentes (A-5)
         sCadImp = ""
         If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
         sCadImp = lsRep.nREpo108706_ImpRepDevengados_Vigentes("dbCmactConsolidada..", "1", Val(TxtTipCambio))
        End If
         
    Case 108707: ' Reporte de Intereses Devengados Vencidos y Cobranza Judicial (A-6)
         sCadImp = ""
         If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
         sCadImp = lsRep.nRepo108707_ImpRepDevengados_Vencidos("dbCmactConsolidada..", "1", Val(TxtTipCambio))
        End If
         
    Case 108708:  ' Resumen de Garantias  (A-7)
        sCadImp = ""
        If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
        sCadImp = lsRep.nRepo108708_ImpRepResumenGarantias("dbCmactConsolidada..", "1", Val(TxtTipCambio))
        End If
    'Case 108709:  ' Cartera de Alto Riesgo  (A-8)

    Case 108710: '  Colocaciones x Sectores Economicos  (A-9)
          sCadImp = ""
          If Not IsNumeric(TxtTipCambio) Then
            MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
            Exit Sub
        Else
          sCadImp = lsRep.nRepo108710_ImpRepColocxSectEcon("dbCmactConsolidada..", "1", Val(TxtTipCambio))
        End If
    Case 108711: ' Reporte de Intereses de Créditos (A-4)
            sCadImp = ""
            sCadImp = lsRep.nRepo108711_ImpCarteraReclasificacion("dbCmactConsolidada..", "1", "nIntDev")
    
    Case 108712: ' Reporte Revision  de Provision de Cartera de Creditos
                sCadImp = ""
                sCadImp = lsRep.nRepo108712_ImpReversionIntDeveng("dbCmactConsolidada..", "1")
    
    Case 108713:
                sCadImp = ""
                sCadImp = lsRep.nRepo108713_ImpRepCarteraAgencia_Prod("dbCmactConsolidada..", "1", Val(TxtTipCambio), "D")
    
    Case 108714:
                sCadImp = ""
                sCadImp = lsRep.nRepo108713_ImpRepCarteraAgencia_Prod("dbCmactConsolidada..", "1", Val(TxtTipCambio), "S")
    Case 108715:
                Call lsRep.nRepo108715_ImpRepSaldoCartera_Rango("dbCmactConsolidada..", "1", Val(Me.TxtTipCambio))
    
    Case 108721: ' Creditos Vigentes(Garantia) - Pyme

            lsCadenaPar = Replace(ValorMoneda, "Credito", "CS") & ValorNorRefPar & _
            Replace(ValorProducto, "Credito", "CS") & _
            " And CS.nDiasAtraso >= " & Val(Me.TxtDiaAtrIni) & _
            " And CS.nDiasAtraso <= " & Val(Me.TxtDiasAtrFin)

             lsCadenaDesPar = DescCondSeleccionado & DescProdSeleccionado
            'AgenciaSeleccionada (False)
            If MatAgencias(0) = "" Then
                MsgBox "No se ha seleccionado agencia", vbInformation, "Aviso"
               Exit Sub
             End If
             sCadImp = ""
             sCadImp = lsRep.nRepo108721_fgImpCredVigGarant("dbCmactConsolidada..", Val(TxtTipCambio), lsCadenaPar, lsCadenaDesPar, Val(Me.TxtDiaAtrIni), Val(Me.TxtDiasAtrFin), MatAgencias(0), gdFecSis)

    
    Case 108722: ' Creditos Vigentes (Garantia) - Consumo
            
            lsCadenaPar = Replace(ValorMoneda, "Credito", "CS") & ValorNorRefPar & _
            Replace(ValorProdConsumo, "Credito", "CS") & _
            " And CS.nDiasAtraso >= " & Val(TxtDiaAtrIni.Text) & _
            " And CS.nDiasAtraso <= " & Val(TxtDiasAtrFin.Text)

            lsCadenaDesPar = DescProdConsumoSeleccionado & DescCondSeleccionado
            If MatAgencias(0) = "" Then
                MsgBox "No se ha seleccionado agencia", vbInformation, "Aviso"
                Exit Sub
            End If
            sCadImp = ""
             sCadImp = lsRep.nRepo108722_fgImpCredPersVigentesGarant("dbCmactConsolidada..", Val(TxtTipCambio), lsCadenaPar, lsCadenaDesPar, Val(Me.TxtDiaAtrIni), Val(Me.TxtDiasAtrFin), MatAgencias(0), gdFecSis)
    Case 108723: ' Creditos PIGNORATICIO - Vigentes
            If MatAgencias(0) = "" Then
                MsgBox "No se ha seleccionado agencia", vbInformation, "Aviso"
                Exit Sub
            End If
            sCadImp = ""
            sCadImp = lsRep.nRepo108723_fgImprimeCredPigIntDev("dbCmactConsolidada..", Val(Me.TxtDiaAtrIni), Val(Me.TxtDiasAtrFin), MatAgencias(0))
    
    Case 108724:
    
            lsCadenaPar = Replace(ValorMoneda, "Credito", "CS") & _
            Replace(ValorProducto, "Credito", "CS")
            lsCadenaDesPar = DescProdSeleccionado
         If MatAgencias(0) = "" Then
            MsgBox "No se ha seleccionado agencia", vbInformation, "Aviso"
            Exit Sub
         End If
         
         sCadImp = ""
         sCadImp = lsRep.nRepo108724_fgImpCredRefinan("dbCmactConsolidada..", Val(Me.TxtTipCambio), lsCadenaPar, lsCadenaDesPar, MatAgencias(0), gdFecSis)
         sCadImp = ""
         'If fgImpCredRefinan(Val(Me.txtTipoCambio.Text), lsCadenaPar, lsCadenaDesPar, matAgencia(0)) Then
         '   rtf.Text = lsRtfImpG
         '   frmPrevio.Previo rtf, "Creditos PIGNORATICIO - Vigentes ", True, 66
         'End If

       
                
                
    Case 108801:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            Fechaini = "01" & Mid(CStr(CredRepoMEs.GEtFechaCierreMes), 3, 10)
            sCadImp = lsRep.nRepo108801_("dbCmactConsolidada..", Fechaini, CredRepoMEs.GEtFechaCierreMes, Val(TxtTipCambio))
            End If
    Case 108802:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
                sCadImp = lsRep.nRepo108802_("dbCmactConsolidada..", Val(TxtTipCambio))
            End If
    Case 108803:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            sCadImp = lsRep.nRepo108803_("dbCmactConsolidada..", Val(TxtTipCambio))
            End If
    Case 108804:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            sCadImp = lsRep.nRepo108804_("dbCmactConsolidada..", Val(TxtTipCambio))
            End If
    Case 108806:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
                sCadImp = lsRep.nRepo108806_("dbCmactConsolidada..", Val(TxtTipCambio))
            End If
            
    Case 108808:
            If Not IsNumeric(TxtTipCambio) Then
                MsgBox "Ingerese Correctamente el Tipo de Cambio", vbInformation, "AVISO"
                Exit Sub
            Else
            Fechaini = "01" & Mid(CStr(CredRepoMEs.GEtFechaCierreMes), 3, 10)
           Call lsRep.nRepo108808_("dbCmactConsolidada..", CredRepoMEs.GEtFechaCierreMes, Val(TxtTipCambio), gsNomCmac, gdFecSis, DateAdd("D", -1 * Day(TxtFecIniA02), TxtFecIniA02), TxtFecIniA02)
            End If
            'sCadImp = "111111"
    End Select
    
    If Len(Trim(sCadImp)) <= 1 Then
        MsgBox "No existen datos para el reporte", vbExclamation, "Aviso"
    Else
        Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
            Case gColCredRepDatosReqMora, gColCredRepConsResCartSuper, gColCredRepCartaCobMoro1, _
                 gColCredRepCartaCobMoro2, gColCredRepCartaCobMoro3, gColCredRepCartaCobMoro4, _
                 gColCredRepCartaCobMoro5, gColCredRepCartaInvCredAlt, gColCredRepCartaRecup
            Case Else
                P.Show sCadImp, "Reportes de Creditos", True
        End Select
    End If
    
    Set P = Nothing
    Set oNCredDoc = Nothing
    Set CredRepoMEs = Nothing
    Set lsRep = Nothing
End Sub
Public Sub nRepo108808_(ByVal psServConsol As String, ByVal pdFechaDesde As Date, ByVal pdFechaHasta As Date, _
ByVal pnTipoCambio As Double)
Dim Co As nCredRepoFinMes
'Dim xlAplicacion As Excel.Application
'Dim xlLibro As Excel.Workbook
'Dim xlHoja1 As Excel.Worksheet
Dim xlHojaP As Excel.Worksheet
Dim sql As String
Dim Rs As New ADODB.Recordset
Dim fs As Scripting.FileSystemObject
Dim Total As Double
Dim Tabula As Integer

Dim lsCond1(11) As String, lsCond2(11) As String
Dim Det As Integer
Dim lnFil As Integer, lnCol As Integer

Dim lnNroCreFMS As Currency, lnNroCreFMD As Currency
Dim lnMonCreFMS As Currency, lnMonCreFMD As Currency
Dim lnNroCreOtorgS As Currency, lnNroCreOtorgD As Currency
Dim lnMonCreOtorgS As Currency, lnMonCreOtorgD As Currency
Dim lnNroCreCancelS As Currency, lnNroCreCancelD As Currency
Dim lnMonCreCancelS As Currency, lnMonCreCancelD As Currency
Dim lnNroCredS As Currency, lnNroCredD As Currency
Dim lnMonCredS As Currency, lnMonCredD As Currency

Dim Titulo As String
Dim lsCreditosVigentes As String
Dim lsPignoraticio As String
Dim lsVig As String
'Dim Tabula As Integer

Set Co = New nCredRepoFinMes
lsCreditosVigentes = gColocEstVigNorm & "," & gColocEstVigMor & "," & gColocEstVigVenc & "," & gColocEstRefNorm & "," & gColocEstRefMor & "," & gColocEstRefVenc
lsPignoraticio = gColPEstDesem & "," & gColPEstVenci & "," & gColPEstPRema & "," & gColPEstRenov
lsVig = gColocEstRecVigJud & "," & gColocEstRecVigCast & "," & gColocEstRecCanJud & "," & gColocEstRecCanCast
'On Error GoTo ErrorExcel
Screen.MousePointer = 11

Total = 4 * 25
'Me.barra.Max = Total
'rtf.Text = ""
Tabula = 20
ReDim Lineas(20)
Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.Path & "\SPOOLER\INFORME_COLOC_BCR.xls") Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\INFO4.xls")
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If
Set xlHoja1 = xlLibro.Worksheets.Add

'--************************** CLIENTES NUEVOS Y CONOCIDOS *****************************
'EncabezadoInfo4Excel
Titulo = " C R E D I T O   E M P R E S A R I A L"
xlAplicacion.Range("A1:E7").Font.Bold = True
xlAplicacion.Range("A1:E7").Font.Size = 9
xlAplicacion.Range("A5:P15").Font.Size = 8
xlAplicacion.Range("A4:E20").Font.Size = 8
xlAplicacion.Range("A7:E7").HorizontalAlignment = xlHAlignCenter
xlAplicacion.Range("A11:E11").HorizontalAlignment = xlHAlignCenter
xlAplicacion.Range("A11:E12").Font.Bold = True
xlHoja1.Cells(1, 3) = "R E P O R T E   I N F O 4"
xlHoja1.Cells(2, 2) = gsNomCmac
xlHoja1.Range("B2:E3").MergeCells = True
xlHoja1.Cells(3, 2) = "INFORMACION AL " & Format(gdFecSis, "dd/mm/yyyy")
xlHoja1.Cells(4, 2) = "T.C.F. :" & Format(pnTipoCambio, "#,#0.000")
xlHoja1.Cells(5, 3) = Titulo

'---------------------------------------
For Det = 1 To 11
    Select Case Det
        Case 1
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('101','201') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 2
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('101','201') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) not in('01') "
        Case 3
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('301') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 4
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('302','303') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 5
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('304') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 6
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('401','423') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in('01') "
        Case 7
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('301') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 8
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('302','303') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 9
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('304') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 10
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('401','423') "
            lsCond2(Det) = " AND Substring(C.cLineacred,1,2) in ('03','05','06','07','08') "
        Case 11
            lsCond1(Det) = " AND Substring(C.cCtaCod,6,3) in('305') "
            lsCond2(Det) = "  "
    End Select

    sql = "SELECT Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.cCtaCod End ) NumOtorgS , " _
        & " Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.cCtaCod End ) NumOtorgD , " _
        & " Isnull(Sum ( CASE  WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.NMONTODESEMB End ),  0 ) SKOtorgS,  " _
        & " Isnull(Sum ( CASE  WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.NMONTODESEMB*" & pnTipoCambio & "  End ),  0 ) SKOtorgD  " _
        & " From " & psServConsol & "CreditoConsol  C " _
        & " WHERE C.nPrdEstado in (" & lsCreditosVigentes & "," & gColocEstCancelado & "," & lsVig & ") " _
        & "  AND  C.DFECVIG BETWEEN '" & Format(pdFechaDesde, "mm/dd/yyyy") & "' AND '" & Format(pdFechaHasta, "mm/dd/yyyy") & " 23:59' " _
        & lsCond1(Det) & lsCond2(Det)
    
    'rs.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Set Rs = Co.GetQuery(sql)
    lnNroCreOtorgS = Rs!NumOtorgS
    lnNroCreOtorgD = Rs!NumOtorgD
    lnMonCreOtorgS = Rs!SKOtorgS
    lnMonCreOtorgS = Rs!SKOtorgS
    
    Rs.Close

    sql = "SELECT Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.cCtaCod End ) NumFinMesS , " _
        & " Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.cCtaCod End ) NumFinMesD , " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '1' THEN (C.nSaldoCap) End ), 0 ) SKFinMesS , " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '2' THEN (C.nSaldoCap * " & pnTipoCambio & ") End ), 0 ) SKFinMesD " _
        & " From " & psServConsol & "CreditoSaldoConsol C " _
        & " JOIN " & psServConsol & "CreditoConsol CC on C.cCtaCod = CC.cCtaCod " _
        & " WHERE C.nPrdEstado in (" & lsCreditosVigentes & ") And Datediff(d,dFecha,'" & Format(pdFechaDesde, "mm/dd/yyyy") & "') = 0 " _
        & "  " _
        & lsCond1(Det) & Replace(lsCond2(Det), "C", "CC")

    'rs.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Set Rs = Co.GetQuery(sql)
    lnNroCreFMS = Rs!NumFinMesS
    lnNroCreFMD = Rs!NumFinMesD
    lnMonCreFMS = Rs!SKFinMesS
    lnMonCreFMD = Rs!SKFinMesD
    
    Rs.Close

    sql = "SELECT Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='1' THEN C.cCtaCod End ) NumCredS ,  " _
        & " Count( CASE WHEN SUBSTRING(C.cCtaCod,9,1)='2' THEN C.cCtaCod End ) NumCredD , " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '1' THEN (C.nSaldoCap) End ), 0 ) SKCredS, " _
        & " Isnull(SUM( CASE WHEN substring(C.cCtaCod,9,1) = '2' THEN (C.nSaldoCap * " & pnTipoCambio & " ) End ), 0 ) SKCredD " _
        & " From " & psServConsol & "CreditoSaldoConsol C " _
        & " JOIN " & psServConsol & "CreditoConsol CC on C.cCtaCod = CC.cCtaCod " _
        & " WHERE C.nPrdEstado in (" & lsCreditosVigentes & ") And Datediff(d,dFecha,'" & Format(pdFechaHasta, "mm/dd/yyyy") & "')=0" _
        & "  " _
        & lsCond1(Det) & Replace(lsCond2(Det), "C", "CC")
    
    'rs.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
    Set Rs = Co.GetQuery(sql)
    lnNroCredS = Rs!NumCredS
    lnNroCredD = Rs!NumCredD
    lnMonCredS = Rs!SKCredS
    lnMonCredD = Rs!SKCredD
    
    Rs.Close

    lnNroCreCancelS = lnNroCreFMS + lnNroCreOtorgS - lnNroCredS
    lnNroCreCancelD = lnNroCreFMD + lnNroCreOtorgD - lnNroCredD
    lnMonCreCancelS = lnMonCreFMS + lnMonCreOtorgS - lnMonCredS
    lnMonCreCancelD = lnMonCreFMD + lnMonCreOtorgD - lnMonCredD
    
    If Det = 1 Or Det = 3 Or Det = 7 Or Det = 10 Or Det = 11 Then
        lnFil = lnFil + 3
        lnCol = 1
        
        xlHoja1.Cells(lnFil, lnCol) = "Nro Cred. Vigentes " & Format(pdFechaDesde, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 1, lnCol) = "Nro Cred. Otorgados   " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 2, lnCol) = "Nro Cred. Cancelados  " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 3, lnCol) = "Nro Cred. Vigentes    " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 4, lnCol) = "Saldo Cred. Vigentes  " & Format(pdFechaDesde, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 5, lnCol) = "Monto Cred. Otorgados " & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 6, lnCol) = "Monto Cred. Cancelados" & Format(pdFechaHasta, "dd/mm/yyyy")
        xlHoja1.Cells(lnFil + 7, lnCol) = "Saldo Cred. Vigentes  " & Format(pdFechaHasta, "dd/mm/yyyy")
    End If
    
    xlHoja1.Cells(lnFil, lnCol + 1) = lnNroCreFMS
    xlHoja1.Cells(lnFil, lnCol + 2) = lnNroCreFMD
    xlHoja1.Cells(lnFil + 1, lnCol + 1) = lnNroCreOtorgS
    xlHoja1.Cells(lnFil + 1, lnCol + 2) = lnNroCreOtorgD
    xlHoja1.Cells(lnFil + 2, lnCol + 1) = lnNroCreCancelS
    xlHoja1.Cells(lnFil + 2, lnCol + 2) = lnNroCreCancelD
    xlHoja1.Cells(lnFil + 3, lnCol + 1) = lnNroCredS
    xlHoja1.Cells(lnFil + 3, lnCol + 2) = lnNroCredD
    xlHoja1.Cells(lnFil + 4, lnCol + 1) = lnMonCreFMS
    xlHoja1.Cells(lnFil + 4, lnCol + 2) = lnMonCreFMD
    xlHoja1.Cells(lnFil + 5, lnCol + 1) = lnMonCreOtorgS
    xlHoja1.Cells(lnFil + 5, lnCol + 2) = lnMonCreOtorgD
    xlHoja1.Cells(lnFil + 6, lnCol + 1) = lnMonCreCancelS
    xlHoja1.Cells(lnFil + 6, lnCol + 2) = lnMonCreCancelD
    xlHoja1.Cells(lnFil + 7, lnCol + 1) = lnMonCredS
    xlHoja1.Cells(lnFil + 7, lnCol + 2) = lnMonCredD

Next Det

xlHoja1.SaveAs App.Path & "\SPOOLER\INFO4.xls"
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Screen.MousePointer = 0
MsgBox "Se ha Generado el Archivo INFO4.XLS Satisfactoriamente", vbInformation, "Aviso"
Exit Sub

ErrorExcel:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
    xlLibro.Close
    ' Cierra Microsoft Excel con el método Quit.
    xlAplicacion.Quit
    'Libera los objetos.
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing

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
Unload frmColRecReporte
    ReDim MatAgencias(0)
    ReDim MatProductos(0)
    ReDim matAnalista(0)
    ReDim MatInstitucion(0)
    Set Progress = New clsProgressBar
    Set Progreso = New clsProgressBar
    Dim oTipCambio As nTipoCambio
    
    Set oTipCambio = New nTipoCambio
    TxtTipCambio.Text = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "0.00")
    Set oTipCambio = Nothing
    
    
End Sub

Private Function ValorProducto() As String
Dim i As Integer
Dim lsCad As String

lsCad = ""
' Cred. Comercial
For i = 0 To chkComercial.Count - 1
    If chkComercial(i).value Then
        lsCad = lsCad & "'" & chkComercial(i).Tag & "',"
    End If
Next i

' Cred. MicroEmpresarial
For i = 0 To chkMicroEmpresa.Count - 1
    If chkMicroEmpresa(i).value Then
        lsCad = lsCad & "'" & chkMicroEmpresa(i).Tag & "',"
    End If
Next i

' Cred. Consumo
For i = 0 To chkConsumo.Count - 1
    If chkConsumo(i).value Then
        lsCad = lsCad & "'" & chkConsumo(i).Tag & "',"
    End If
Next i

 '  Cred. Hipotecario
For i = 0 To chkHipotecario.Count - 1
    If chkHipotecario(i).value Then
        lsCad = lsCad & "'" & chkHipotecario(i).Tag & "',"
    End If
Next i

If Len(lsCad) > 0 Then
    lsCad = Mid(lsCad, 1, (Len(lsCad) - 1))
    ValorProducto = " AND substring(Credito.cCtaCod,6,3) IN (" & lsCad & ") "
Else
    ValorProducto = ""
End If
End Function


Private Sub Form_Unload(Cancel As Integer)
    Unload frmSelectAgencias
    Unload frmSelectAnalistas
    Unload frmUbicacionGeo
    Set frmCredReportes = Nothing
End Sub



Private Sub lsRep_CloseProgress()
    Progreso.CloseForm Me
End Sub

Private Sub lsRep_Progress(pnValor As Long, pnTotal As Long)
    Progreso.Max = pnTotal
    Progreso.Progress pnValor, "Generando Reporte", "Procesando ..."
End Sub

Private Sub lsRep_ShowProgress()
    Progreso.ShowForm Me
End Sub

Private Sub optCredVig_Click(Index As Integer)
    CmdAnalistas.Visible = IIf(Index = 2, True, False)
End Sub

Private Sub optEstadistica_Click(Index As Integer)
txtLineaCredito.Enabled = IIf(Index = 2, True, False)
txtLineaCredito.BackColor = IIf(Index = 2, &H80000005, &H8000000F)
txtLineaCredito.Text = ""
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
Dim i As Integer
Dim sTipo As String
    Limpia
    Me.Caption = "Reportes de Créditos " & Mid(TVRep.SelectedItem.Text, 8, Len(TVRep.SelectedItem.Text) - 7)
    Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
        Case gColCredRepIngxPagoCred
            Call HabilitaControleFrame1(False, True, True, False)
            CmdSelecAge.Visible = True
        Case gColCredRepDesemEfect
            Call HabilitaControleFrame1(True, True, True, False, , , , , , , , , , , , , , , True)
            CmdSelecAge.Visible = True
        Case gColCredRepSalCarVig
            Call HabilitaControleFrame1(False, False, True, False)
        Case gColCredRepCredCancel
            Call HabilitaControleFrame1(True, True, True, True)
        Case gColCredRepResSalCarxAna
            Call HabilitaControleFrame1(False, False, True, False, True, True, , , , , , , , , , , , , True)
        Case gColCredRepMoraInst
            Call HabilitaControleFrame1(True, False, True, False, True, True, True, , , , , , , , , , , , True)
        
        '(Se agrego una segunda opcion con una bandera)
        Case gColCredRepMoraxAna, gColCredRepAtraPagoCuotaLib
            Call HabilitaControleFrame1(False, True, True, False, False, True, False, True, , , , , , , , , , , True)
        '''''''''''
        
        Case gColCredRepCredProtes
            Call HabilitaControleFrame1(False, False, True, False, False, False, False, False)
        Case gColCredRepCredRetir
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, False, , , , , , , , , , , True)
        Case gColCredRepCredxUbiGeo
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, False, False, False, True, , , , , , , , True)
        
        '(Se agrego una segunda opcion con una bandera)
        Case gColCredRepCredVig, gColCredRepCredVigconCuoLibre
            
            Call HabilitaControleFrame1(False, True, True, False, False, True, False, False, True, False, False, True, , , , , IIf(Mid(TVRep.SelectedItem.Text, 1, 6) = gColCredRepCredVig, True, False), , True)
        '''''''''''''''''''''''
        
        Case gColCredRepCredxInst
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, True, True, True)
        Case gColCredRepMoraxInst
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, False, False, True)
        Case gColCredRepResSalCartxAna
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, False, False, False, False, False, False, False, True, True, , True)
        Case gColCredRepResSaldeCartxInst
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, True, , True)
            'Ubica el Boton para Seleccionar la Institucion en la posicion Inferior
            CmdInstitucion.Left = 1965
            CmdInstitucion.Top = 5535
        Case gColCredRepLisDesctoPlanilla
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, True, True, True)
        Case gColCredRepPagosconCheque
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, False, False, False, False, False, False, False, False, False, False, True, True, False, False)
            CmdSelecAge.Visible = True
        Case gColCredRepPagosdeOtrasAgen
            Call HabilitaControleFrame1(True, True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case gColCredRepPagosEnOtrasAgen
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False)
        Case gColCredRepIntEnSusp
            Call HabilitaControleFrame1(True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True)
        
        ''''''''''''''''
        Case gColCredRepProgPagosxCuota
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, True, False, False, False, False, False, False, False, False, False, False, True, True, False)
        Case gColCredRepDatosReqMora
            Call HabilitaControleFrame1(False, False, True, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, True, False, False)
        Case gColCredRepConsCartAltoRiesgoxAna, gColCredRepConsColocxAnalista, gColCredRepConsMetAlcanzxAna, gColCredRepConsColocyMoraxAna, gColCredRepConsResCartSuper
            Call HabilitaControleFrame1(False, True, True, False, False, False, False, True, False, False, False, True, False, False, False, False, False, False, True, False)
        Case gColCredRepConsColocxAgencia
            Call HabilitaControleFrame1(False, True, True, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, True, False)
        Case gColCredRepConsColocxFteFinan
            Call HabilitaControleFrame1(False, False, True, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, True, False)
        
        Case gColCredRepCartaCobMoro1, gColCredRepCartaCobMoro2, gColCredRepCartaCobMoro3, _
             gColCredRepCartaCobMoro4, gColCredRepCartaCobMoro5, gColCredRepCartaRecup
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, True, False, False, False, False, False, False, False, False, False, True, , False)
        Case gColCredRepCartaInvCredAlt
                        Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, True, True, False, False, False, False, False, False, False, False, True, , False)
        
        Case gColCredRepCredVigArqueo
            Call HabilitaControleFrame1(False, True, True, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, True)
        Case gColCredRepVisitaCobroCuotas
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, False, False, False, False, False, False, False, False, False, False, True, False, False, False)
            
        Case gColCredRepClientesNCuotasPend
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, False, True, False, False, False, False, False, False, False, False, True, , False)
         
        Case gColCredRepIngresosxGasto
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, True, False, False, False, False, False, False, False, False, False, False, True)
        Case gColCredRepCredVigIntDeven
            Call HabilitaControleFrame1(False, False, True, False, False, True, False, True, True, False, False, False, False, False, False, False, False, False, True, False, False)
        Case gColCredRepEstMensual
            Call HabilitaControleFrame1(True, True, True, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, True)
        Case gColCredRepCredDesmMayores
            Call HabilitaControleFrame1(True, True, True, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False, True, False, True, , , True)
      
      '***********************************
        Case 108701:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108702:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108703:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108704:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108705:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108706:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108707:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108708:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        'Case 108709:
        Case 108710:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108711:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108712:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108713:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108714:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108715:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        
        Case 108721:
                Call HabilitaControleFrame1(False, False, True, False, False, True, False, False, True, False, False, True, False, False, False, False, False, False, True, False, True, False, False, False)
                chkProducto(0).Enabled = False
                chkProducto(1).Enabled = False
                chkProducto(2).Enabled = False
                chkProducto(3).Enabled = False
            
                For i = 0 To chkComercial.Count - 1
                    chkComercial(i).Enabled = False
                Next i
                For i = 0 To chkMicroEmpresa.Count - 1
                    chkMicroEmpresa(i).Enabled = False
                Next i
                For i = 0 To chkMicroEmpresa.Count - 1
                    chkMicroEmpresa(i).Enabled = True
                Next i
                For i = 0 To chkConsumo.Count - 1
                    chkConsumo(i).Enabled = False
                Next i
                For i = 0 To chkHipotecario.Count - 1
                    chkHipotecario(i).Enabled = False
                Next i
                   
                
        Case 108722:
        
                Call HabilitaControleFrame1(False, False, True, False, False, True, False, False, True, False, False, True, False, False, False, False, True, False, True, False, True, False, False, False)

        Case 108723:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, True, False, False, False)
        
        Case 108724:
                Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, True, False, False, False)
        
        Case 108801:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108802:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108803:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108804:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108806:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, True, False, False, False, False, False, False, False, False, False, False, False, False)
        Case 108808:
            Call HabilitaControleFrame1(False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, False)
        Case Else
            Me.Caption = "Reportes de Creditos "
    End Select
End Sub

Private Sub Limpia()
Dim i As Integer
    
    Call HabilitaControleFrame1(False, False, False, False)
        
    For i = 0 To 3
        chkProducto(i).value = 0
        If i < 3 Then
            ChkCond(i).value = 1
            If i < 2 Then
                ChkMonA02(i).value = 0
            End If
        End If
    Next
    TxtFecIniA02.Text = Format(gdFecSis, "dd/MM/YYYY")
    TxtFecFinA02.Text = Format(gdFecSis, "dd/MM/YYYY")
     
    TxtDiaAtrIni.Text = 0
    TxtDiasAtrFin.Text = 999
    optCredVig(0).value = True
    optEstadistica(0).value = True
    
    chkProducto(0).Enabled = True
    chkProducto(0).value = 0
    chkProducto(1).Enabled = True
    chkProducto(1).value = 0
    chkProducto(2).Enabled = True
    chkProducto(2).value = 0
    chkProducto(3).Enabled = True
    chkProducto(3).value = 0
            
    For i = 0 To chkComercial.Count - 1
        chkComercial(i).Enabled = True
        chkComercial(i).value = 0
    Next i
    For i = 0 To chkMicroEmpresa.Count - 1
        chkMicroEmpresa(i).Enabled = True
        chkMicroEmpresa(i).value = 0
    Next i
    For i = 0 To chkConsumo.Count - 1
        chkConsumo(i).Enabled = True
        chkConsumo(i).value = 0
    Next i
    For i = 0 To chkHipotecario.Count - 1
        chkHipotecario(i).Enabled = True
        chkHipotecario(i).value = 0
    Next i
    
 End Sub

Private Sub TVRep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TVRep_Click
End If
End Sub

Private Sub TVRep_NodeClick(ByVal Node As MSComctlLib.Node)
    TVRep_Click
End Sub
 
Private Function Genera_Reporte108306(ByVal cSubTit As String, ByVal psMoneda As String, ByVal psProducto As String, ByVal psAnalistas As String) As String
  
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
        Genera_Reporte108306 = ""
        Exit Function
    Else
        Genera_Reporte108306 = "Reporte_Generado"
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
            nFila = 4
            xlHoja1.Cells(nFila, 1) = cSubTit
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
            xlHoja1.Range("A4:M4").MergeCells = True

            xlHoja1.Range("A1:B3").Font.Bold = True
            xlHoja1.Range("A4").Font.Bold = True
                        
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

Private Function Genera_Reporte108607(ByVal cSubTit As String, ByVal pnTipoCambio_ As Currency, ByVal pdFechaFin_ As String, ByVal psMoneda_ As String, ByVal psProductos_ As String, ByVal psAgencias_ As String, ByVal psAnalistas_ As String) As String
     
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
    Dim J As Long
    
    Dim sFilaTitulos() As Long 'En este arreglo se graba el numero de la fila en donde escribir los titulos de las agencias
    Dim sFilaSubTotal() As Long 'En este arreglo se graba el numero de filas en donde se deberan llenar con subtotales
    Dim sFilaTextoSubTotal() As String 'En este arreglo se graba el texto de el numero de la filas que se incluiran en las formulas del subtotal
    Dim sFilaTotal As Long 'Fila del total general
    Dim nContadorSubFila As Long 'cuantas subfilas existen
    Dim nFila As Long 'La fila actual
    Dim sTextoTotal As String 'el texto donde se grabara las filas a sumar en el total
    
    Dim sTempoAgencia As String
    Dim sTempoAnalista As String
    
    Set loExc = New DCredReporte
    Set reg = loExc.Recupera_ConsolidadoCarteraxAnalista(pnTipoCambio_, pdFechaFin_, psMoneda_, psProductos_, psAgencias_, psAnalistas_)
    If reg.BOF Then
        Genera_Reporte108607 = ""
        Exit Function
    Else
        Genera_Reporte108607 = "Reporte Generado"
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
                cMatAgencia(nContador) = reg!cAgencia
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
            xlHoja1.Cells(nFila, 1) = cSubTit
            xlHoja1.Cells(nFila, 23) = "T.C.F.= " & Format(pnTipoCambio_, "#,##0.00")
            
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
             
            Dim sRef As Byte
            sRef = 1
            nContadorSubFila = 0
            
            'Dimensiono los arreglos para las variables y el subtotal
            ReDim Preserve sFilaSubTotal(nContadorSubFila) As Long
            ReDim Preserve sFilaTitulos(nContadorSubFila) As Long
            ReDim Preserve sFilaTextoSubTotal(nContadorSubFila) As String
            
            'Grabo la fila del primer titulo
            sFilaTitulos(nContadorSubFila) = nFila
            
            For i = 0 To nContador
                If sTempoAgencia <> cMatAgencia(i) Then
                    'Agregar fila en blanco que diga subtotal
                    nFila = nFila + 1
                    xlHoja1.Cells(nFila, 1) = "SubTotal"
                    'Actualizo el valor de la fila para el subtotal
                    sFilaSubTotal(nContadorSubFila) = nFila
                     
                    nContadorSubFila = nContadorSubFila + 1
                    ReDim Preserve sFilaSubTotal(nContadorSubFila) As Long
                    ReDim Preserve sFilaTextoSubTotal(nContadorSubFila) As String
                       
                    sTempoAgencia = cMatAgencia(i)
                    'Grabo la fila en la que se encuentra el titulo
                    ReDim Preserve sFilaTitulos(nContadorSubFila) As Long
                    
                    'Agrego fila en blanco
                    nFila = nFila + 1
                    
                    'Agrego fila que diga el nombre de la agencia
                    nFila = nFila + 1
                    xlHoja1.Cells(nFila, 1) = cMatAgencia(i) & " " & matAgencia(i)
                    sFilaTitulos(nContadorSubFila) = nFila
                      
                End If
                 
                nFila = nFila + 1
                
                'LLeno la posicion de cada una de las celdas a sumar en la formula de subtotal
                sFilaTextoSubTotal(nContadorSubFila) = sFilaTextoSubTotal(nContadorSubFila) & "+*" & Trim(Str(nFila))
                
                'Imprimo la fila con los valores normales
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
                     
                With xlHoja1.Range("A" & Trim(Str(nFila)) & ":X" & Trim(Str(nFila)))
                    '.Font.Bold = True
                    .Borders.LineStyle = xlDash
                    .Borders.Weight = xlThin
                    .Borders.ColorIndex = 0
                    '.Interior.ColorIndex = 19
                End With
                  
            Next
             
            'Imprimo fila del ultimo subtotal
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
            
            'SubTitulo/Comentario
            xlHoja1.Range("A5").HorizontalAlignment = xlLeft
            
            'De los subtotales
            For i = 0 To nContadorSubFila
                 
                 sTextoTotal = sTextoTotal & sFilaTextoSubTotal(i)
                 
                 xlHoja1.Range("A" & Trim(Str(sFilaSubTotal(i))) & ":X" & Trim(Str(sFilaSubTotal(i)))).Font.Bold = True
                 
                 'Negrita y Color de los Titulos x Agencia
                 xlHoja1.Range("A" & Trim(Str(sFilaTitulos(i)))).Font.Bold = True
                 xlHoja1.Range("A" & Trim(Str(sFilaTitulos(i))) & ":C" & Trim(Str(sFilaTitulos(i)))).Interior.ColorIndex = 38
                             
                 'Formulas
                For J = Asc("B") To Asc("X")
                    xlHoja1.Range(Trim(Chr(J)) & Trim(Str(sFilaSubTotal(i)))).Formula = "=" & Replace(sFilaTextoSubTotal(i), "*", Trim(Chr(J)))
                Next J
            Next
            'Bordes y Colores del Total
            With xlHoja1.Range("A" & Trim(Str(sFilaTotal)) & ":X" & Trim(Str(sFilaTotal)))
                .Font.Bold = True
                .Borders.LineStyle = xlDash
                .Borders.Weight = xlThin
                .Borders.ColorIndex = 0
                .Interior.ColorIndex = 24
            End With
            
            'Calculo de Formulas de la Fila Total
            
            For i = Asc("B") To Asc("X")
            
                xlHoja1.Range(Trim(Chr(i)) & Trim(Str(sFilaTotal))).Formula = "=" & Replace(sTextoTotal, "*", Trim(Chr(i)))
             Next
            
            'Bordes del Titulo
            With xlHoja1.Range("A6:X8")
                .Font.Bold = True
                .Borders.LineStyle = xlDash
                .Borders.Weight = xlThin
                .Borders.ColorIndex = 0
            End With
    
            'Colores del Titulo
            xlHoja1.Range("A6:K8").Interior.ColorIndex = 19
            xlHoja1.Range("L6:P8").Interior.ColorIndex = 15
            xlHoja1.Range("Q6:X8").Interior.ColorIndex = 35
    
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

Private Function Genera_ReporteWORD(ByVal psModeloCarta As Long, ByVal psMensaje As String, ByVal psCondicion As String, ByVal psMoneda As String, ByVal psProductos As String, ByVal psAnalistas As String, ByVal pnDiasIni As Integer, ByVal pnDiasFin As Integer, ByVal psNota1 As Integer, ByVal psNota2 As Integer, ByVal psTipoCuotas As Integer, ByVal psCuotasPend As Integer)
Dim aLista() As String
Dim vFilas As Integer
Dim vFecAviso As Date
Dim k As Integer
Dim CadenaAna As String

Dim psCtaCod As String

Dim lnDeudaFecha As Currency
 
'A la Fecha
Dim lnSaldoKFecha As Currency
Dim lnIntCompFecha As Currency
Dim lnGastoFecha As Currency
Dim lnIntMorFecha As Currency
Dim lnPenalidadFecha As Currency
 
Dim oNegCred As NCredito
Dim MatCalend As Variant
Dim J As Integer

Dim loExc As DCredReporte

Dim rsCarta As New ADODB.Recordset

Dim lsModeloPlantilla As String
Dim vCont As Integer
Dim lnDeuda As Currency
 
Select Case psModeloCarta
    Case gColCredRepCartaCobMoro1
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso1
    Case gColCredRepCartaCobMoro2
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso2
    Case gColCredRepCartaCobMoro3
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso3
    Case gColCredRepCartaCobMoro4
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso4
    Case gColCredRepCartaCobMoro5
        lsModeloPlantilla = App.Path & cPlantillaCartaAMoroso5
    Case gColCredRepCartaInvCredAlt
        lsModeloPlantilla = App.Path & cPlantillaCartaInvCredParalelo
    Case gColCredRepCartaRecup
        lsModeloPlantilla = App.Path & cPlantillaCartaRecup
    Case Else
        MsgBox " Error en la definicion de la Plantilla"
        Genera_ReporteWORD = "Error en la definicion de la plantilla"
        Exit Function
End Select

    Set loExc = New DCredReporte
    
    Set rsCarta = loExc.RecuperaDatosCartasWORD(IIf(psModeloCarta = gColCredRepCartaCobMoro1, 0, IIf(psModeloCarta = gColCredRepCartaInvCredAlt, 2, 1)), psCondicion, psMoneda, psProductos, psAnalistas, pnDiasIni, pnDiasFin, psNota1, psNota2, psTipoCuotas, psCuotasPend)
     
    If rsCarta.BOF Then
        Genera_ReporteWORD = ""
        Exit Function
    Else
        Genera_ReporteWORD = "Reporte Generado"
    End If
    
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
    
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
    
    'Crea Nuevo Documento
    wApp.Documents.Add
    
    Do While Not rsCarta.EOF
        vFilas = vFilas + 1
          
        psCtaCod = rsCarta!cCtaCod
        
        'Obtener la deuda A LA FECHA
        '===========================
        Set oNegCred = New NCredito
        MatCalend = oNegCred.RecuperaMatrizCalendarioPendiente(psCtaCod)
        
        lnSaldoKFecha = Format(oNegCred.MatrizCapitalAFecha(psCtaCod, MatCalend), "#0.00")
        If UBound(MatCalend) > 0 Then
            lnIntCompFecha = Format(oNegCred.MatrizInteresTotalesAFechaSinMora(psCtaCod, MatCalend, gdFecSis), "#0.00")
            lnGastoFecha = Format(oNegCred.MatrizInteresGastosAFecha(psCtaCod, MatCalend, gdFecSis), "#0.00")
            lnIntMorFecha = Format(oNegCred.MatrizInteresMorFecha(psCtaCod, MatCalend), "#0.00")
            lnPenalidadFecha = Format(oNegCred.CalculaGastoPenalidadCancelacion(CDbl(lnSaldoKFecha), CInt(Mid(psCtaCod, 9, 1))), "#0.00")
            lnDeudaFecha = Format(CDbl(lnSaldoKFecha) + CDbl(lnIntCompFecha) + CDbl(lnGastoFecha) + CDbl(lnIntMorFecha) + CDbl(lnPenalidadFecha), "#0.00")
        End If
 
        '===========================

        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.Paste
        wApp.Application.Selection.InsertBreak
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd
        
        With wApp.Selection.Find
            .Text = "CampFecha"
            .Replacement.Text = Trim(ImpreFormat(Format(gdFecSis, "dddd, d mmmm yyyy"), 25))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
       With wApp.Selection.Find
            .Text = "CampTitNombre"
            .Replacement.Text = Trim(PstaNombre(rsCarta!cPersNombre, True))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "CampTitDireccion"
            .Replacement.Text = Trim(rsCarta!cPersDireccDomicilio) & " - " & Trim(rsCarta!cUbiGeoDescripcion)
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "CampTitZonaDir"
            .Replacement.Text = Trim(rsCarta!cUbiGeoDescripcion)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "CampNroCredito"
            .Replacement.Text = Trim(psCtaCod)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "CampAnalista"
            .Replacement.Text = Trim(PstaNombre(rsCarta!cDesAnalista, True))
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        If psModeloCarta <> gColCredRepCartaCobMoro1 Then
            With wApp.Selection.Find
                .Text = "CampTitDirNegocio"
                .Replacement.Text = Trim("" & rsCarta!cRazSocDirecc)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampFiadorNombre"
                .Replacement.Text = Trim(PstaNombre(rsCarta!cDesFiador, True))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampFiadorDireccion"
                .Replacement.Text = Trim(PstaNombre(rsCarta!cDireccionFiador, True))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampDiasAtraso"
                .Replacement.Text = Trim(Str(rsCarta!nDiasAtraso))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampCuotasVenc"
                .Replacement.Text = Trim(Str(rsCarta!nCuota))
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampMoneda"
                .Replacement.Text = Trim(rsCarta!cDesMoneda)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            With wApp.Selection.Find
                .Text = "CampDeuda"
                .Replacement.Text = Format(lnDeudaFecha, "#,###.00")
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
             End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            
            If psModeloCarta = gColCredRepCartaInvCredAlt Then
             
                With wApp.Selection.Find
                    .Text = "CampCuotasPend"
                    .Replacement.Text = Trim(Str(rsCarta!nCuotasPend))
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                End With
                wApp.Selection.Find.Execute Replace:=wdReplaceAll
            
                With wApp.Selection.Find
                    .Text = "CampNota"
                    .Replacement.Text = Trim(Str(rsCarta!nColocNota))
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                End With
                wApp.Selection.Find.Execute Replace:=wdReplaceAll
            End If
            
        End If
           
        rsCarta.MoveNext
    Loop
    rsCarta.Close
    Set rsCarta = Nothing

 
wAppSource.ActiveDocument.Close
wApp.Visible = True

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




