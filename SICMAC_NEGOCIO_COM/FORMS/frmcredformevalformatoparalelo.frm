VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredFormEvalFormatoParalelo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creditos - Evaluacion - Formato Paralelo"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10830
   Icon            =   "frmcredformevalformatoparalelo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   9000
      TabIndex        =   79
      Top             =   360
      Width           =   1815
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
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
         Left            =   120
         TabIndex        =   82
         Top             =   720
         Width           =   1600
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   81
         Top             =   300
         Width           =   1600
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
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
         Left            =   120
         TabIndex        =   80
         Top             =   300
         Width           =   1600
      End
   End
   Begin VB.CommandButton cmdCajaFlujo 
      Caption         =   "Generar Flujo Caja"
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
      Left            =   9000
      TabIndex        =   78
      Top             =   2400
      Width           =   1810
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Hoja Evaluaci?n"
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
      Left            =   9000
      TabIndex        =   76
      Top             =   1680
      Width           =   1810
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   0
      TabIndex        =   18
      Top             =   2790
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Evaluaci?n"
      TabPicture(0)   =   "frmcredformevalformatoparalelo.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frDatosCredVig"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frDatos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frEstimacionMonto"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frResumen"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Propuesta del Cr?dito"
      TabPicture(1)   =   "frmcredformevalformatoparalelo.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frPropuesta"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   " Gastos del Negocio :"
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
         Height          =   1560
         Left            =   5040
         TabIndex        =   85
         Top             =   2970
         Width           =   5175
         Begin SICMACT.FlexEdit feGastosNegocio 
            Height          =   1200
            Left            =   75
            TabIndex        =   86
            Top             =   180
            Width           =   4935
            _extentx        =   8705
            _extenty        =   2117
            cols0           =   5
            highlight       =   1
            encabezadosnombres=   "-N-Concepto-Monto-Aux"
            encabezadosanchos=   "0-300-3090-1400-0"
            font            =   "frmcredformevalformatoparalelo.frx":0342
            fontfixed       =   "frmcredformevalformatoparalelo.frx":036A
            columnasaeditar =   "X-X-X-3-X"
            listacontroles  =   "0-0-0-0-0"
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            encabezadosalineacion=   "C-C-L-R-C"
            formatosedit    =   "0-0-0-2-0"
            lbeditarflex    =   -1  'True
            lbultimainstancia=   -1  'True
            tipobusqueda    =   6
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Gastos Familiares : "
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
         Height          =   1635
         Left            =   5040
         TabIndex        =   83
         Top             =   4905
         Width           =   5175
         Begin SICMACT.FlexEdit feGastosFamiliares 
            Height          =   1275
            Left            =   105
            TabIndex        =   84
            Top             =   240
            Width           =   4905
            _extentx        =   8652
            _extenty        =   2249
            cols0           =   5
            highlight       =   1
            encabezadosnombres=   "-N-Concepto-Monto-Aux"
            encabezadosanchos=   "0-300-3090-1400-0"
            font            =   "frmcredformevalformatoparalelo.frx":0390
            fontfixed       =   "frmcredformevalformatoparalelo.frx":03B8
            columnasaeditar =   "X-X-X-3-X"
            listacontroles  =   "0-0-0-0-0"
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            backcolorcontrol=   -2147483643
            encabezadosalineacion=   "C-C-L-R-C"
            formatosedit    =   "0-0-0-2-0"
            lbeditarflex    =   -1  'True
            lbultimainstancia=   -1  'True
            tipobusqueda    =   6
            lbbuscaduplicadotext=   -1  'True
            rowheight0      =   300
         End
      End
      Begin VB.Frame frResumen 
         Caption         =   "Resumen"
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
         Height          =   2440
         Left            =   5040
         TabIndex        =   25
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txtMonPropuesto 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   3640
            TabIndex        =   56
            Text            =   "0.00"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtMonParalelo 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H8000000D&
            Height          =   300
            Left            =   3640
            TabIndex        =   55
            Text            =   "0.00"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox txtResumenIncIngresos 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   3640
            TabIndex        =   54
            Text            =   "0.00"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox txtIngresos 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   3640
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtResuMargenBrutoCaja 
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
            Height          =   300
            Left            =   3640
            TabIndex        =   37
            Text            =   "0.00"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label29 
            Caption         =   "Monto Propuesto:"
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
            Height          =   300
            Left            =   2040
            TabIndex        =   61
            Top             =   2100
            Width           =   1575
         End
         Begin VB.Label Label28 
            Caption         =   "Monto Calculado Paralelo:"
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
            Height          =   300
            Left            =   1320
            TabIndex        =   60
            Top             =   1620
            Width           =   2295
         End
         Begin VB.Label Label18 
            Caption         =   "Incremento de Ingresos %:"
            Height          =   300
            Left            =   1700
            TabIndex        =   59
            Top             =   1000
            Width           =   1935
         End
         Begin VB.Label Label17 
            Caption         =   "Ingresos:"
            Height          =   300
            Left            =   2920
            TabIndex        =   58
            Top             =   645
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "Margen Bruto de Caja:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1680
            TabIndex        =   57
            Top             =   280
            Width           =   2055
         End
      End
      Begin VB.Frame frEstimacionMonto 
         Caption         =   "Estimaci?n Monto"
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
         Height          =   2775
         Left            =   480
         TabIndex        =   24
         Top             =   3780
         Width           =   4095
         Begin VB.TextBox txtEstMonOtrosIngresos 
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
            Height          =   300
            Left            =   2040
            TabIndex        =   36
            Text            =   "0.00"
            Top             =   2400
            Width           =   1335
         End
         Begin VB.TextBox txtCutCredVigente 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtEstMonConsFamiliar 
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
            Height          =   300
            Left            =   2040
            TabIndex        =   35
            Text            =   "0.00"
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtEstMonOtrosGasto 
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
            Height          =   300
            Left            =   2040
            TabIndex        =   34
            Text            =   "0.00"
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtMagBruto 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtIncIngreso 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtEstMonIngreso 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label15 
            Caption         =   "Cuota Cred. Vigente:"
            Height          =   300
            Left            =   240
            TabIndex        =   52
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "Otros Ingresos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   51
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label13 
            Caption         =   "Consumo Familiar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   240
            TabIndex        =   50
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label12 
            Caption         =   "Otros Gastos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   49
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "% Margen Bruto:"
            Height          =   300
            Left            =   240
            TabIndex        =   48
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "Incremento Ingresos:"
            Height          =   300
            Left            =   240
            TabIndex        =   47
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label9 
            Caption         =   "Ingresos:"
            Height          =   300
            Left            =   240
            TabIndex        =   46
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame frDatos 
         Caption         =   "Datos"
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
         Height          =   855
         Left            =   480
         TabIndex        =   23
         Top             =   2880
         Width           =   4095
         Begin SICMACT.uSpinner spnDatosIncrIngreso 
            Height          =   300
            Left            =   2400
            TabIndex        =   33
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "MS Sans Serif"
            FontSize        =   9.75
         End
         Begin VB.Label Label34 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3520
            TabIndex        =   90
            Top             =   420
            Width           =   300
         End
         Begin VB.Label Label3 
            Caption         =   "Incremento de Ingreso:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame frPropuesta 
         Caption         =   "Propuesta del Cr?dito"
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
         Height          =   6135
         Left            =   -74760
         TabIndex        =   20
         Top             =   360
         Width           =   10335
         Begin VB.TextBox txtDestino 
            Height          =   650
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   69
            Top             =   5400
            Width           =   9850
         End
         Begin VB.TextBox txtGarantias 
            Height          =   650
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   68
            Top             =   4490
            Width           =   9850
         End
         Begin VB.TextBox txtFormalidadNegocio 
            Height          =   650
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   67
            Top             =   3600
            Width           =   9850
         End
         Begin VB.TextBox txtCrediticia 
            Height          =   650
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   66
            Top             =   2680
            Width           =   9850
         End
         Begin VB.TextBox txtGiroUbicacion 
            Height          =   650
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   65
            Top             =   1800
            Width           =   9850
         End
         Begin VB.TextBox txtEntornoFamiliar 
            Height          =   650
            Left            =   240
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   64
            Top             =   840
            Width           =   9850
         End
         Begin MSMask.MaskEdBox txtFechaVista 
            Height          =   345
            Left            =   8760
            TabIndex        =   63
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   609
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label33 
            Caption         =   "Sobre el Destino y el Impacto del Mismo"
            Height          =   300
            Left            =   240
            TabIndex        =   74
            Top             =   5160
            Width           =   4575
         End
         Begin VB.Label Label32 
            Caption         =   "Sobre los Colaterales o Garant?as"
            Height          =   300
            Left            =   240
            TabIndex        =   73
            Top             =   4280
            Width           =   3975
         End
         Begin VB.Label Label31 
            Caption         =   "Sobre la Consistencia de la Informaci?n y la Formalidad del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   72
            Top             =   3360
            Width           =   6255
         End
         Begin VB.Label Label30 
            Caption         =   "Sobre la Experiencia Crediticia"
            Height          =   300
            Left            =   240
            TabIndex        =   71
            Top             =   2460
            Width           =   4215
         End
         Begin VB.Label Label27 
            Caption         =   "Sobre el Giro y la Ubicaci?n del Negocio"
            Height          =   300
            Left            =   240
            TabIndex        =   70
            Top             =   1560
            Width           =   4095
         End
         Begin VB.Label Label2 
            Caption         =   "Sobre el Entorno Familiar del Cliente o Representante"
            Height          =   300
            Left            =   240
            TabIndex        =   22
            Top             =   600
            Width           =   4695
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha de Visita:"
            Height          =   300
            Left            =   7560
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame frDatosCredVig 
         Caption         =   "Datos Cr?dito Vigente"
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
         Height          =   2460
         Left            =   480
         TabIndex        =   19
         Top             =   320
         Width           =   4095
         Begin VB.TextBox txtIngNeto 
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
            Height          =   300
            Left            =   2040
            TabIndex        =   31
            Text            =   "0.00"
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtCapPago 
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
            Height          =   300
            Left            =   2040
            TabIndex        =   30
            Text            =   "0.00"
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox txtVentas 
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
            Height          =   300
            Left            =   2040
            TabIndex        =   29
            Text            =   "0.00"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox txtSaldoActual 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   28
            Text            =   "0.00"
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtMonAprobado 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   300
            Left            =   2040
            TabIndex        =   27
            Text            =   "0.00"
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label8 
            Caption         =   "Ingreso Neto:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   41
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Excedente :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   40
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Ventas:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   240
            TabIndex        =   39
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Saldo Actual:"
            Height          =   300
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label4 
            Caption         =   "Monto Aprobado:"
            Height          =   300
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Width           =   1335
         End
      End
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   3700
      TabIndex        =   14
      Top             =   320
      Width           =   5220
      Begin VB.TextBox txtActividad 
         Enabled         =   0   'False
         Height          =   320
         Left            =   720
         TabIndex        =   16
         Top             =   200
         Width           =   4455
      End
      Begin VB.Label Label26 
         Caption         =   "Actividad:"
         Height          =   255
         Left            =   40
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdInfromeVista 
      Caption         =   "Informe de Visita"
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
      Left            =   9000
      TabIndex        =   6
      Top             =   2040
      Width           =   1810
   End
   Begin VB.Frame Frame1 
      Height          =   1830
      Left            =   120
      TabIndex        =   0
      Top             =   940
      Width           =   8775
      Begin VB.TextBox txtFechaEduSBS 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Left            =   7320
         TabIndex        =   77
         Top             =   1000
         Width           =   1200
      End
      Begin VB.TextBox txtFechaExpeCaja 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Left            =   1920
         TabIndex        =   75
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox txtUltimoEduSBS 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Left            =   7320
         TabIndex        =   4
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox txtCampana 
         Enabled         =   0   'False
         Height          =   320
         Left            =   1920
         TabIndex        =   3
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtNCredito 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Left            =   1920
         TabIndex        =   2
         Top             =   1000
         Width           =   1200
      End
      Begin VB.TextBox txtNombCliente 
         Enabled         =   0   'False
         Height          =   320
         Left            =   1920
         TabIndex        =   1
         Top             =   200
         Width           =   6735
      End
      Begin VB.TextBox txtExpCredito 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   320
         Left            =   7320
         TabIndex        =   5
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label25 
         Caption         =   "Exposici?n con este Cr?dito:"
         Height          =   255
         Left            =   5280
         TabIndex        =   13
         Top             =   1500
         Width           =   2055
      End
      Begin VB.Label Label24 
         Caption         =   "Fecha ultimo endeud. RCC:"
         Height          =   255
         Left            =   5280
         TabIndex        =   12
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label23 
         Caption         =   "Ultimo endeudamiento RCC:"
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   645
         Width           =   2055
      End
      Begin VB.Label Label22 
         Caption         =   "Campa?a :"
         Height          =   255
         Left            =   1140
         TabIndex        =   10
         Top             =   1500
         Width           =   855
      End
      Begin VB.Label Label21 
         Caption         =   "Num. Cr?ditos Vigentes :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1750
      End
      Begin VB.Label Label20 
         Caption         =   "Exp. en la Caja (Desde) :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   640
         Width           =   1800
      End
      Begin VB.Label Label19 
         Caption         =   "Cliente :"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   650
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2820
      Left            =   45
      TabIndex        =   17
      Top             =   0
      Width           =   8925
      _ExtentX        =   15743
      _ExtentY        =   4974
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Informaci?n del Negocio"
      TabPicture(0)   =   "frmcredformevalformatoparalelo.frx":03DE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ActXCodCta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   495
         Left            =   40
         TabIndex        =   62
         Top             =   480
         Width           =   3580
         _extentx        =   7223
         _extenty        =   873
         texto           =   "Cr?dito:"
      End
   End
   Begin TabDlg.SSTab SSTabRatios1 
      Height          =   675
      Left            =   0
      TabIndex        =   87
      Top             =   9435
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   1191
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ratios e Indicadores"
      TabPicture(0)   =   "frmcredformevalformatoparalelo.frx":03FA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCapaAceptable"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCapacidadPago"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtCapacidadNeta"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin SICMACT.EditMoney txtCapacidadNeta 
         Height          =   300
         Left            =   8790
         TabIndex        =   88
         Top             =   330
         Width           =   945
         _extentx        =   1667
         _extenty        =   529
         font            =   "frmcredformevalformatoparalelo.frx":0416
         forecolor       =   8421504
         text            =   "0"
      End
      Begin VB.Label lblCapacidadPago 
         Caption         =   "Capacidad de Pago:"
         Height          =   200
         Left            =   7240
         TabIndex        =   91
         Top             =   380
         Width           =   1455
      End
      Begin VB.Label lblCapaAceptable 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aceptable"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   9870
         TabIndex        =   89
         Top             =   420
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmCredFormEvalFormatoParalelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************************************************
'*  Nombre:         frmCredFormEvalFormatoPralelo                                                       *
'*  Descripcion:    Formulario para Evaluacion de Creditos que tiene el tipo de Evaluacion Paralelo     *
'*  Creado:         TI-ERS004-2016                                                                      *
'*  Autor:          JOEP, 25-06-2016                                                                    *
'********************************************************************************************************

Option Explicit

Dim gsOpeCod As String

Dim fnTipoCliente As Integer
Dim sCtaCod As String
Dim fnTipoRegMant As Integer
Dim fnTipoPermiso As Integer
Dim fbPermiteGrabar As Boolean
Dim fbBloqueaTodo As Boolean

'Cabecera - Formato Paralelo
Dim fsActividad As String
Dim fsCliente As String
Dim fdPersIng As Date
Dim fnUltimoEduSBS As Double
Dim fnNCred As Integer
Dim fdUltimaEduSBS As Date
Dim fsCampana As String
Dim fnExpCred As Double

'Evaluacion - Formato Paralelo
Dim fnMonAprobado As Double
Dim fnSalActual As Double
Dim fnVentas As Double
Dim fnCapPago As Double
Dim fnIngNeto As Double

Dim fnDatosIncIngreso As Double
Dim fnEstMontoIngresos As Double
Dim fnEstMontoIncIngreso As Double
Dim fnMagBruto As Double
Dim fnOtrGastos As Double
Dim fnConsFamiliar As Double
Dim fnCutCredVigent As Double
Dim fnOtrIngresos As Double

Dim fnMagBrutoCaja As Double
Dim fnIngresos As Double
Dim fnIncIngresos As Double
Dim fnMonParalelo As Double
Dim fnMonPropuesto As Double

'Propuesta del Credito - Formato Paralelo
Dim fdFechaVista As Date
Dim fsSustIncVenta As String

Dim cSPrd As String, cPrd As String
Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
Dim objPista As COMManejador.Pista

Dim nFormato As Integer
Dim fnMontoIni As Double
Dim lnMin As Double, lnMax As Double
Dim lnMinDol As Double, lnMaxDol As Double
Dim nTC As Double
Dim fbGrabar As Boolean

Dim nEstado As Integer

Dim nValorMagBrutoDec As Currency

Dim nValorMagBruto1 As Double
Dim nValorMagBruto2 As Double

Dim Cal1 As Currency
Dim Cal2 As Currency
Dim Cal3 As Currency

Dim nPersPersoneria As Integer

Dim Por As String

'Fjulo de Caja
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
'Fjulo de Caja

'Dim lnColocCondi As Integer
Dim lcMovNro As String 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
'CTI320200110 ERS003-2020. Agreg?
Dim rsFeGastoNeg As ADODB.Recordset
Dim rsFeDatGastoFam As ADODB.Recordset
Dim i As Integer
Dim j As Integer

Dim MatIfiGastoNego As Variant
Dim MatIfiGastoFami As Variant
Dim MatIfiNoSupervisadaGastoNego As Variant
Dim MatIfiNoSupervisadaGastoFami As Variant

Dim fnTotalRefGastoNego As Currency
Dim fnTotalRefGastoFami As Currency

Dim rsGastoNeg As ADODB.Recordset
Dim rsGastoFam As ADODB.Recordset

Dim rsDatGastoFam As ADODB.Recordset
Dim rsDatGastoNeg As ADODB.Recordset
Dim rsDatIfiGastoFami As ADODB.Recordset
Dim rsDatIfiGastoNego As ADODB.Recordset
Dim rsDatIfiNoSupervisadaGastoFami As ADODB.Recordset
Dim rsDatIfiNoSupervisadaGastoNego As ADODB.Recordset
Dim lnFila As Integer

Dim rsDatRatios As ADODB.Recordset
Dim rsRatiosActual As ADODB.Recordset
Dim rsRatiosAceptableCritico As ADODB.Recordset
Dim rsAceptableCritico As ADODB.Recordset
Dim fbImprimirVB As Boolean
Dim pnMontoOtrasIfisConsumo As Double
Dim pnMontoOtrasIfisEmpresarial As Double

'Fin CTI320200110

Public Function Inicio(ByVal psTipoRegMant As Integer, ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnProducto As Integer, _
                     ByVal pnSubProducto As Integer, ByVal pnMontoExpEsteCred As Double, ByVal pbImprimir As Boolean, ByVal pnEstado As Integer, _
                     Optional ByVal pbImprimirVB As Boolean = False) As Boolean
    
    gsOpeCod = ""
    lcMovNro = "" 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
    nPersPersoneria = 0
    fbImprimirVB = pbImprimirVB 'CTI3ERS0032020
    Call LimpiaFormulario
    Call LLenarFormulario
    
    nFormato = pnFormato
    nEstado = pnEstado
    fbGrabar = False
    
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsDCredito As ADODB.Recordset
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim rsDLLenarEvaluacion As ADODB.Recordset
        
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
        
    sCtaCod = psCtaCod
    
    fnTipoRegMant = psTipoRegMant
    ActXCodCta.NroCuenta = sCtaCod
    
    If nEstado = 2001 Then
        cmdImprimir.Enabled = True
        cmdInfromeVista.Enabled = True
        cmdCajaFlujo.Enabled = True 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Else
        cmdCajaFlujo.Enabled = False 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        cmdImprimir.Enabled = False
        cmdInfromeVista.Enabled = False
    End If
    
    'Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    'Set rsDCredito = oDCOMFormatosEval.RecuperarDatosCredEvalFormatoParalelo(sCtaCod) ' Llenar Datos en la Cabecera Informacion de Negocio
    
    'lnColocCondi = rsDCredito!nColocCondicion
    
'    If lnColocCondi = 4 Then
'        SSTab1.TabEnabled(1) = False
'    Else
'        SSTab1.TabEnabled(1) = True
'    End If
    
    '(3: Analista, 2: Coordinador, 1: JefeAgencia)
    fnTipoPermiso = oNCOMFormatosEval.ObtieneTipoPermisoCredEval(gsCodCargo) ' Obtener el tipo de Permiso, Segun Cargo
    Call CargarFlexEdit 'CTI320200110 ERS003-2020. Agreg?
    
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsDCredito = oDCOMFormatosEval.RecuperarDatosCredEvalFormatoParalelo(sCtaCod) 'Llenar Datos en la Cabecera Informacion de Negocio
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsDLLenarEvaluacion = oDCOMFormatosEval.RecuperarDatosCredEvalFPEvaluacion(sCtaCod) 'Llenar Datos en Evaluacion
    Set rsAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(psCtaCod) 'Obtenemos Datos, Aceptable / Critico de los Ratios'CTI320200110 ERS003-2020. Agreg?
    
    If Not (rsDCredito.EOF And rsDCredito.BOF) Then
        nPersPersoneria = rsDCredito!nPersPersoneria 'Para saber si la persona es Natural o Juridica
    End If
    
    
    If CargaControlesTipoPermiso(fnTipoPermiso) Then
    
        If fnTipoRegMant = 1 Then
            If (rsDLLenarEvaluacion.EOF And rsDLLenarEvaluacion.BOF) Then
                MsgBox "El Cliente no cumple condiciones para Cr?dito Paralelo:" & Chr(13) & _
                       "No Aplica a Clientes Nuevos.", vbInformation, "Alerta"
                Exit Function
            End If
            
            '------------------------------------------------------------------------
            'Validacion de Tipo de Credito
            '========================================================================
            '30% Tipo de Credito
            If Left(rsDLLenarEvaluacion!cTpoCredCod, 1) = 4 Or Left(rsDLLenarEvaluacion!cTpoCredCod, 1) = 5 Then
                Por = "30%"
                Cal1 = (rsDLLenarEvaluacion!nMontoCol) * 0.3
                Cal2 = Cal1 + (rsDLLenarEvaluacion!nMontoCol)
                Cal3 = Cal2 - (rsDLLenarEvaluacion!nSaldo)
                
            '40% Tipo de Credito
            Else
                Por = "40%"
                Cal1 = (rsDLLenarEvaluacion!nMontoCol) * 0.4
                Cal2 = Cal1 + (rsDLLenarEvaluacion!nMontoCol)
                Cal3 = Cal2 - (rsDLLenarEvaluacion!nSaldo)
            
            End If
            
            If (rsDLLenarEvaluacion!nMontoPro) <= Cal3 Then
            Else
                MsgBox "El cliente no cumple con las condiciones necesarias para continuar con el proceso de Cr?dito:" & Chr(13) & "" & Chr(13) & _
                       "El Monto Propuesto: " & Format((rsDLLenarEvaluacion!nMontoPro), "#,##0.00") & "" & Chr(13) & _
                       "Es mayor " & Chr(13) & _
                       "Al Monto Calculado: " & Format(Cal3, "#,##0.00") & " al " & Por & "" & Chr(13) & "" & Chr(13) & _
                       "Se tom? en consideraci?n el Saldo Disponible con relaci?n al Monto Inicial del Cr?dito Vigente.", vbInformation, "Alerta"
                Exit Function
            End If
            
            '-----------------------------------------------------------------------
                        
            If Not (rsDCredito.EOF And rsDCredito.BOF) Then
                
                If (rsDCredito!cActiGiro) = "" Then
                    MsgBox ("Por favor, actualizar los datos del cliente. " & Chr(13) & "(Actividad o Giro del negocio)"), vbInformation, "Alerta"
                    Exit Function
                End If
            
                fsActividad = IIf((rsDCredito!cActiGiro) = "", "", (rsDCredito!cActiGiro))
                fsCliente = Trim(rsDCredito!cPersNombre)
                fdPersIng = Trim(rsDCredito!dPersIng)
                fnUltimoEduSBS = Trim(rsDCredito!nUltimoEduSBS)
                fnNCred = Trim(rsDCredito!nCreditos)
                fdUltimaEduSBS = Trim(rsDCredito!dFechaUltEnduSBS)
                fsCampana = Trim(rsDCredito!CDescripcion)
                fnExpCred = Trim(rsDCredito!nExpoCred)
                
                '============================================
                    If nPersPersoneria = 2 Then
                        txtEstMonConsFamiliar.Enabled = False ' si la Persona es Juridica se pondra solo lectura Monto Consumo Familiar
                    End If
                '=============================================
            End If
            
            If Not (rsDLLenarEvaluacion.BOF And rsDLLenarEvaluacion.EOF) Then
                fnMonAprobado = Trim(rsDLLenarEvaluacion!nMontoCol)
                fnSalActual = Trim(rsDLLenarEvaluacion!nSaldo)
                fnCutCredVigent = Trim(rsDLLenarEvaluacion!nMontoCuota)
                fnMonPropuesto = Trim(rsDLLenarEvaluacion!nMontoPro)
            End If
                     
            If Not (rsDCredito.EOF And rsDCredito.BOF) Then
                txtActividad.Text = fsActividad
                txtNombCliente.Text = fsCliente
                txtFechaExpeCaja.Text = fdPersIng
                txtUltimoEduSBS.Text = Format(fnUltimoEduSBS, "#,##0.00")
                txtNCredito.Text = Format(fnNCred, "0#")
                txtFechaEduSBS.Text = fdUltimaEduSBS
                txtCampana.Text = fsCampana
                txtExpCredito.Text = Format(pnMontoExpEsteCred, "#,##0.00")
            End If
            
            If Not (rsDLLenarEvaluacion.BOF And rsDLLenarEvaluacion.EOF) Then
                txtMonAprobado.Text = Format(fnMonAprobado, "#,##0.00")
                txtSaldoActual.Text = Format(fnSalActual, "#,##0.00")
                txtEstMonIngreso.Text = Format(fnVentas, "#,##0.00")
                txtCutCredVigente.Text = Format(fnCutCredVigent, "#,##0.00")
                txtIngresos.Text = Format(fnVentas, "#,##0.00")
                txtMonPropuesto.Text = Format(fnMonPropuesto, "#,##0.00")
            End If
            
            cmdActualizar.Visible = False
            cmdGuardar.Visible = True
            
            'CTI3 ERS0032020
            'Carga de rsDatIfiGastoNego (Ifis Gastos Negocio)
'            Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
            Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    
            Set rsDatIfiGastoNego = oDCOMFormatosEval.RecuperaDatosIfiCuota(sCtaCod, nFormato, gFormatoGastosNego, gCodCuotaIfiGastoNego)
            Set rsDatIfiGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(sCtaCod, nFormato, gFormatoGastosFami, gCodCuotaIfiGastoFami)

            ReDim MatIfiGastoNego(rsDatIfiGastoNego.RecordCount, 4)
            i = 0
            Do While Not rsDatIfiGastoNego.EOF
                MatIfiGastoNego(i, 0) = rsDatIfiGastoNego!nNroCuota
                MatIfiGastoNego(i, 1) = rsDatIfiGastoNego!CDescripcion
                MatIfiGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiGastoNego!nMonto), 0, rsDatIfiGastoNego!nMonto), "#0.00")
                rsDatIfiGastoNego.MoveNext
                  i = i + 1
            Loop
            rsDatIfiGastoNego.Close
            Set rsDatIfiGastoNego = Nothing
            
            'Carga de rsDatIfiGastoFami (Ifis Gastos Familiares)
            ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
            j = 0
            Do While Not rsDatIfiGastoFami.EOF
                MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
                MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!CDescripcion
                MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#0.00")
                rsDatIfiGastoFami.MoveNext
            j = j + 1
            Loop
            rsDatIfiGastoFami.Close
            Set rsDatIfiGastoFami = Nothing
            
        ElseIf fnTipoRegMant = 2 Then
        
            If fnTipoRegMant = 2 And Mantenimineto(IIf(fnTipoRegMant = 2, False, True)) = False Then
                    MsgBox "Cliente no ha Hasido Evaluado", vbInformation, "Aviso"
                    Exit Function
            End If
            
             '============================================
             If nPersPersoneria = 2 Then
                txtEstMonConsFamiliar.Enabled = False ' si la Persona es Juridica se pondra solo lectura Monto Consumo Familiar
             End If
                '=============================================
                
             cmdGuardar.Visible = False
             cmdActualizar.Visible = True
        
        ElseIf fnTipoRegMant = 3 Then
            
            Call Mantenimineto(IIf(fnTipoRegMant = 3, False, True))
            Call Consultar
            
            If pnEstado = 2001 Or pnEstado = 2002 Then
                cmdInfromeVista.Enabled = True
                cmdImprimir.Enabled = True
            End If
            
            'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
            gsOpeCod = gCredConsultarEvaluacionCred
            lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gConsultar, "Evaluacion Credito Formato 9 - Paralelo", sCtaCod, gCodigoCuenta
            Set objPista = Nothing
            'Fin LUCV20181220
            
        End If
    Else
        Unload Me
        Exit Function
    End If
       
    
        fbGrabar = False
        
        If Not pbImprimir Then
            If fbImprimirVB Then
                Call cmdActualizar_Click
            End If
            Me.Show 1
        Else
            cmdImprimir_Click
        End If
        
        Inicio = fbGrabar
        
End Function

'Actualizar Datos
Private Sub cmdActualizar_Click()
    Dim oCredFormEval As COMNCredito.NCOMFormatosEval
    Dim ActualizarFormatoParalelo As Boolean
   
    If ValidarDatosFormatoParalelo Then
       gsOpeCod = gCredMantenimientoEvaluacionCred
       Set objPista = New COMManejador.Pista
       Set oCredFormEval = New COMNCredito.NCOMFormatosEval
       lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
       
       'CTI320200110 ERS003-2020. Agreg?
       Set rsGastoNeg = IIf(feGastosNegocio.rows - 1 > 0, feGastosNegocio.GetRsNew(), Nothing)
       Set rsGastoFam = IIf(feGastosFamiliares.rows - 1 > 0, feGastosFamiliares.GetRsNew(), Nothing)
       'Fin CTI320200110
       
       ActualizarFormatoParalelo = oCredFormEval.ActualizarfrmCredFormEvalFormatoParalelo(sCtaCod, nFormato, CDate(txtFechaExpeCaja.Text), _
                                                                                          txtMonAprobado.Text, txtSaldoActual.Text, txtVentas.Text, txtCapPago.Text, txtIngNeto.Text, _
                                                                                          spnDatosIncrIngreso.valor, _
                                                                                          txtEstMonIngreso.Text, txtIncIngreso.Text, txtMagBruto.Text, txtEstMonOtrosGasto.Text, txtEstMonConsFamiliar.Text, txtCutCredVigente.Text, txtEstMonOtrosIngresos.Text, _
                                                                                          txtResuMargenBrutoCaja.Text, txtIngresos.Text, txtResumenIncIngresos.Text, txtMonParalelo.Text, txtMonPropuesto.Text, _
                                                                                          CDate(txtFechaVista.Text), txtDestino.Text, txtEntornoFamiliar.Text, txtGiroUbicacion.Text, txtCrediticia.Text, txtFormalidadNegocio.Text, txtGarantias.Text, _
                                                                                          rsGastoNeg, rsGastoFam, MatIfiGastoNego, MatIfiGastoFami, MatIfiNoSupervisadaGastoNego, MatIfiNoSupervisadaGastoFami)
                                                                                          'rsGastoNeg, rsGastoFam, MatIfiGastoNego, MatIfiGastoFami, MatIfiNoSupervisadaGastoNego, MatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agreg?
                                                                                          'IIf(txtFechaVista.Text = "__/__/____", CDate(gdFecSis), txtFechaVista.Text)
        
        
        If ActualizarFormatoParalelo Then
            'CTI320200110 ERS003-2020. Agreg?
            Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval 'CTI320200110 ERS003-2020. Agreg?
            Call oDCOMFormatosEval.RecalculaIndicadoresyRatiosEvaluacion(sCtaCod)
            Set rsRatiosActual = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
            Set rsRatiosAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod)
            'Fin CTI320200110
            fbGrabar = True
            'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato Paralelo", sCtaCod, gCodigoCuenta 'LUCV20181220 Coment?, Anexo01 de Acta 199-2018
            objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 9 - Paralelo", sCtaCod, gCodigoCuenta 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
            Set objPista = Nothing 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
            If Not fbImprimirVB Then
                MsgBox "Los Datos se Actualizaron Correctamente", vbInformation, "Aviso"
            End If
                Dim objCredito As COMDCredito.DCOMCredito
                Set objCredito = New COMDCredito.DCOMCredito
                Call objCredito.ActualizarEstadoxVB(ActXCodCta.NroCuenta, 1)
        Else
        
            MsgBox "Hubo error al grabar la informacion", vbError, "Error"
            
        End If
            
            'If lnColocCondi <> 4 Then
                cmdInfromeVista.Enabled = True
            'End If
            
        cmdActualizar.Enabled = False
        cmdGuardar.Visible = False
    
        If (nEstado = 2001) Then
            cmdImprimir.Enabled = True
        End If
        
        'CTI320200110 ERS003-2020. Agreg?
        'Actualizacion de los Ratios
        txtCapacidadNeta.Text = CStr(rsRatiosActual!nCapPagNeta * 100) & "%"

        'Ratios: Aceptable / Critico ->*****
        If Not (rsRatiosAceptableCritico.EOF Or rsRatiosAceptableCritico.BOF) Then
            If rsRatiosAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                Me.lblCapaAceptable.Caption = "Aceptable"
                Me.lblCapaAceptable.ForeColor = &H8000&
            Else
                Me.lblCapaAceptable.Caption = "Cr?tico"
                Me.lblCapaAceptable.ForeColor = vbRed
            End If
            
        Else
            lblCapaAceptable.Visible = False
        End If
        'Fin Ratios <-****
        
        Set rsRatiosActual = Nothing
        Set rsRatiosAceptableCritico = Nothing
        
        If fbImprimirVB Then
           cmdActualizar.Enabled = True
           fbImprimirVB = False
        End If
            
        'Fin CTI320200110
        
    End If
    
End Sub

''Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
Private Sub cmdCajaFlujo_Click()
Dim lsArchivo As String
Dim lbLibroOpen As Boolean

    lsArchivo = App.Path & "\Spooler\FlujoCaja_FormatoParalelo" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls"
    lbLibroOpen = ExcelInicio(lsArchivo, xlAplicacion, xlLibro)
    
    If lbLibroOpen Then
        If generaExcel = True Then
            ExcelFin lsArchivo, xlAplicacion, xlLibro, xlHoja1
            AbrirArchivo "FlujoCaja_FormatoParalelo" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls", App.Path & "\Spooler"
        End If
    End If
End Sub

Public Function generaExcel() As Boolean
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim oCont As COMConecta.DCOMConecta
    Dim i As Integer
    Dim nCon As Integer
    Dim nIncre As Integer
    
    Dim cNombClie As String
    Dim cUserAnal As String
    Dim nVentas As Double
    Dim nMontoPro As Double
    Dim nIncrPorc As Double
    Dim nMargenBruto As Double
    Dim nIngrNeto As Double
    Dim nTem As Double
    Dim nPlazo As Integer
    
    generaExcel = True
    
    nVentas = 0
    nMontoPro = 0
    nIncrPorc = 0
    nMargenBruto = 0
    nIngrNeto = 0
    nTem = 0
    nPlazo = 0
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosCabecera  '" & ActXCodCta.NroCuenta & "'"
    
    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rs = oCont.CargaRecordSet(ssql)
        oCont.CierraConexion
    Set oCont = Nothing
    
    If Not (rs.EOF And rs.BOF) Then
        cNombClie = rs!cNombClie
        cUserAnal = rs!cUserAnal
        nVentas = rs!nVentas
        nMontoPro = rs!nMontoPropuesto
        nIncrPorc = rs!nIncreIngreResumen
        nMargenBruto = rs!nMargenBrutoCaja
        nIngrNeto = rs!nIngresoNeto
        nTem = rs!nTasaInteres
        nPlazo = rs!Plazo
    Else
        MsgBox "Error, Comun?quese con el ?rea de TI", vbInformation, "!Error!"
        generaExcel = False
        Exit Function
    End If
    
    'proteger Libro
    xlAplicacion.ActiveWorkbook.Protect (123)
    
    'Adiciona una hoja
    ExcelAddHoja "Hoja1", xlLibro, xlHoja1, True
               
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    
           
    xlHoja1.Cells(2, 2) = "FLUJO DE CAJA PARA PARALELO POR CAMPA?A"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).Font.Bold = True
    
    xlHoja1.Cells(4, 1) = "CLIENTE: "
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(4, 2) = cNombClie
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 6)).Font.Bold = True
            
    xlHoja1.Cells(5, 1) = "ANALISTA: "
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 1)).HorizontalAlignment = xlLeft
            
    xlHoja1.Cells(5, 2) = cUserAnal
    xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 6)).Font.Bold = True
            
    xlHoja1.Cells(7, 2) = "Datos Financieros del Clientes (De la Evaluacion)"
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 7)).Cells.Interior.Color = RGB(141, 180, 226)
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 7)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 7)).Font.Bold = True
    CuadroExcel xlHoja1, 2, 7, 7, 7
    CuadroExcel xlHoja1, 2, 7, 2, 11, False
    CuadroExcel xlHoja1, 2, 7, 7, 11, True
    
    xlHoja1.Cells(8, 2) = "Ventas mensuales promedio"
    xlHoja1.Range(xlHoja1.Cells(8, 2), xlHoja1.Cells(8, 5)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(8, 2), xlHoja1.Cells(8, 5)).HorizontalAlignment = xlLeft
    xlHoja1.Cells(8, 6) = "S/."
    xlHoja1.Cells(8, 7) = Format(nVentas, "#,00.00")
    
    xlHoja1.Cells(9, 2) = "Margen bruto (1-cv/v (en %))"
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 5)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(9, 5)).HorizontalAlignment = xlLeft
    xlHoja1.Cells(9, 6) = "%"
    xlHoja1.Cells(9, 7) = Format(((nMargenBruto / nVentas) * 100), "#00")
    
    xlHoja1.Range(xlHoja1.Cells(10, 2), xlHoja1.Cells(10, 5)).MergeCells = True
        
    xlHoja1.Range(xlHoja1.Cells(11, 2), xlHoja1.Cells(11, 5)).MergeCells = True
    xlHoja1.Cells(11, 2) = "Margen disponible prom. sd/v (en %)"
    xlHoja1.Cells(11, 6) = "%"
    xlHoja1.Cells(11, 7) = Format(((nIngrNeto / nVentas) * 100), "#00")
       
    xlHoja1.Cells(13, 2) = "Datos de la Operacion"
    xlHoja1.Range(xlHoja1.Cells(13, 2), xlHoja1.Cells(13, 7)).Cells.Interior.Color = RGB(141, 180, 226)
    xlHoja1.Range(xlHoja1.Cells(13, 2), xlHoja1.Cells(13, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(13, 2), xlHoja1.Cells(13, 7)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(13, 2), xlHoja1.Cells(13, 7)).Font.Bold = True
    CuadroExcel xlHoja1, 2, 13, 7, 13
    CuadroExcel xlHoja1, 2, 13, 2, 18, False
    CuadroExcel xlHoja1, 2, 13, 7, 18, True
    
    xlHoja1.Range(xlHoja1.Cells(14, 2), xlHoja1.Cells(14, 5)).MergeCells = True
    xlHoja1.Cells(14, 2) = "Plazo de la operaci?n en meses"
    xlHoja1.Cells(14, 7) = nPlazo
    
    xlHoja1.Range(xlHoja1.Cells(15, 2), xlHoja1.Cells(15, 5)).MergeCells = True
    xlHoja1.Cells(15, 2) = "Monto de credito exp. En MN"
    xlHoja1.Cells(15, 6) = "S/."
    xlHoja1.Cells(15, 7) = Format(nMontoPro, "#,00.00")
    
    xlHoja1.Range(xlHoja1.Cells(16, 2), xlHoja1.Cells(16, 5)).MergeCells = True
    xlHoja1.Cells(16, 2) = "TEM"
    xlHoja1.Cells(16, 6) = "%"
    xlHoja1.Cells(16, 7) = nTem
    
    xlHoja1.Range(xlHoja1.Cells(17, 2), xlHoja1.Cells(17, 5)).MergeCells = True
    xlHoja1.Cells(17, 2) = "Numero de Mes del Desembolso"
    xlHoja1.Cells(17, 7) = 1
    
    xlHoja1.Range(xlHoja1.Cells(18, 2), xlHoja1.Cells(18, 5)).MergeCells = True
    xlHoja1.Cells(18, 2) = "Numero de Mes de compra campa?a"
    xlHoja1.Cells(18, 7) = 1
       
    xlHoja1.Cells(20, 2) = "Datos de la Campa?a"
    xlHoja1.Range(xlHoja1.Cells(20, 2), xlHoja1.Cells(20, 7)).Cells.Interior.Color = RGB(141, 180, 226)
    xlHoja1.Range(xlHoja1.Cells(20, 2), xlHoja1.Cells(20, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(20, 2), xlHoja1.Cells(20, 7)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(20, 2), xlHoja1.Cells(20, 7)).Font.Bold = True
    CuadroExcel xlHoja1, 2, 20, 7, 20
    CuadroExcel xlHoja1, 2, 20, 2, 22, False
    CuadroExcel xlHoja1, 2, 20, 7, 22, True
    
    xlHoja1.Range(xlHoja1.Cells(21, 2), xlHoja1.Cells(21, 5)).MergeCells = True
    xlHoja1.Cells(21, 2) = "Mes de la campa?a"
    xlHoja1.Cells(21, 7) = nPlazo
    
    xlHoja1.Range(xlHoja1.Cells(22, 2), xlHoja1.Cells(22, 5)).MergeCells = True
    xlHoja1.Cells(22, 2) = "Increm. % esperado en ventas por campa?a"
    xlHoja1.Cells(22, 6) = "%"
    xlHoja1.Cells(22, 7) = nIncrPorc

    xlHoja1.Cells(24, 2) = "ADICIONALES"
    xlHoja1.Range(xlHoja1.Cells(24, 2), xlHoja1.Cells(24, 7)).Cells.Interior.Color = RGB(141, 180, 226)
    xlHoja1.Range(xlHoja1.Cells(24, 2), xlHoja1.Cells(24, 7)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(24, 2), xlHoja1.Cells(24, 7)).HorizontalAlignment = xlLeft
    xlHoja1.Range(xlHoja1.Cells(24, 2), xlHoja1.Cells(24, 7)).Font.Bold = True
    CuadroExcel xlHoja1, 2, 24, 7, 24
    CuadroExcel xlHoja1, 2, 24, 2, 26, False
    CuadroExcel xlHoja1, 2, 24, 7, 26, True
          
    xlHoja1.Range(xlHoja1.Cells(25, 2), xlHoja1.Cells(25, 5)).MergeCells = True
    xlHoja1.Cells(25, 2) = "MARGEN BRUTO"
    xlHoja1.Cells(25, 7) = Format(nMargenBruto, "#,00.00")
    
    xlHoja1.Range(xlHoja1.Cells(26, 2), xlHoja1.Cells(26, 5)).MergeCells = True
    xlHoja1.Cells(26, 2) = "INGRESO NETO"
    xlHoja1.Cells(26, 7) = Format(nIngrNeto, "#,00.00")

'===Columna de Calculo Dinamico
nCon = 12
nIncre = 1
For i = 1 To nPlazo

If i <> 1 Then
    If nIncre = i Then
        nCon = nCon + 1
        nIncre = nIncre + 1
    End If
Else
  nIncre = nIncre + 1
End If

'Cuadro 1
    xlHoja1.Cells(4, nCon) = "Mes"
    xlHoja1.Cells(5, nCon) = i
    xlHoja1.Range(xlHoja1.Cells(4, 12), xlHoja1.Cells(4, nCon)).Cells.Interior.Color = RGB(141, 180, 226)
    CuadroExcel xlHoja1, 12, 4, nCon, 5

'Cuadro 2
    xlHoja1.Cells(7, 9) = "Ventas mensuales promedio"
    xlHoja1.Range(xlHoja1.Cells(7, 9), xlHoja1.Cells(7, 11)).MergeCells = True
    xlHoja1.Cells(7, nCon) = nVentas

    xlHoja1.Cells(8, 9) = "Ventas de la campa?a"
    xlHoja1.Range(xlHoja1.Cells(8, 9), xlHoja1.Cells(8, 11)).MergeCells = True
    xlHoja1.Cells(8, nCon) = IIf(i = nPlazo, (nVentas * nIncrPorc) / 100, 0)

    xlHoja1.Cells(9, 9) = "Venta Totales"
    xlHoja1.Range(xlHoja1.Cells(9, 9), xlHoja1.Cells(9, 11)).MergeCells = True
    xlHoja1.Cells(9, nCon) = xlHoja1.Cells(7, nCon) + xlHoja1.Cells(8, nCon)

    CuadroExcel xlHoja1, 9, 7, 9, 9, False
    CuadroExcel xlHoja1, 9, 7, nCon, 9, True

'Cuadro 3
    xlHoja1.Cells(11, 9) = "Costo de Ventas promedio"
    xlHoja1.Range(xlHoja1.Cells(11, 9), xlHoja1.Cells(11, 10)).MergeCells = True
    xlHoja1.Cells(11, nCon) = (nVentas * (1 - ((xlHoja1.Cells(25, 7) / xlHoja1.Cells(8, 7) * 100) / 100)))

    xlHoja1.Cells(12, 9) = "Compra de ventas de campa?a"
    xlHoja1.Range(xlHoja1.Cells(12, 9), xlHoja1.Cells(12, 11)).MergeCells = True
    xlHoja1.Cells(12, nCon) = IIf(i = xlHoja1.Cells(18, 7), ((nVentas * (nIncrPorc / 100)) * (1 - ((xlHoja1.Cells(25, 7) / xlHoja1.Cells(8, 7) * 100) / 100))), 0)

    xlHoja1.Cells(13, 9) = "Otros movimientos de KW"
    xlHoja1.Range(xlHoja1.Cells(13, 9), xlHoja1.Cells(13, 11)).MergeCells = True
    xlHoja1.Cells(13, nCon) = 0
    
    xlHoja1.Cells(14, 9) = "Otras compras/gastos no previsto"
    xlHoja1.Range(xlHoja1.Cells(14, 9), xlHoja1.Cells(14, 11)).MergeCells = True
    xlHoja1.Cells(14, nCon) = 0

    xlHoja1.Cells(15, 9) = "Otros Egresos"
    xlHoja1.Range(xlHoja1.Cells(15, 9), xlHoja1.Cells(15, 11)).MergeCells = True
    xlHoja1.Cells(15, nCon) = nVentas - (nVentas * (xlHoja1.Cells(11, 7) / 100)) - xlHoja1.Cells(11, nCon)

    xlHoja1.Cells(16, 9) = "Total Egreso"
    xlHoja1.Range(xlHoja1.Cells(16, 9), xlHoja1.Cells(16, 11)).MergeCells = True
    xlHoja1.Cells(16, nCon) = xlHoja1.Cells(11, nCon) + xlHoja1.Cells(12, nCon) + xlHoja1.Cells(13, nCon) + xlHoja1.Cells(14, nCon) + xlHoja1.Cells(15, nCon)

    CuadroExcel xlHoja1, 9, 11, 9, 16, False
    CuadroExcel xlHoja1, 9, 11, nCon, 16, True

'Cuadro 4
    xlHoja1.Range(xlHoja1.Cells(18, 9), xlHoja1.Cells(18, 11)).MergeCells = True
    xlHoja1.Cells(18, 9) = "Saldo Disponible Operativo"
    xlHoja1.Cells(18, nCon) = xlHoja1.Cells(9, nCon) - xlHoja1.Cells(16, nCon)

    CuadroExcel xlHoja1, 9, 18, 9, 18, False
    CuadroExcel xlHoja1, 9, 18, nCon, 18, True

'Cuadro 5
    xlHoja1.Cells(20, 9) = "Prestamo de la Caja"
    xlHoja1.Range(xlHoja1.Cells(20, 9), xlHoja1.Cells(20, 11)).MergeCells = True
    xlHoja1.Cells(20, nCon) = IIf(i = xlHoja1.Cells(17, 7), xlHoja1.Cells(15, 7), 0)

    xlHoja1.Cells(21, 9) = "Pago a la Caja"
    xlHoja1.Range(xlHoja1.Cells(21, 9), xlHoja1.Cells(21, 11)).MergeCells = True
    xlHoja1.Cells(21, nCon) = "-" & IIf(nPlazo = i, Format((xlHoja1.Cells(15, 7) * ((1 + nTem / 100) ^ (nPlazo))), "#0"), 0)

    xlHoja1.Cells(22, 9) = "Caja Inicial"
    xlHoja1.Range(xlHoja1.Cells(22, 9), xlHoja1.Cells(22, 11)).MergeCells = True
    xlHoja1.Cells(22, nCon) = 0

    xlHoja1.Cells(23, 9) = "Caja minima"
    xlHoja1.Range(xlHoja1.Cells(23, 9), xlHoja1.Cells(23, 11)).MergeCells = True
    xlHoja1.Cells(23, nCon) = 0

    xlHoja1.Cells(24, 9) = "Saldo Disponible Financiero"
    xlHoja1.Range(xlHoja1.Cells(24, 9), xlHoja1.Cells(24, 11)).MergeCells = True
    xlHoja1.Cells(24, nCon) = Format(xlHoja1.Cells(20, nCon) + xlHoja1.Cells(21, nCon) + xlHoja1.Cells(22, nCon) + xlHoja1.Cells(23, nCon), "#0")

    CuadroExcel xlHoja1, 9, 20, 9, 24, False
    CuadroExcel xlHoja1, 9, 20, nCon, 24, True

'Cuadro 6
    xlHoja1.Cells(26, 9) = "Saldo Final"
    xlHoja1.Range(xlHoja1.Cells(26, 9), xlHoja1.Cells(26, 10)).MergeCells = True
    xlHoja1.Cells(26, nCon) = xlHoja1.Cells(18, nCon) + xlHoja1.Cells(24, nCon)

    xlHoja1.Cells(27, 9) = "Saldo Acumulado"
    xlHoja1.Range(xlHoja1.Cells(27, 9), xlHoja1.Cells(27, 10)).MergeCells = True
    
    xlHoja1.Cells(27, nCon) = IIf(nPlazo < i, 0, IIf(xlHoja1.Cells(26, nCon - 1) = "", 0, xlHoja1.Cells(27, nCon - 1)) + xlHoja1.Cells(26, nCon))

    CuadroExcel xlHoja1, 9, 26, 9, 27, False
    CuadroExcel xlHoja1, 9, 26, nCon, 27, True
         
 Next i
          
   rs.Close
   xlHoja1.Cells.Select
   xlHoja1.Cells.Font.Name = "Arial"
   xlHoja1.Cells.Font.Size = 9
   xlHoja1.Cells.EntireColumn.AutoFit
    
    'xlAplicacion.Worksheets("Hoja1").Protect ("123")
      
   MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "!Exito!"
    
End Function
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja

'Salir
Private Sub cmdCancelar_Click()
    Unload Me

    Set MatIfiGastoNego = Nothing 'CTI320200110 ERS003-2020. Agreg?
    Set MatIfiGastoFami = Nothing 'CTI320200110 ERS003-2020. Agreg?
    Set MatIfiNoSupervisadaGastoNego = Nothing 'CTI320200110 ERS003-2020. Agreg?
    Set MatIfiNoSupervisadaGastoFami = Nothing 'CTI320200110 ERS003-2020. Agreg?
End Sub

'Guardar Datos
Private Sub Cmdguardar_Click()
    Dim oCredFormEval As COMNCredito.NCOMFormatosEval
    Dim GrabarFormatoParalelo As Boolean
    Dim rsEvaluacion As ADODB.Recordset
    
    If ValidarDatosFormatoParalelo Then
        gsOpeCod = gCredRegistrarEvaluacionCred
        Set objPista = New COMManejador.Pista
        Set rsEvaluacion = LenarRecordset_Evaluacion
        Set oCredFormEval = New COMNCredito.NCOMFormatosEval
                
        'CTI320200110 ERS003-2020. Agreg?
        Set rsGastoNeg = IIf(feGastosNegocio.rows - 1 > 0, feGastosNegocio.GetRsNew(), Nothing)
        Set rsGastoFam = IIf(feGastosFamiliares.rows - 1 > 0, feGastosFamiliares.GetRsNew(), Nothing)
        'Fin CTI320200110
       
        If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        GrabarFormatoParalelo = oCredFormEval.GrabarfrmCredFormEvalFormatoParalelo(sCtaCod, nFormato, Trim(txtActividad.Text), CDate(txtFechaExpeCaja.Text), _
                                                                                   txtUltimoEduSBS.Text, txtNCredito.Text, CDate(txtFechaEduSBS.Text), _
                                                                                   Trim(txtCampana.Text), txtExpCredito.Text, _
                                                                                   rsEvaluacion, txtVentas, txtCapPago, txtIngNeto, _
                                                                                   spnDatosIncrIngreso.valor, txtEstMonIngreso.Text, txtIncIngreso.Text, txtMagBruto.Text, txtEstMonOtrosGasto.Text, txtEstMonConsFamiliar.Text, txtCutCredVigente.Text, _
                                                                                   txtEstMonOtrosIngresos.Text, txtResuMargenBrutoCaja.Text, txtIngresos.Text, txtResumenIncIngresos, txtMonParalelo, _
                                                                                   CDate(txtFechaVista.Text), txtDestino.Text, txtEntornoFamiliar.Text, txtGiroUbicacion.Text, txtCrediticia.Text, txtFormalidadNegocio.Text, txtGarantias.Text, _
                                                                                   rsGastoNeg, rsGastoFam, MatIfiGastoNego, MatIfiGastoFami, MatIfiNoSupervisadaGastoNego, MatIfiNoSupervisadaGastoFami)
                                                                                   
                                                                                'rsGastoNeg, rsGastoFam, MatIfiGastoNego, MatIfiGastoFami, MatIfiNoSupervisadaGastoNego, MatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agreg?
                                                                                'IIf(txtFechaVista.Text = "__/__/____", CDate(gdFecSis), txtFechaVista.Text)
        If GrabarFormatoParalelo Then
            'CTI320200110 ERS003-2020. Agreg?
            Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval 'CTI320200110 ERS003-2020. Agreg?
            Call oDCOMFormatosEval.RecalculaIndicadoresyRatiosEvaluacion(sCtaCod)
            Set rsRatiosActual = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
            Set rsRatiosAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod)
            'Fin CTI320200110
            fbGrabar = True
            'RECO20161020 ERS060-2016 **********************************************************
            Dim oNCOMColocEval As New NCOMColocEval
            'Dim lcMovNro As String 'LUCV20181220 Coment?, Anexo01 de Acta 199-2018
            lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
            
            If Not ValidaExisteRegProceso(sCtaCod, gTpoRegCtrlEvaluacion) Then
               'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Coment?, Anexo01 de Acta 199-2018
               'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato Paralelo", sCtaCod, gCodigoCuenta 'LUCV20181220 Coment?, Anexo01 de Acta 199-2018
               Call oNCOMColocEval.insEstadosExpediente(sCtaCod, "Evaluacion de Credito", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlEvaluacion)
               Set oNCOMColocEval = Nothing
            End If
            'RECO FIN **************************************************************************
            'objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato Paralelo", sCtaCod, gCodigoCuenta 'RECO20161020 ERS060-2016
            objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 9 - Paralelo", sCtaCod, gCodigoCuenta 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
            Set objPista = Nothing 'LUCV20181220 Agreg?, Anexo01 de Acta 199-2018
            
            'CTI320200110 ERS003-2020. Coment?
'            'FondoCrecerBitacora
'            Dim objFCBS_UP As COMDCredito.DCOMCredito
'            Set objFCBS_UP = New COMDCredito.DCOMCredito
'            objFCBS_UP.FondoCrecerBitacora IIf(fnTipoRegMant = 1, gCredRegistrarEvaluacionCred, gCredMantenimientoEvaluacionCred), lcMovNro, gsCodPersUser, sCtaCod, "Formato de Evaluaci?n (Sicmac Negocio)"
'            Set objFCBS_UP = Nothing
'            'FondoCrecerBitacora
            'Fin CTI320200110
            
            MsgBox "Los Datos se Grabaron Correctamente", vbInformation, "Aviso"
        Else
            MsgBox "Hubo error al grabar la informacion", vbError, "Error"
        End If
        
            'If lnColocCondi <> 4 Then
                cmdInfromeVista.Enabled = True
            'End If
            
        cmdActualizar.Visible = False
        cmdGuardar.Enabled = False
        
        If (nEstado = 2001) Then
            cmdImprimir.Enabled = True
        End If
        
        'CTI320200110 ERS003-2020. Agreg?
        'Actualizacion de los Ratios
        txtCapacidadNeta.Text = CStr(rsRatiosActual!nCapPagNeta * 100) & "%"

        'Ratios: Aceptable / Critico ->*****
        If Not (rsRatiosAceptableCritico.EOF Or rsRatiosAceptableCritico.BOF) Then
            If rsRatiosAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                Me.lblCapaAceptable.Caption = "Aceptable"
                Me.lblCapaAceptable.ForeColor = &H8000&
            Else
                Me.lblCapaAceptable.Caption = "Cr?tico"
                Me.lblCapaAceptable.ForeColor = vbRed
            End If
            
        Else
            lblCapaAceptable.Visible = False
        End If
        'Fin Ratios <-****
        
        Set rsRatiosActual = Nothing
        Set rsRatiosAceptableCritico = Nothing
        'Fin CTI320200110
    End If
End Sub

Public Sub Controles()

'txtFechaExpeCaja.Enabled = False
'txtVentas.Enabled = False
'txtCapPago.Enabled = False
'txtIngNeto.Enabled = False
'spnDatosIncrIngreso.Enabled = False
'txtEstMonOtrosGasto.Enabled = False
'txtEstMonConsFamiliar.Enabled = False
'txtEstMonOtrosIngresos.Enabled = False
'txtResuMargenBrutoCaja.Enabled = False
'
'txtFechaVista.Enabled = False
'txtDestino.Enabled = False
'txtEntornoFamiliar.Enabled = False
'txtGiroUbicacion.Enabled = False
'txtCrediticia.Enabled = False
'txtFormalidadNegocio.Enabled = False
'txtGarantias.Enabled = False

'cmdGuardar.Enabled = False
'cmdActualizar = False

End Sub

Public Function LenarRecordset_Evaluacion() As ADODB.Recordset

    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsEvaluacion As ADODB.Recordset

    Set rsEvaluacion = New ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsEvaluacion = oDCOMFormatosEval.RecuperarDatosCredEvalFPEvaluacion(sCtaCod) ' llenar mi formato evaluacion
    
        If Not (rsEvaluacion.BOF And rsEvaluacion.EOF) Then

            fnMonAprobado = Trim(rsEvaluacion!nMontoCol)
            fnSalActual = Trim(rsEvaluacion!nSaldo)
            fnMonPropuesto = Trim(rsEvaluacion!nMontoPro)

            Set LenarRecordset_Evaluacion = rsEvaluacion
    
        End If

End Function

Public Function Mantenimineto(ByVal pbMantenimiento As Boolean) As Boolean
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsMantenimientoFormatoParalelo As ADODB.Recordset

    Mantenimineto = False

    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsMantenimientoFormatoParalelo = New ADODB.Recordset
    Set rsMantenimientoFormatoParalelo = oDCOMFormatosEval.RecuperarDatosTotalFormatoParalelo(sCtaCod)
    pnMontoOtrasIfisConsumo = 0
    pnMontoOtrasIfisEmpresarial = 0
    'CTI320200110 ERS003-2020. Agreg?:
    'Obtener valores
    Set rsDatGastoNeg = oDCOMFormatosEval.RecuperaDatosCredEvalGastosNeg(sCtaCod)
    Set rsDatIfiGastoNego = oDCOMFormatosEval.RecuperaDatosIfiCuota(sCtaCod, nFormato, gFormatoGastosNego, gCodCuotaIfiGastoNego)
    If Not (rsDatIfiGastoNego.BOF Or rsDatIfiGastoNego.EOF) Then
        For i = 1 To rsDatIfiGastoNego.RecordCount
           pnMontoOtrasIfisEmpresarial = pnMontoOtrasIfisEmpresarial + rsDatIfiGastoNego!nMonto
           rsDatIfiGastoNego.MoveNext
        Next i
        rsDatIfiGastoNego.MoveFirst
    End If
        
    Set rsDatIfiNoSupervisadaGastoNego = oDCOMFormatosEval.RecuperaDatosIfiCuota(sCtaCod, nFormato, gFormatoGastosNego, gCodCuotaIfiNoSupervisadaGastoNego)
    
    Set rsDatGastoFam = oDCOMFormatosEval.RecuperaDatosCredEvalGastosFam(sCtaCod)
    Set rsDatIfiGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(sCtaCod, nFormato, gFormatoGastosFami, gCodCuotaIfiGastoFami)
    If Not (rsDatIfiGastoFami.BOF Or rsDatIfiGastoFami.EOF) Then
        For i = 1 To rsDatIfiGastoFami.RecordCount
           pnMontoOtrasIfisConsumo = pnMontoOtrasIfisConsumo + rsDatIfiGastoFami!nMonto
           rsDatIfiGastoFami.MoveNext
        Next i
        rsDatIfiGastoFami.MoveFirst
    End If
    Set rsDatIfiNoSupervisadaGastoFami = oDCOMFormatosEval.RecuperaDatosIfiCuota(sCtaCod, nFormato, gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami)
    Set rsDatRatios = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
    
    'Asignar Valores
    'Gastos Negocio
    Call LimpiaFlex(feGastosNegocio)
    Do While Not rsDatGastoNeg.EOF
        feGastosNegocio.AdicionaFila
        lnFila = feGastosNegocio.row
        feGastosNegocio.TextMatrix(lnFila, 1) = rsDatGastoNeg!nConsValor
        feGastosNegocio.TextMatrix(lnFila, 2) = rsDatGastoNeg!cConsDescripcion
        feGastosNegocio.TextMatrix(lnFila, 3) = Format(rsDatGastoNeg!nMonto, "#,##0.00")
        
        If fbImprimirVB And rsDatGastoNeg!nConsValor = 9 Then
            feGastosNegocio.TextMatrix(lnFila, 3) = Format(pnMontoOtrasIfisEmpresarial, "#,##0.00")
        End If
        
            Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
                Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego
                    Me.feGastosNegocio.BackColorRow &HC0FFFF, True
                    Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                    Me.feGastosNegocio.ForeColorRow vbBlack, True
                Case Else
                    Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
            End Select
        rsDatGastoNeg.MoveNext
    Loop
    rsDatGastoNeg.Close
    Set rsDatGastoNeg = Nothing
    
    'Gastos Familiares
    Call LimpiaFlex(feGastosFamiliares)
    Do While Not rsDatGastoFam.EOF
        feGastosFamiliares.AdicionaFila
        lnFila = feGastosFamiliares.row
        feGastosFamiliares.TextMatrix(lnFila, 1) = rsDatGastoFam!nConsValor
        feGastosFamiliares.TextMatrix(lnFila, 2) = rsDatGastoFam!cConsDescripcion
        feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsDatGastoFam!nMonto, "#,##0.00")
        If fbImprimirVB And rsDatGastoFam!nConsValor = 5 Then
            feGastosFamiliares.TextMatrix(lnFila, 3) = Format(pnMontoOtrasIfisConsumo, "#,##0.00")
        End If
         Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
            Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami
                Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                Me.feGastosFamiliares.ForeColorRow vbBlack, True
            Case gCodDeudaLCNUGastoFami
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                Me.feGastosFamiliares.ForeColorRow vbBlack, True
            Case Else
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
         End Select
        rsDatGastoFam.MoveNext
    Loop
    rsDatGastoFam.Close
    Set rsDatGastoFam = Nothing
    
    'Carga de rsDatIfiGastoNego (Ifis Gastos Negocio)
    ReDim MatIfiGastoNego(rsDatIfiGastoNego.RecordCount, 4)
    i = 0
    Do While Not rsDatIfiGastoNego.EOF
        MatIfiGastoNego(i, 0) = rsDatIfiGastoNego!nNroCuota
        MatIfiGastoNego(i, 1) = rsDatIfiGastoNego!CDescripcion
        MatIfiGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiGastoNego!nMonto), 0, rsDatIfiGastoNego!nMonto), "#0.00")
        rsDatIfiGastoNego.MoveNext
          i = i + 1
    Loop
    rsDatIfiGastoNego.Close
    Set rsDatIfiGastoNego = Nothing
    
    'Carga de rsDatIfiGastoFami (Ifis Gastos Familiares)
    ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
    j = 0
    Do While Not rsDatIfiGastoFami.EOF
        MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
        MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!CDescripcion
        MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#0.00")
        rsDatIfiGastoFami.MoveNext
    j = j + 1
    Loop
    rsDatIfiGastoFami.Close
    Set rsDatIfiGastoFami = Nothing
    
    '(Carga de rsDatIfiNoSupervisadaGastoNego
    ReDim MatIfiNoSupervisadaGastoNego(rsDatIfiNoSupervisadaGastoNego.RecordCount, 4)
    i = 0
    Do While Not rsDatIfiNoSupervisadaGastoNego.EOF
        MatIfiNoSupervisadaGastoNego(i, 0) = rsDatIfiNoSupervisadaGastoNego!nNroCuota
        MatIfiNoSupervisadaGastoNego(i, 1) = rsDatIfiNoSupervisadaGastoNego!CDescripcion
        MatIfiNoSupervisadaGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiNoSupervisadaGastoNego!nMonto), 0, rsDatIfiNoSupervisadaGastoNego!nMonto), "#0.00")
        rsDatIfiNoSupervisadaGastoNego.MoveNext
          i = i + 1
    Loop
    rsDatIfiNoSupervisadaGastoNego.Close
    Set rsDatIfiNoSupervisadaGastoNego = Nothing
    
    'Carga de rsDatIfiNoSupervisadaGastoFami
    ReDim MatIfiNoSupervisadaGastoFami(rsDatIfiNoSupervisadaGastoFami.RecordCount, 4)
    j = 0
    Do While Not rsDatIfiNoSupervisadaGastoFami.EOF
        MatIfiNoSupervisadaGastoFami(j, 0) = rsDatIfiNoSupervisadaGastoFami!nNroCuota
        MatIfiNoSupervisadaGastoFami(j, 1) = rsDatIfiNoSupervisadaGastoFami!CDescripcion
        MatIfiNoSupervisadaGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiNoSupervisadaGastoFami!nMonto), 0, rsDatIfiNoSupervisadaGastoFami!nMonto), "#0.00")
        rsDatIfiNoSupervisadaGastoFami.MoveNext
    j = j + 1
    Loop
    rsDatIfiNoSupervisadaGastoFami.Close
    Set rsDatIfiNoSupervisadaGastoFami = Nothing
    
    'Ratios: Aceptable / Critico ->*****
    If Not (rsAceptableCritico.EOF Or rsAceptableCritico.BOF) Then
        If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
            Me.lblCapaAceptable.Caption = "Aceptable"
            Me.lblCapaAceptable.ForeColor = &H8000&
        Else
            Me.lblCapaAceptable.Caption = "Cr?tico"
            Me.lblCapaAceptable.ForeColor = vbRed
        End If
    Else
        Me.lblCapaAceptable.Visible = False
    End If
    
    'Ratios e Indicadores
    If CDbl(rsDatRatios!nCapPagNeta) > 0 Then
        txtCapacidadNeta.Text = CStr(rsDatRatios!nCapPagNeta * 100) & "%"
        lblCapacidadPago.Visible = True
        txtCapacidadNeta.Visible = True
        lblCapaAceptable.Visible = True
    Else
        lblCapacidadPago.Visible = False
        txtCapacidadNeta.Visible = False
        lblCapaAceptable.Visible = False
    End If
    'Fin CTI320200110

    If Not (rsMantenimientoFormatoParalelo.BOF And rsMantenimientoFormatoParalelo.EOF) Then
        txtActividad.Text = rsMantenimientoFormatoParalelo!cActividad
        txtNombCliente.Text = rsMantenimientoFormatoParalelo!cPersNombre
        txtFechaExpeCaja.Text = rsMantenimientoFormatoParalelo!dFechaExpeCaja
        txtUltimoEduSBS.Text = Format(rsMantenimientoFormatoParalelo!nUltEndeSBS, "#,##0.00")
        txtNCredito.Text = Format(rsMantenimientoFormatoParalelo!nNCreditos, "0#")
        txtFechaEduSBS.Text = rsMantenimientoFormatoParalelo!dUltEndeuSBS
        txtCampana.Text = rsMantenimientoFormatoParalelo!cCampa?a
        txtExpCredito.Text = Format(rsMantenimientoFormatoParalelo!nExposiCred, "#,##0.00")
        
        txtMonAprobado.Text = Format(rsMantenimientoFormatoParalelo!nMontoApro, "#,##0.00")
        txtSaldoActual.Text = Format(rsMantenimientoFormatoParalelo!nSaldoActual, "#,##0.00")
        txtVentas.Text = Format(rsMantenimientoFormatoParalelo!nVentas, "#,##0.00")
        txtCapPago.Text = Format(rsMantenimientoFormatoParalelo!nCapPago, "#,##0.00")
        txtIngNeto.Text = Format(rsMantenimientoFormatoParalelo!nIngresoNeto, "#,##0.00")
        
        spnDatosIncrIngreso.valor = rsMantenimientoFormatoParalelo!nIncreIngresoDatos
        
        txtEstMonIngreso.Text = Format(rsMantenimientoFormatoParalelo!nIngreEstMontos, "#,##0.00")
        txtIncIngreso.Text = Format(rsMantenimientoFormatoParalelo!nIncreIngresoEstiMontos, "#,##0.00")
        txtMagBruto.Text = Format(rsMantenimientoFormatoParalelo!nMargenBruto, "#,##0.00")
        txtEstMonOtrosGasto.Text = Format(rsMantenimientoFormatoParalelo!nOtrsoGastos, "#,##0.00")
        txtEstMonConsFamiliar.Text = Format(rsMantenimientoFormatoParalelo!nConsuFamili, "#,##0.00")
        txtCutCredVigente.Text = Format(rsMantenimientoFormatoParalelo!nCuotaCredVig, "#,##0.00")
        txtEstMonOtrosIngresos.Text = Format(rsMantenimientoFormatoParalelo!nOtrosIng, "#,##0.00")
        
        txtResuMargenBrutoCaja.Text = Format(rsMantenimientoFormatoParalelo!nMargenBrutoCaja, "#,##0.00")
        txtIngresos.Text = Format(rsMantenimientoFormatoParalelo!nIngreResumen, "#,##0.00")
        txtResumenIncIngresos.Text = Format(rsMantenimientoFormatoParalelo!nIncreIngreResumen, "#,##0.00")
        txtMonParalelo.Text = Format(rsMantenimientoFormatoParalelo!nMontoParalelo, "#,##0.00")
        txtMonPropuesto.Text = Format(rsMantenimientoFormatoParalelo!nMontoPropuesto, "#,##0.00")
        
        txtFechaVista.Text = rsMantenimientoFormatoParalelo!dFecVisita
        txtDestino.Text = rsMantenimientoFormatoParalelo!cDestino
        txtEntornoFamiliar.Text = rsMantenimientoFormatoParalelo!cEntornoFami
        txtGiroUbicacion.Text = rsMantenimientoFormatoParalelo!cGiroUbica
        txtCrediticia.Text = rsMantenimientoFormatoParalelo!cExpeCrediticia
        txtFormalidadNegocio.Text = rsMantenimientoFormatoParalelo!cFormalNegocio
        txtGarantias.Text = rsMantenimientoFormatoParalelo!cColateGarantia
        
        Mantenimineto = True
    End If
    
    cmdGuardar.Visible = pbMantenimiento
    cmdActualizar.Visible = Not pbMantenimiento

End Function

'validar Datos
Public Function ValidarDatosFormatoParalelo() As Boolean

ValidarDatosFormatoParalelo = True


    If txtFechaExpeCaja.Text = "__/__/____" Then
        MsgBox "Ingrese Fecha de Experiencia en la Caja ", vbInformation, "Aviso"
        SSTab1.Tab = 0
        txtFechaExpeCaja.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If txtVentas.Text = 0 Then
        MsgBox "Ingrese Ventas", vbInformation, "Aviso"
        SSTab1.Tab = 0
        txtVentas.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If txtCapPago.Text = 0 Then
        MsgBox "Ingrese Excedente ", vbInformation, "Aviso"
        SSTab1.Tab = 0
        txtCapPago.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If txtIngNeto.Text = 0 Then
        MsgBox "Ingrese Ingreso Neto ", vbInformation, "Aviso"
        SSTab1.Tab = 0
        txtIngNeto.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If val(spnDatosIncrIngreso.valor) = 0 Then
        MsgBox "Ingrese Incremento de Ingreso ", vbInformation, "Aviso"
        SSTab1.Tab = 0
        spnDatosIncrIngreso.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
'    If Trim(txtEstMonOtrosGasto.Text) = 0 Then
'        MsgBox "Falta Ingresar Otros Gastos", vbInformation, "Aviso"
'        SSTab1.Tab = 0
'        txtEstMonOtrosGasto.SetFocus
'        ValidarDatosFormatoParalelo = False
'        Exit Function
'    End If
'    If Trim(txtEstMonConsFamiliar.Text) = 0 Then
'        MsgBox "Falta Ingresar Consumo Familiar", vbInformation, "Aviso"
'        SSTab1.Tab = 0
'        txtEstMonConsFamiliar.SetFocus
'        ValidarDatosFormatoParalelo = False
'        Exit Function
'    End If
'    If Trim(txtEstMonOtrosIngresos.Text) = 0 Then
'        MsgBox "Falta Ingresar Otros Ingresos", vbInformation, "Aviso"
'        SSTab1.Tab = 0
'        txtEstMonOtrosIngresos.SetFocus
'        ValidarDatosFormatoParalelo = False
'        Exit Function
'    End If
    
    If Trim(txtResuMargenBrutoCaja.Text) = 0 Then
        MsgBox "Falta Ingresar Margen Bruto", vbInformation, "Aviso"
        SSTab1.Tab = 0
        txtResuMargenBrutoCaja.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
'If lnColocCondi <> 4 Then
    If txtFechaVista.Text = "__/__/____" Then
        MsgBox "Ingresar Fecha de Vista", vbInformation, "Aviso"
        SSTab1.Tab = 1
        txtFechaVista.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If Trim(txtEntornoFamiliar.Text) = "" Then
        MsgBox "Falta Ingresar Sobre el Entorno Familiar del Cliente o Representante", vbInformation, "Aviso"
        SSTab1.Tab = 1
        txtEntornoFamiliar.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If Trim(txtGiroUbicacion.Text) = "" Then
        MsgBox "Falta Ingresar Sobre el Giro y la Ubicacion del Negocio", vbInformation, "Aviso"
        SSTab1.Tab = 1
        txtGiroUbicacion.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If Trim(txtCrediticia.Text) = "" Then
        MsgBox "Falta Ingresar Sobre la Experiencia Crediticia", vbInformation, "Aviso"
        SSTab1.Tab = 1
        txtCrediticia.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If Trim(txtFormalidadNegocio.Text) = "" Then
        MsgBox "Falta Sobre la Consistencia de la Informacion y la Formalidad del Negocio", vbInformation, "Aviso"
        SSTab1.Tab = 1
        txtFormalidadNegocio.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If Trim(txtGarantias.Text) = "" Then
        MsgBox "Falta Ingresar Sobre los Colaterales o Garantias", vbInformation, "Aviso"
        SSTab1.Tab = 1
        txtGarantias.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
    If Trim(txtDestino.Text) = "" Then
        MsgBox "Falta Ingresar Sobre el Destino y el Impacto del Mismo", vbInformation, "Aviso"
        SSTab1.Tab = 1
        txtDestino.SetFocus
        ValidarDatosFormatoParalelo = False
        Exit Function
    End If
    
'End If

End Function

Private Sub cmdImprimir_Click()
    Call ImprimeFormato
End Sub

Private Sub cmdInfromeVista_Click()

    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
               
    If (rsInfVisita.EOF And rsInfVisita.BOF) Then
        Set oDCOMFormatosEval = Nothing
        MsgBox "No existe datos para este reporte.", vbOKOnly, "Atenci?n"
        Exit Sub
    End If
    Call CargaInformeVisitaPDF(rsInfVisita) 'gCredReportes

End Sub

Private Sub ImprimeFormato()

Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsImformeVisitaFormatoParalelo As ADODB.Recordset
    Dim oDoc  As cPDF
    Dim psCtaCod As String
    Set oDoc = New cPDF
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsImformeVisitaFormatoParalelo = New ADODB.Recordset
    Set rsImformeVisitaFormatoParalelo = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormatoParalelo(sCtaCod)
    Dim a As Integer
    Dim B As Integer
    Dim nFila As Integer
    Dim nFila1 As Integer
    a = 50
    B = 29

    'Creaci?n del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Cr?dito de Maynas S.A."
    oDoc.Subject = "Informe de Visita N? " & sCtaCod
    oDoc.Title = "Informe de Visita N? " & sCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\FormatoParalelo_HojaEvaluacion" & sCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
        
    If Not (rsImformeVisitaFormatoParalelo.BOF Or rsImformeVisitaFormatoParalelo.EOF) Then

    'Tama?o de hoja A4
    oDoc.NewPage A4_Vertical


    '---------- cabecera ---------------
    oDoc.WImage 45, 45, 45, 113, "Logo"
    oDoc.WTextBox 40, 60, 35, 390, UCase(rsImformeVisitaFormatoParalelo!cAgeDescripcion), "F2", 7.5, hLeft

    oDoc.WTextBox 40, 30, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F2", 7.5, hRight
    oDoc.WTextBox 60, 440, 10, 200, "USUARIO: " & Trim(gsCodUser), "F2", 7.5, hLeft
    oDoc.WTextBox 70, 440, 10, 200, "ANALISTA: " & Trim(rsImformeVisitaFormatoParalelo!cUser), "F2", 7.5, hLeft
       
    oDoc.WTextBox 65, 100, 10, 400, "HOJA DE EVALUACION", "F2", 12, hCenter
    oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & sCtaCod, "F2", 7.5, hLeft
    oDoc.WTextBox 90, 440, 10, 300, "MONEDAD: " & IIf(Mid(sCtaCod, 9, 1) = "1", "SOLES", "DOLARES"), "F2", 7.5, hLeft
    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsImformeVisitaFormatoParalelo!cPersCod), "F2", 7.5, hLeft
    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsImformeVisitaFormatoParalelo!cPersNombre), "F2", 7.5, hLeft
    oDoc.WTextBox 100, 440, 10, 200, "DNI: " & Trim(rsImformeVisitaFormatoParalelo!cPersDni) & "   ", "F2", 7.5, hLeft
    oDoc.WTextBox 110, 440, 10, 200, "RUC: " & Trim(IIf(rsImformeVisitaFormatoParalelo!cPersRuc = "-", Space(11), rsImformeVisitaFormatoParalelo!cPersRuc)), "F2", 7.5, hLeft

    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 120, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 130, 55, 1, 160, "Datos Credito Vigente", "F2", 7.5, hjustify
    oDoc.WTextBox 140, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    
    oDoc.WTextBox 150, 55, 1, 190, "Monto Aprobado", "F1", 7.5, hjustify
    oDoc.WTextBox 150, 70, 1, 190, txtMonAprobado.Text, "F1", 7.5, hRight
    oDoc.WTextBox 160, 55, 1, 190, "Saldo Actual", "F1", 7.5, hjustify
    oDoc.WTextBox 160, 70, 1, 190, txtSaldoActual.Text, "F1", 7.5, hRight
    oDoc.WTextBox 170, 55, 1, 190, "Ventas", "F1", 7.5, hjustify
    oDoc.WTextBox 170, 70, 1, 190, txtVentas.Text, "F1", 7.5, hRight
    oDoc.WTextBox 180, 55, 1, 190, "Cap. Pago", "F1", 7.5, hjustify
    'oDoc.WTextBox 180, 70, 1, 190, txtCapPago.Text, "F1", 7.5, hRight 'CTI320200110 ERS003-2020. Coment?
    oDoc.WTextBox 180, 70, 1, 190, txtCapacidadNeta, "F1", 7.5, hRight 'CTI320200110 ERS003-2020. Agreg?
    oDoc.WTextBox 190, 55, 1, 190, "Ingresos Neto Empresarial", "F1", 7.5, hjustify
    oDoc.WTextBox 190, 70, 1, 190, txtIngNeto.Text, "F1", 7.5, hRight
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 220, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 230, 55, 1, 190, "Datos", "F2", 7.5, hjustify
    oDoc.WTextBox 240, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft

    oDoc.WTextBox 250, 55, 1, 190, "Incremento de Ingreso", "F1", 7.5, hjustify
    oDoc.WTextBox 250, 70, 1, 190, spnDatosIncrIngreso.valor & "%", "F1", 7.5, hRight
    
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 270, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 280, 55, 1, 190, "Estimacion Monto", "F2", 7.5, hjustify
    oDoc.WTextBox 290, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    
    oDoc.WTextBox 300, 55, 1, 190, "Ingresos", "F1", 7.5, hjustify
    oDoc.WTextBox 300, 70, 1, 190, txtEstMonIngreso.Text, "F1", 7.5, hRight
    oDoc.WTextBox 310, 55, 1, 190, "% Incremento Ingresos", "F1", 7.5, hjustify
    oDoc.WTextBox 310, 70, 1, 190, txtIncIngreso.Text, "F1", 7.5, hRight
    oDoc.WTextBox 320, 55, 1, 190, "% Margen Bruto", "F1", 7.5, hjustify
    oDoc.WTextBox 320, 70, 1, 190, txtMagBruto.Text, "F1", 7.5, hRight
    oDoc.WTextBox 330, 55, 1, 190, "Otros Gastos", "F1", 7.5, hjustify
    oDoc.WTextBox 330, 70, 1, 190, txtEstMonOtrosGasto.Text, "F1", 7.5, hRight
    oDoc.WTextBox 340, 55, 1, 190, "Consumo Familiar", "F1", 7.5, hjustify
    oDoc.WTextBox 340, 70, 1, 190, txtEstMonConsFamiliar.Text, "F1", 7.5, hRight
    oDoc.WTextBox 350, 55, 1, 190, "Cuota Cred. Vigente", "F1", 7.5, hjustify
    oDoc.WTextBox 350, 70, 1, 190, txtCutCredVigente.Text, "F1", 7.5, hRight
    oDoc.WTextBox 360, 55, 1, 190, "Otros Ingresos", "F1", 7.5, hjustify
    oDoc.WTextBox 360, 70, 1, 190, txtEstMonOtrosIngresos.Text, "F1", 7.5, hRight
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox 380, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    oDoc.WTextBox 390, 55, 1, 190, "Resumen", "F2", 7.5, hjustify
    oDoc.WTextBox 400, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        
    oDoc.WTextBox 410, 55, 1, 190, "Margen Bruto de Caja", "F1", 7.5, hjustify
    oDoc.WTextBox 410, 70, 1, 190, txtResuMargenBrutoCaja.Text, "F1", 7.5, hRight
    oDoc.WTextBox 420, 55, 1, 190, "Ingresos", "F1", 7.5, hjustify
    oDoc.WTextBox 420, 70, 1, 190, txtIngresos.Text, "F1", 7.5, hRight
    oDoc.WTextBox 430, 55, 1, 190, "% Incremento de Ingresos", "F1", 7.5, hjustify
    oDoc.WTextBox 430, 70, 1, 190, txtResumenIncIngresos.Text, "F1", 7.5, hRight
    oDoc.WTextBox 440, 55, 1, 190, "Monto Calculado Paralelo", "F2", 7.5, hjustify
    oDoc.WTextBox 440, 70, 1, 190, txtMonParalelo.Text, "F1", 7.5, hRight
    oDoc.WTextBox 450, 55, 1, 190, "Monto Propuesto", "F2", 7.5, hjustify
    oDoc.WTextBox 450, 70, 1, 190, txtMonPropuesto.Text, "F1", 7.5, hRight
    
    oDoc.PDFClose
    oDoc.Show
    Else
        MsgBox "Los Datos de Hoja de Evaluacion se mostrara despues de GRABAR la Sugerencia", vbInformation, "Aviso"
    End If

End Sub

'CTI320200110 ERS003-2020
Private Sub feGastosFamiliares_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feGastosFamiliares.ColumnasAEditar, "-")
    
    If Me.feGastosFamiliares.row <> 1 Then
        If Editar(pnCol) = "X" Then
            MsgBox "Esta celda no es editable", vbInformation, "Aviso"
            SendKeys "{TAB}", True
            Cancel = False
            Exit Sub
        End If
    End If
End Sub
Private Sub feGastosFamiliares_Click()
    If feGastosFamiliares.col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami _
        Or CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
    If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodDeudaLCNUGastoFami Then
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
        Me.feGastosFamiliares.ForeColorRow vbBlack, True
    Else
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End If
End Sub
Private Sub feGastosFamiliares_EnterCell()
    If feGastosFamiliares.col = 3 Or (feGastosFamiliares.col = 3 And feGastosFamiliares.row = 1) Then
            feGastosFamiliares.AvanceCeldas = Vertical
    Else
            feGastosFamiliares.AvanceCeldas = Horizontal
    End If
        
    If feGastosFamiliares.col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami _
            Or (CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami) Then
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
        
    If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodDeudaLCNUGastoFami Then
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
    Else
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End If
End Sub
Private Sub feGastosFamiliares_RowColChange()
    If feGastosFamiliares.col = 3 Or (feGastosFamiliares.col = 3 And feGastosFamiliares.row = 1) Then
        feGastosFamiliares.AvanceCeldas = Vertical
    Else
        feGastosFamiliares.AvanceCeldas = Horizontal
    End If
    
    If feGastosFamiliares.col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiGastoFami _
        Or (CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiNoSupervisadaGastoFami) Then 'CTI320200110 ERS003-2020, Agreg?
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
        feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodDeudaLCNUGastoFami Then
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
       
    Else
        Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End If
End Sub
Private Sub feGastosFamiliares_OnClickTxtBuscar(psMontoIfiGastoFami As String, psDescripcion As String)
    psMontoIfiGastoFami = 0
    psDescripcion = ""
    psDescripcion = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoFami = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3) 'Monto
    
     If CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then 'CTI320200110 ERS003-2020. Agreg?
        If psMontoIfiGastoFami = 0 Then
            fnTotalRefGastoFami = 0
            Set MatIfiGastoFami = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        End If
    Else
        If psMontoIfiGastoFami = 0 Then
            fnTotalRefGastoFami = 0
            Set MatIfiNoSupervisadaGastoFami = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiNoSupervisadaGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), _
                                             gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agreg?
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiNoSupervisadaGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), _
                                            gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agreg?
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        End If
    End If
End Sub

Private Sub feGastosFamiliares_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feGastosFamiliares.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feGastosFamiliares.TextMatrix(pnRow, pnCol) < 0 Then
            feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feGastosFamiliares.TextMatrix(pnRow, pnCol) = 0
    End If

    If (Me.feGastosFamiliares.col = 3 And Me.feGastosFamiliares.row = 4) Then
        SSTab1.Tab = 2
        SendKeys "{TAB}"
    End If
End Sub

Private Sub feGastosNegocio_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feGastosNegocio.ColumnasAEditar, "-")
    If Editar(pnCol) = "X" Then
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        Cancel = False
        SendKeys "{TAB}", True
        Exit Sub
    End If
End Sub
Private Sub feGastosNegocio_Click()
    If feGastosNegocio.col = 3 Then
        If CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1)) = gCodCuotaIfiGastoNego _
            Or (CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1)) = gCodCuotaIfiNoSupervisadaGastoNego) Then 'CTI320200110 ERS003-2020, Agreg?
            feGastosNegocio.ListaControles = "0-0-0-1-0"
        Else
            feGastosNegocio.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
        Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego
            Me.feGastosNegocio.BackColorRow &HC0FFFF, True
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
            Me.feGastosNegocio.ForeColorRow vbBlack, True
        Case Else
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub

Private Sub feGastosNegocio_EnterCell()
    If feGastosNegocio.col = 3 Or (feGastosNegocio.col = 3 And feGastosNegocio.row = 1) Then
        feGastosNegocio.AvanceCeldas = Vertical
    Else
        feGastosNegocio.AvanceCeldas = Horizontal
    End If
    
    If feGastosNegocio.col = 3 Then
        If (CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1)) = gCodCuotaIfiGastoNego) _
        Or (CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1)) = gCodCuotaIfiNoSupervisadaGastoNego) Then 'CTI320200110 ERS003-2020, Agreg?: gCodCuotaIfiNoSupervisadaGastoNego
            feGastosNegocio.ListaControles = "0-0-0-1-0"
        Else
            feGastosNegocio.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
        Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego
            Me.feGastosNegocio.BackColorRow &HC0FFFF, True
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
            Me.feGastosNegocio.ForeColorRow vbBlack, True
        Case Else
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub

Private Sub feGastosNegocio_RowColChange()
    If feGastosNegocio.col = 3 Or (feGastosNegocio.col = 3 And feGastosNegocio.row = 1) Then
        feGastosNegocio.AvanceCeldas = Vertical
    Else
        feGastosNegocio.AvanceCeldas = Horizontal
    End If
    
    If feGastosNegocio.col = 3 Then
        If CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1)) = gCodCuotaIfiGastoNego _
        Or (CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1)) = gCodCuotaIfiNoSupervisadaGastoNego) Then 'CTI320200110 ERS003-2020, Agreg?: gCodCuotaIfiNoSupervisadaGastoNego
        feGastosNegocio.ListaControles = "0-0-0-1-0"
        Else
        feGastosNegocio.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
        Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego
            Me.feGastosNegocio.BackColorRow &HC0FFFF, True
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
            Me.feGastosNegocio.ForeColorRow vbBlack, True
        Case Else
            Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub
Private Sub feGastosNegocio_OnClickTxtBuscar(psMontoIfiGastoNego As String, psDescripcion As String) 'GastosNegocio
    psDescripcion = ""
    psDescripcion = feGastosNegocio.TextMatrix(feGastosNegocio.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoNego = 0
    psMontoIfiGastoNego = feGastosNegocio.TextMatrix(feGastosNegocio.row, 3) 'Monto
    
    If feGastosNegocio.TextMatrix(feGastosNegocio.row, 1) = gCodCuotaIfiGastoNego Then 'CTI320200110 ERS003-2020. Agreg?
        If psMontoIfiGastoNego = 0 Then
            fnTotalRefGastoNego = 0
            Set MatIfiGastoNego = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2), gFormatoGastosNego, gCodCuotaIfiGastoNego ', sCtaCod
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2), gFormatoGastosNego, gCodCuotaIfiGastoNego ', sCtaCod
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        End If
    Else
        If psMontoIfiGastoNego = 0 Then
            fnTotalRefGastoNego = 0
            Set MatIfiNoSupervisadaGastoNego = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiNoSupervisadaGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2), _
                                              gFormatoGastosNego, gCodCuotaIfiNoSupervisadaGastoNego 'CTI320200110 ERS003-2020. Agreg?
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiNoSupervisadaGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2), _
                                              gFormatoGastosNego, gCodCuotaIfiNoSupervisadaGastoNego 'CTI320200110 ERS003-2020. Agreg?
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        End If
    End If
    
    feGastosNegocio.TopRow = 1
    feGastosNegocio.row = 1
End Sub

Private Sub feGastosNegocio_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feGastosNegocio.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feGastosNegocio.TextMatrix(pnRow, pnCol) < 0 Then
            feGastosNegocio.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feGastosNegocio.TextMatrix(pnRow, pnCol) = 0
    End If

    If Me.feGastosNegocio.col = 3 And Me.feGastosNegocio.row = 12 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Me.feGastosFamiliares.SetFocus
        feGastosFamiliares.row = 1
        feGastosFamiliares.col = 3
        SendKeys "{TAB}"
        SendKeys "{F2}"
    End If
End Sub

'Fin CTI320200110
Private Sub txtDestino_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtGiroUbicacion
    End If
End Sub

Private Sub txtFechaVista_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtEntornoFamiliar
    End If
End Sub

Private Sub txtEntornoFamiliar_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtGiroUbicacion
    End If
End Sub

Private Sub txtGiroUbicacion_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtCrediticia
    End If
End Sub

Private Sub txtCrediticia_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtFormalidadNegocio
        End If
End Sub

Private Sub txtFormalidadNegocio_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtGarantias
    End If
End Sub

Private Sub txtGarantias_KeyPress(KeyAscii As Integer)
KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl txtDestino
    End If
End Sub

Public Sub Form_Load()

fbGrabar = False
'cmdInfromeVista.Enabled = False
cmdActualizar.Visible = False

SSTab1.Tab = 0

CentraForm Me

End Sub

Private Sub Consultar()
    txtMonAprobado.Enabled = False
    txtSaldoActual.Enabled = False
    txtVentas.Enabled = False
    txtCapPago.Enabled = False
    txtIngNeto.Enabled = False

    spnDatosIncrIngreso.Enabled = False

    txtEstMonIngreso.Enabled = False
    txtIncIngreso.Enabled = False
    txtMagBruto.Enabled = False
    txtEstMonOtrosGasto.Enabled = False
    txtEstMonConsFamiliar.Enabled = False
    txtCutCredVigente.Enabled = False
    txtEstMonOtrosIngresos.Enabled = False

    txtResuMargenBrutoCaja.Enabled = False
    txtIngresos.Enabled = False
    txtResumenIncIngresos.Enabled = False
    txtMonParalelo.Enabled = False
    txtMonPropuesto.Enabled = False

    txtFechaVista.Enabled = False
    txtEntornoFamiliar.Enabled = False
    txtGiroUbicacion.Enabled = False
    txtCrediticia.Enabled = False
    txtFormalidadNegocio.Enabled = False
    txtGarantias.Enabled = False
    txtDestino.Enabled = False
    
    cmdGuardar.Enabled = False
    cmdActualizar.Enabled = False
End Sub

Private Sub LimpiaFormulario()
  
    txtMonAprobado.Text = ""
    txtSaldoActual.Text = ""
    txtVentas.Text = ""
    txtCapPago.Text = ""
    txtIngNeto.Text = ""
        
    spnDatosIncrIngreso.valor = 0
        
    txtEstMonIngreso.Text = ""
    txtIncIngreso.Text = ""
    txtMagBruto.Text = ""
    txtEstMonOtrosGasto.Text = ""
    txtEstMonConsFamiliar.Text = ""
    txtCutCredVigente.Text = ""
    txtEstMonOtrosIngresos.Text = ""
        
    txtResuMargenBrutoCaja.Text = ""
    txtIngresos.Text = ""
    txtResumenIncIngresos.Text = ""
    txtMonParalelo.Text = ""
    txtMonPropuesto.Text = ""
        
    txtFechaVista.Text = "__/__/____"
    txtDestino.Text = ""
End Sub

Private Sub LLenarFormulario()
           
    txtMonAprobado.Text = "0.00"
    txtSaldoActual.Text = "0.00"
    
    txtVentas.Text = "0.00"
    txtCapPago.Text = "0.00"
    txtIngNeto.Text = "0.00"
    
    spnDatosIncrIngreso.valor = "00"
        
    txtMagBruto.Text = "0.00"
    txtIncIngreso.Text = "0.00"
    txtEstMonOtrosGasto.Text = "0.00"
    txtEstMonConsFamiliar.Text = "0.00"
    txtEstMonOtrosIngresos.Text = "0.00"
    
    txtCutCredVigente.Text = "0.00"
        
    txtResuMargenBrutoCaja.Text = "0.00"
    txtResumenIncIngresos.Text = "0.00"
    txtMonParalelo.Text = "0.00"
    
    txtMonPropuesto.Text = "0.00"
End Sub

Private Sub CalculoTotal(ByVal pnTipo As Integer)

Dim nValorMagBruto As Currency

On Error GoTo ErrorCalculo
    
    Select Case pnTipo
    Case 1:
            If txtIngresos.Text <> 0 Then
                nValorMagBruto1 = CDbl(txtResuMargenBrutoCaja.Text) / CDbl(txtIngresos.Text)
                nValorMagBruto2 = nValorMagBruto1
                txtMagBruto.Text = Format(nValorMagBruto2, "#,##0.00")
            End If
    Case 2:
            txtMonParalelo.Text = Format(((CDbl(txtEstMonIngreso.Text) * CDbl(txtIncIngreso.Text) * CDbl(nValorMagBruto2)) - CDbl(txtEstMonOtrosGasto.Text) - CDbl(txtEstMonConsFamiliar.Text) - CDbl(txtCutCredVigente.Text) + CDbl(txtEstMonOtrosIngresos.Text)), "#,##0.00")
    Case 3:
            txtIncIngreso.Text = Format(1 + CDbl(spnDatosIncrIngreso.valor) / 100, "#,##0.00")
    Case 4:
            If CCur(txtMonPropuesto.Text) > CCur(txtMonParalelo.Text) Then
                    MsgBox "El Monto Propuesto es Mayor Al Monto Calculado ", vbInformation, "Aviso"
                    
                    cmdGuardar.Enabled = False
                    cmdActualizar.Enabled = False
            Else
            
            End If
    End Select
    Exit Sub
    
ErrorCalculo:
'MsgBox "Error: Ingrese los datos Correctamente." & Chr(13) & "Detalles de error: " & Err.Description, vbCritical, "Error"

'Select Case pnTipo
'    Case 1:
'            txtResuMargenBrutoCaja.Text = "0.00"
'            txtIngresos.Text = "0.00"
'    Case 2:
'            txtIngresos.Text = "0.00"
'            txtResumenIncIngresos.Text = "0.00"
'            txtMagBruto.Text = "0.00"
'            txtEstMonOtrosGasto.Text = "0.00"
'            txtEstMonConsFamiliar.Text = "0.00"
'            txtCutCredVigente.Text = "0.00"
'            txtEstMonOtrosIngresos.Text = "0.00"
'    Case 3:
'            spnDatosIncrIngreso.valor = 0
'
'End Select
'Call CalculoTotal(pnTipo)
     
End Sub

Private Sub Form_Activate()
   'txtVentas.SetFocus
End Sub

Private Sub txtFechaVista_LostFocus()

If Not IsDate(txtFechaVista) Then
    MsgBox "Verifique Dia,Mes,A?o , Fecha Incorrecta", vbInformation, "Aviso"
End If
    EnfocaControl txtEntornoFamiliar
End Sub

'*********************************************************************************************************
'Datos Credito Vigente
Private Sub txtVentas_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtVentas, KeyAscii, 10, , True)
                If KeyAscii = 13 Then
                        If txtVentas.Text = "" Then
                            txtVentas.Text = "0.00"
                        End If
                        If IsNumeric(txtVentas.Text) Then
                        Else
                                txtVentas.Text = "0.00"
                        End If
                        If CCur(txtVentas.Text) <= 0 Then
                            txtVentas.Text = "0.00"
                        End If
                    EnfocaControl txtCapPago
                    fEnfoque txtCapPago
                End If
                    
End Sub

Private Sub txtVentas_GotFocus()
''Me.txtFechaExpeCaja.SelStart = 0
Me.txtVentas.SelLength = Len(txtVentas.Text)
End Sub

Private Sub txtVentas_LostFocus()
            If CCur(txtVentas.Text) > 0 Then
                    txtVentas.Text = Format(txtVentas.Text, "#,##0.00")
                    txtEstMonIngreso.Text = txtVentas.Text
                    txtIngresos.Text = txtVentas.Text
                    Call CalculoTotal(2)
            Else
                txtVentas.Text = "0.00"
            End If
End Sub

Private Sub txtCapPago_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCapPago, KeyAscii, 10, , True)
    If KeyAscii = 13 Then
        If txtCapPago.Text = "" Then
            txtCapPago.Text = "0.00"
        End If
        If IsNumeric(txtCapPago.Text) Then
            Else
                txtCapPago.Text = "0.00"
        End If
        If CCur(txtCapPago.Text) <= 0 Then
            txtCapPago.Text = "0.00"
        End If
        EnfocaControl txtIngNeto
        fEnfoque txtIngNeto
    End If
        
End Sub

Private Sub txtCapPago_LostFocus()
        If CCur(txtCapPago.Text) > 0 Then
            txtCapPago.Text = Format(txtCapPago.Text, "#,##0.00")
            Call CalculoTotal(2)
        Else
            txtCapPago.Text = "0.00"
        End If
End Sub

Private Sub txtIngNeto_KeyPress(KeyAscii As Integer)

    KeyAscii = NumerosDecimales(txtIngNeto, KeyAscii, 10, , True)
    
        If KeyAscii = 13 Then
            If txtIngNeto.Text = "" Then
                txtIngNeto.Text = "0.00"
            End If
            If IsNumeric(txtIngNeto.Text) Then
                Else
                    txtIngNeto.Text = "0.00"
            End If
            If CCur(txtIngNeto.Text) <= 0 Then
                txtIngNeto.Text = "0.00"
            End If
        
        EnfocaControl spnDatosIncrIngreso
        spnDatosIncrIngreso.SetFocus
        End If
        
        
End Sub

Private Sub txtIngNeto_LostFocus()

        If CCur(txtIngNeto.Text) > 0 Then
                
                txtIngNeto.Text = Format(txtIngNeto.Text, "#,##0.00")
                Call CalculoTotal(2)
        Else
            txtIngNeto.Text = "0.00"
        End If
End Sub

'Datos
Private Sub spnDatosIncrIngreso_LostFocus()

            If CInt(spnDatosIncrIngreso.valor) > 0 Then
                
                txtEstMonOtrosGasto.SetFocus
                txtResumenIncIngresos.Text = Format(CDbl(spnDatosIncrIngreso.valor), "#,#0.00")
                
            Else
                spnDatosIncrIngreso.valor = 0
            End If
        Call CalculoTotal(3)
End Sub

Private Sub spnDatosIncrIngreso_KeyPress(KeyAscii As Integer)

          If KeyAscii = 13 Then
            If spnDatosIncrIngreso.valor = "" Then
                spnDatosIncrIngreso.valor = "00"
            End If
            If IsNumeric(spnDatosIncrIngreso.valor) Then
                Else
                    spnDatosIncrIngreso.valor = "00"
            End If
            If CInt(spnDatosIncrIngreso.valor) <= 0 Then
                spnDatosIncrIngreso.valor = "00"
            End If
            If val(spnDatosIncrIngreso.valor) > 100 Then
                MsgBox "El valor no Puede ser Mayor de 100", vbInformation, "Aviso"
                spnDatosIncrIngreso.valor = "00"
            End If
            EnfocaControl txtEstMonOtrosGasto
            fEnfoque txtEstMonOtrosGasto
          End If
            
End Sub

'Estimacion Monto
Private Sub txtEstMonOtrosGasto_KeyPress(KeyAscii As Integer)

                KeyAscii = NumerosDecimales(txtEstMonOtrosGasto, KeyAscii, 10, , True) 'FRHU 20150611
                    If KeyAscii = 13 Then
                        If txtEstMonOtrosGasto.Text = "" Then
                            txtEstMonOtrosGasto.Text = "0.00"
                        End If
                        If IsNumeric(txtEstMonOtrosGasto.Text) Then
                            Else
                                txtEstMonOtrosGasto.Text = "0.00"
                        End If
                        If CCur(txtEstMonOtrosGasto.Text) <= 0 Then
                            txtEstMonOtrosGasto.Text = "0.00"
                        End If
                            EnfocaControl txtEstMonConsFamiliar
                            fEnfoque txtEstMonConsFamiliar
                    End If
                            
                            
End Sub

Private Sub txtEstMonOtrosGasto_LostFocus()
 If CCur(txtEstMonOtrosGasto.Text) > 0 Then
    txtEstMonOtrosGasto.Text = Format(txtEstMonOtrosGasto.Text, "#,##0.00")
    Call CalculoTotal(2)
End If
End Sub

Private Sub txtEstMonConsFamiliar_KeyPress(KeyAscii As Integer)

    KeyAscii = NumerosDecimales(txtEstMonConsFamiliar, KeyAscii, 10, , True)
        If KeyAscii = 13 Then
            If txtEstMonConsFamiliar.Text = "" Then
                txtEstMonConsFamiliar.Text = "0.00"
            End If
            If IsNumeric(txtEstMonConsFamiliar.Text) Then
                Else
                    txtEstMonConsFamiliar.Text = "0.00"
            End If
            If CCur(txtEstMonConsFamiliar.Text) <= 0 Then
                txtEstMonConsFamiliar.Text = "0.00"
            End If
            EnfocaControl txtEstMonOtrosIngresos
            fEnfoque txtEstMonOtrosIngresos
        End If
            
End Sub

Private Sub txtEstMonConsFamiliar_LostFocus()

If CCur(txtEstMonConsFamiliar.Text) > 0 Then
        txtEstMonConsFamiliar.Text = Format(txtEstMonConsFamiliar.Text, "#,##0.00")
        Call CalculoTotal(2)
Else
    txtEstMonConsFamiliar.Text = "0.00"
End If
End Sub

Private Sub txtEstMonOtrosIngresos_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtEstMonOtrosIngresos, KeyAscii, 10, , True)
        If KeyAscii = 13 Then
            If txtEstMonOtrosIngresos.Text = "" Then
                txtEstMonOtrosIngresos.Text = "0.00"
            End If
            If IsNumeric(txtEstMonOtrosIngresos.Text) Then
                Else
                    txtEstMonOtrosIngresos.Text = "0.00"
            End If
            If CCur(txtEstMonOtrosIngresos.Text) <= 0 Then
                txtEstMonOtrosIngresos.Text = "0.00"
            End If
            EnfocaControl txtResuMargenBrutoCaja
            fEnfoque txtResuMargenBrutoCaja
        End If
            
End Sub

Private Sub txtEstMonOtrosIngresos_LostFocus()

        If CCur(txtEstMonOtrosIngresos.Text) > 0 Then
             txtEstMonOtrosIngresos.Text = Format(txtEstMonOtrosIngresos.Text, "#,##0.00")
                Call CalculoTotal(2)
        Else
            txtEstMonOtrosIngresos.Text = "0.00"
        End If
        
End Sub

'Resumen
Private Sub txtResuMargenBrutoCaja_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtResuMargenBrutoCaja, KeyAscii, 10, , True)
        If KeyAscii = 13 Then
            If txtResuMargenBrutoCaja.Text = "" Then
                txtResuMargenBrutoCaja.Text = "0.00"
            End If
            If IsNumeric(txtResuMargenBrutoCaja.Text) Then
                Else
                    txtResuMargenBrutoCaja.Text = "0.00"
            End If
            If CCur(txtResuMargenBrutoCaja.Text) <= 0 Then
                txtResuMargenBrutoCaja.Text = "0.00"
            End If
            
                'SendKeys "{Tab}", True
                SSTab1.Tab = 1
                EnfocaControl txtFechaVista
        End If
End Sub

Private Sub txtResuMargenBrutoCaja_LostFocus()
            If CCur(txtResuMargenBrutoCaja.Text) > 0 Then
                txtResuMargenBrutoCaja.Text = Format(txtResuMargenBrutoCaja.Text, "###," & String(15, "#") & "#,##0.00")
                Call CalculoTotal(1)
                Call CalculoTotal(2)
                Call CalculoTotal(4)
            Else
                txtResuMargenBrutoCaja.Text = "0.00"
            End If
End Sub
'*****************************************************************

Private Function CargaControlesTipoPermiso(ByVal TipoPermiso As Integer) As Boolean
    '1: JefeAgencia->
    If TipoPermiso = 1 Then
        Call HabilitaControles(False)
        CargaControlesTipoPermiso = True
     '2: Coordinador->
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False)
        CargaControlesTipoPermiso = True
     '3: Analista ->
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True)
        CargaControlesTipoPermiso = True
     'Usuario sin Permisos al formato
    Else
        MsgBox "No tiene Permisos para este m?dulo", vbInformation, "Aviso"
        Call HabilitaControles(False)
        CargaControlesTipoPermiso = False
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean)
    frDatosCredVig.Enabled = pbHabilitaA
    frDatos.Enabled = pbHabilitaA
    frEstimacionMonto.Enabled = pbHabilitaA
    frResumen.Enabled = pbHabilitaA
    frPropuesta.Enabled = pbHabilitaA
    cmdGuardar.Enabled = pbHabilitaA
    cmdActualizar.Enabled = pbHabilitaA
End Function
'CTI320200110 ERS003-2020. Agreg?
Private Sub CargarFlexEdit() 'Registrar New Formato Evaluacion
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim nMonto As Double
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    nMonto = Format(0, "00.00")
    
    CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(nFormato, sCtaCod, rsFeGastoNeg, rsFeDatGastoFam)
                                                            
    'Gastos Negocio
    feGastosNegocio.Clear
    feGastosNegocio.FormaCabecera
    feGastosNegocio.rows = 2
    Call LimpiaFlex(feGastosNegocio)
        Do While Not rsFeGastoNeg.EOF
            feGastosNegocio.AdicionaFila
            lnFila = feGastosNegocio.row
            feGastosNegocio.TextMatrix(lnFila, 1) = rsFeGastoNeg!nConsValor
            feGastosNegocio.TextMatrix(lnFila, 2) = rsFeGastoNeg!cConsDescripcion
            feGastosNegocio.TextMatrix(lnFila, 3) = Format(rsFeGastoNeg!nMonto, "#,##0.00")

                Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
                    Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego
                        Me.feGastosNegocio.BackColorRow &HC0FFFF, True
                        Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                        Me.feGastosNegocio.ForeColorRow vbBlack, True
                    Case gCodCuotaCmac
                        Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
                        Me.feGastosNegocio.ForeColorRow vbBlack, True
                    Case Else
                        Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                End Select
            rsFeGastoNeg.MoveNext
        Loop
    rsFeGastoNeg.Close
    Set rsFeGastoNeg = Nothing
    
    'Gastos Familiares
    feGastosFamiliares.Clear
    feGastosFamiliares.FormaCabecera
    feGastosFamiliares.rows = 2
    Call LimpiaFlex(feGastosFamiliares)
        Do While Not rsFeDatGastoFam.EOF
            feGastosFamiliares.AdicionaFila
            lnFila = feGastosFamiliares.row
            feGastosFamiliares.TextMatrix(lnFila, 1) = rsFeDatGastoFam!nConsValor
            feGastosFamiliares.TextMatrix(lnFila, 2) = rsFeDatGastoFam!cConsDescripcion
            feGastosFamiliares.TextMatrix(lnFila, 3) = Format(rsFeDatGastoFam!nMonto, "#,##0.00")
                
            Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
                Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami
                   Me.feGastosFamiliares.BackColorRow &HC0FFFF
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                   Me.feGastosFamiliares.ForeColorRow vbBlack, True
                Case gCodDeudaLCNUGastoFami, gCodCuotaCmac
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                   Me.feGastosFamiliares.ForeColorRow vbBlack, True
                Case Else
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
             End Select
            rsFeDatGastoFam.MoveNext
        Loop
    rsFeDatGastoFam.Close
    Set rsFeDatGastoFam = Nothing
End Sub
'Fin CTI320200110

