VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmCredFormEvalFormato3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Créditos - Evaluación - Formato 3"
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredFormEvalFormato3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdMNME 
      Caption         =   "MN - ME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   9540
      TabIndex        =   146
      Top             =   1800
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Hoja Evaluación"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   9540
      TabIndex        =   36
      Top             =   450
      Width           =   1740
   End
   Begin VB.CommandButton cmdVerCar 
      Caption         =   "&Ver CAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   9540
      TabIndex        =   35
      Top             =   1140
      Width           =   1740
   End
   Begin VB.CommandButton cmdInformeVisita 
      Caption         =   "Infor&me de Visita"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   9540
      TabIndex        =   34
      Top             =   800
      Width           =   1740
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   9560
      TabIndex        =   32
      Top             =   40
      Width           =   1720
   End
   Begin VB.CommandButton cmdGenerarFlujoForm3 
      Caption         =   "Generar &Flujo Caja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   9540
      TabIndex        =   38
      Top             =   1470
      Width           =   1740
   End
   Begin TabDlg.SSTab SSTabInfoNego 
      Height          =   2145
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   9520
      _ExtentX        =   16801
      _ExtentY        =   3784
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Información del Negocio"
      TabPicture(0)   =   "frmCredFormEvalFormato3.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtFechaEvaluacion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "ActXCodCta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtGiroNeg"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.Frame Frame1 
         Height          =   1430
         Left            =   120
         TabIndex        =   39
         Top             =   680
         Width           =   9350
         Begin VB.TextBox txtNombreCliente 
            Height          =   300
            Left            =   1750
            TabIndex        =   3
            Top             =   120
            Width           =   3915
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Propia"
            Height          =   255
            Index           =   1
            Left            =   1800
            TabIndex        =   10
            Top             =   1100
            Width           =   855
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Alquilada"
            Height          =   255
            Index           =   2
            Left            =   2640
            TabIndex        =   11
            Top             =   1100
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Ambulante"
            Height          =   255
            Index           =   3
            Left            =   3720
            TabIndex        =   12
            Top             =   1100
            Width           =   1095
         End
         Begin VB.OptionButton OptCondLocal 
            Caption         =   "Otros"
            Height          =   255
            Index           =   4
            Left            =   4920
            TabIndex        =   13
            Top             =   1100
            Width           =   855
         End
         Begin VB.TextBox txtCondLocalOtros 
            Height          =   285
            Left            =   5760
            MaxLength       =   250
            TabIndex        =   14
            Top             =   1100
            Visible         =   0   'False
            Width           =   2955
         End
         Begin MSMask.MaskEdBox txtFecUltEndeuda 
            Height          =   300
            Left            =   8050
            TabIndex        =   7
            Top             =   120
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   8421504
            Enabled         =   0   'False
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin Spinner.uSpinner spnTiempoLocalAnio 
            Height          =   315
            Left            =   1750
            TabIndex        =   0
            Top             =   780
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   99
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin Spinner.uSpinner spnTiempoLocalMes 
            Height          =   315
            Left            =   3000
            TabIndex        =   8
            Top             =   780
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   12
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
         End
         Begin SICMACT.EditMoney txtExposicionCredito 
            Height          =   300
            Left            =   8040
            TabIndex        =   9
            Top             =   765
            Width           =   1220
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin Spinner.uSpinner spnExpEmpAnio 
            Height          =   315
            Left            =   1750
            TabIndex        =   5
            Top             =   450
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   99
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin Spinner.uSpinner spnExpEmpMes 
            Height          =   315
            Left            =   3000
            TabIndex        =   6
            Top             =   450
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            Max             =   12
            MaxLength       =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontName        =   "MS Sans Serif"
            FontSize        =   8.25
            ForeColor       =   8421504
         End
         Begin SICMACT.EditMoney txtUltEndeuda 
            Height          =   300
            Left            =   8040
            TabIndex        =   4
            Top             =   450
            Width           =   1220
            _ExtentX        =   2143
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
            Height          =   195
            Left            =   1200
            TabIndex        =   50
            Top             =   150
            Width           =   555
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exp. como empresario :"
            Height          =   195
            Left            =   80
            TabIndex        =   49
            Top             =   465
            Width           =   1695
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo en el local :"
            Height          =   195
            Left            =   400
            TabIndex        =   48
            Top             =   795
            Width           =   1365
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Condición local :"
            Height          =   255
            Left            =   610
            TabIndex        =   47
            Top             =   1095
            Width           =   1215
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exposición con este crédito:"
            Height          =   195
            Left            =   6000
            TabIndex        =   46
            Top             =   795
            Width           =   2010
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "años"
            Height          =   255
            Left            =   2570
            TabIndex        =   45
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "años"
            Height          =   255
            Left            =   2570
            TabIndex        =   44
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   3795
            TabIndex        =   43
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "meses"
            Height          =   255
            Left            =   3795
            TabIndex        =   42
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ult. endeudamiento RCC:"
            Height          =   195
            Left            =   6200
            TabIndex        =   41
            Top             =   510
            Width           =   1830
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Ult. endeudamiento RCC:"
            Height          =   195
            Left            =   5720
            TabIndex        =   40
            Top             =   150
            Width           =   2310
         End
      End
      Begin VB.TextBox txtGiroNeg 
         Height          =   300
         Left            =   5400
         TabIndex        =   2
         Top             =   360
         Width           =   4035
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Crédito"
      End
      Begin MSMask.MaskEdBox txtFechaEvaluacion 
         Height          =   300
         Left            =   8175
         TabIndex        =   124
         Top             =   0
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de evaluación al :"
         Height          =   195
         Left            =   6360
         TabIndex        =   125
         Top             =   45
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Giro del Negocio :"
         Height          =   255
         Left            =   4080
         TabIndex        =   51
         Top             =   390
         Width           =   1335
      End
   End
   Begin TabDlg.SSTab SSTabIngresos 
      Height          =   6960
      Left            =   0
      TabIndex        =   52
      Top             =   2160
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   12277
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ingresos y Egresos"
      TabPicture(0)   =   "frmCredFormEvalFormato3.frx":0326
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraBalanceGeneral3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Propuesta del Crédito"
      TabPicture(1)   =   "frmCredFormEvalFormato3.frx":0342
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "framePropuesta"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Comentarios y Referidos"
      TabPicture(2)   =   "frmCredFormEvalFormato3.frx":035E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frameReferido"
      Tab(2).Control(1)=   "frameComentario"
      Tab(2).ControlCount=   2
      Begin VB.Frame frameReferido 
         Caption         =   "Referidos :"
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
         Height          =   3375
         Left            =   -74880
         TabIndex        =   119
         Top             =   3120
         Width           =   9855
         Begin VB.CommandButton cmdQuitar3 
            Caption         =   "&Quitar"
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
            Left            =   8640
            TabIndex        =   31
            Top             =   2880
            Width           =   1170
         End
         Begin VB.CommandButton cmdAgregarRef3 
            Caption         =   "&Agregar"
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
            Left            =   7440
            TabIndex        =   30
            Top             =   2880
            Width           =   1170
         End
         Begin SICMACT.FlexEdit feReferidos3 
            Height          =   2535
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   4471
            Cols0           =   7
            HighLight       =   1
            EncabezadosNombres=   "N-Nombres-DNI-Teléfono-Comentario-NroDNI-Aux"
            EncabezadosAnchos=   "350-3250-960-1260-3650-0-0"
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
            ColumnasAEditar =   "X-1-2-3-4-X-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-L-L-L-L-C"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "N"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
         End
      End
      Begin VB.Frame frameComentario 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74880
         TabIndex        =   118
         Top             =   360
         Width           =   9855
         Begin VB.TextBox txtComentario3 
            Height          =   2250
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   3000
            MultiLine       =   -1  'True
            TabIndex        =   28
            Top             =   240
            Width           =   9615
         End
      End
      Begin VB.Frame framePropuesta 
         Caption         =   "Propuesta del Credito:"
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
         Height          =   6375
         Left            =   -74760
         TabIndex        =   102
         Top             =   480
         Width           =   10695
         Begin VB.TextBox txtDestino3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   27
            Top             =   5520
            Width           =   10455
         End
         Begin VB.TextBox txtColaterales3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   4560
            Width           =   10455
         End
         Begin VB.TextBox txtFormalidadNegocio3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   3600
            Width           =   10455
         End
         Begin VB.TextBox txtGiroUbicacion3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   1680
            Width           =   10455
         End
         Begin VB.TextBox txtExperiencia3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   2640
            Width           =   10455
         End
         Begin VB.TextBox txtEntornoFamiliar3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   720
            Width           =   10455
         End
         Begin MSMask.MaskEdBox txtFechaVisita3 
            Height          =   300
            Left            =   9360
            TabIndex        =   21
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el destino y el impacto del mismo:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   109
            Top             =   5280
            Width           =   2850
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   240
            TabIndex        =   108
            Top             =   4320
            Width           =   2400
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   107
            Top             =   3360
            Width           =   4770
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la experiencia Crediticia:"
            Height          =   195
            Left            =   240
            TabIndex        =   106
            Top             =   2400
            Width           =   2220
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   105
            Top             =   1440
            Width           =   2820
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   104
            Top             =   480
            Width           =   3795
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Visita:"
            Height          =   195
            Left            =   8160
            TabIndex        =   103
            Top             =   300
            Width           =   1140
         End
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
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
         Left            =   -73200
         TabIndex        =   101
         Top             =   6120
         Width           =   1170
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
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
         Index           =   0
         Left            =   -74640
         TabIndex        =   100
         Top             =   6120
         Width           =   1170
      End
      Begin VB.Frame Frame8 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   98
         Top             =   3360
         Width           =   9975
         Begin SICMACT.FlexEdit FlexEdit1 
            Height          =   1935
            Left            =   120
            TabIndex        =   99
            Top             =   360
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   3413
            Cols0           =   6
            HighLight       =   1
            EncabezadosNombres=   "N°-Nombre-DNI-Telef.-Referido-DNI"
            EncabezadosAnchos=   "1000-2800-1000-1500-2300-1000"
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
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-R-L-C-C-C"
            FormatosEdit    =   "0-2-0-0-0-0"
            TextArray0      =   "N°"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   1005
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   96
         Top             =   360
         Width           =   9975
         Begin VB.TextBox Text1 
            Height          =   2010
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   97
            Top             =   240
            Width           =   9735
         End
      End
      Begin VB.Frame frmCredEvalFormato1 
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
         Height          =   6015
         Left            =   -74880
         TabIndex        =   83
         Top             =   360
         Width           =   9975
         Begin VB.TextBox txtDestino 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   89
            Top             =   5280
            Width           =   9735
         End
         Begin VB.TextBox txtColaterales 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   88
            Top             =   4320
            Width           =   9735
         End
         Begin VB.TextBox txtFormalidadNegocio 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   87
            Top             =   3360
            Width           =   9735
         End
         Begin VB.TextBox txtExperiencia 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   86
            Top             =   2400
            Width           =   9735
         End
         Begin VB.TextBox txtGiroUbicacion 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   85
            Top             =   1440
            Width           =   9735
         End
         Begin VB.TextBox txtEntornoFamiliar 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   84
            Top             =   480
            Width           =   9735
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   5040
            Width           =   2400
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   94
            Top             =   4080
            Width           =   2400
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   3120
            Width           =   4770
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sólo la experiencia crediticia:"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   2160
            Width           =   2070
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Left            =   120
            TabIndex        =   91
            Top             =   1200
            Width           =   2820
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   3795
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Otros Ingresos :"
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
         Height          =   2055
         Left            =   6640
         TabIndex        =   82
         Top             =   4740
         Width           =   4600
         Begin SICMACT.FlexEdit feOtrosIngresos 
            Height          =   1815
            Left            =   80
            TabIndex        =   20
            Top             =   200
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   3201
            Cols0           =   5
            HighLight       =   1
            EncabezadosNombres=   "-N-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-300-2800-1300-0"
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
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-C"
            FormatosEdit    =   "0-0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame fraBalanceGeneral3 
         Caption         =   "Balance General :"
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
         Height          =   3495
         Left            =   1580
         TabIndex        =   81
         Top             =   315
         Width           =   5040
         Begin SICMACT.FlexEdit feBalanceGeneral 
            Height          =   3255
            Left            =   60
            TabIndex        =   17
            Top             =   240
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   5741
            Cols0           =   7
            HighLight       =   1
            EncabezadosNombres=   "-nConsCod-nConsValor-N-Descripcion-Monto-Aux"
            EncabezadosAnchos=   "0-0-0-0-3400-1400-0"
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
            ColumnasAEditar =   "X-X-X-X-X-5-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-C-L-R-C"
            FormatosEdit    =   "0-0-0-0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame4 
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
         Height          =   3015
         Left            =   1580
         TabIndex        =   80
         Top             =   3780
         Width           =   5020
         Begin SICMACT.FlexEdit feGastosFamiliares 
            Height          =   2775
            Left            =   60
            TabIndex        =   19
            Top             =   195
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   4895
            Cols0           =   5
            HighLight       =   1
            EncabezadosNombres=   "-N-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-300-3200-1300-0"
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
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-C"
            FormatosEdit    =   "0-0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
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
         Height          =   4420
         Left            =   6640
         TabIndex        =   79
         Top             =   315
         Width           =   4600
         Begin SICMACT.FlexEdit feGastosNegocio 
            Height          =   4170
            Left            =   45
            TabIndex        =   18
            Top             =   195
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   7355
            Cols0           =   5
            HighLight       =   1
            EncabezadosNombres=   "-N-Concepto-Monto-Aux"
            EncabezadosAnchos=   "0-300-2890-1200-0"
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
            ColumnasAEditar =   "X-X-X-3-X"
            ListaControles  =   "0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-C"
            FormatosEdit    =   "0-0-0-2-0"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " Ventas y Cost."
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
         Height          =   2580
         Left            =   60
         TabIndex        =   74
         Top             =   320
         Width           =   1480
         Begin VB.TextBox txtIngresoNegocio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Text            =   "0"
            Top             =   480
            Width           =   1180
         End
         Begin VB.TextBox txtEgresoNegocio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Text            =   "0"
            Top             =   1200
            Width           =   1180
         End
         Begin SICMACT.EditMoney txtMargenBruto 
            Height          =   300
            Left            =   120
            TabIndex        =   75
            Top             =   2160
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   8421504
            Text            =   "0"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Bruto:"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   1920
            Width           =   1035
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Egreso Venta:"
            Height          =   195
            Left            =   120
            TabIndex        =   77
            Top             =   945
            Width           =   1020
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ingreso del Neg.:"
            Height          =   195
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   1260
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Propuesta del Credito:"
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
         Height          =   6375
         Left            =   -74880
         TabIndex        =   59
         Top             =   360
         Width           =   9855
         Begin VB.TextBox txtDestino2 
            Height          =   645
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   65
            Top             =   5520
            Width           =   9615
         End
         Begin VB.TextBox txtColaterales2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   64
            Top             =   4560
            Width           =   9615
         End
         Begin VB.TextBox txtFormalidadNegocio2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   63
            Top             =   3600
            Width           =   9615
         End
         Begin VB.TextBox txtGiroUbicacion2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   62
            Top             =   1680
            Width           =   9615
         End
         Begin VB.TextBox txtExperiencia2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   61
            Top             =   2640
            Width           =   9615
         End
         Begin VB.TextBox txtEntornoFamiliar2 
            Height          =   570
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   60
            Top             =   720
            Width           =   9615
         End
         Begin MSMask.MaskEdBox txtFechaVisita 
            Height          =   300
            Left            =   8520
            TabIndex        =   66
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el destino y el impacto del mismo:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   73
            Top             =   5280
            Width           =   2850
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre los colaterales y garantías:"
            Height          =   195
            Left            =   120
            TabIndex        =   72
            Top             =   4320
            Width           =   2400
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la consistencia de la información y la formalidad del negocio:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   71
            Top             =   3360
            Width           =   4770
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre la experiencia Crediticia:"
            Height          =   195
            Left            =   120
            TabIndex        =   70
            Top             =   2400
            Width           =   2220
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el giro y la ubicación del negocio:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   69
            Top             =   1440
            Width           =   2820
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sobre el entorno familiar del cliente o representante:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   480
            Width           =   3795
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Visita:"
            Height          =   195
            Left            =   7320
            TabIndex        =   67
            Top             =   300
            Width           =   1140
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Comentarios :"
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
         Height          =   2655
         Left            =   -74880
         TabIndex        =   57
         Top             =   360
         Width           =   9855
         Begin VB.TextBox txtComentario 
            Height          =   2250
            IMEMode         =   3  'DISABLE
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   58
            Top             =   240
            Width           =   9615
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Referidos :"
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
         Height          =   4095
         Left            =   -74880
         TabIndex        =   54
         Top             =   3120
         Width           =   9855
         Begin VB.CommandButton cmdAgregarRef 
            Caption         =   "&Agregar"
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
            Left            =   8520
            TabIndex        =   55
            Top             =   3600
            Width           =   1170
         End
         Begin SICMACT.FlexEdit feReferidos 
            Height          =   3255
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   9675
            _ExtentX        =   17066
            _ExtentY        =   5741
            Cols0           =   7
            HighLight       =   1
            EncabezadosNombres=   "N-Nombres-DNI-Telefono-Referido-NroDNI-Aux"
            EncabezadosAnchos=   "350-3000-960-1260-3000-960-0"
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
            ColumnasAEditar =   "X-1-2-3-4-5-X"
            ListaControles  =   "0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "L-L-L-L-L-L-C"
            FormatosEdit    =   "0-0-0-0-0-0-0"
            TextArray0      =   "N"
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   345
            RowHeight0      =   300
         End
      End
      Begin VB.CommandButton cmdQuitar2 
         Caption         =   "&Quitar"
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
         Left            =   -67680
         TabIndex        =   53
         Top             =   6720
         Width           =   1170
      End
   End
   Begin TabDlg.SSTab SSTabRatios 
      Height          =   1095
      Left            =   0
      TabIndex        =   110
      Top             =   9105
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Ratios e Indicadores"
      TabPicture(0)   =   "frmCredFormEvalFormato3.frx":037A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label13(2)"
      Tab(0).Control(1)=   "Label32"
      Tab(0).Control(2)=   "Label33"
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(4)=   "lblCapaAceptable"
      Tab(0).Control(5)=   "lblEndeAceptable"
      Tab(0).Control(6)=   "lblEndeudamiento"
      Tab(0).Control(7)=   "lblRentabilidad"
      Tab(0).Control(8)=   "lblLiquidez"
      Tab(0).Control(9)=   "txtRentabilidad"
      Tab(0).Control(10)=   "txtLiquidezCte"
      Tab(0).Control(11)=   "txtExcedenteMensual"
      Tab(0).Control(12)=   "txtIngresoNeto"
      Tab(0).Control(13)=   "txtEndeudamiento"
      Tab(0).Control(14)=   "txtCapacidadNeta"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Datos Flujo Caja Proyectada"
      TabPicture(1)   =   "frmCredFormEvalFormato3.frx":0396
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label13(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label13(4)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label13(6)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label13(7)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label13(5)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label34"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label13(8)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label13(9)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label35"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label13(10)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label36"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label13(11)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label39"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label13(12)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label40"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "EditMoneyIncC3"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "EditMoneyIncGV3"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "EditMoneyIncPP3"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "EditMoneyIncCM3"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "EditMoneyIncVC3"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      Begin SICMACT.EditMoney txtCapacidadNeta 
         Height          =   300
         Left            =   -73440
         TabIndex        =   111
         Top             =   450
         Width           =   850
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtEndeudamiento 
         Height          =   300
         Left            =   -71320
         TabIndex        =   112
         Top             =   450
         Width           =   850
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtIngresoNeto 
         Height          =   300
         Left            =   -66480
         TabIndex        =   113
         Top             =   570
         Width           =   1100
         _ExtentX        =   1931
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtExcedenteMensual 
         Height          =   300
         Left            =   -64920
         TabIndex        =   114
         Top             =   570
         Width           =   1100
         _ExtentX        =   1931
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtLiquidezCte 
         Height          =   300
         Left            =   -68160
         TabIndex        =   120
         Top             =   570
         Width           =   1100
         _ExtentX        =   1931
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtRentabilidad 
         Height          =   300
         Left            =   -69735
         TabIndex        =   121
         Top             =   570
         Width           =   1100
         _ExtentX        =   1931
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   8421504
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncVC3 
         Height          =   300
         Left            =   120
         TabIndex        =   126
         ToolTipText     =   "Incremento de ventas al contado - Anual"
         Top             =   600
         Width           =   1305
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "0.00"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncCM3 
         Height          =   300
         Left            =   2880
         TabIndex        =   127
         ToolTipText     =   "Incremento de ventas al contado - Anual"
         Top             =   600
         Width           =   1305
         _ExtentX        =   1508
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "0.00"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncPP3 
         Height          =   300
         Left            =   5160
         TabIndex        =   128
         ToolTipText     =   "Incremento de ventas al contado - Anual"
         Top             =   600
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "0.00"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncGV3 
         Height          =   300
         Left            =   7440
         TabIndex        =   129
         ToolTipText     =   "Incremento de ventas al contado - Anual"
         Top             =   600
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483640
         Text            =   "0.00"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney EditMoneyIncC3 
         Height          =   300
         Left            =   9480
         TabIndex        =   130
         ToolTipText     =   "Incremento de ventas al contado - Anual"
         Top             =   600
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   -2147483645
         Text            =   "0.00"
      End
      Begin VB.Label lblLiquidez 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Liquidez Cte:"
         Height          =   195
         Left            =   -68160
         TabIndex        =   149
         Top             =   360
         Width           =   930
      End
      Begin VB.Label lblRentabilidad 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rentabilidad Pat.:"
         Height          =   195
         Left            =   -69720
         TabIndex        =   148
         Top             =   380
         Width           =   1290
      End
      Begin VB.Label lblEndeudamiento 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Endeudamiento:"
         Height          =   195
         Left            =   -72480
         TabIndex        =   147
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label Label40 
         Caption         =   "Anual"
         Height          =   255
         Left            =   10800
         TabIndex        =   145
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   12
         Left            =   10560
         TabIndex        =   144
         Top             =   645
         Width           =   165
      End
      Begin VB.Label Label39 
         Caption         =   "Anual"
         Height          =   255
         Left            =   8880
         TabIndex        =   143
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   11
         Left            =   8640
         TabIndex        =   142
         Top             =   645
         Width           =   165
      End
      Begin VB.Label Label36 
         Caption         =   "Anual"
         Height          =   255
         Left            =   6600
         TabIndex        =   141
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   10
         Left            =   6360
         TabIndex        =   140
         Top             =   645
         Width           =   165
      End
      Begin VB.Label Label35 
         Caption         =   "Anual"
         Height          =   255
         Left            =   4440
         TabIndex        =   139
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   9
         Left            =   4200
         TabIndex        =   138
         Top             =   645
         Width           =   165
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   8
         Left            =   1440
         TabIndex        =   137
         Top             =   650
         Width           =   165
      End
      Begin VB.Label Label34 
         Caption         =   "Anual"
         Height          =   255
         Left            =   1680
         TabIndex        =   136
         Top             =   650
         Width           =   495
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. de Consumo:"
         Height          =   195
         Index           =   5
         Left            =   9480
         TabIndex        =   135
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. Gasto Ventas:"
         Height          =   195
         Index           =   7
         Left            =   7320
         TabIndex        =   134
         Top             =   360
         Width           =   1410
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. Pago Personal:"
         Height          =   195
         Index           =   6
         Left            =   5160
         TabIndex        =   133
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incr. compra mercaderias:"
         Height          =   195
         Index           =   4
         Left            =   2880
         TabIndex        =   132
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Incremento ventas contado:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   131
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblEndeAceptable 
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
         Left            =   -71300
         TabIndex        =   123
         Top             =   795
         Width           =   750
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
         Height          =   150
         Left            =   -73400
         TabIndex        =   122
         Top             =   800
         Width           =   750
      End
      Begin VB.Line Line1 
         X1              =   -70080
         X2              =   -70080
         Y1              =   360
         Y2              =   960
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Excedente:"
         Height          =   195
         Left            =   -64920
         TabIndex        =   117
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso Neto:"
         Height          =   195
         Left            =   -66480
         TabIndex        =   116
         Top             =   375
         Width           =   1005
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad de Pago:"
         Height          =   195
         Index           =   2
         Left            =   -74880
         TabIndex        =   115
         Top             =   520
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   10620
      TabIndex        =   33
      Top             =   40
      Visible         =   0   'False
      Width           =   700
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9960
      Y1              =   10320
      Y2              =   10320
   End
End
Attribute VB_Name = "frmCredFormEvalFormato3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmCredFormEvalFormato3
'** Descripción : Formulario para evaluación de Creditos que tienen el tipo de evaluación 2
'** Referencia  : ERS004-2016
'** Creación    : LUCV, 20160525 09:00:00 AM
'**********************************************************************************************
Option Explicit
    Dim sCtaCod As String
    Dim sPersCod As String
    Dim gsOpeCod As String
    Dim fnTipoRegMant As Integer
    Dim fnTipoPermiso As Integer
    Dim fbPermiteGrabar As Boolean
    Dim fbBloqueaTodo As Boolean
    Dim fnTotalRefGastoNego As Currency
    Dim fnTotalRefGastoFami As Currency
    Dim fsCliente As String
    Dim fsGiroNego As String
    Dim fsAnioExp As Integer
    Dim fsMesExp As Integer
    Dim fsUserAnalista  As String
    Dim fnEstado As Integer
    Dim fnMontoDeudaSbs As Currency
    Dim fnFechaDeudaSbs As Currency
    Dim fnPlazo As Integer
    Dim fnProducto As Integer
    
    Dim lnCondLocal As Integer
    Dim MatIfiGastoNego As Variant
    Dim MatIfiGastoFami As Variant
    Dim MatReferidos As Variant
    
    Dim MatIfiNoSupervisadaGastoNego As Variant 'CTI320200110 ERS003-2020. Agregó
    Dim MatIfiNoSupervisadaGastoFami As Variant 'CTI320200110 ERS003-2020. Agregó
    
    Dim rsFeGastoNeg As ADODB.Recordset
    Dim rsFeDatGastoFam As ADODB.Recordset
    Dim rsFeDatOtrosIng As ADODB.Recordset
    Dim rsFeDatBalanGen As ADODB.Recordset
    Dim rsFeDatActivos As ADODB.Recordset
    Dim rsFeDatPasivos As ADODB.Recordset
    Dim rsFeDatPasivosNo As ADODB.Recordset
    Dim rsFeDatPatrimonio As ADODB.Recordset
    Dim rsFeDatRef As ADODB.Recordset
    
    Dim rsCredEval As ADODB.Recordset
    Dim rsDCredito As ADODB.Recordset
    Dim rsAceptableCritico As ADODB.Recordset
    Dim rsCapacPagoNeta As ADODB.Recordset
    Dim rsCuotaIFIs As ADODB.Recordset
    Dim rsPropuesta As ADODB.Recordset
        
    Dim rsDatPasivosNo As ADODB.Recordset
    Dim rsDatActivoPasivo As ADODB.Recordset
    Dim rsDatGastoNeg As ADODB.Recordset
    Dim rsDatGastoFam As ADODB.Recordset
    Dim rsDatOtrosIng As ADODB.Recordset
    Dim rsDatRef As ADODB.Recordset
    Dim rsDatRatioInd As ADODB.Recordset
    Dim rsDatIfiGastoNego As ADODB.Recordset
    Dim rsDatIfiGastoFami As ADODB.Recordset
    Dim rsDatVentaCosto As ADODB.Recordset
    Dim rsDatActivos As ADODB.Recordset
    Dim rsDatPasivos As ADODB.Recordset
    
    Dim nMontoAct As Currency
    Dim nMontoPas As Currency
    Dim nMontoPat As Currency
    Dim nMargenBruto As Currency
    
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim objPista As COMManejador.Pista
    Dim fnFormato As Integer
    Dim fnMontoIni As Double
    Dim lnMin As Double
    Dim lnMax As Double
    Dim lnMinDol As Double
    Dim lnMaxDol As Double
    Dim nTC As Double
    Dim i As Integer
    Dim j As Integer
    Dim K As Integer
    Dim fbGrabar As Boolean
    Dim fnColocCondi As Integer
    Dim fbTieneReferido6Meses As Boolean 'LUCV20171115, Agregó segun correo: RUSI

    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Dim rsDatParamFlujoCajaForm3 As ADODB.Recordset
    Dim nMaximo As Integer
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    
    Dim MatBalActCorr As Variant
    Dim MatBalActNoCorr As Variant
    Dim rsDatIfiBalActCorri As ADODB.Recordset
    Dim rsDatIfiBalActNoCorri As ADODB.Recordset

    Dim lcMovNro As String 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
    Dim rsDatIfiNoSupervisadaGastoNego As ADODB.Recordset 'CTI320200110 ERS003-2020
    Dim rsDatIfiNoSupervisadaGastoFami As ADODB.Recordset 'CTI320200110 ERS003-2020
    Dim fbImprimirVB As Boolean 'CTI320200110 ERS003-2020
    Dim pnMontoOtrasIfisConsumo As Double
    Dim pnMontoOtrasIfisEmpresarial As Double
    
Private Sub cmdGenerarFlujoForm3_Click()

On Error GoTo ErrorInicioExcel ' agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM

Dim lsArchivo As String
Dim lbLibroOpen As Boolean
Dim bGeneraExcel As Boolean 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM

    'lsArchivo = App.Path & "\Spooler\FlujoCaja_Formato3" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls"  'comentado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
    lsArchivo = App.Path & "\Spooler\FlujoCaja_Formato3" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls"  'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
    lbLibroOpen = ExcelInicio(lsArchivo, xlAplicacion, xlLibro)
    
    If lbLibroOpen Then
    bGeneraExcel = False 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
        bGeneraExcel = generaExcelForm3 'agregado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
        If bGeneraExcel Then 'modificado pti1 20180726
            ExcelFin lsArchivo, xlAplicacion, xlLibro, xlHoja1
            'AbrirArchivo "FlujoCaja_Formato3" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls", App.Path & "\Spooler" 'comentado pti1 20180726 Memorandum Nº 1602-2018-GM-DI_CMACM
            AbrirArchivo "FlujoCaja_Formato3" & gsCodUser & Format(gdFecSis, "DDMMYYYY") & ".xls", App.Path & "\Spooler"
        End If
    End If
    
Exit Sub 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
ErrorInicioExcel:         ' agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
MsgBox Err.Description + "Error 1: Error al iniciar la creación del excel comunicar a TI", vbInformation, "Error" 'pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
    
End Sub

Public Function generaExcelForm3() As Boolean

    On Error GoTo ErrorInicioExcel 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
    
    
    Dim ssql As String
    Dim rs As New ADODB.Recordset
    Dim rsCabcera As New ADODB.Recordset
    Dim rsCuotas As New ADODB.Recordset
    Dim rsParFlujoCaja As New ADODB.Recordset
    Dim oCont As COMConecta.DCOMConecta
    Dim i As Integer
    Dim nCon As Integer
    Dim nFila As Integer
    Dim nCol As Integer
    Dim nColFin As Integer
    Dim a As Integer
    Dim nColInicio As Integer
    Dim Z As Integer
    
    Dim dFechaEval As Date
    
    generaExcelForm3 = True
    
    'proteger Libro
    'xlAplicacion.ActiveWorkbook.Protect (123) 'pti comentado
    
    'Adiciona una hoja
    ExcelAddHoja "Hoja1", xlLibro, xlHoja1, True
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    
    xlHoja1.Cells(2, 2) = "FLUJO DE CAJA MENSUAL PRESUPUESTADO"
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 12)).Font.Bold = True
    
    xlHoja1.Cells(4, 1) = "CLIENTE: "
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(5, 1) = "ANALISTA: "
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(6, 1) = "DNI: "
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 1)).HorizontalAlignment = xlLeft
    
    xlHoja1.Cells(7, 1) = "RUC: "
    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(7, 1)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 1), xlHoja1.Cells(7, 1)).HorizontalAlignment = xlLeft
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosCabecera  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rsCabcera = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    
    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosCuotas  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rsCuotas = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    

    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosConceptos  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rs = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    

    ssql = "exec stp_sel_ERS0512017_FlujoCajaRptObtieneDatosParametros  '" & ActXCodCta.NroCuenta & "'"

    Set oCont = New COMConecta.DCOMConecta
    oCont.AbreConexion
    Set rsParFlujoCaja = oCont.CargaRecordSet(ssql)
    oCont.CierraConexion
    Set oCont = Nothing
    
    

'Cabecera
If Not (rsCabcera.EOF And rsCabcera.BOF) Then
    dFechaEval = rsCabcera!fechaEval
    
    xlHoja1.Cells(4, 2) = rsCabcera!NombreClie
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(4, 2), xlHoja1.Cells(4, 6)).Font.Bold = True

    xlHoja1.Cells(5, 2) = rsCabcera!NombreAnal
    xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 6)).Font.Bold = True
    
    xlHoja1.Cells(6, 2) = rsCabcera!nDoc
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 6)).Font.Bold = True

    xlHoja1.Cells(7, 2) = rsCabcera!nDocTrib
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(7, 6)).Font.Bold = True
    
   
Else
     
        MsgBox "Error, Comuníquese con el Área de TI", vbInformation, "!Error!"
        generaExcelForm3 = False
        Exit Function
End If
    
    
    xlHoja1.Cells(9, 2) = "Conceptos / Meses"
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 2)).MergeCells = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 2)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 2)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(10, 2)).Cells.Interior.Color = RGB(141, 180, 226)
        
    xlHoja1.Cells(9, 3) = "Flujo Mensual"
    xlHoja1.Cells(10, 3) = Format(dFechaEval, "mmm-yyyy")
    xlHoja1.Range(xlHoja1.Cells(9, 3), xlHoja1.Cells(10, 3)).Cells.Interior.Color = RGB(141, 180, 226)
    xlHoja1.Range(xlHoja1.Cells(9, 3), xlHoja1.Cells(10, 3)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(9, 3), xlHoja1.Cells(10, 3)).HorizontalAlignment = xlCenter
    
    CuadroExcel xlHoja1, 2, 9, 3, 10, True
    CuadroExcel xlHoja1, 2, 9, 3, 10, False
    
    nCon = 11
    
 
    
'Conceptos
    If Not (rs.EOF And rs.BOF) Then
        For i = 1 To rs.RecordCount
        
            CuadroExcel xlHoja1, 2, nCon, 3, nCon
        
            xlHoja1.Cells(nCon, 2) = rs!Descripcion
            xlHoja1.Cells(nCon, 3) = rs!Monto
            
            If rs!Descripcion = "INVERSION" Then
                nCon = nCon + 2
            Else
                nCon = nCon + 1
            End If
            
            xlHoja1.Range(xlHoja1.Cells(11, 2), xlHoja1.Cells(11, nCon)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(14, 2), xlHoja1.Cells(14, nCon)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(27, 2), xlHoja1.Cells(27, nCon)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(31, 2), xlHoja1.Cells(31, nCon)).Font.Bold = True
            
            CuadroExcel xlHoja1, 2, nCon, 3, nCon - 1
            
            rs.MoveNext
        Next i
        
         
         
    Else
      
    
        MsgBox "Error, Comuníquese con el Área de TI", vbInformation, "!Error!"
        generaExcelForm3 = False
        Exit Function
    End If
  
'Pie
    If Not (rsParFlujoCaja.EOF And rsParFlujoCaja.BOF) Then

        xlHoja1.Cells(nCon + 1, 2) = "DATOS ADICIONALES"
        xlHoja1.Range(xlHoja1.Cells(nCon + 1, 2), xlHoja1.Cells(nCon + 1, 2)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(nCon + 1, 2), xlHoja1.Cells(nCon + 1, 2)).HorizontalAlignment = xlCenter
        CuadroExcel xlHoja1, 2, nCon + 1, 3, nCon + 1
        xlHoja1.Cells(nCon + 2, 2) = "Fecha de Pago"
        xlHoja1.Cells(nCon + 2, 3) = Format(rsParFlujoCaja!dFechaPago, "YYYY/mm/dd")
        CuadroExcel xlHoja1, 2, nCon + 2, 3, nCon + 2

        xlHoja1.Cells(nCon + 4, 3) = "Mes"
        xlHoja1.Cells(nCon + 4, 4) = "Anual"
        CuadroExcel xlHoja1, 3, nCon + 4, 4, nCon + 4
        xlHoja1.Range(xlHoja1.Cells(nCon + 4, 3), xlHoja1.Cells(nCon + 4, 4)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(nCon + 4, 3), xlHoja1.Cells(nCon + 4, 4)).HorizontalAlignment = xlCenter

        xlHoja1.Cells(nCon + 5, 2) = "Incremento de ventas al contado "
        xlHoja1.Cells(nCon + 6, 2) = "Incremento de Compra de Mercaderias"
        xlHoja1.Cells(nCon + 7, 2) = "Incremento de Consumo"
        xlHoja1.Cells(nCon + 8, 2) = "Incremento de Pago Personal"
        xlHoja1.Cells(nCon + 9, 2) = "Ingremento de Gastos de Ventas"

        xlHoja1.Cells(nCon + 5, 3) = Format(((1 + rsParFlujoCaja!nIncVentCont / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 6, 3) = Format(((1 + rsParFlujoCaja!nIncCompMerc / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 7, 3) = Format(((1 + rsParFlujoCaja!nIncConsu / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 8, 3) = Format(((1 + rsParFlujoCaja!nIncPagPers / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"
        xlHoja1.Cells(nCon + 9, 3) = Format(((1 + rsParFlujoCaja!nIncGastvent / 100) ^ (1 / 12) - 1) * 100, "#0.00") & "%"

        xlHoja1.Cells(nCon + 5, 4) = Format(rsParFlujoCaja!nIncVentCont, "#0.0") & "%"
        xlHoja1.Cells(nCon + 6, 4) = Format(rsParFlujoCaja!nIncCompMerc, "#0.0") & "%"
        xlHoja1.Cells(nCon + 7, 4) = Format(rsParFlujoCaja!nIncConsu, "#0.0") & "%"
        xlHoja1.Cells(nCon + 8, 4) = Format(rsParFlujoCaja!nIncPagPers, "#0.0") & "%"
        xlHoja1.Cells(nCon + 9, 4) = Format(rsParFlujoCaja!nIncGastvent, "#0.0") & "%"

        CuadroExcel xlHoja1, 2, nCon + 5, 4, nCon + 9, True
        CuadroExcel xlHoja1, 2, nCon + 5, 4, nCon + 9, False
        xlHoja1.Range(xlHoja1.Cells(nCon + 5, 3), xlHoja1.Cells(nCon + 9, 4)).HorizontalAlignment = xlCenter
        
        
    Else
         
        MsgBox "Registre los Datos de Flujo de Caja Proyectada, y dar click en Guardar", vbInformation, "!Aviso!"
        generaExcelForm3 = False
        Exit Function
    End If
    
'Obtener las Letras del Abecedario A-Z
    Dim MatAZ As Variant
    Dim P As Integer
    P = 1
    Set MatAZ = Nothing
    ReDim MatAZ(1, 140)
    For i = 65 To 90
        MatAZ(1, P) = ChrW(i)
        P = P + 1
    Next i
           
              
    Dim MatLetrasRep As Variant
    Dim Y As Integer
    Set MatLetrasRep = Nothing
    Y = 1
    ReDim MatLetrasRep(1, 131)
    For a = 1 To 130
        If a <= 26 Then
                MatLetrasRep(1, Y) = ChrW(65) & MatAZ(1, Y) 'AA,AB,AC......AZ
            Y = Y + 1
        ElseIf (a >= 27 And a <= 52) Then
            If a = 27 Then
                P = 1
            End If
                MatLetrasRep(1, Y) = ChrW(66) & MatAZ(1, P) 'BA,BB,BC......BZ
            Y = Y + 1
            P = P + 1
        ElseIf (a >= 53 And a <= 78) Then
            If a = 53 Then
                P = 1
            End If
                MatLetrasRep(1, Y) = ChrW(67) & MatAZ(1, P) 'CA,CB,CC......CZ
            Y = Y + 1
            P = P + 1
        ElseIf (a >= 79 And a <= 104) Then
            If a = 79 Then
                P = 1
            End If
                MatLetrasRep(1, Y) = ChrW(68) & MatAZ(1, P) 'DA,DB,DC......DZ
            Y = Y + 1
            P = P + 1
        End If
    Next a
    

''Cuotas
i = 0
Y = 0
Z = 0
nFila = 34
nCol = 4
nColInicio = 4
nColFin = 0
   If Not (rsCuotas.EOF And rsCuotas.BOF) Then
   
        For i = 1 To rsCuotas.RecordCount

            If i >= 24 Then
                Y = Y + 1
            End If

            xlHoja1.Cells(9, nCol) = rsCuotas!nCuota
            xlHoja1.Range(xlHoja1.Cells(9, 4), xlHoja1.Cells(9, nCol)).Cells.Interior.Color = RGB(141, 180, 226)
            xlHoja1.Range(xlHoja1.Cells(9, nCol), xlHoja1.Cells(9, nCon)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(9, nCol), xlHoja1.Cells(9, nCon)).HorizontalAlignment = xlCenter
            
            xlHoja1.Cells(10, nCol) = Format(rsCuotas!dFechaCuotas, "mmm-yyyy")
            xlHoja1.Range(xlHoja1.Cells(10, 4), xlHoja1.Cells(10, nCol)).Cells.Interior.Color = RGB(141, 180, 226)
            xlHoja1.Range(xlHoja1.Cells(10, nCol), xlHoja1.Cells(10, nCon)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(10, nCol), xlHoja1.Cells(10, nCon)).HorizontalAlignment = xlCenter

            'Ingresos Operativos
            xlHoja1.Range(xlHoja1.Cells(11, 2), xlHoja1.Cells(11, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            xlHoja1.Cells(11, nCol) = "=SUM(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "12" & ":" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "13)"
            
            'Ventas al Contado
            xlHoja1.Cells(12, nCol) = Round((xlHoja1.Cells(12, nCol - 1) * ((1 + rsParFlujoCaja!nIncVentCont / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(12, nCol - 1)))
            
            'Otros Ingresos
            xlHoja1.Cells(13, nCol) = "=C13"
            
            'Engresos Operativos
            xlHoja1.Range(xlHoja1.Cells(14, 2), xlHoja1.Cells(14, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            xlHoja1.Cells(14, nCol) = "=SUM(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "15" & ":" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "26)"
            
            'Compra de Mercaderia
            xlHoja1.Cells(15, nCol) = Round((xlHoja1.Cells(15, nCol - 1) * ((1 + rsParFlujoCaja!nIncCompMerc / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(15, nCol - 1)))
            
            'Planilla
            xlHoja1.Cells(16, nCol) = Round((xlHoja1.Cells(16, nCol - 1) * ((1 + rsParFlujoCaja!nIncPagPers / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(16, nCol - 1)))
            
            'calculo Alquiler de Locales
            xlHoja1.Cells(17, nCol) = "=C17"
            
            'calculo Utiles de oficinas
            xlHoja1.Cells(18, nCol) = "=C18"

            'calculo Rep y Mtto de Equipos
            xlHoja1.Cells(19, nCol) = "=C19"

            'calculo Rep y Mtto de Vehiculo
            xlHoja1.Cells(20, nCol) = "=C20"

            'Seguro y Flete
            xlHoja1.Cells(21, nCol) = "=C21"

            'calculo Transporte/Combustible/ Gas
            xlHoja1.Cells(22, nCol) = "=C22"

            'calculo Sunat + Impuestos
            xlHoja1.Cells(23, nCol) = "=C23"

            'calculo Publicidad y otros gastos de ventas (**Nuevo)
            xlHoja1.Cells(24, nCol) = Round((xlHoja1.Cells(24, nCol - 1) * ((1 + rsParFlujoCaja!nIncGastvent / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(24, nCol - 1)))
            
            'calculo Otros
            xlHoja1.Cells(25, nCol) = "=C25"

            'calculo Consumo Per.Nat.
            xlHoja1.Cells(26, nCol) = Round((xlHoja1.Cells(26, nCol - 1) * ((1 + rsParFlujoCaja!nIncConsu / 100) ^ (1 / 12) - 1) + xlHoja1.Cells(26, nCol - 1)))

            'calculo Flujo Operativo
            xlHoja1.Range(xlHoja1.Cells(27, 2), xlHoja1.Cells(27, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            xlHoja1.Cells(27, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "11" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "14)"

            'Cobro de Prestamo y dividendos
            xlHoja1.Cells(28, nCol) = "=C28"

            'Pago de cuota Prestamos vigentes
            xlHoja1.Cells(29, nCol) = "=C29"

            'Pago de cuotas de prestamos solicitado
            xlHoja1.Cells(30, nCol) = "=C30"

            'calculo Flujo Financiero
            xlHoja1.Range(xlHoja1.Cells(31, 2), xlHoja1.Cells(31, nCol)).Cells.Interior.Color = RGB(190, 190, 190)
            xlHoja1.Cells(31, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "27" & "+" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "28" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "29" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "30)"
            
            'calculo Inversion
            xlHoja1.Cells(32, nCol) = "=C32"
            
            'calculo Saldo
            xlHoja1.Cells(34, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "31" & "-" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "32)"
            'Si los datos son numero negativos se pone rojo SALDO
            If xlHoja1.Cells(34, nCol) < 0 Then
                xlHoja1.Range(xlHoja1.Cells(34, nCol), xlHoja1.Cells(34, nCol)).Cells.Interior.Color = RGB(255, 0, 0)
            End If
            
            'calculo Saldo Disponible
            If i >= 25 Then
                Z = Z + 1
            End If
            xlHoja1.Cells(35, nCol) = "=(" & IIf(i >= 25, MatLetrasRep(1, Z), MatAZ(1, i + 2)) & "36)"
            'Si los datos son numero negativos se pone rojo SALDO
            If xlHoja1.Cells(35, nCol) < 0 Then
                xlHoja1.Range(xlHoja1.Cells(35, nCol), xlHoja1.Cells(35, nCol)).Cells.Interior.Color = RGB(255, 0, 0)
            End If
            
            'calculo Saldo Acumulado
            xlHoja1.Cells(36, nCol) = "=(" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "34" & "+" & IIf(i >= 24, MatLetrasRep(1, Y), MatAZ(1, i + 3)) & "35)"
            'Si los datos son numero negativos se pone rojo SALDO
            If xlHoja1.Cells(36, nCol) < 0 Then
                xlHoja1.Range(xlHoja1.Cells(36, nCol), xlHoja1.Cells(36, nCol)).Cells.Interior.Color = RGB(255, 0, 0)
            End If
            
            nCol = nCol + 1

            If (i Mod 12) = 0 Then
                nColFin = nCol - 1
                    xlHoja1.Cells(8, nColInicio) = "Año" & (i / 12)
                    xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nColFin)).HorizontalAlignment = xlCenter
                    xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nColFin)).MergeCells = True
                    xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nColFin)).Font.Bold = True
                nColInicio = nColFin + 1
            End If
            rsCuotas.MoveNext
        Next i
      
        If nCol <> nColInicio Then
        'Para la celda si no cumple un año
        xlHoja1.Range(xlHoja1.Cells(8, nColInicio), xlHoja1.Cells(8, nCol - 1)).MergeCells = True
        End If
        
        xlHoja1.Range(xlHoja1.Cells(8, 4), xlHoja1.Cells(8, nCol - 1)).Cells.Interior.Color = RGB(141, 180, 226)
        CuadroExcel xlHoja1, 4, 8, nCol - 1, 8
        
        For i = 0 To 29
            If i <= 23 Then
                CuadroExcel xlHoja1, 4, 9 + i, nCol - 1, 9 + i
            ElseIf i >= 27 Then
                CuadroExcel xlHoja1, 4, nFila, nCol - 1, nFila
                nFila = nFila + 1
            End If
        Next i
          
        
    Else
        MsgBox "Error al crear el Excel, Comuníquese con el Área de TI", vbInformation, "!Error!"
        generaExcelForm3 = False
        Exit Function
    End If

xlHoja1.Cells.Select
xlHoja1.Cells.Font.Name = "Arial"
xlHoja1.Cells.Font.Size = 9
xlHoja1.Cells.EntireColumn.AutoFit

'xlAplicacion.Worksheets("Hoja1").Protect ("123")

MsgBox "Reporte Generado Satisfactoriamente", vbInformation, "!Exito!"

rs.Close
rsCabcera.Close
rsParFlujoCaja.Close
rsCuotas.Close


Exit Function 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
ErrorInicioExcel: 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM
MsgBox Err.Description + "Error 2: Error al iniciar la creación del excel, Comuníquese con el Área de TI", vbInformation, "Error" 'agregado pti1 26072018 Memorandum Nº 1602-2018-GM-DI_CMACM

End Function

'JOEP20180725 ERS034-2018
Private Sub cmdMNME_Click()
    Call frmCredFormEvalCredCel.Inicio(ActXCodCta.NroCuenta, 11)
End Sub
'JOEP20180725 ERS034-2018

Private Sub EditMoneyIncCM3_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        EditMoneyIncPP3.SetFocus
        fEnfoque EditMoneyIncPP3
    End If
End Sub

Private Sub EditMoneyIncPP3_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        EditMoneyIncGV3.SetFocus
        fEnfoque EditMoneyIncGV3
    End If
End Sub

Private Sub EditMoneyIncVC3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EditMoneyIncCM3.SetFocus
        fEnfoque EditMoneyIncCM3
    End If
End Sub

Private Sub feBalanceGeneral_Click()
If fnTipoRegMant = 2 Then
        If feBalanceGeneral.col = 5 Then
                If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
                    feBalanceGeneral.ListaControles = "0-0-0-0-1"
                ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
                    feBalanceGeneral.ListaControles = "0-0-0-0-1"
                Else
                    feBalanceGeneral.ListaControles = "0-0-0-0-0"
                End If
        End If
Else
    If feBalanceGeneral.col = 5 Then
        If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        Else
            feBalanceGeneral.ListaControles = "0-0-0-0-0"
        End If
    End If
End If
End Sub

Private Sub feBalanceGeneral_EnterCell()
If fnTipoRegMant = 2 Then
    If feBalanceGeneral.col = 5 Then
        If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        Else
            feBalanceGeneral.ListaControles = "0-0-0-0-0"
        End If
    End If
Else
    If feBalanceGeneral.col = 5 Then
        If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        Else
            feBalanceGeneral.ListaControles = "0-0-0-0-0"
        End If
    End If
End If
End Sub

Private Sub feBalanceGeneral_OnClickTxtBuscar(psCodigo As String, psDescripcion As String)
    Dim fnTotalBalanceActCorriente As Currency
    Dim fnTotalBalanceActNoCorriente As Currency

    psCodigo = 0
    psDescripcion = ""
    psDescripcion = feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4) 'Cuotas Otras IFIs
    psCodigo = feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5) 'Monto
    
If feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 Then
    If psCodigo = 0 Then
        fnTotalBalanceActCorriente = 0
        Set MatBalActCorr = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CLng(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5))), fnTotalBalanceActCorriente, MatBalActCorr, feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4)
        psCodigo = Format(fnTotalBalanceActCorriente, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CLng(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5))), fnTotalBalanceActCorriente, MatBalActCorr, feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4)
        psCodigo = Format(fnTotalBalanceActCorriente, "#,##0.00")
    End If
ElseIf feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 Then
    If psCodigo = 0 Then
         fnTotalBalanceActNoCorriente = 0
        Set MatBalActNoCorr = Nothing
        frmCredFormEvalCuotasIfis.Inicio (CLng(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5))), fnTotalBalanceActNoCorriente, MatBalActNoCorr, feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4)
        psCodigo = Format(fnTotalBalanceActNoCorriente, "#,##0.00")
    Else
        frmCredFormEvalCuotasIfis.Inicio (CLng(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 5))), fnTotalBalanceActNoCorriente, MatBalActNoCorr, feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 4)
        psCodigo = Format(fnTotalBalanceActNoCorriente, "#,##0.00")
    End If
End If
End Sub
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja

'_____________________________________________________________________________________________________________
'******************************************LUCV20160525: EVENTOS Varios***************************************
Private Sub Form_Load()
    fbGrabar = False
    CentraForm Me
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    EnfocaControl spnTiempoLocalAnio

'JOEP20180725 ERS034-2018
    If fnTipoRegMant = 3 Then
        If Not ConsultaRiesgoCamCred(sCtaCod) Then
            cmdMNME.Visible = True
        End If
    End If
'JOEP20180725 ERS034-2018
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set MatIfiGastoNego = Nothing 'LUCV20161115
    Set MatIfiGastoFami = Nothing 'LUCV20161115
    Set MatIfiNoSupervisadaGastoNego = Nothing 'CTI320200110 ERS003-2020. Agregó
    Set MatIfiNoSupervisadaGastoFami = Nothing 'CTI320200110 ERS003-2020. Agregó
End Sub
Private Sub Cmdguardar_Click()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim GrabarDatos As Boolean
    Dim rsGastoNeg As ADODB.Recordset
    Dim rsGastoFam As ADODB.Recordset
    Dim rsOtrosIng As ADODB.Recordset
    Dim rsBalGen As ADODB.Recordset
    Dim MatActiPasivo As Variant
    Dim MatActiPasivoDet As Variant
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsRatiosActual As ADODB.Recordset
    Dim rsRatiosAceptableCritico As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval

    'Para Contar los totales y los detalles de los activos pasivos
    Dim nContadorTotal As Integer
    Dim nContadorDet As Integer
    Dim nContador As Integer
    Set rsGastoNeg = IIf(feGastosNegocio.rows - 1 > 0, feGastosNegocio.GetRsNew(), Nothing)
    Set rsGastoFam = IIf(feGastosFamiliares.rows - 1 > 0, feGastosFamiliares.GetRsNew(), Nothing)
    Set rsOtrosIng = IIf(feOtrosIngresos.rows - 1 > 0, feOtrosIngresos.GetRsNew(), Nothing)
'Contar Totales y Detalles (ActivoPasivo) -> Filas ******
     nContadorTotal = 0
     nContadorDet = 0
     For i = 1 To feBalanceGeneral.rows - 1
        If feBalanceGeneral.TextMatrix(i, 3) = "" Then
        nContadorTotal = nContadorTotal + 1
        Else
        nContadorDet = nContadorDet + 1
        End If
    Next i
    'Fin Filas <-**********
    
    'Flex a Matriz Referidos **********->
        ReDim MatReferidos(feReferidos3.rows - 1, 6)
        For i = 1 To feReferidos3.rows - 1
            MatReferidos(i, 1) = feReferidos3.TextMatrix(i, 0)
            MatReferidos(i, 2) = feReferidos3.TextMatrix(i, 1)
            MatReferidos(i, 3) = feReferidos3.TextMatrix(i, 2)
            MatReferidos(i, 4) = feReferidos3.TextMatrix(i, 3)
            MatReferidos(i, 5) = feReferidos3.TextMatrix(i, 4)
            MatReferidos(i, 6) = feReferidos3.TextMatrix(i, 5)
         Next i
    'Fin Referidos
    
    'LUCV20162606, Carga Matriz Activo, Pasivo, Patrimonio, Totales **********->
    i = 0
    j = 0
    K = 0
    nContador = 0
    ReDim MatActiPasivo(nContadorTotal + 1, 5)
    ReDim MatActiPasivoDet(nContadorDet + 1, 5)
    While feBalanceGeneral.rows - 1 > nContador
        i = i + 1
        'Para Cargar Datos en Matriz-> CredFormEvalActivoPasivo
        If feBalanceGeneral.TextMatrix(i, 3) = "" Then
            j = j + 1
            MatActiPasivo(j, 1) = feBalanceGeneral.TextMatrix(i, 1)
            MatActiPasivo(j, 2) = feBalanceGeneral.TextMatrix(i, 2)
            MatActiPasivo(j, 3) = feBalanceGeneral.TextMatrix(i, 3)
            MatActiPasivo(j, 4) = feBalanceGeneral.TextMatrix(i, 4)
            MatActiPasivo(j, 5) = CDbl(feBalanceGeneral.TextMatrix(i, 5))
         Else 'Para Cargar Datos en Matriz-> CredFormEvalActivoPasivoDet
             K = K + 1
            MatActiPasivoDet(K, 1) = feBalanceGeneral.TextMatrix(i, 1)
            MatActiPasivoDet(K, 2) = feBalanceGeneral.TextMatrix(i, 2)
            MatActiPasivoDet(K, 3) = feBalanceGeneral.TextMatrix(i, 3)
            MatActiPasivoDet(K, 4) = feBalanceGeneral.TextMatrix(i, 4)
            MatActiPasivoDet(K, 5) = CDbl(feBalanceGeneral.TextMatrix(i, 5))
        End If
             nContador = nContador + 1
    Wend
    'Fin LUCV20162606 <-**********
    
    If ValidaDatos Then
        
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    Dim MatFlujoCaja As Variant
    Set MatFlujoCaja = Nothing
    ReDim MatFlujoCaja(1, 5)
        For i = 1 To 1
            MatFlujoCaja(i, 1) = EditMoneyIncVC3
            MatFlujoCaja(i, 2) = EditMoneyIncCM3
            MatFlujoCaja(i, 3) = EditMoneyIncPP3
            MatFlujoCaja(i, 4) = EditMoneyIncGV3
            MatFlujoCaja(i, 5) = EditMoneyIncC3
        Next i
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        If Not fbImprimirVB Then
            If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        End If
        If txtUltEndeuda.Text = "__/__/____" Then
            txtUltEndeuda.Text = "01/01/1900"
        End If

        Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
        Set objPista = New COMManejador.Pista
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
        If fnTipoPermiso = 3 Then
    GrabarDatos = oNCOMFormatosEval.GrabarCredFormEvalFormato1_5(sCtaCod, fnFormato, fnTipoRegMant, _
                                                                    Trim(txtGiroNeg.Text), CInt(spnExpEmpAnio.valor), CInt(spnExpEmpMes.valor), CInt(spnTiempoLocalAnio.valor), _
                                                                    CInt(spnTiempoLocalMes.valor), CDbl(txtUltEndeuda.Text), Format(txtFecUltEndeuda.Text, "yyyymmdd"), _
                                                                    lnCondLocal, IIf(txtCondLocalOtros.Visible = False, "", txtCondLocalOtros.Text), CDbl(txtExposicionCredito.Text), _
                                                                    Format(txtFechaEvaluacion.Text, "yyyymmdd"), _
                                                                    Format(txtFechaVisita3.Text, "yyyymmdd"), _
                                                                    txtEntornoFamiliar3.Text, txtGiroUbicacion3.Text, _
                                                                    txtExperiencia3.Text, txtFormalidadNegocio3.Text, _
                                                                    txtColaterales3, txtDestino3.Text, _
                                                                    txtComentario3.Text, MatReferidos, MatIfiGastoNego, MatIfiGastoFami, _
                                                                    rsGastoFam, rsOtrosIng, rsGastoNeg, _
                                                                    CDbl(txtIngresoNegocio.Text), _
                                                                    CDbl(txtEgresoNegocio.Text), _
                                                                    CDbl(txtMargenBruto.Text), _
                                                                    MatActiPasivo, MatActiPasivoDet, , , _
                                                                    gRatioCapacidadPago, _
                                                                    CDbl(Replace(txtCapacidadNeta.Text, "%", "")), _
                                                                    gRatioEndeudamiento, _
                                                                    CDbl(Replace(txtEndeudamiento.Text, "%", "")), _
                                                                    gRatioIngresoNetoNego, _
                                                                    CDbl(txtIngresoNeto.Text), _
                                                                    gRatioExcedenteMensual, _
                                                                    CDbl(txtExcedenteMensual.Text), , , , , , , fnColocCondi, MatFlujoCaja, _
                                                                    MatBalActCorr, MatBalActNoCorr, _
                                                                    MatIfiNoSupervisadaGastoNego, MatIfiNoSupervisadaGastoFami)
                                                                    
                                                                    'MatIfiNoSupervisadaGastoNego, MatIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó
                                                                    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja MatFlujoCaja,MatBalActCorr, MatBalActNoCorr
                                                                    
        Call oDCOMFormatosEval.RecalculaIndicadoresyRatiosEvaluacion(sCtaCod)
        Set rsRatiosActual = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
        Set rsRatiosAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod)
        'JOEP20180725 ERS034-2018
            Call EmiteFormRiesgoCamCred(sCtaCod)
        'JOEP20180725 ERS034-2018
        Else
        'GrabarDatos = oNCOMFormatosEval.GrabarCredEvaluacionVerif(sCtaCod, Trim(txtVerif.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        End If
            If GrabarDatos Then
                fbGrabar = True
                'RECO20161020 ERS060-2016 **********************************************************
                Dim oNCOMColocEval As New NCOMColocEval
                'Dim lcMovNro As String 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
                'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Comentó, Anexo01 de Acta 199-2018
                
                If Not ValidaExisteRegProceso(sCtaCod, gTpoRegCtrlEvaluacion) Then
                   'lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser) 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                   'objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 3", sCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                   Call oNCOMColocEval.insEstadosExpediente(sCtaCod, "Evaluacion de Credito", lcMovNro, "", "", "", 1, 2001, gTpoRegCtrlEvaluacion)
                   Set oNCOMColocEval = Nothing
                End If
                'RECO FIN **************************************************************************
                If fnTipoRegMant = 1 Then
                    objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Evaluacion Credito Formato 3", sCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    If Not fbImprimirVB Then
                        MsgBox "Los datos se grabaron correctamente", vbInformation, "Aviso"
                    End If
                Else
                    objPista.InsertarPista gsOpeCod, lcMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Evaluacion Credito Formato 3", sCtaCod, gCodigoCuenta 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    Set objPista = Nothing 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
                    If Not fbImprimirVB Then
                        MsgBox "Los datos se actualizaron correctamente", vbInformation, "Aviso"
                    End If
                    Dim objCredito As COMDCredito.DCOMCredito
                    Set objCredito = New COMDCredito.DCOMCredito
                    Call objCredito.ActualizarEstadoxVB(ActXCodCta.NroCuenta, 1)
                End If
                
                'CTI320200110 ERS003-2020. Comentó, método sin finalidad o existencia
'                'FondoCrecerBitacora
'                Dim objFCBS_UP As COMDCredito.DCOMCredito
'                Set objFCBS_UP = New COMDCredito.DCOMCredito
'                objFCBS_UP.FondoCrecerBitacora IIf(fnTipoRegMant = 1, gCredRegistrarEvaluacionCred, gCredMantenimientoEvaluacionCred), lcMovNro, gsCodPersUser, sCtaCod, "Formato de Evaluación (Sicmac Negocio)"
'                Set objFCBS_UP = Nothing
'                'FondoCrecerBitacora
                'Fin CTI320200110
                
                'Habilita / Deshabilita Botones - Text
                If fnEstado = 2000 Then               '*****-> Si es Solicitado
                    If fnColocCondi <> 4 Then
                        Me.cmdInformeVisita.Enabled = True
                        Me.cmdVerCar.Enabled = False
                    Else
                        Me.cmdInformeVisita.Enabled = False
                        Me.cmdVerCar.Enabled = False
                    End If
                    Me.cmdGuardar.Enabled = False
                    Me.cmdImprimir.Enabled = False
                Else                                  '*****-> Sugerido +
                    Me.cmdImprimir.Enabled = True
                    Me.cmdGuardar.Enabled = False
                    If fnColocCondi <> 4 Then
                        Me.cmdVerCar.Enabled = True 'No refinanciado
                        Me.cmdInformeVisita.Enabled = True
                    Else
                        Me.cmdVerCar.Enabled = False
                        Me.cmdInformeVisita.Enabled = False
                    End If
                End If
                
                '*****->No Refinanciados (Propuesta Credito)
                    If fnColocCondi <> 4 Then
                        txtFechaVisita3.Enabled = True
                        txtEntornoFamiliar3.Enabled = True
                        txtGiroUbicacion3.Enabled = True
                        txtExperiencia3.Enabled = True
                        txtFormalidadNegocio3.Enabled = True
                        txtColaterales3.Enabled = True
                        txtDestino3.Enabled = True
                     Else
                        framePropuesta.Enabled = False
                        txtFechaVisita3.Enabled = False
                        txtEntornoFamiliar3.Enabled = False
                        txtGiroUbicacion3.Enabled = False
                        txtExperiencia3.Enabled = False
                        txtFormalidadNegocio3.Enabled = False
                        txtColaterales3.Enabled = False
                        txtDestino3.Enabled = False
                    End If
                '*****->Fin No Refinanciados
                    
                'Actualización de los Ratios
                    txtCapacidadNeta.Text = CStr(rsRatiosActual!nCapPagNeta * 100) & "%"
                    txtEndeudamiento.Text = CStr(rsRatiosActual!nEndeuPat * 100) & "%"
                    txtLiquidezCte.Text = Format(rsRatiosActual!nLiquidezCte, "#,##0.00")
                    txtRentabilidad.Text = CStr(rsRatiosActual!nRentaPatri * 100) & "%"
                    txtIngresoNeto.Text = Format(rsRatiosActual!nIngreNeto, "#,##0.00")
                    txtExcedenteMensual.Text = Format(rsRatiosActual!nExceMensual, "#,##0.00")
                    
                'Ratios: Aceptable / Critico ->*****
                    If Not (rsRatiosAceptableCritico.EOF Or rsRatiosAceptableCritico.BOF) Then
                    If rsRatiosAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
                        Me.lblCapaAceptable.Caption = "Aceptable"
                        Me.lblCapaAceptable.ForeColor = &H8000&
                    Else
                        Me.lblCapaAceptable.Caption = "Crítico"
                        Me.lblCapaAceptable.ForeColor = vbRed
                    End If
                    
                    If rsRatiosAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
                        Me.lblEndeAceptable.Caption = "Aceptable"
                        Me.lblEndeAceptable.ForeColor = &H8000&
                    Else
                        Me.lblEndeAceptable.Caption = "Crítico"
                        Me.lblEndeAceptable.ForeColor = vbRed
                    End If
                    Else
                        lblCapaAceptable.Visible = False
                        lblEndeAceptable.Visible = False
                    End If
                'Fin Ratios <-****
                    Set rsRatiosActual = Nothing
                    Set rsRatiosAceptableCritico = Nothing
            Else
                MsgBox "Hubo errores al grabar la información", vbError, "Error"
            End If
    'Else
    'MsgBox "Ha Ocurrido un Problema o Faltan Ingresar Datos", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdVerCar_Click()
    Call GeneraVerCar
End Sub
Private Sub cmdImprimir_Click()
    Call ImprimirFormatoEvaluacion
End Sub
Private Sub cmdInformeVisita_Click()
    'Call CargaInformeVisitaPDF
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
               
    cmdInformeVisita.Enabled = False
    If (rsInfVisita.EOF And rsInfVisita.BOF) Then
        Set oDCOMFormatosEval = Nothing
        MsgBox "No existe datos para este reporte.", vbOKOnly, "Atención"
        Exit Sub
    End If
    Call CargaInformeVisitaPDF(rsInfVisita) 'gCredReportes
    Set rsInfVisita = Nothing
    cmdInformeVisita.Enabled = True
End Sub
Private Sub cmdCancelar_Click()
    Unload frmCredFormEvalCuotasIfis
    Unload Me
    
    Set MatIfiNoSupervisadaGastoNego = Nothing 'CTI320200110 ERS003-2020. Agregó
    Set MatIfiNoSupervisadaGastoFami = Nothing 'CTI320200110 ERS003-2020. Agregó
End Sub
Private Sub cmdAgregarRef3_Click()
    If feReferidos3.rows - 1 < 25 Then
        feReferidos3.lbEditarFlex = True
        feReferidos3.AdicionaFila
        feReferidos3.SetFocus
        feReferidos3.AvanceCeldas = Horizontal
        SendKeys "{Enter}"
    Else
        MsgBox "No puede agregar mas de 25 registros", vbInformation, "Aviso"
    End If
End Sub
Private Sub cmdQuitar3_Click()
    If MsgBox("Esta Seguro de Eliminar Registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        feReferidos3.EliminaFila (feReferidos3.row)
    End If
End Sub

'LUCV20160620, KeyPress / GotFocus / LostFocus ->**********
    'TAB0 -> Ingresos/Egresos
Private Sub spnTiempoLocalAnio_KeyPress(KeyAscii As Integer) 'TiempoMismoLocal
    If KeyAscii = 13 Then
        spnTiempoLocalMes.SetFocus
    End If
End Sub
Private Sub spnTiempoLocalMes_KeyPress(KeyAscii As Integer) 'TiempoMismoLocal
    If KeyAscii = 13 Then
        OptCondLocal(1).SetFocus
    End If
End Sub
Private Sub OptCondLocal_KeyPress(Index As Integer, KeyAscii As Integer) 'CondicionLocal
    If KeyAscii = 13 Then
        SSTabIngresos.Tab = 0
        txtIngresoNegocio.SetFocus
    End If
End Sub
Private Sub txtCondLocalOtros_KeyPress(KeyAscii As Integer) 'OtroCondicionLocal
    If KeyAscii = 13 Then
        SSTabIngresos.Tab = 0
        txtIngresoNegocio.SetFocus
    End If
End Sub

Private Sub txtIngresoNegocio_KeyPress(KeyAscii As Integer) 'Ingresos
   KeyAscii = NumerosDecimales(txtIngresoNegocio, KeyAscii, 10, , True)
    If KeyAscii = 45 Then KeyAscii = 0
    If KeyAscii = 13 Then
        txtEgresoNegocio.SetFocus
    End If
End Sub
Private Sub txtEgresoNegocio_KeyPress(KeyAscii As Integer) 'EgresoVenta
    KeyAscii = NumerosDecimales(txtEgresoNegocio, KeyAscii, 10, , True)
    If KeyAscii = 45 Then KeyAscii = 0
    If fnProducto <> "800" Then 'CTI320200110 ERS003-2020
        If KeyAscii = 13 Then
            Me.feBalanceGeneral.SetFocus
            Me.feBalanceGeneral.row = 1
            Me.feBalanceGeneral.col = 5
            SendKeys "{F2}"
        End If
    Else
        If KeyAscii = 13 Then
            Me.feGastosNegocio.SetFocus
            Me.feGastosNegocio.row = 1
            Me.feGastosNegocio.col = 3
            SendKeys "{F2}"
        End If
    End If
End Sub

   'TAB1 ->PropuestaCredito
Private Sub txtFechaVisita3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtEntornoFamiliar3.SetFocus
        
        If Not IsDate(txtFechaVisita3) Then
            MsgBox "Verifique Dia,Mes,Año , Fecha Incorrecta", vbInformation, "Aviso"
            txtFechaVisita3.SetFocus
        End If
    End If
End Sub

Private Sub txtEntornoFamiliar3_KeyPress(KeyAscii As Integer) 'Entornofamiliar
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtGiroUbicacion3.SetFocus
    End If
End Sub
Private Sub txtGiroUbicacion3_KeyPress(KeyAscii As Integer) 'SobreGiro
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtExperiencia3.SetFocus
    End If
End Sub
Private Sub txtExperiencia3_KeyPress(KeyAscii As Integer) 'ExperienciaCrediticia
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtFormalidadNegocio3.SetFocus
    End If
End Sub
Private Sub txtFormalidadNegocio3_KeyPress(KeyAscii As Integer) 'ConsistenciaInformacion
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtColaterales3.SetFocus
    End If
End Sub
Private Sub txtColaterales3_KeyPress(KeyAscii As Integer) 'Colaterales_Garantias
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        txtDestino3.SetFocus
    End If
End Sub
Private Sub txtDestino3_KeyPress(KeyAscii As Integer) 'Destino del crédito
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        SSTabIngresos.Tab = 2
        'If fnColocCondi = 1 Then 'LUCV20171115, Agregó segun correo: RUSI
        If Not fbTieneReferido6Meses Then
            txtComentario3.SetFocus
        Else
            cmdGuardar.SetFocus
        End If
    End If
End Sub
    'TAB1 ->ComentarioReferido
Private Sub txtComentario3_KeyPress(KeyAscii As Integer) 'Referidos/ ComentariosReferidos
    KeyAscii = SoloLetras3(KeyAscii, True)
    If KeyAscii = 13 Then
        If fnColocCondi = 1 Then
            cmdAgregarRef3.SetFocus
        End If
    End If
End Sub

    'GotFocus / LostFocus
Private Sub txtIngresoNegocio_GotFocus()
    fEnfoque txtEgresoNegocio
End Sub
Private Sub txtIngresoNegocio_LostFocus()
    If Trim(txtIngresoNegocio.Text) = "" Then
        txtIngresoNegocio.Text = "0.00"
    Else
        txtIngresoNegocio.Text = Format(txtIngresoNegocio.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub
Private Sub txtEgresoNegocio_GotFocus()
 fEnfoque txtEgresoNegocio
End Sub

Private Sub txtEgresoNegocio_LostFocus()
    If Trim(txtEgresoNegocio.Text) = "" Then
        txtEgresoNegocio.Text = "0.00"
    Else
        txtEgresoNegocio.Text = Format(txtEgresoNegocio.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub
'LUCV20160620, KeyPress / GotFocus / LostFocus Fin <-**********

'Para Buscar Cuotas IFIs (GastosNegocio / GastosFamiliares)**********->
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
Private Sub feGastosNegocio_Click() 'GastosNegocio
    If feGastosNegocio.col = 3 Then
        If CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 0)) = gCodCuotaIfiGastoNego _
            Or (CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 0)) = gCodCuotaIfiNoSupervisadaGastoNego) Then 'CTI320200110 ERS003-2020, Agregó
            feGastosNegocio.ListaControles = "0-0-0-1-0"
        Else
            feGastosNegocio.ListaControles = "0-0-0-0-0"
        End If
    End If
    
        Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
            Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoNego
                'Me.feGastosNegocio.CellBackColor = &HC0FFFF
                Me.feGastosNegocio.BackColorRow &HC0FFFF, True
                Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                Me.feGastosNegocio.ForeColorRow vbBlack, True
            Case gCodCuotaCmac
                Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
                Me.feGastosNegocio.ForeColorRow vbBlack, True
            Case Else
                Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
        End Select
End Sub
Private Sub feGastosNegocio_EnterCell() 'LUCV20160525 - Me permite Buscar OtrasCuotasIFIs (GastosNegocio)
    If feGastosNegocio.col = 3 Then
        If CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 0)) = gCodCuotaIfiGastoNego _
            Or (CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 0)) = gCodCuotaIfiNoSupervisadaGastoNego) Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoNego
            feGastosNegocio.ListaControles = "0-0-0-1-0"
        Else
            feGastosNegocio.ListaControles = "0-0-0-0-0"
        End If
    End If
    
        Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
            Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoNego
                'Me.feGastosNegocio.CellBackColor = &HC0FFFF
                Me.feGastosNegocio.BackColorRow &HC0FFFF, True
                Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                Me.feGastosNegocio.ForeColorRow vbBlack, True
            Case gCodCuotaCmac
                Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
                Me.feGastosNegocio.ForeColorRow vbBlack, True
            Case Else
                Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
        End Select
End Sub
Private Sub feGastosNegocio_RowColChange() 'PresionarEnter:Monto
    If feGastosNegocio.col = 3 Then
        feGastosNegocio.AvanceCeldas = Vertical
    Else
        feGastosNegocio.AvanceCeldas = Horizontal
    End If
    
    If feGastosNegocio.col = 3 Then
        If CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 0)) = gCodCuotaIfiGastoNego _
            Or (CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 0)) = gCodCuotaIfiNoSupervisadaGastoNego) Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoNego
            feGastosNegocio.ListaControles = "0-0-0-1-0"
        Else
            feGastosNegocio.ListaControles = "0-0-0-0-0"
        End If
    End If
    
        Select Case CInt(feGastosNegocio.TextMatrix(feGastosNegocio.row, 1))
            Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoNego
                'Me.feGastosNegocio.CellBackColor = &HC0FFFF
                Me.feGastosNegocio.BackColorRow &HC0FFFF, True
                Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                Me.feGastosNegocio.ForeColorRow vbBlack, True
            Case gCodCuotaCmac
                Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
                Me.feGastosNegocio.ForeColorRow vbBlack, True
            Case Else
                Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
        End Select
End Sub
Private Sub feGastosNegocio_OnClickTxtBuscar(psMontoIfiGastoNego As String, psDescripcion As String) 'GastosNegocio
    psMontoIfiGastoNego = 0
    psDescripcion = ""
    psDescripcion = feGastosNegocio.TextMatrix(feGastosNegocio.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoNego = feGastosNegocio.TextMatrix(feGastosNegocio.row, 3) 'Monto
    
    If feGastosNegocio.TextMatrix(feGastosNegocio.row, 1) = gCodCuotaIfiGastoNego Then 'CTI320200110 ERS003-2020. Agregó
        If psMontoIfiGastoNego = 0 Then
            Set MatIfiGastoNego = Nothing
            fnTotalRefGastoNego = 0
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2), gFormatoGastosNego, gCodCuotaIfiGastoNego 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2), gFormatoGastosNego, gCodCuotaIfiGastoNego 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        End If
    Else
        If psMontoIfiGastoNego = 0 Then
            fnTotalRefGastoNego = 0
            Set MatIfiNoSupervisadaGastoNego = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiNoSupervisadaGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2), _
                                              gFormatoGastosNego, gCodCuotaIfiNoSupervisadaGastoNego 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosNegocio.TextMatrix(feGastosNegocio.row, 3))), fnTotalRefGastoNego, MatIfiNoSupervisadaGastoNego, feGastosNegocio.TextMatrix(feGastosNegocio.row, 2), _
                                              gFormatoGastosNego, gCodCuotaIfiNoSupervisadaGastoNego 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoNego = Format(fnTotalRefGastoNego, "#,##0.00")
        End If
    End If
End Sub
Private Sub feGastosNegocio_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feGastosNegocio.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feGastosNegocio.TextMatrix(pnRow, pnCol) < 0 Then
            feGastosNegocio.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feGastosNegocio.TextMatrix(pnRow, pnCol) = 0
    End If
    
    'If Me.feGastosNegocio.Col = 3 And Me.feGastosNegocio.row = 11 Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If Me.feGastosNegocio.col = 3 And Me.feGastosNegocio.row = 12 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Me.feGastosFamiliares.SetFocus
        feGastosFamiliares.row = 1
        feGastosFamiliares.col = 3
        SendKeys "{TAB}"
        SendKeys "{F2}"
    End If
    
End Sub
Private Sub feGastosFamiliares_KeyPress(KeyAscii As Integer)
        If (feGastosFamiliares.col = 1 And feGastosFamiliares.row = 1) Or (feGastosFamiliares.col = 3 And feGastosFamiliares.row = 7) Then
        If KeyAscii = 13 Then
            feOtrosIngresos.row = 1
            feOtrosIngresos.col = 3
            EnfocaControl feOtrosIngresos
            SendKeys "{Enter}", True
        End If
    End If
End Sub
Private Sub feGastosFamiliares_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feGastosFamiliares.ColumnasAEditar, "-")
    If Me.feGastosFamiliares.row <> 1 Then
        If Editar(pnCol) = "X" Then
            MsgBox "Esta celda no es editable", vbInformation, "Aviso"
            Cancel = False
            SendKeys "{TAB}", True
            Exit Sub
        End If
    End If
End Sub
Private Sub feGastosFamiliares_Click() 'GastosFamiliares
    If feGastosFamiliares.col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami _
            Or CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then 'CTI320200110 ERS003-2020, Agregó
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
    
        Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
            Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
                Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                Me.feGastosFamiliares.ForeColorRow vbBlack, True
            Case gCodDeudaLCNUGastoFami
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                Me.feGastosFamiliares.ForeColorRow vbBlack, True
            Case Else
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        End Select
End Sub
Private Sub feGastosFamiliares_EnterCell() 'LUCV20160525 - Me permite Buscar CuotasIFIs(GastosFamiliares)
    If feGastosNegocio.col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami _
            Or CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiNoSupervisadaGastoFami Then 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
    
        Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
            Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
                Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                Me.feGastosFamiliares.ForeColorRow vbBlack, True
            Case gCodDeudaLCNUGastoFami
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                Me.feGastosFamiliares.ForeColorRow vbBlack, True
            Case Else
                Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
        End Select
End Sub
Private Sub feGastosFamiliares_RowColChange() 'PresionarEnter:Monto
    If feGastosFamiliares.col = 3 Then
        feGastosFamiliares.AvanceCeldas = Vertical
    Else
        feGastosFamiliares.AvanceCeldas = Horizontal
    End If
    
    If feGastosFamiliares.col = 3 Then
        If CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiGastoFami _
            Or (CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 0)) = gCodCuotaIfiNoSupervisadaGastoFami) Then 'CTI320200110 ERS003-2020, Agregó
            feGastosFamiliares.ListaControles = "0-0-0-1-0"
        Else
            feGastosFamiliares.ListaControles = "0-0-0-0-0"
        End If
    End If
    
    Select Case CInt(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1))
        Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020, Agregó: gCodCuotaIfiNoSupervisadaGastoFami
            'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
            Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            Me.feGastosFamiliares.ForeColorRow vbBlack, True
        Case gCodDeudaLCNUGastoFami
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
            Me.feGastosFamiliares.ForeColorRow vbBlack, True
        Case Else
            Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
    End Select
End Sub
Private Sub feGastosFamiliares_OnClickTxtBuscar(psMontoIfiGastoFami As String, psDescripcion As String) 'GastosFamiliares
    psMontoIfiGastoFami = 0
    psDescripcion = ""
    psDescripcion = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2) 'Cuotas Otras IFIs
    psMontoIfiGastoFami = feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3) 'Monto
    
    If CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 1)) = gCodCuotaIfiGastoFami Then 'CTI320200110 ERS003-2020. Agregó
        If psMontoIfiGastoFami = 0 Then
            fnTotalRefGastoFami = 0
            Set MatIfiGastoFami = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), gFormatoGastosFami, gCodCuotaIfiGastoFami 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        End If
    Else
        If psMontoIfiGastoFami = 0 Then
            fnTotalRefGastoFami = 0
            Set MatIfiNoSupervisadaGastoFami = Nothing
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiNoSupervisadaGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), _
                                             gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó
            psMontoIfiGastoFami = Format(fnTotalRefGastoFami, "#,##0.00")
        Else
            frmCredFormEvalCuotasIfis.Inicio (CLng(feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 3))), fnTotalRefGastoFami, MatIfiNoSupervisadaGastoFami, feGastosFamiliares.TextMatrix(feGastosFamiliares.row, 2), _
                                            gFormatoGastosFami, gCodCuotaIfiNoSupervisadaGastoFami  'CTI320200110 ERS003-2020. Agregó
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
End Sub

Private Sub OptCondLocal_Click(Index As Integer)
    Select Case Index
    Case 1, 2, 3
        Me.txtCondLocalOtros.Visible = False
        Me.txtCondLocalOtros.Text = ""
    Case 4
        Me.txtCondLocalOtros.Visible = True
        Me.txtCondLocalOtros.Text = ""
    End Select
    lnCondLocal = Index
End Sub

'***** LUCV20160528 - OnCellChange / RowColChange
Private Sub feReferidos3_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Or pnCol = 4 Then
        feReferidos3.TextMatrix(pnRow, pnCol) = UCase(feReferidos3.TextMatrix(pnRow, pnCol))
    End If
    
    Select Case pnCol
    Case 2
        If IsNumeric(feReferidos3.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos3.TextMatrix(pnRow, pnCol)) = 8 Then
                Select Case CCur(feReferidos3.TextMatrix(pnRow, pnCol))
                    Case Is > 0
                    Case Else
                        MsgBox "Por favor, verifique el DNI", vbInformation, "Alerta"
                        feReferidos3.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "El DNI, tiene que ser 8 dígitos.", vbInformation, "Alerta"
                feReferidos3.TextMatrix(pnRow, pnCol) = 0
            End If
            
        Else
            MsgBox "El DNI, tiene que ser numérico.", vbInformation, "Alerta"
            feReferidos3.TextMatrix(pnRow, pnCol) = 0
        End If
    Case 3
        If IsNumeric(feReferidos3.TextMatrix(pnRow, pnCol)) Then
            If Len(feReferidos3.TextMatrix(pnRow, pnCol)) = 9 Then
                Select Case CCur(feReferidos3.TextMatrix(pnRow, pnCol))
                Case Is > 0
                Case Else
                    MsgBox "Teléfono Mal Ingresado", vbInformation, "Alerta"
                    feReferidos3.TextMatrix(pnRow, pnCol) = 0
                    Exit Sub
                End Select
            Else
                MsgBox "Faltan caracteres en el teléfono / celular.", vbInformation, "Alerta"
                feReferidos3.TextMatrix(pnRow, pnCol) = 0
            End If
        Else
            MsgBox "El telefono, solo permite ingreso de datos tipo numérico." & Chr(10) & "Ejemplo: 065404040, 984047523 ", vbInformation, "Alerta"
            feReferidos3.TextMatrix(pnRow, pnCol) = 0
        End If
'    Case 5
'        If IsNumeric(feReferidos3.TextMatrix(pnRow, pnCol)) Then
'            If Len(feReferidos3.TextMatrix(pnRow, pnCol)) = 8 Then
'                Select Case CCur(feReferidos3.TextMatrix(pnRow, pnCol))
'                Case Is > 0
'                Case Else
'                    MsgBox "El DNI del referido, tiene que contener 8 dígitos", vbInformation, "Alerta"
'                    feReferidos3.TextMatrix(pnRow, pnCol) = 0
'                    Exit Sub
'                End Select
'            Else
'                MsgBox "El DNI del referido, tiene que ser 8 dígitos", vbInformation, "Alerta"
'                feReferidos3.TextMatrix(pnRow, pnCol) = 0
'            End If
'        Else
'            MsgBox "El DNI del referido, sólo permite ingreso de datos tipo numérico.", vbInformation, "Alerta"
'            feReferidos3.TextMatrix(pnRow, pnCol) = 0
'        End If
    End Select
End Sub

Private Sub feReferidos3_RowColChange()
    If feReferidos3.col = 1 Then
        feReferidos3.MaxLength = "200"
    ElseIf feReferidos3.col = 2 Then
        feReferidos3.MaxLength = "8"
    ElseIf feReferidos3.col = 3 Then
        feReferidos3.MaxLength = "9"
    ElseIf feReferidos3.col = 4 Then
        feReferidos3.MaxLength = "200"
    ElseIf feReferidos3.col = 5 Then
        feReferidos3.MaxLength = "8"
    End If
End Sub
Private Sub feBalanceGeneral_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim Editar() As String
    Editar = Split(Me.feBalanceGeneral.ColumnasAEditar, "-")
    If Me.feBalanceGeneral.row <> 1 Then
        If Editar(pnCol) = "X" Then
            MsgBox "Esta celda no es editable", vbInformation, "Aviso"
            Cancel = False
            SendKeys "{TAB}", True
            Exit Sub
        End If
    End If
End Sub
Private Sub feBalanceGeneral_KeyPress(KeyAscii As Integer)
    'If feBalanceGeneral.Col = 5 And feBalanceGeneral.row = 8 Then'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If feBalanceGeneral.col = 5 And feBalanceGeneral.row = 10 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        If KeyAscii = 13 Then
            EnfocaControl feGastosNegocio
            feGastosNegocio.row = 1
            feGastosNegocio.col = 3
            SendKeys "{Enter}"
        End If
    End If
End Sub
Private Sub feBalanceGeneral_OnCellChange(pnRow As Long, pnCol As Long)
    'If pnRow = 3 Or pnRow = 6 Or pnRow = 7 Then 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If pnRow = 4 Or pnRow = 8 Or pnRow = 9 Then 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        MsgBox "No se puede Editar este Registro", vbInformation, "Aviso"
        feBalanceGeneral.TextMatrix(pnRow, pnCol) = ""
    End If
    
    If IsNumeric(feBalanceGeneral.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feBalanceGeneral.TextMatrix(pnRow, pnCol) < 0 Then
            feBalanceGeneral.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feBalanceGeneral.TextMatrix(pnRow, pnCol) = 0
    End If


    Select Case feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2)
        Case 1000, 1001
            Me.feBalanceGeneral.BackColorRow (&H80000000)
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
         Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 100, 0)
            Me.feBalanceGeneral.BackColorRow (&HC0FFFF)
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
         Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 200, 0)
            Me.feBalanceGeneral.BackColorRow (&HC0FFFF)
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 206
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.feBalanceGeneral.BackColorRow (&HFFFFFF)
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
    End Select
Call CalculoTotal(2)
Call CalculoTotal(1)
End Sub
Private Sub feBalanceGeneral_RowColChange() 'PresionarEnter:Monto
    If feBalanceGeneral.col = 5 Then
        feBalanceGeneral.AvanceCeldas = Vertical
    Else
        feBalanceGeneral.AvanceCeldas = Horizontal
    End If

'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
If fnTipoRegMant = 2 Then
        If feBalanceGeneral.col = 5 Then
                If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
                    feBalanceGeneral.ListaControles = "0-0-0-0-1"
                ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
                    feBalanceGeneral.ListaControles = "0-0-0-0-1"
                Else
                    feBalanceGeneral.ListaControles = "0-0-0-0-0"
                End If
         End If
Else
    If feBalanceGeneral.col = 5 Then
        If CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 5 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        ElseIf CInt(feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 0)) = 6 Then
            feBalanceGeneral.ListaControles = "0-0-0-0-1"
        Else
            feBalanceGeneral.ListaControles = "0-0-0-0-0"
        End If
    End If
End If
'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja

    Select Case feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2)
        Case 1000, 1001
            Me.feBalanceGeneral.BackColorRow (&H80000000)
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
         Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 100, 0)
            Me.feBalanceGeneral.BackColorRow (&HC0FFFF)
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
         Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 200, 0)
            Me.feBalanceGeneral.BackColorRow (&HC0FFFF)
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        Case 206
            Me.feBalanceGeneral.ForeColorRow vbBlack, True
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.feBalanceGeneral.BackColorRow (&HFFFFFF)
            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
    End Select
End Sub

Private Sub feOtrosIngresos_RowColChange() 'PresionarEnter:Monto
    If feOtrosIngresos.col = 3 Then
        feOtrosIngresos.AvanceCeldas = Vertical
    Else
        feOtrosIngresos.AvanceCeldas = Horizontal
    End If
End Sub
Private Sub feOtrosIngresos_OnCellChange(pnRow As Long, pnCol As Long)
    If IsNumeric(feOtrosIngresos.TextMatrix(pnRow, pnCol)) Then 'Valida valores no Negativos
        If feOtrosIngresos.TextMatrix(pnRow, pnCol) < 0 Then
            feOtrosIngresos.TextMatrix(pnRow, pnCol) = 0
        End If
    Else
        feOtrosIngresos.TextMatrix(pnRow, pnCol) = 0
    End If
    
    If Me.feOtrosIngresos.col = 3 And Me.feOtrosIngresos.row = 5 Then
        Me.SSTabIngresos.Tab = 1
        SendKeys "{TAB}"
   End If
    
End Sub
'Fin <- LUCV20160528 - OnCellChange / RowColChange *****

'________________________________________________________________________________________________________________________
'*************************************************LUCV20160525: METODOS Varios **************************************************
Public Function Inicio(ByVal psTipoRegMant As Integer, ByVal psCtaCod As String, ByVal pnFormato As Integer, ByVal pnProducto As Integer, _
                       ByVal pnSubProducto As Integer, ByVal pnMontoExpEsteCred As Double, ByVal pbImprimir As Boolean, ByVal pnEstado As Integer, _
                       Optional ByVal pbImprimirVB As Boolean = False) As Boolean
                     
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsDCredEval As ADODB.Recordset
    Dim rsDColCred As ADODB.Recordset
    Dim oTipoCam  As COMDConstSistema.NCOMTipoCambio
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    Set oTipoCam = New COMDConstSistema.NCOMTipoCambio
    fbImprimirVB = pbImprimirVB 'CTI3ERS0032020

    If psCtaCod <> -1 Then 'cCtaCod -> **********
        gsOpeCod = ""
        lcMovNro = "" 'LUCV20181220 Agregó, Anexo01 de Acta 199-2018
        nTC = oTipoCam.EmiteTipoCambio(gdFecSis, TCFijoDia)
        sCtaCod = psCtaCod
        fnTipoRegMant = psTipoRegMant
        ActXCodCta.NroCuenta = sCtaCod
        
        '(3: Analista, 2: Coordinador, 1: JefeAgencia)
        fnTipoPermiso = oNCOMFormatosEval.ObtieneTipoPermisoCredEval(gsCodCargo)  ' Obtener el tipo de Permiso, Segun Cargo
        Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
        Set rsDCredito = oDCOMFormatosEval.RecuperaSolicitudDatoBasicosEval(sCtaCod) ' Datos Basicos del Credito Solicitado
        
        If (rsDCredito!cActiGiro) = "" Then
            MsgBox "Por favor, actualizar los datos del cliente. " & Chr(13) & " (Actividad o Giro del negocio)", vbInformation, "Alerta"
            Exit Function
        End If
        
        '***** Datos básicos de cabecera de Formato
        fsGiroNego = IIf((rsDCredito!cActiGiro) = "", "", (rsDCredito!cActiGiro)) 'Giro y
        fsCliente = Trim(rsDCredito!cPersNombre)
        fnMontoIni = Trim(rsDCredito!nMonto)
        fsAnioExp = CInt(rsDCredito!nAnio)
        fnColocCondi = rsDCredito!nColocCondicion
        fbTieneReferido6Meses = rsDCredito!bTieneReferido6Meses   'Si tiene evaluacion registrada 6 meses (LUCV20171115, agregó según correo: RUSI)
        fsMesExp = CInt(rsDCredito!nMes)
        fnFechaDeudaSbs = IIf(rsDCredito!dFechaUltimaDeudaSBS = "", "__/__/____", rsDCredito!dFechaUltimaDeudaSBS)
        fnMontoDeudaSbs = Format(CCur(rsDCredito!nMontoUltimaDeudaSBS), "#,##0.00")
        fnPlazo = CInt(rsDCredito!nPlazo)
        fnProducto = pnProducto
        
        spnExpEmpAnio.valor = fsAnioExp
        spnExpEmpMes.valor = fsMesExp
        txtUltEndeuda.Text = Format(fnMontoDeudaSbs, "#,##0.00")
        txtFecUltEndeuda.Text = Format(fnFechaDeudaSbs, "dd/mm/yyyy")
        txtExposicionCredito.Text = Format(pnMontoExpEsteCred, "#,##0.00")
        txtFechaEvaluacion.Text = Format(gdFecSis, "dd/mm/yyyy")
        '***** Fin datos de cabecera
        
        Set rsDCredEval = oDCOMFormatosEval.RecuperaColocacCredEval(sCtaCod) 'Ojo: Recuperar Credito Si ha sido Registrado el Form. Eval.
        Set rsAceptableCritico = oDCOMFormatosEval.RecuperaDatosRatiosAceptableCritico(sCtaCod) 'Obtenemos Datos, Aceptable / Critico de los Ratios
        If fnTipoPermiso = 2 Then
           If rsDCredEval.RecordCount = 0 Then ' Si no hay credito registrado
                MsgBox "El analista no ha registrado la Evaluacion respectiva", vbInformation, "Aviso"
                fbPermiteGrabar = False
            Else
                fbPermiteGrabar = True
             End If
        End If
        Set rsDCredito = Nothing
        Set rsDCredEval = Nothing
        
        SSTabIngresos.Tab = 0
        fnEstado = pnEstado
        fnFormato = pnFormato
        fbPermiteGrabar = False
        fbBloqueaTodo = False
        'frameLinea.Visible = False 'Ocultar Tab->LineaCreditoAutomatica 'CTI320200110 ERS003-2020. Agregó
    Else
        MsgBox "No se ha registrado el número de cuenta del crédito a evaluar ", vbInformation, "Aviso"
    End If 'Fin CtaCod <-**********
    
    Set oDCOMFormatosEval = Nothing
    Set oTipoCam = Nothing
    Call CargaControlesInicio
    
    If fnTipoRegMant = 3 Then
        fbBloqueaTodo = True
    End If

    'Carga de Datos Segun Evento: (Registrar / Mantenimiento) *****->
    If CargaDatos Then
        If CargaControlesTipoPermiso(fnTipoPermiso, fbPermiteGrabar, fbBloqueaTodo) Then
            If fnTipoRegMant = 1 Then   'Para el Evento: "Registrar"
                If Not rsCredEval.EOF Then
                    Call Mantenimiento
                    fnTipoRegMant = 2
                Else
                    Call Registro
                    fnTipoRegMant = 1
                End If
            ElseIf fnTipoRegMant = 2 Then 'Para el Evento. "Mantenimiento"
                If rsCredEval.EOF Then
                    Call Registro
                    fnTipoRegMant = 1
                Else
                    Call Mantenimiento
                    fnTipoRegMant = 2
                End If
            ElseIf fnTipoRegMant = 3 Then  ' Para el Evento. "Consulta"
                    Call Mantenimiento
                    fnTipoRegMant = 3
            End If
        Else
            Unload Me
            Exit Function
        End If
    Else
        If CargaControlesTipoPermiso(1, False) Then
        End If
    End If
    'Fin Carga <-*****
    
    'Habilita / Deshabilita Botones - Text
        If fnEstado = 2000 Then             '*****-> Si es Solicitado
            'Me.cmdGuardar.Enabled = True
            Me.cmdImprimir.Enabled = False
            Me.cmdInformeVisita.Enabled = False
            cmdGenerarFlujoForm3.Enabled = False 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            If fnColocCondi <> 4 Then
                Me.cmdVerCar.Enabled = False
            Else
                Me.cmdVerCar.Enabled = False
            End If
        Else                                '*****-> Sugerido +
            'Me.cmdGuardar.Enabled = True
            Me.cmdImprimir.Enabled = True
            cmdGenerarFlujoForm3.Enabled = True 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            If fnColocCondi <> 4 Then
                Me.cmdVerCar.Enabled = True 'No refinanciado
                Me.cmdInformeVisita.Enabled = True
            Else
                Me.cmdVerCar.Enabled = False
                Me.cmdInformeVisita.Enabled = False
            End If
        End If
                     
        '*****->No Refinanciados (Propuesta Credito)
          If fnColocCondi <> 4 Then
              txtFechaVisita3.Enabled = True
              txtEntornoFamiliar3.Enabled = True
              txtGiroUbicacion3.Enabled = True
              txtExperiencia3.Enabled = True
              txtFormalidadNegocio3.Enabled = True
              txtColaterales3.Enabled = True
              txtDestino3.Enabled = True
           Else
              framePropuesta.Enabled = False
              txtFechaVisita3.Enabled = False
              txtEntornoFamiliar3.Enabled = False
              txtGiroUbicacion3.Enabled = False
              txtExperiencia3.Enabled = False
              txtFormalidadNegocio3.Enabled = False
              txtColaterales3.Enabled = False
              txtDestino3.Enabled = False
          End If
        '*****->Fin No Refinanciados
    
    Set rsAceptableCritico = Nothing
    fbGrabar = False
    Call CalculoTotal(1)
    If Not pbImprimir Then
        If fbImprimirVB Then
             Call Cmdguardar_Click
             cmdGuardar.Enabled = True
             fbImprimirVB = False
        End If
        Me.Show 1
    Else
        cmdImprimir_Click
    End If
    Inicio = fbGrabar
End Function

'***** LUCV20160529 / feReferidos32
Public Function ValidaDatosReferencia() As Boolean
    Dim i As Integer, j As Integer
    ValidaDatosReferencia = False
    If feReferidos3.rows - 1 < 2 Then
        MsgBox "Debe registrar por lo menos 2 referencias para continuar", vbInformation, "Aviso"
        cmdAgregarRef3.SetFocus
        ValidaDatosReferencia = False
        Exit Function
    End If
    For i = 1 To feReferidos3.rows - 1  'Verfica Tipo de Valores del DNI
        If Trim(feReferidos3.TextMatrix(i, 1)) <> "" Then
            For j = 1 To Len(Trim(feReferidos3.TextMatrix(i, 2)))
                If (Mid(feReferidos3.TextMatrix(i, 2), j, 1) < "0" Or Mid(feReferidos3.TextMatrix(i, 2), j, 1) > "9") Then
                   MsgBox "Uno de los Digitos del primer DNI de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                   feReferidos3.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next j
        End If
    Next i
    For i = 1 To feReferidos3.rows - 1  'Verfica Longitud del DNI
        If Trim(feReferidos3.TextMatrix(i, 1)) <> "" Then
            If Len(Trim(feReferidos3.TextMatrix(i, 2))) <> gnNroDigitosDNI Then
                MsgBox "Primer DNI de la fila " & i & " no es de " & gnNroDigitosDNI & " digitos", vbInformation, "Aviso"
                feReferidos3.SetFocus
                ValidaDatosReferencia = False
                Exit Function
            End If
        End If
    Next i
    For i = 1 To feReferidos3.rows - 1  'Verfica Tipo de Valores del Telefono
        If Trim(feReferidos3.TextMatrix(i, 1)) <> "" Then
            For j = 1 To Len(Trim(feReferidos3.TextMatrix(i, 3)))
                If (Mid(feReferidos3.TextMatrix(i, 3), j, 1) < "0" Or Mid(feReferidos3.TextMatrix(i, 3), j, 1) > "9") Then
                   MsgBox "Uno de los Digitos del teléfono de la fila " & i & " no es un Numero", vbInformation, "Aviso"
                   feReferidos3.SetFocus
                   ValidaDatosReferencia = False
                   Exit Function
                End If
            Next j
        End If
    Next i

    For i = 1 To feReferidos3.rows - 1 'Verfica ambos DNI que no sean iguales
        For j = 1 To feReferidos3.rows - 1
            If i <> j Then
                If feReferidos3.TextMatrix(i, 2) = feReferidos3.TextMatrix(j, 2) Then
                    MsgBox "No se puede ingresar el mismo DNI mas de una vez en los referidos", vbInformation, "Alerta"
                    ValidaDatosReferencia = False
                    Exit Function
                End If
            End If
        Next
    Next
    ValidaDatosReferencia = True
End Function

Public Function ValidaGrillas(ByVal Flex As FlexEdit) As Boolean
    Dim i As Integer
    ValidaGrillas = False
    For i = 1 To Flex.rows - 1
        If Flex.TextMatrix(i, 0) <> "" Then
            If Trim(Flex.TextMatrix(i, 1)) = "" Or Trim(Flex.TextMatrix(i, 3)) = "" Then
                ValidaGrillas = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrillas = True
End Function

Public Function ValidaDatos() As Boolean
ValidaDatos = False
Dim nIndice As Integer
Dim i As Integer
Dim lsMensajeIfi As String 'LUCV20161115
    If fnTipoPermiso = 3 Then
      '********** Para TAB:0 -> Ingresos y Egresos
        If spnTiempoLocalAnio.valor = "" Then
        MsgBox "Ingrese Tiempo en el mismo local: Años", vbInformation, "Aviso"
            ValidaDatos = False
            SSTabIngresos.Tab = 0
            Exit Function
        End If
        If spnTiempoLocalMes.valor = "" Then
        MsgBox "Ingrese Tiempo en el mismo local: Meses", vbInformation, "Aviso"
            spnTiempoLocalMes.SetFocus
            SSTabIngresos.Tab = 0
            ValidaDatos = False
            Exit Function
        End If
        If OptCondLocal(1).value = 0 And OptCondLocal(2).value = 0 And OptCondLocal(3).value = 0 And OptCondLocal(4).value = 0 Then
            MsgBox "Falta elegir la Condicion del local", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            ValidaDatos = False
            Exit Function
        End If
        If txtCondLocalOtros.Visible = True Then
            If txtCondLocalOtros.Text = "" Then
            MsgBox "Ingrese la Descripcion de la Opcion: Otro Local", vbInformation, "Aviso"
                SSTabIngresos.Tab = 0
                ValidaDatos = False
                Exit Function
            End If
        End If
        If Trim(txtGiroNeg.Text) = "" Then
            MsgBox "Falta ingresar el Giro del Negocio, Favor Actualizar los Datos del Cliente", vbInformation, "Aviso"
            SSTabIngresos.Tab = 1
            txtGiroNeg.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If Trim(txtFechaEvaluacion.Text) = "__/__/____" Then
            MsgBox "Falta Ingresar la Fecha de Evaluacion", vbInformation, "Aviso"
            SSTabIngresos.Tab = 1
            txtFechaEvaluacion.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtIngresoNegocio.Text = "" Then
            MsgBox "Falta Ingresar el Ingreso del Negocio", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            txtIngresoNegocio.SetFocus
            ValidaDatos = False
            Exit Function
        End If
        If txtEgresoNegocio.Text = "" Then
            MsgBox "Falta Ingresar el Egreso del Negocio", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            txtEgresoNegocio.SetFocus
            ValidaDatos = False
            Exit Function
        End If
    '********** Para TAB:1 -> Propuesta del Credito
        If fnColocCondi <> 4 Then 'Valida, si el credito no es refinanciado
            If Trim(txtFechaVisita3.Text) = "__/__/____" Or Not IsDate(Trim(txtFechaVisita3.Text)) Then
                MsgBox "Falta ingresar la fecha de visita o el formato de la fecha no es el correcto." & Chr(10) & " Formato: DD/MM/YYY", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtFechaVisita.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtEntornoFamiliar3.Text = "" Then
                MsgBox "Por favor Ingrese, El Entorno Familiar del Cliente o Representante", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtEntornoFamiliar3.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtGiroUbicacion3.Text = "" Then
                MsgBox "Por favor Ingrese, El Giro y la Ubicacion del Negocio", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtGiroUbicacion3.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtExperiencia3.Text = "" Then
                MsgBox "Por favor Ingrese, Sobre la Experiencia Crediticia", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtExperiencia3.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtFormalidadNegocio3.Text = "" Then
                MsgBox "Por favor Ingrese, La Formalidad del Negocio", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtFormalidadNegocio3.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtColaterales3.Text = "" Then
                MsgBox "Por favor Ingrese, Sobre las Garantias y Colaterales", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtColaterales3.SetFocus
                ValidaDatos = False
                Exit Function
            End If
            If txtDestino3.Text = "" Then
                MsgBox "Por favor Ingrese, El destino del Credito", vbInformation, "Aviso"
                SSTabIngresos.Tab = 1
                txtDestino3.SetFocus
                ValidaDatos = False
                Exit Function
            End If
        End If
            
    '********** PARA TAB2 -> Comentarios y Referidos
        'LUCV25072016->*****, Si el cliente es Nuevo -> Referente es Obligatorio
        'If fnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
        If Not fbTieneReferido6Meses Then
            frameReferido.Enabled = True
            frameComentario.Enabled = True
                For i = 0 To feReferidos3.rows - 1
                    If feReferidos3.TextMatrix(i, 0) <> "" Then
                        If Trim(feReferidos3.TextMatrix(i, 0)) = "" Or Trim(feReferidos3.TextMatrix(i, 1)) = "" _
                            Or Trim(feReferidos3.TextMatrix(i, 2)) = "" Or Trim(feReferidos3.TextMatrix(i, 3)) = "" Or Trim(feReferidos3.TextMatrix(i, 4)) = "" Then
                            MsgBox "Faltan datos en la lista de Referencias", vbInformation, "Aviso"
                            SSTabIngresos.Tab = 2
                            ValidaDatos = False
                            Exit Function
                        End If
                    End If
                Next i
        
                If ValidaDatosReferencia = False Then 'Contenido de feReferidos2: Referidos
                    SSTabIngresos.Tab = 2
                    ValidaDatos = False
                    Exit Function
                End If
                
                If txtComentario3.Text = "" Then 'Comentarios
                    MsgBox "Por favor Ingrese, Comentarios", vbInformation, "Aviso"
                    SSTabIngresos.Tab = 2
                    txtComentario3.SetFocus
                    ValidaDatos = False
                    Exit Function
                End If
                
        Else
            'si el cliente es nuevo-> referido obligatorio
                frameReferido.Enabled = False
                feReferidos3.Enabled = False
                cmdAgregarRef3.Enabled = False
                cmdQuitar3.Enabled = False
                txtComentario3.Enabled = False 'Comentarios
                frameComentario.Enabled = False
        End If
            'Fin LUCV25072016 <-*****
            
        '********** Para TAB:0 -> Validacion Grillas: GastosNegocio, OtrosIngresos, GastosFamiliares
        If ValidaGrillas(feGastosNegocio) = False Then
            MsgBox "Faltan datos en la lista de Gastos del Negocio", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            ValidaDatos = False
            Exit Function
        End If
        If ValidaGrillas(feOtrosIngresos) = False Then
            MsgBox "Faltan datos en la lista de Otros Ingresos", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            ValidaDatos = False
            Exit Function
        End If
        If ValidaGrillas(feGastosFamiliares) = False Then
            MsgBox "Faltan datos en la lista de Gastos Familiares", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            ValidaDatos = False
            Exit Function
        End If
        
        '********** Para TAB:0 -> Grilla Balance General
        If fnProducto <> "800" Then 'CTI320200110 ERS003-2020
            For nIndice = 1 To feBalanceGeneral.rows - 1
                'Activos
                If feBalanceGeneral.TextMatrix(nIndice, 2) = 1000 And feBalanceGeneral.TextMatrix(nIndice, 1) = 7025 Then 'Activo
                    If val(Replace(feBalanceGeneral.TextMatrix(nIndice, 5), ",", "")) <= 0 Then
                        MsgBox "No se ingresaron datos en el Activo", vbInformation, "Alerta"
                        ValidaDatos = False
                        SSTabIngresos.Tab = 0
                        Exit Function
                    End If
                End If
                If feBalanceGeneral.TextMatrix(nIndice, 2) = 100 And feBalanceGeneral.TextMatrix(nIndice, 1) = 7025 Then 'Activo Corriente
                    If val(Replace(feBalanceGeneral.TextMatrix(nIndice, 5), ",", "")) <= 0 Then
                        MsgBox "No se ingresaron datos en el Activo Corriente", vbInformation, "Alerta"
                        ValidaDatos = False
                        SSTabIngresos.Tab = 0
                        Exit Function
                    End If
                End If
                
                'Pasivos
                If feBalanceGeneral.TextMatrix(nIndice, 2) = 1000 And feBalanceGeneral.TextMatrix(nIndice, 1) = 7026 Then 'Pasivo
                    If val(Replace(feBalanceGeneral.TextMatrix(nIndice, 5), ",", "")) <= 0 Then
                        MsgBox "No se ingresaron datos en el pasivo", vbInformation, "Alerta"
                        ValidaDatos = False
                        SSTabIngresos.Tab = 0
                        Exit Function
                    End If
                End If
                'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                If (feBalanceGeneral.TextMatrix(nIndice, 2) = 100 Or feBalanceGeneral.TextMatrix(nIndice, 2) = 107) And feBalanceGeneral.TextMatrix(nIndice, 1) = 7026 Then  'Pasivo Corriente
                    If (val(Replace(feBalanceGeneral.TextMatrix(5, 5), ",", "")) + val(Replace(feBalanceGeneral.TextMatrix(7, 5), ",", ""))) <= 0 Then
                        MsgBox "No se ingresaron datos en el pasivo Corriente", vbInformation, "Alerta"
                        ValidaDatos = False
                        SSTabIngresos.Tab = 0
                        Exit Function
                    End If
                End If
                'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                'Patrimonio
                If feBalanceGeneral.TextMatrix(nIndice, 2) = 1001 And feBalanceGeneral.TextMatrix(nIndice, 1) = 7026 Then 'Patrimonio
                    If val(Replace(feBalanceGeneral.TextMatrix(nIndice, 5), ",", "")) <= 0 Then
                        MsgBox "Patrimonio = (Total Activo - Total Pasivo) " & Chr(10) & "- No se ingresaron datos en el patrimonio." & Chr(10) & "- El patrimonio no debe ser menor o igual que cero", vbInformation, "Alerta"
                        ValidaDatos = False
                        SSTabIngresos.Tab = 0
                        Exit Function
                    End If
                End If
            Next
        End If 'Fin CTI320200110
      
        '********** Para TAB:0 -> Valida Margen Bruto
        If nMargenBruto <= 0 Then
             MsgBox "Margen Bruto = (Ingresos del Negocio) - (Egreso por Venta)" & Chr(10) & "El Margen Bruto no debe ser menor o igual que cero.", vbInformation, "Alerta"
             ValidaDatos = False
             SSTabIngresos.Tab = 0
             Exit Function
        End If
        
        'LUCV20161115, Agregó->Según ERS068-2016
        If Not ValidaIfiExisteCompraDeuda(sCtaCod, MatIfiGastoFami, MatIfiGastoNego, lsMensajeIfi, MatIfiNoSupervisadaGastoFami, MatIfiNoSupervisadaGastoNego) Or Len(Trim(lsMensajeIfi)) > 0 Then
            MsgBox "Ifi y Cuota registrada en detalle de cambio de estructura de pasivos no coincide:  " & Chr(10) & Chr(10) & " " & lsMensajeIfi & " ", vbInformation, "Aviso"
            SSTabIngresos.Tab = 0
            Exit Function
        End If
   End If
    ValidaDatos = True
End Function

Private Function CargaControlesTipoPermiso(ByVal TipoPermiso As Integer, ByVal pPermiteGrabar As Boolean, Optional ByVal pBloqueaTodo As Boolean = False) As Boolean
    '1: JefeAgencia->
    If TipoPermiso = 1 Then
        Call HabilitaControles(False, False, False)
        CargaControlesTipoPermiso = True
     '2: Coordinador->
    ElseIf TipoPermiso = 2 Then
        Call HabilitaControles(False, False, pPermiteGrabar)
        CargaControlesTipoPermiso = True
     '3: Analista ->
    ElseIf TipoPermiso = 3 Then
        Call HabilitaControles(True, False, True)
        CargaControlesTipoPermiso = True
     'Usuario sin Permisos al formato
    Else
        MsgBox "No tiene Permisos para este módulo", vbInformation, "Aviso"
        CargaControlesTipoPermiso = False
    End If
    
    If pBloqueaTodo Then 'Para el Caso despues de dar Verificacion
        Call HabilitaControles(True, True, False)
        CargaControlesTipoPermiso = True
    End If
End Function

Private Function HabilitaControles(ByVal pbHabilitaA As Boolean, ByVal pbHabilitaRatios As Boolean, ByVal pbHabilitaGuardar As Boolean)
'HabilitacionControlesAnalistas:     pbHabilitaA = True
    'Tab0: Ingresos/Egresos
    spnTiempoLocalAnio.Enabled = pbHabilitaA
    spnTiempoLocalMes.Enabled = pbHabilitaA
    OptCondLocal(1).Enabled = pbHabilitaA
    OptCondLocal(2).Enabled = pbHabilitaA
    OptCondLocal(3).Enabled = pbHabilitaA
    OptCondLocal(4).Enabled = pbHabilitaA
    txtCondLocalOtros.Enabled = pbHabilitaA
    'txtFechaEvaluacion.Enabled = pbHabilitaA
    txtIngresoNegocio.Enabled = pbHabilitaA
    txtEgresoNegocio.Enabled = pbHabilitaA
    feGastosNegocio.Enabled = pbHabilitaA
    feBalanceGeneral.Enabled = pbHabilitaA
    feOtrosIngresos.Enabled = pbHabilitaA
    feGastosFamiliares.Enabled = pbHabilitaA

    'Tab1: Propuesta/Credito
    txtFechaVisita.Enabled = pbHabilitaA
    txtEntornoFamiliar3.Enabled = pbHabilitaA
    txtGiroUbicacion3.Enabled = pbHabilitaA
    txtExperiencia3.Enabled = pbHabilitaA
    txtFormalidadNegocio3.Enabled = pbHabilitaA
    txtColaterales3.Enabled = pbHabilitaA
    txtDestino3.Enabled = pbHabilitaA

    'Tab2: Comentarios/Referidos
    txtComentario3.Enabled = pbHabilitaA
    feReferidos3.Enabled = pbHabilitaA
    cmdAgregarRef3.Enabled = pbHabilitaA
    cmdQuitar3.Enabled = pbHabilitaA
    frameReferido.Enabled = pbHabilitaA

    'txtVerif.Enabled = pbHabilitaB
    If fnEstado = 2000 Then
        SSTabRatios.Visible = False
    Else
        SSTabRatios.Visible = pbHabilitaRatios
    End If
    
    'cmdInformeVisita.Enabled = pbHabilitaRatios
    'cmdVerCar.Enabled = pbHabilitaRatios
    'cmdImprimir.Enabled = pbHabilitaRatios
    cmdGuardar.Enabled = pbHabilitaGuardar

End Function
Private Sub CargaControlesInicio()
    Call CargarFlexEdit
    'DesHabilita la CargaInicial de Controles
    ActXCodCta.Enabled = False
    txtNombreCliente.Enabled = False
    txtExposicionCredito.Enabled = False
    txtGiroNeg.Enabled = False
    txtUltEndeuda.Enabled = False
    txtFecUltEndeuda.Enabled = False
    spnExpEmpAnio.Enabled = False
    spnExpEmpMes.Enabled = False
    txtMargenBruto.Enabled = False
    
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtRentabilidad.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
    txtLiquidezCte.Enabled = False
    
    txtIngresoNegocio.Text = "0.00"
    txtEgresoNegocio.Text = "0.00"
    SSTabRatios.Visible = False
End Sub
Private Sub CargarFlexEdit() 'Registrar New Formato Evaluacion
    Dim lnFila As Integer
    Dim CargarFlexEdit As Boolean
    Dim nMonto As Double
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    nMonto = Format(0, "00.00")
    
   CargarFlexEdit = oNCOMFormatosEval.CargaDatosFlexEdit(fnFormato, _
                                                        sCtaCod, _
                                                        rsFeGastoNeg, _
                                                        rsFeDatGastoFam, _
                                                        rsFeDatOtrosIng, _
                                                        rsFeDatBalanGen, _
                                                        rsFeDatActivos, _
                                                        rsFeDatPasivos, _
                                                        rsFeDatPasivosNo, _
                                                        rsFeDatPatrimonio, _
                                                        rsFeDatRef)
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
                Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaGastoNego
                    'Me.feGastosNegocio.CellBackColor = &HC0FFFF
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

    'Otros Ingresos
    feOtrosIngresos.Clear
    feOtrosIngresos.FormaCabecera
    feOtrosIngresos.rows = 2
    Call LimpiaFlex(feOtrosIngresos)
        Do While Not rsFeDatOtrosIng.EOF
            feOtrosIngresos.AdicionaFila
            lnFila = feOtrosIngresos.row
            feOtrosIngresos.TextMatrix(lnFila, 1) = rsFeDatOtrosIng!nConsValor
            feOtrosIngresos.TextMatrix(lnFila, 2) = rsFeDatOtrosIng!cConsDescripcion
            feOtrosIngresos.TextMatrix(lnFila, 3) = Format(rsFeDatOtrosIng!nMonto, "#,##0.00")
            rsFeDatOtrosIng.MoveNext
        Loop
    rsFeDatOtrosIng.Close
    Set rsFeDatOtrosIng = Nothing

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
                Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                   'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
                   Me.feGastosFamiliares.BackColorRow &HC0FFFF, True
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
                   Me.feGastosFamiliares.ForeColorRow vbBlack, True
                Case gCodDeudaLCNUGastoFami
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-X-X"
                   Me.feGastosFamiliares.ForeColorRow vbBlack, True
                Case Else
                   Me.feGastosFamiliares.ColumnasAEditar = "X-X-X-3-X"
            End Select
            rsFeDatGastoFam.MoveNext
        Loop
    rsFeDatGastoFam.Close
    Set rsFeDatGastoFam = Nothing
    
    'Balance General
    feBalanceGeneral.Clear
    feBalanceGeneral.FormaCabecera
    feBalanceGeneral.rows = 2
    Call LimpiaFlex(feBalanceGeneral)
        Do While Not rsFeDatBalanGen.EOF
            feBalanceGeneral.AdicionaFila
            lnFila = feBalanceGeneral.row
            feBalanceGeneral.TextMatrix(lnFila, 1) = rsFeDatBalanGen!nConsCod
            feBalanceGeneral.TextMatrix(lnFila, 2) = rsFeDatBalanGen!nConsValor
            feBalanceGeneral.TextMatrix(lnFila, 3) = rsFeDatBalanGen!nNumAut
            feBalanceGeneral.TextMatrix(lnFila, 4) = rsFeDatBalanGen!cConsDescripcion
            feBalanceGeneral.TextMatrix(lnFila, 5) = Format(rsFeDatBalanGen!nMonto, "#,##0.00")
            
           Select Case feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2)
                    Case 1000, 1001
                        Me.feBalanceGeneral.BackColorRow (&H80000000)
                        Me.feBalanceGeneral.ForeColorRow vbBlack, True
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
                   'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                     Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 100, 0)
                        Me.feBalanceGeneral.BackColorRow (&HC0FFFF)
                        Me.feBalanceGeneral.ForeColorRow vbBlack, True
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
                     Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 200, 0)
                        Me.feBalanceGeneral.BackColorRow (&HC0FFFF)
                        Me.feBalanceGeneral.ForeColorRow vbBlack, True
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
                    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                     Case 206
                        'Me.feBalanceGeneral.ForeColorRow vbBlack, True 'CTI320200110 ERS003-2020. Comentó
                        'Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X" 'CTI320200110 ERS003-2020. Comentó
                        'CTI320200110 ERS003-2020. Agregó
                        If (CDbl(feBalanceGeneral.TextMatrix(lnFila, 5)) > 0) Then
                            Me.feBalanceGeneral.ForeColorRow vbBlack, True
                            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
                        Else
                            Me.feBalanceGeneral.ForeColorRow vbBlack, True
                            Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
                            Me.feBalanceGeneral.RowHeight(lnFila) = 1
                        End If
                        'Fin CTI320200110 ERS003-2020
                     Case Else
                        Me.feBalanceGeneral.BackColorRow (&HFFFFFF)
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
             End Select
            rsFeDatBalanGen.MoveNext
        Loop
    rsFeDatBalanGen.Close
    Set rsFeDatBalanGen = Nothing
End Sub
Private Function CargaDatos() As Boolean 'Mantenimiento Formatos
On Error GoTo ErrorCargaDatos
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim i As Integer
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
 
    CargaDatos = oNCOMFormatosEval.CargaDatosCredEvaluacion2(sCtaCod, _
                                                            fnFormato, _
                                                            rsCredEval, _
                                                            rsDatGastoNeg, _
                                                            rsDatGastoFam, _
                                                            rsDatOtrosIng, _
                                                            rsDatRef, _
                                                            rsDatActivos, _
                                                            rsDatPasivos, _
                                                            rsCuotaIFIs, _
                                                            rsPropuesta, _
                                                            rsCapacPagoNeta, _
                                                            rsDatRatioInd, _
                                                            rsDatActivoPasivo, _
                                                            rsDatIfiGastoNego, _
                                                            rsDatIfiGastoFami, _
                                                            rsDatVentaCosto, , , , , , _
                                                            rsDatParamFlujoCajaForm3, _
                                                            rsDatIfiBalActCorri, _
                                                            rsDatIfiBalActNoCorri, _
                                                            rsDatIfiNoSupervisadaGastoNego, rsDatIfiNoSupervisadaGastoFami)
                    
                    'rsDatIfiNoSupervisadaGastoNego, rsDatIfiNoSupervisadaGastoFami CTI320200110 ERS003-2020. Agregó
                    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja rsDatParamFlujoCajaForm3,rsDatIfiBalActCorri,rsDatIfiBalActNoCorri
    Exit Function
ErrorCargaDatos:
    CargaDatos = False
    MsgBox Err.Description + ": Error al carga datos", vbInformation, "Error"
End Function

Private Sub CalculoTotal(ByVal pnTipo As Integer)
nMontoAct = 0
nMontoPas = 0
nMontoPat = 0
nMargenBruto = 0
On Error GoTo ErrorCalculo
Select Case pnTipo
    Case 1:
            nMargenBruto = Format(CCur((txtIngresoNegocio.Text)) - CCur(txtEgresoNegocio.Text), "###," & String(15, "#") & "#0.00")
            txtMargenBruto.Text = Format(nMargenBruto, "#,##0.00")
    Case 2:
            'Activo Total
            'For i = 1 To 2 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            For i = 1 To 3 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            nMontoAct = nMontoAct + CCur(IIf(Trim(feBalanceGeneral.TextMatrix(i, 5)) = "", 0, Trim(feBalanceGeneral.TextMatrix(i, 5))))
            'feBalanceGeneral.TextMatrix(3, 5) = Format(nMontoAct, "#,##0.00") 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            feBalanceGeneral.TextMatrix(4, 5) = Format(nMontoAct, "#,##0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Next i
            
            'Pasivo Total
            'For i = 4 To 6 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            For i = 5 To 8 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            nMontoPas = nMontoPas + CCur(IIf(Trim(feBalanceGeneral.TextMatrix(i, 5)) = "", 0, Trim(feBalanceGeneral.TextMatrix(i, 5))))
            'feBalanceGeneral.TextMatrix(7, 5) = Format(nMontoPas, "#,##0.00")'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            feBalanceGeneral.TextMatrix(9, 5) = Format(nMontoPas, "#,##0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            Next i

            'Patrimonio
            'feBalanceGeneral.TextMatrix(8, 5) = Format((nMontoAct - nMontoPas), "#,##0.00")'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            feBalanceGeneral.TextMatrix(10, 5) = Format((nMontoAct - nMontoPas), "#,##0.00") 'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
End Select

Exit Sub
ErrorCalculo:
MsgBox "Informacion: Ingrese los datos Correctamente." & Chr(13) & "Detalles de error: " & Err.Description, vbInformation, "Error"
Select Case pnTipo
    Case 1:
            txtIngresoNegocio.Text = "0.00"
            txtEgresoNegocio.Text = "0.00"
End Select
 Call CalculoTotal(pnTipo)
End Sub

Private Function Registro()
    gsOpeCod = gCredRegistrarEvaluacionCred
    txtNombreCliente.Text = fsCliente
    txtGiroNeg.Text = fsGiroNego
    cmdInformeVisita.Enabled = False
    cmdVerCar.Enabled = False
    
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtRentabilidad.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
    txtLiquidezCte.Enabled = False
    
    'si el cliente es nuevo-> referido obligatorio
    'If fnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
    If Not fbTieneReferido6Meses Then
        frameReferido.Enabled = True
        feReferidos3.Enabled = True
        cmdAgregarRef3.Enabled = True
        cmdQuitar3.Enabled = True
        frameComentario.Enabled = True 'Comentarios
        txtComentario.Enabled = True
    Else
        frameReferido.Enabled = False
        feReferidos3.Enabled = False
        cmdAgregarRef3.Enabled = False
        cmdQuitar3.Enabled = False
        txtComentario3.Enabled = False 'Comentarios
        frameComentario.Enabled = False
    End If
    
    'Ratios: Aceptable / Critico ->*****
    If Not (rsAceptableCritico.BOF Or rsAceptableCritico.EOF) Then
        If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
            Me.lblCapaAceptable.Caption = "Aceptable"
            Me.lblCapaAceptable.ForeColor = &H8000&
        Else
            Me.lblCapaAceptable.Caption = "Crítico"
            Me.lblCapaAceptable.ForeColor = vbRed
        End If
        
        If rsAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
            Me.lblEndeAceptable.Caption = "Aceptable"
            Me.lblEndeAceptable.ForeColor = &H8000&
        Else
            Me.lblEndeAceptable.Caption = "Crítico"
            Me.lblEndeAceptable.ForeColor = vbRed
        End If
    Else
        lblCapaAceptable.Visible = False
        lblCapaAceptable.Visible = False
    End If
    'Fin Ratios <-****
    
    '*****->No Refinanciados (Propuesta Credito)
    If fnColocCondi <> 4 Then
        txtFechaVisita3.Enabled = True
        txtEntornoFamiliar3.Enabled = True
        txtGiroUbicacion3.Enabled = True
        txtExperiencia3.Enabled = True
        txtFormalidadNegocio3.Enabled = True
        txtColaterales3.Enabled = True
        txtDestino3.Enabled = True
     Else
        framePropuesta.Enabled = False
        txtFechaVisita3.Enabled = False
        txtEntornoFamiliar3.Enabled = False
        txtGiroUbicacion3.Enabled = False
        txtExperiencia3.Enabled = False
        txtFormalidadNegocio3.Enabled = False
        txtColaterales3.Enabled = False
        txtDestino3.Enabled = False
    End If
    '*****->Fin No Refinanciados
    
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    If Not (rsDatParamFlujoCajaForm3.BOF And rsDatParamFlujoCajaForm3.EOF) Then
        EditMoneyIncVC3.Text = Format(rsDatParamFlujoCajaForm3!nIncVentCont, "#0.00")
        EditMoneyIncCM3.Text = Format(rsDatParamFlujoCajaForm3!nIncCompMerc, "#0.00")
        EditMoneyIncPP3.Text = Format(rsDatParamFlujoCajaForm3!nIncPagPers, "#0.00")
        EditMoneyIncGV3.Text = Format(rsDatParamFlujoCajaForm3!nIncGastvent, "#0.00")
        EditMoneyIncC3.Text = Format(rsDatParamFlujoCajaForm3!nIncConsu, "#0.00")
    End If
    rsDatParamFlujoCajaForm3.Close
    Set rsDatParamFlujoCajaForm3 = Nothing
   'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
   'CTI3 ERS0032020
   'Carga de rsDatIfiGastoNego -> Matrix
    ReDim MatIfiGastoNego(rsDatIfiGastoNego.RecordCount, 4)
    i = 0
    Do While Not rsDatIfiGastoNego.EOF
        MatIfiGastoNego(i, 0) = rsDatIfiGastoNego!nNroCuota
        MatIfiGastoNego(i, 1) = rsDatIfiGastoNego!CDescripcion
        MatIfiGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiGastoNego!nMonto), 0, rsDatIfiGastoNego!nMonto), "#,##0.00")
        rsDatIfiGastoNego.MoveNext
          i = i + 1
    Loop
    rsDatIfiGastoNego.Close
    Set rsDatIfiGastoNego = Nothing
    
    'Carga de rsDatIfiGastoFami -> Matrix
    ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
    j = 0
    Do While Not rsDatIfiGastoFami.EOF
        MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
        MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!CDescripcion
        MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#,##0.00")
        rsDatIfiGastoFami.MoveNext
        j = j + 1
    Loop
    rsDatIfiGastoFami.Close
    Set rsDatIfiGastoFami = Nothing
   
    If fnProducto = "800" Then
        fraBalanceGeneral3.Visible = False
    End If
    'Fin CTI320200110
    
End Function

Private Function Mantenimiento()
Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval

pnMontoOtrasIfisConsumo = 0
pnMontoOtrasIfisEmpresarial = 0
Dim lnFila As Integer
    If fnTipoPermiso = 3 Then
        gsOpeCod = gCredMantenimientoEvaluacionCred
    Else
        'gsOpeCod = gCredVerificacionEvaluacionCred
    End If
    
    'Para Botones *****
    If Not fbBloqueaTodo Then
        cmdInformeVisita.Enabled = False
        cmdVerCar.Enabled = False
        cmdImprimir.Enabled = False
    End If
    
    'Ver Ratios
    If fnEstado > 2000 Then
        SSTabRatios.Visible = True
    Else
        SSTabRatios.Visible = False
        cmdInformeVisita.Enabled = False
        cmdVerCar.Enabled = False
        cmdImprimir.Enabled = False
    End If
    
    'Ratios/ Indicadores
    txtCapacidadNeta.Enabled = False
    txtEndeudamiento.Enabled = False
    txtRentabilidad.Enabled = False
    txtIngresoNeto.Enabled = False
    txtExcedenteMensual.Enabled = False
    txtLiquidezCte.Enabled = False
    
    'si el cliente es nuevo-> referido obligatorio
    'If fnColocCondi = 1 Then 'LUCV2017115, Según correo: RUSI
    If Not fbTieneReferido6Meses Then
        frameReferido.Enabled = True
        feReferidos3.Enabled = True
        cmdAgregarRef3.Enabled = True
        cmdQuitar3.Enabled = True
        txtComentario3.Enabled = True 'Comentarios
        frameComentario.Enabled = True
    Else
        frameReferido.Enabled = False
        feReferidos3.Enabled = False
        cmdAgregarRef3.Enabled = False
        cmdQuitar3.Enabled = False
        txtComentario3.Enabled = False 'Comentarios
        frameComentario.Enabled = False
    End If
    
    'Ratios: Aceptable / Critico ->*****
     If Not (rsAceptableCritico.EOF Or rsAceptableCritico.BOF) Then
        If rsAceptableCritico!nCapPag = 1 Then 'Capacidad Pago
            Me.lblCapaAceptable.Caption = "Aceptable"
            Me.lblCapaAceptable.ForeColor = &H8000&
        Else
            Me.lblCapaAceptable.Caption = "Crítico"
            Me.lblCapaAceptable.ForeColor = vbRed
        End If
        
        If rsAceptableCritico!nEndeud = 1 Then 'Endeudamiento Pat.
            Me.lblEndeAceptable.Caption = "Aceptable"
            Me.lblEndeAceptable.ForeColor = &H8000&
        Else
            Me.lblEndeAceptable.Caption = "Crítico"
            Me.lblEndeAceptable.ForeColor = vbRed
        End If
    Else
        Me.lblCapaAceptable.Visible = False
        Me.lblEndeAceptable.Visible = False
    End If
    'Fin Ratios <-****
    
    '*****->No Refinanciados (Propuesta Credito)
    If fnColocCondi <> 4 Then
        txtFechaVisita3.Enabled = True
        txtEntornoFamiliar3.Enabled = True
        txtGiroUbicacion3.Enabled = True
        txtExperiencia3.Enabled = True
        txtFormalidadNegocio3.Enabled = True
        txtColaterales3.Enabled = True
        txtDestino3.Enabled = True
     Else
        framePropuesta.Enabled = False
        txtFechaVisita3.Enabled = False
        txtEntornoFamiliar3.Enabled = False
        txtGiroUbicacion3.Enabled = False
        txtExperiencia3.Enabled = False
        txtFormalidadNegocio3.Enabled = False
        txtColaterales3.Enabled = False
        txtDestino3.Enabled = False
    End If
    '*****->Fin No Refinanciados
    
    'LUCV20160626, Para CARGAR CABECERA->**********
    Set rsDCredito = oDCOMFormatosEval.RecuperaSolicitudDatoBasicosEval(sCtaCod) ' Datos Basicos del Credito Solicitado
    ActXCodCta.NroCuenta = sCtaCod
    txtGiroNeg.Text = rsCredEval!cActividad
    txtNombreCliente.Text = fsCliente
    spnExpEmpAnio.valor = rsCredEval!nExpEmpAnio
    spnExpEmpMes.valor = rsCredEval!nExpEmpMes
    spnTiempoLocalAnio.valor = rsCredEval!nTmpoLocalAnio
    spnTiempoLocalMes.valor = rsCredEval!nTmpoLocalMes
    OptCondLocal(rsCredEval!nCondiLocal).value = 1
    txtCondLocalOtros.Text = rsCredEval!cCondiLocalOtro
    txtExposicionCredito.Text = Format(rsCredEval!nExposiCred, "#,##0.00")
    txtFechaEvaluacion.Text = Format(rsCredEval!dFecEval, "dd/mm/yyyy")
    txtUltEndeuda.Text = Format(rsCredEval!nUltEndeSBS, "#,##0.00")
    txtFecUltEndeuda.Text = Format(rsCredEval!dUltEndeuSBS, "dd/mm/yyyy")
    txtComentario3.Text = Trim(rsCredEval!cComentario)
    
    txtIngresoNegocio.Text = Format(rsDatVentaCosto!nIngNegocio, "#,##0.00")
    txtEgresoNegocio.Text = Format(rsDatVentaCosto!nEgrVenta, "#,##0.00")
    txtMargenBruto.Text = Format(rsDatVentaCosto!nMargBruto, "#,##0.00")
                 
    'LUCV20160626, Para CARGAR PROPUESTA->**********
    If fnColocCondi <> 4 Then
        txtFechaVisita3.Text = Format(rsPropuesta!dFecVisita, "dd/mm/yyyy")
        txtEntornoFamiliar3.Text = Trim(rsPropuesta!cEntornoFami)
        txtGiroUbicacion3.Text = Trim(rsPropuesta!cGiroUbica)
        txtExperiencia3.Text = Trim(rsPropuesta!cExpeCrediticia)
        txtFormalidadNegocio3.Text = Trim(rsPropuesta!cFormalNegocio)
        txtColaterales3.Text = Trim(rsPropuesta!cColateGarantia)
        txtDestino3.Text = Trim(rsPropuesta!cDestino)
    End If
    'LUCV20160626, Para la CARGAR FLEX - Mantenimiento **********->
        If Not (rsDatIfiGastoFami.BOF Or rsDatIfiGastoFami.EOF) Then
            For i = 1 To rsDatIfiGastoFami.RecordCount
               pnMontoOtrasIfisConsumo = pnMontoOtrasIfisConsumo + rsDatIfiGastoFami!nMonto
               rsDatIfiGastoFami.MoveNext
            Next i
            rsDatIfiGastoFami.MoveFirst
        End If
        If Not (rsDatIfiGastoNego.BOF Or rsDatIfiGastoNego.EOF) Then
            For i = 1 To rsDatIfiGastoNego.RecordCount
               pnMontoOtrasIfisEmpresarial = pnMontoOtrasIfisEmpresarial + rsDatIfiGastoNego!nMonto
               rsDatIfiGastoNego.MoveNext
            Next i
            rsDatIfiGastoNego.MoveFirst
        End If
    'Call FormatearGrillas(feGastosNegocio2)
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
                Case gCodCuotaIfiGastoNego, gCodCuotaIfiNoSupervisadaGastoNego 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaGastoNego
                    'Me.feGastosNegocio.CellBackColor = &HC0FFFF
                    Me.feGastosNegocio.BackColorRow &HC0FFFF, True
                    Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
                    Me.feGastosNegocio.ForeColorRow vbBlack, True
                Case gCodCuotaCmac
                    Me.feGastosNegocio.ColumnasAEditar = "X-X-X-X-X"
                    Me.feGastosNegocio.ForeColorRow vbBlack, True
                Case Else
                    Me.feGastosNegocio.ColumnasAEditar = "X-X-X-3-X"
            End Select
            rsDatGastoNeg.MoveNext
        Loop
    rsDatGastoNeg.Close
    Set rsDatGastoNeg = Nothing
    
    'Call FormatearGrillas(feGastosFamiliares2)
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
                Case gCodCuotaIfiGastoFami, gCodCuotaIfiNoSupervisadaGastoFami 'CTI320200110 ERS003-2020. Agregó: gCodCuotaIfiNoSupervisadaGastoFami
                   'Me.feGastosFamiliares.CellBackColor = &HC0FFFF
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
    
    'Call FormatearGrillas(feOtrosIngresos2)
    Call LimpiaFlex(feOtrosIngresos)
        Do While Not rsDatOtrosIng.EOF
            feOtrosIngresos.AdicionaFila
            lnFila = feOtrosIngresos.row
            feOtrosIngresos.TextMatrix(lnFila, 1) = rsDatOtrosIng!nConsValor
            feOtrosIngresos.TextMatrix(lnFila, 2) = rsDatOtrosIng!cConsDescripcion
            feOtrosIngresos.TextMatrix(lnFila, 3) = Format(rsDatOtrosIng!nMonto, "#,##0.00")
            rsDatOtrosIng.MoveNext
        Loop
    rsDatOtrosIng.Close
    Set rsDatOtrosIng = Nothing
    
    'Call FormatearGrillas(feCuotaIfis)
    Call LimpiaFlex(frmCredFormEvalCuotasIfis.feCuotaIfis)
        Do While Not rsCuotaIFIs.EOF
            frmCredFormEvalCuotasIfis.feCuotaIfis.AdicionaFila
            lnFila = frmCredFormEvalCuotasIfis.feCuotaIfis.row
            frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 1) = rsCuotaIFIs!CDescripcion
            frmCredFormEvalCuotasIfis.feCuotaIfis.TextMatrix(lnFila, 2) = Format(rsCuotaIFIs!nMonto, "#,##0.00")
            rsCuotaIFIs.MoveNext
        Loop
    rsCuotaIFIs.Close
    Set rsCuotaIFIs = Nothing
    
    'Call FormatearGrillas(feReferidos32)
    Call LimpiaFlex(feReferidos3)
        Do While Not rsDatRef.EOF
            feReferidos3.AdicionaFila
            lnFila = feReferidos3.row
            feReferidos3.TextMatrix(lnFila, 0) = rsDatRef!nCodRef
            feReferidos3.TextMatrix(lnFila, 1) = rsDatRef!cNombre
            feReferidos3.TextMatrix(lnFila, 2) = rsDatRef!cDniNom
            feReferidos3.TextMatrix(lnFila, 3) = rsDatRef!cTelf
            feReferidos3.TextMatrix(lnFila, 4) = rsDatRef!cReferido
            feReferidos3.TextMatrix(lnFila, 5) = rsDatRef!cDNIRef
            rsDatRef.MoveNext
        Loop
    rsDatRef.Close
    Set rsDatRef = Nothing
    
    'Call FormatearGrillas(feBalanceGeneral2)
    Call LimpiaFlex(feBalanceGeneral)
        Do While Not rsDatActivoPasivo.EOF
            feBalanceGeneral.AdicionaFila
            lnFila = feBalanceGeneral.row
            feBalanceGeneral.TextMatrix(lnFila, 1) = rsDatActivoPasivo!nConsCod
            'feBalanceGeneral.TextMatrix(lnFila, 2) = rsDatActivoPasivo!nConsValor 'Comento JOEP20171015 Segun ERS051-2017 Flujo de Caja
            
        'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Cajaa
            If rsDatActivoPasivo!nConsValor = 201 Then
                feBalanceGeneral.TextMatrix(lnFila, 2) = 107
            Else
                feBalanceGeneral.TextMatrix(lnFila, 2) = rsDatActivoPasivo!nConsValor
            End If
       'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
            
            feBalanceGeneral.TextMatrix(lnFila, 3) = rsDatActivoPasivo!nNumAut
            feBalanceGeneral.TextMatrix(lnFila, 4) = rsDatActivoPasivo!cConsDescripcion
            feBalanceGeneral.TextMatrix(lnFila, 5) = Format(rsDatActivoPasivo!nTotal, "#,##0.00")
                    
           Select Case feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2)
                Case 1000, 1001
                    Me.feBalanceGeneral.BackColorRow (&H80000000)
                    Me.feBalanceGeneral.ForeColorRow vbBlack, True
                    Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
            'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                 Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 100 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 100, 0)
                    Me.feBalanceGeneral.BackColorRow (&HC0FFFF)
                    Me.feBalanceGeneral.ForeColorRow vbBlack, True
                    Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
                 Case IIf((feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 2) = 200 And feBalanceGeneral.TextMatrix(feBalanceGeneral.row, 1) = 7026), 200, 0)
                    Me.feBalanceGeneral.BackColorRow (&HC0FFFF)
                    Me.feBalanceGeneral.ForeColorRow vbBlack, True
                    Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
            'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
                 Case 206
                    'Me.feBalanceGeneral.ForeColorRow vbBlack, True 'CTI320200110 ERS003-2020. Comentó
                    'Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X" 'CTI320200110 ERS003-2020. Comentó
                    'CTI320200110 ERS003-2020. Agregó
                    If (CDbl(feBalanceGeneral.TextMatrix(lnFila, 5)) > 0) Then
                        Me.feBalanceGeneral.ForeColorRow vbBlack, True
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
                    Else
                        Me.feBalanceGeneral.ForeColorRow vbBlack, True
                        Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-X-X"
                        Me.feBalanceGeneral.RowHeight(lnFila) = 1
                    End If
                    'Fin CTI320200110 ERS003-2020
                 Case Else
                    Me.feBalanceGeneral.BackColorRow (&HFFFFFF)
                    Me.feBalanceGeneral.ColumnasAEditar = "X-X-X-X-X-5-X"
            End Select
            rsDatActivoPasivo.MoveNext
        Loop
    rsDatActivoPasivo.Close
    Set rsDatActivoPasivo = Nothing
    'LUCV20160626, Fin Carga Flex <-**********
    
        'Carga de rsDatIfiGastoNego -> Matrix
        ReDim MatIfiGastoNego(rsDatIfiGastoNego.RecordCount, 4)
        i = 0
        Do While Not rsDatIfiGastoNego.EOF
            MatIfiGastoNego(i, 0) = rsDatIfiGastoNego!nNroCuota
            MatIfiGastoNego(i, 1) = rsDatIfiGastoNego!CDescripcion
            MatIfiGastoNego(i, 2) = Format(IIf(IsNull(rsDatIfiGastoNego!nMonto), 0, rsDatIfiGastoNego!nMonto), "#,##0.00")
            rsDatIfiGastoNego.MoveNext
              i = i + 1
        Loop
        rsDatIfiGastoNego.Close
        Set rsDatIfiGastoNego = Nothing

        'Carga de rsDatIfiGastoFami -> Matrix
        ReDim MatIfiGastoFami(rsDatIfiGastoFami.RecordCount, 4)
        j = 0
        Do While Not rsDatIfiGastoFami.EOF
            MatIfiGastoFami(j, 0) = rsDatIfiGastoFami!nNroCuota
            MatIfiGastoFami(j, 1) = rsDatIfiGastoFami!CDescripcion
            MatIfiGastoFami(j, 2) = Format(IIf(IsNull(rsDatIfiGastoFami!nMonto), 0, rsDatIfiGastoFami!nMonto), "#,##0.00")
            rsDatIfiGastoFami.MoveNext
            j = j + 1
        Loop
        rsDatIfiGastoFami.Close
        Set rsDatIfiGastoFami = Nothing
    
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        'Carga de rsDatIfiBalActCorri -> Matrix
            ReDim MatBalActCorr(rsDatIfiBalActCorri.RecordCount, 4)
            i = 0
            Do While Not rsDatIfiBalActCorri.EOF
                MatBalActCorr(i, 0) = rsDatIfiBalActCorri!nNroCuota
                MatBalActCorr(i, 1) = rsDatIfiBalActCorri!CDescripcion
                MatBalActCorr(i, 2) = Format(IIf(IsNull(rsDatIfiBalActCorri!nMonto), 0, rsDatIfiBalActCorri!nMonto), "#,##0.00")
                rsDatIfiBalActCorri.MoveNext
                  i = i + 1
            Loop
            rsDatIfiBalActCorri.Close
            Set rsDatIfiBalActCorri = Nothing
            
        'Carga de rsDatIfiBalActNoCorri -> Matrix
            ReDim MatBalActNoCorr(rsDatIfiBalActNoCorri.RecordCount, 4)
            i = 0
            Do While Not rsDatIfiBalActNoCorri.EOF
                MatBalActNoCorr(i, 0) = rsDatIfiBalActNoCorri!nNroCuota
                MatBalActNoCorr(i, 1) = rsDatIfiBalActNoCorri!CDescripcion
                MatBalActNoCorr(i, 2) = Format(IIf(IsNull(rsDatIfiBalActNoCorri!nMonto), 0, rsDatIfiBalActNoCorri!nMonto), "#,##0.00")
                rsDatIfiBalActNoCorri.MoveNext
                  i = i + 1
            Loop
            rsDatIfiBalActNoCorri.Close
            Set rsDatIfiBalActNoCorri = Nothing
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
    'CTI320200110 ERS003-2020. Agregó:
        '(Carga de rsDatIfiNoSupervisadaGastoNego -> Matrix)
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
        
        'Carga de rsDatIfiNoSupervisadaGastoFami -> Matrix
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
        'Fin CTI320200110 ERS003-2020
    
    'LUCV20160628, Para CARGA RATIOS/INDICADORES
    txtCapacidadNeta.Text = CStr(rsDatRatioInd!nCapPagNeta * 100) & "%"
    txtEndeudamiento.Text = CStr(rsDatRatioInd!nEndeuPat * 100) & "%"
    txtLiquidezCte.Text = Format(rsDatRatioInd!nLiquidezCte, "#,##0.00")
    txtRentabilidad.Text = CStr(rsDatRatioInd!nRentaPatri * 100) & "%"
    txtIngresoNeto.Text = Format(rsDatRatioInd!nIngreNeto, "#,##0.00")
    txtExcedenteMensual.Text = Format(rsDatRatioInd!nExceMensual, "#,##0.00")
    Set rsDCredito = Nothing
    
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
        If Not (rsDatParamFlujoCajaForm3.BOF And rsDatParamFlujoCajaForm3.EOF) Then
            EditMoneyIncVC3.Text = Format(rsDatParamFlujoCajaForm3!nIncVentCont, "#0.00")
            EditMoneyIncCM3.Text = Format(rsDatParamFlujoCajaForm3!nIncCompMerc, "#0.00")
            EditMoneyIncPP3.Text = Format(rsDatParamFlujoCajaForm3!nIncPagPers, "#0.00")
            EditMoneyIncGV3.Text = Format(rsDatParamFlujoCajaForm3!nIncGastvent, "#0.00")
            EditMoneyIncC3.Text = Format(rsDatParamFlujoCajaForm3!nIncConsu, "#0.00")
        End If
        rsDatParamFlujoCajaForm3.Close
        Set rsDatParamFlujoCajaForm3 = Nothing
    'Agrego JOEP20171015 Segun ERS051-2017 Flujo de Caja
    
    
    'CTI320200110 ERS003-2020
        If (fnProducto = "800") And CDbl(rsDatRatioInd!nEndeuPat) <= 0 And CDbl(rsDatRatioInd!nLiquidezCte) <= 0 And CDbl(rsDatRatioInd!nRentaPatri) <= 0 Then
            'Balance
            fraBalanceGeneral3.Visible = False
            'Ratios
            Me.lblEndeudamiento.Visible = False
            Me.txtEndeudamiento.Visible = False
            Me.lblEndeAceptable.Visible = False
            
            Me.lblRentabilidad.Visible = False
            Me.txtRentabilidad.Visible = False
            Me.lblLiquidez.Visible = False
            Me.txtLiquidezCte.Visible = False
            Me.Line1.Visible = False
        End If
    'Fin CTI320200110
End Function

Private Sub GeneraVerCar()
    Dim oCred As COMNCredito.NCOMFormatosEval
    Dim oDCredSbs As COMDCredito.DCOMFormatosEval
    Dim R As ADODB.Recordset
    Dim lcDNI, lcRUC As String
    Dim RSbs, RDatFin1, RCap As ADODB.Recordset
    Set oCred = New COMNCredito.NCOMFormatosEval
    
    Call oCred.RecuperaDatosInformeComercial(ActXCodCta.NroCuenta, R)
    Set oCred = Nothing
    
    If R.EOF And R.BOF Then
        MsgBox "No existen Datos para el Reporte...", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lcDNI = Trim(R!dni_deudor)
    lcRUC = Trim(R!ruc_deudor)
    
    Set oDCredSbs = New COMDCredito.DCOMFormatosEval
    Set RSbs = oDCredSbs.RecuperaCaliSbs(lcDNI, lcRUC)
    Set RDatFin1 = oDCredSbs.RecuperaDatosFinan(ActXCodCta.NroCuenta, fnFormato)
    Set oDCredSbs = Nothing
    Call ImprimeInformeCriteriosAceptacionRiesgoFormatoEval(ActXCodCta.NroCuenta, gsNomAge, gsCodUser, R, RSbs, RDatFin1)
End Sub

Private Sub ImprimirFormatoEvaluacion()
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim rsInfVisita As ADODB.Recordset
    
    Dim rsMostrarCuotasIfis As ADODB.Recordset
    Dim rsMostrarCuotasIfisGF As ADODB.Recordset
    Dim rsRatiosIndicadores As ADODB.Recordset
    Dim rsIngresoEgreso As ADODB.Recordset
    
    Dim oDoc  As cPDF
    Dim psCtaCod As String
    Set oDoc = New cPDF
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Set rsInfVisita = New ADODB.Recordset
    'Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
    Set rsInfVisita = oDCOMFormatosEval.MostrarFormatoSinConvenioInfVisCabecera(sCtaCod, fnFormato)
    
    Set rsMostrarCuotasIfis = oDCOMFormatosEval.MostrarCuotasIfis(sCtaCod, fnFormato, 7022)
    Set rsMostrarCuotasIfisGF = oDCOMFormatosEval.MostrarCuotasIfis(sCtaCod, fnFormato, 7023)
    Set rsRatiosIndicadores = oDCOMFormatosEval.RecuperaDatosRatios(sCtaCod)
    Set rsIngresoEgreso = oDCOMFormatosEval.RecuperaDatosCredEvalVentaCosto(sCtaCod)
    
    Dim a As Currency
    Dim nFila As Integer

    'Creación del Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Informe de Visita Nº " & sCtaCod
    oDoc.Title = "Informe de Visita Nº " & sCtaCod
    
    If Not oDoc.PDFCreate(App.Path & "\Spooler\FormatoEvaluacion_" & sCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
        Exit Sub
    End If
    
    'Contenido
    oDoc.Fonts.Add "F1", "Courier New", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Courier New", TrueType, Bold, WinAnsiEncoding
    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
        
    If Not (rsInfVisita.BOF Or rsInfVisita.EOF) Then

    'Tamaño de hoja A4
    oDoc.NewPage A4_Vertical

    'Call CabeceraImpCuadros(rsInfVisita)

        '---------- cabecera
    oDoc.WImage 45, 45, 45, 113, "Logo"
    oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft

    oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
    oDoc.WTextBox 60, 450, 10, 410, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
    oDoc.WTextBox 70, 450, 10, 490, "ANALISTA: " & UCase(Trim(rsInfVisita!cUser)), "F1", 7.5, hLeft
      
    oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
    oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
    oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
    oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
    oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
    oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft

    nFila = 110
    nFila = nFila + 10
    
    '*****-> LUCV20160913
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "VENTAS Y COSTOS", "F2", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "Ingresos", "F1", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, Format(rsIngresoEgreso!nIngNegocio, "#,##0.00"), "F1", 7.5, hRight
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "Egresos", "F1", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, Format(rsIngresoEgreso!nEgrVenta, "#,##0.00"), "F1", 7.5, hRight
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "Margen Bruto", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, Format(rsIngresoEgreso!nMargBruto, "#,##0.00"), "F2", 7.5, hRight
    nFila = nFila + 10
    '<-***** Fin LUCV20160913
    
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "GASTOS DEL NEGOCIO", "F2", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
    
    a = 0
    For i = 1 To feGastosNegocio.rows - 1
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 15, 250, feGastosNegocio.TextMatrix(i, 2), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(feGastosNegocio.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
        a = a + feGastosNegocio.TextMatrix(i, 3)
    Next i
    nFila = nFila + 10
    oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    
    
    oDoc.WTextBox nFila, 55, 1, 160, "GASTO DE NEGOCIO - CUOTAS IFIS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        a = 0
        If Not (rsMostrarCuotasIfis.BOF And rsMostrarCuotasIfis.EOF) Then
            For i = 1 To rsMostrarCuotasIfis.RecordCount
                'oDoc.WTextBox nFila, 55, 1, 160, rsMostrarCuotasIfis!nNroCuota, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 55, 1, 300, rsMostrarCuotasIfis!CDescripcion, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 140, 1, 160, Format(rsMostrarCuotasIfis!nMonto, "#,##0.00"), "F1", 7.5, hRight
                a = a + rsMostrarCuotasIfis!nMonto
                rsMostrarCuotasIfis.MoveNext
                nFila = nFila + 10
            Next i
            'nFila = nFila + 10
                oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
         End If
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
    
    '--------------------------------------------------------------------------------------------------------------------------
    
    oDoc.WTextBox nFila, 55, 1, 160, "OTROS INGRESOS", "F2", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
    
    a = 0
    For i = 1 To Me.feOtrosIngresos.rows - 1
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 15, 250, feOtrosIngresos.TextMatrix(i, 2), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(feOtrosIngresos.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
        a = a + feOtrosIngresos.TextMatrix(i, 3)
    Next i
    nFila = nFila + 10
    oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES", "F2", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
    
    a = 0
    For i = 1 To Me.feGastosFamiliares.rows - 1
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 15, 250, feGastosFamiliares.TextMatrix(i, 2), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(feGastosFamiliares.TextMatrix(i, 3), "#,#0.00"), "F1", 7.5, hRight
        a = a + feGastosFamiliares.TextMatrix(i, 3)
    Next i
    nFila = nFila + 10
    oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    
    
    oDoc.WTextBox nFila, 55, 1, 160, "GASTOS FAMILIARES  - CUOTAS IFIS", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        
        a = 0
        If Not (rsMostrarCuotasIfisGF.BOF And rsMostrarCuotasIfisGF.EOF) Then
            For i = 1 To rsMostrarCuotasIfisGF.RecordCount
                'oDoc.WTextBox nFila, 55, 1, 160, rsMostrarCuotasIfisGF!nNroCuota, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 55, 1, 300, rsMostrarCuotasIfisGF!CDescripcion, "F1", 7.5, hjustify
                oDoc.WTextBox nFila, 140, 1, 160, Format(rsMostrarCuotasIfisGF!nMonto, "#,##0.00"), "F1", 7.5, hRight
                a = a + rsMostrarCuotasIfisGF!nMonto
                nFila = nFila + 10
                rsMostrarCuotasIfisGF.MoveNext
            Next i
            'nFila = nFila + 10
                oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
         End If
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
    
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.WTextBox nFila, 55, 1, 160, "BALANCE GENERAL", "F2", 7.5, hjustify
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    oDoc.WTextBox nFila, 55, 1, 160, "CONCEPTO", "F2", 7.5, hjustify
    oDoc.WTextBox nFila, 140, 1, 160, "MONTO", "F2", 7.5, hRight
    
    a = 0
    For i = 1 To Me.feBalanceGeneral.rows - 1
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 15, 250, feBalanceGeneral.TextMatrix(i, 4), "F1", 7.5, hLeft
        oDoc.WTextBox nFila, 150, 15, 150, Format(feBalanceGeneral.TextMatrix(i, 5), "#,#0.00"), "F1", 7.5, hRight
        a = a + feBalanceGeneral.TextMatrix(i, 5)
    Next i
    nFila = nFila + 10
    oDoc.WTextBox nFila, 140, 1, 160, "TOTAL" & Space(10) & Format(a, "#,##0.00"), "F2", 7.5, hRight
    nFila = nFila + 10
    oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
    nFila = nFila + 10
    '--------------------------------------------------------------------------------------------------------------------------
    
    If nFila >= 770 Then
    
        'Tamaño de hoja A4
        oDoc.NewPage A4_Vertical
        
        oDoc.WImage 45, 45, 45, 113, "Logo"
        oDoc.WTextBox 40, 60, 35, 390, UCase(rsInfVisita!cAgeDescripcion), "F2", 7.5, hLeft
    
        oDoc.WTextBox 40, 60, 35, 490, "FECHA: " & Format(gdFecSis, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss"), "F1", 7.5, hRight
        oDoc.WTextBox 60, 450, 10, 410, "USUARIO: " & Trim(gsCodUser), "F1", 7.5, hLeft
        oDoc.WTextBox 70, 450, 10, 490, "ANALISTA: " & UCase(Trim(rsInfVisita!cUser)), "F1", 7.5, hLeft
          
        oDoc.WTextBox 80, 100, 10, 400, "HOJA DE EVALUACION", "F2", 10, hCenter
        oDoc.WTextBox 90, 55, 10, 300, "CODIGO CUENTA: " & Trim(rsInfVisita!cCtaCod), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 55, 10, 300, "CODIGO CLIENTE: " & Trim(rsInfVisita!cPersCod), "F1", 7.5, hLeft
        oDoc.WTextBox 110, 55, 10, 300, "CLIENTE: " & Trim(rsInfVisita!cPersNombre), "F1", 7.5, hLeft
        oDoc.WTextBox 100, 450, 10, 200, "DNI: " & Trim(rsInfVisita!cPersDni) & "   ", "F1", 7.5, hLeft
        oDoc.WTextBox 110, 450, 10, 200, "RUC: " & Trim(IIf(rsInfVisita!cPersRuc = "-", Space(11), rsInfVisita!cPersRuc)), "F1", 7.5, hLeft
        
        nFila = 110
        
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de Pago", "F1", 7.5, hjustify
        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Pat.", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Corriente", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
            oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
            oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
        Else
             oDoc.WTextBox nFila + 10, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
             oDoc.WTextBox nFila + 20, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
             oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
             oDoc.WTextBox nFila + 10, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
             oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
             oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
        End If
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Pat.", "F1", 7.5, hjustify
'            oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Corriente", "F1", 7.5, hjustify
'            oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
'        End If
'        oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
'        oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
'
'        oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
'            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
'            oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
'        End If
'        oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
'        oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
'
'        'Para el caso que sea Creditos Agropecuarios
'        'If fnPlazo > 30 And (fnProducto = 601 Or fnProducto = 602) Then
'            oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
'          If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
'          End If
        'End If

    Else
        
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "RATIOS E INDICADORES", "F2", 7.5, hjustify
        nFila = nFila + 10
        oDoc.WTextBox nFila, 50, 1, 500, "--------------------------------------------------------------------------------------------------", "F1", 7.5, hLeft
        nFila = nFila + 10
        oDoc.WTextBox nFila, 55, 1, 160, "Capacidad de Pago", "F1", 7.5, hjustify
        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Pat.", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Corriente", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
            oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
            oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
            oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
            oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
            oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
        Else
             oDoc.WTextBox nFila + 10, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
             oDoc.WTextBox nFila + 20, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
             oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
             oDoc.WTextBox nFila + 10, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
             oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
             oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
        End If
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 55, 1, 160, "Endeudamiento Pat.", "F1", 7.5, hjustify
'            oDoc.WTextBox nFila + 20, 55, 1, 160, "Liquidez Corriente", "F1", 7.5, hjustify
'            oDoc.WTextBox nFila + 30, 55, 1, 160, "Rentabilidad Pat.", "F1", 7.5, hjustify
'        End If
'        oDoc.WTextBox nFila + 40, 55, 1, 160, "Ingreso Neto Empresarial", "F1", 7.5, hjustify
'        oDoc.WTextBox nFila + 50, 55, 1, 160, "Excedente", "F1", 7.5, hjustify
'
'        oDoc.WTextBox nFila, 150, 15, 150, CStr(rsRatiosIndicadores!nCapPagNeta * 100) & "%", "F1", 7.5, hRight
'        If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'            oDoc.WTextBox nFila + 10, 150, 15, 150, CStr(rsRatiosIndicadores!nEndeuPat * 100) & "%", "F1", 7.5, hRight
'            oDoc.WTextBox nFila + 20, 150, 15, 150, Format(rsRatiosIndicadores!nLiquidezCte, "#,#0.00"), "F1", 7.5, hRight
'            oDoc.WTextBox nFila + 30, 150, 15, 150, CStr(rsRatiosIndicadores!nRentaPatri * 100) & "%", "F1", 7.5, hRight
'        End If
'        oDoc.WTextBox nFila + 40, 150, 15, 150, Format(rsRatiosIndicadores!nIngreNeto, "#,#0.00"), "F1", 7.5, hRight
'        oDoc.WTextBox nFila + 50, 150, 15, 150, Format(rsRatiosIndicadores!nExceMensual, "#,#0.00"), "F1", 7.5, hRight
'
'        'Para el caso que sea Creditos Agropecuarios
'        '   If fnPlazo > 30 And (fnProducto = 601 Or fnProducto = 602) Then
'                oDoc.WTextBox nFila, 320, 1, 250, "EN RELACION A SU EXCEDENTE", "F1", 7.5, hLeft
'         If Not (Left(rsInfVisita!cTpoProdCod, 1) = "7" Or Left(rsInfVisita!cTpoProdCod, 1) = "8") Then
'                oDoc.WTextBox nFila + 10, 320, 1, 250, "EN RELACION A SU PATRIMONIO TOTAL", "F1", 7.5, hLeft
'         End If
        '    End If
    End If
    '--------------------------------------------------------------------------------------------------------------------------
    oDoc.PDFClose
    oDoc.Show
    Else
        MsgBox "Los Datos de la propuesta del Credito no han sido Registrados Correctamente", vbInformation, "Aviso"
    End If
    Set rsInfVisita = Nothing
End Sub

'Private Sub CargaInformeVisitaPDF()
'    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
'    Dim rsInfVisita As ADODB.Recordset
'    Dim oDoc  As cPDF
'    Dim psCtaCod As String
'
'    Set oDoc = New cPDF
'    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
'    Set rsInfVisita = New ADODB.Recordset
'    Set rsInfVisita = oDCOMFormatosEval.RecuperarDatosInformeVisitaFormato1_6(sCtaCod)
'    Dim A As Integer
'    Dim B As Integer
'    Dim nContador As Integer
'    Dim nFilaAdicional As Integer
'    A = 50
'    B = 29
'
'    'Creación del Archivo
'    oDoc.Author = gsCodUser
'    oDoc.Creator = "SICMACT - Negocio"
'    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
'    oDoc.Subject = "Informe de Visita Nº " & sCtaCod
'    oDoc.Title = "Informe de Visita Nº " & sCtaCod
'
'    If Not oDoc.PDFCreate(App.Path & "\Spooler\FormatoEvaluacion_" & sCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
'        Exit Sub
'    End If
'
'    'Contenido
'    oDoc.Fonts.Add "F1", "Arial", TrueType, Normal, WinAnsiEncoding
'    oDoc.Fonts.Add "F2", "Arial", TrueType, Bold, WinAnsiEncoding
'
'    oDoc.LoadImageFromFile App.Path & "\logo_cmacmaynas.bmp", "Logo"
'
'    If Not (rsInfVisita.BOF Or rsInfVisita.EOF) Then
'
'    'Tamaño de hoja A4
'    oDoc.NewPage A4_Vertical
'
'    oDoc.WImage 40, 60, 35, 105, "Logo"
'    oDoc.WTextBox 40, 60, 35, 390, rsInfVisita!Agencia, "F2", 10, hLeft
'
'    oDoc.WTextBox 40, 60, 35, 390, "FECHA", "F2", 10, hRight
'    oDoc.WTextBox 40, 60, 35, 450, Format(gdFecSis, "dd/mm/yyyy"), "F2", 10, hRight
'    oDoc.WTextBox 40, 60, 35, 490, Format(Time, "hh:mm:ss"), "F2", 10, hRight
'
'    oDoc.WTextBox 90 - B, 60, 15, 160, "Cliente", "F2", 10, hLeft
'    oDoc.WTextBox 90 - B, 60, 15, 80, ":", "F2", 10, hRight
'    oDoc.WTextBox 90 - B, 150, 15, 500, rsInfVisita!cPersNombre, "F1", 10, hjustify
'
'    oDoc.WTextBox 90 - B, 400, 15, 160, "Analista", "F2", 10, hjustify
'    oDoc.WTextBox 90 - B, 460, 15, 80, ":", "F2", 10, hjustify
'    oDoc.WTextBox 90 - B, 470, 15, 500, UCase(rsInfVisita!UserAnalista), "F1", 10, hjustify
'
'    oDoc.WTextBox 100 - B, 60, 15, 160, "Usuario", "F2", 10, hLeft
'    oDoc.WTextBox 100 - B, 60, 15, 80, ":", "F2", 10, hRight
'    oDoc.WTextBox 100 - B, 150, 15, 118, gsCodUser, "F1", 10, hjustify
'
'    oDoc.WTextBox 100 - B, 400, 15, 160, "Producto", "F2", 10, hjustify
'    oDoc.WTextBox 100 - B, 460, 15, 80, ":", "F2", 10, hjustify
'    oDoc.WTextBox 100 - B, 470, 15, 118, rsInfVisita!cConsDescripcion, "F1", 10, hjustify
'
'    oDoc.WTextBox 110 - B, 60, 15, 160, "Credito", "F2", 10, hLeft
'    oDoc.WTextBox 110 - B, 60, 15, 80, ":", "F2", 10, hRight
'    oDoc.WTextBox 110 - B, 150, 15, 500, rsInfVisita!cCtaCod, "F1", 10, hjustify
'
'    oDoc.WTextBox 120 - B, 60, 15, 160, "Cod. Cliente", "F2", 10, hLeft
'    oDoc.WTextBox 120 - B, 60, 15, 80, ":", "F2", 10, hRight
'    oDoc.WTextBox 120 - B, 150, 15, 500, rsInfVisita!cPersCod, "F1", 10, hjustify
'
'    oDoc.WTextBox 120 - B, 270, 15, 160, "Doc. Natural", "F2", 10, hjustify
'    oDoc.WTextBox 120 - B, 328, 15, 80, ":", "F2", 10, hjustify
'    oDoc.WTextBox 120 - B, 335, 15, 500, rsInfVisita!DNI, "F1", 10, hjustify
'
'    oDoc.WTextBox 120 - B, 400, 15, 160, "Doc. Juridico", "F2", 10, hjustify
'    oDoc.WTextBox 120 - B, 460, 15, 80, ":", "F2", 10, hjustify
'    oDoc.WTextBox 120 - B, 470, 15, 500, IIf(rsInfVisita!Ruc = "NULL", "-", rsInfVisita!Ruc), "F1", 10, hjustify
'
'                 'Top  Left Alto Ancho
'    oDoc.WTextBox 110, 100, 15, 400, "INFORME DE VISITA AL CLIENTE", "F2", 12, hCenter
'
'    'cuadro de Fecha de visita
'    oDoc.WTextBox 130, 50, 80, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'    '135
'    oDoc.WTextBox 185 - A, 55, 15, 160, "Fecha de Visita :", "F1", 10, hLeft
'    oDoc.WTextBox 185 - A, 190, 15, 500, Format(rsInfVisita!dFecVisita, "dd/mm/yyyy"), "F1", 10, hjustify
'
'    oDoc.WTextBox 200 - A, 55, 15, 160, "Fecha de ultima visita :", "F1", 10, hLeft
'    oDoc.WTextBox 215 - A, 55, 15, 160, "Persona(s) Entrevistada(s) :", "F1", 10, hLeft
'
'    oDoc.WTextBox 230 - A, 55, 15, 160, "Sr.(a) :", "F1", 10, hLeft
'    oDoc.WTextBox 230 - A, 300, 15, 160, "Cargo/Parentesco :", "F1", 10, hjustify
'    oDoc.WTextBox 245 - A, 55, 15, 160, "Sr.(a) :", "F1", 10, hLeft
'    oDoc.WTextBox 245 - A, 300, 15, 160, "Cargo/Parentesco :", "F1", 10, hjustify
'
'    'cuadro de Tipo de Visita
'    oDoc.WTextBox 260 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'    oDoc.WTextBox 262 - A, 55, 15, 500, "Tipo de Visita :", "F2", 10, hLeft
'
'    'cuadro de Tipo de Visita: Contenido
'    oDoc.WTextBox 275 - A, 50, 40, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'
'    oDoc.WTextBox 280 - A, 55, 15, 10, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 280 - A, 75, 15, 500, "1° Evaluacion (Cliente Nuevo)", "F1", 10, hjustify
'    oDoc.WTextBox 280 - A, 250, 15, 500, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 280 - A, 270, 15, 500, "Paralelo", "F1", 10, hjustify
'    oDoc.WTextBox 280 - A, 400, 15, 700, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 280 - A, 420, 15, 800, "Inspeccion de Garantias", "F1", 10, hjustify
'    oDoc.WTextBox 295 - A, 55, 15, 900, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 295 - A, 75, 15, 110, "Represtamo", "F1", 10, hjustify
'    oDoc.WTextBox 295 - A, 250, 15, 120, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 295 - A, 270, 15, 130, "Ampliacion", "F1", 10, hjustify
'
'    'cuadro de Sobre el Entorno Familiar del Cliente o Representante
'    nFilaAdicional = CalculaFilaAdicional(rsInfVisita!cEntornoFami)
'    nContador = nFilaAdicional
'    oDoc.WTextBox 315 - A, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 317 - A, 55, 15, 500, "Sobre el Entorno Familiar del Cliente o Representante:", "F2", 10, hLeft
'    'cuadro de Sobre el Entorno Familiar del Cliente o Representante : CONTENIDO
'    oDoc.WTextBox 330 - A, 50, 50 + nFilaAdicional, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 335 - A, 55, 10, 500, rsInfVisita!cEntornoFami, "F1", 10, hjustify
'
'    'cuadro de Sobre el giro y la Ubicacion del Negocio
'    nFilaAdicional = CalculaFilaAdicional(rsInfVisita!cEntornoFami)
'    oDoc.WTextBox 380 - A + nContador, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 382 - A + nContador, 55, 15, 500, "Sobre el Giro y la Ubicacion del Negocio:", "F2", 10, hLeft
'    'cuadro de Sobre el giro y la Ubicacion del Negocio : CONTENIDO
'    oDoc.WTextBox 395 - A + nContador, 50, 50 + nFilaAdicional, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 397 - A + nContador, 55, 10, 500, rsInfVisita!cGiroUbica, "F1", 10, hjustify
'    nContador = nContador + nFilaAdicional
'
'    'cuadro de Sobre la Experiencia Crediticia
'    nFilaAdicional = CalculaFilaAdicional(rsInfVisita!cExpeCrediticia)
'    oDoc.WTextBox 445 - A + nContador, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 447 - A + nContador, 55, 15, 500, "Sobre la Experiencia Crediticia:", "F2", 10, hLeft
'    'cuadro de Sobre la Experiencia Crediticia : CONTENIDO
'    oDoc.WTextBox 460 - A + nContador, 50, 50 + nFilaAdicional, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 462 - A + nContador, 55, 10, 500, rsInfVisita!cExpeCrediticia, "F1", 10, hjustify
'    nContador = nContador + nFilaAdicional
'
'    'cuadro de Sobre la consistencia de la informacion y la formalidad del negocio
'    nFilaAdicional = CalculaFilaAdicional(rsInfVisita!cFormalNegocio)
'    oDoc.WTextBox 510 - A + nContador, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 512 - A + nContador, 55, 15, 500, "Sobre la Consistencia de la Informacion y la Formalidad del Negocio:", "F2", 10, hLeft
'    'cuadro de Sobre la consistencia de la informacion y la formalidad del negocio : CONTENIDO
'    oDoc.WTextBox 525 - A + nContador, 50, 50 + nFilaAdicional, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 527 - A + nContador, 55, 10, 500, rsInfVisita!cFormalNegocio, "F1", 10, hjustify
'    nContador = nContador + nFilaAdicional
'
'    'cuadro de Sobre la Colaterales o Garantias
'    nFilaAdicional = CalculaFilaAdicional(rsInfVisita!cFormalNegocio)
'    oDoc.WTextBox 575 - A + nContador, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 577 - A + nContador, 55, 15, 500, "Sobre los Colaterales o Garantias:", "F2", 10, hLeft
'    'cuadro de Sobre la Colaterales o Garantias : CONTENIDO
'    oDoc.WTextBox 590 - A + nContador, 50, 50 + nFilaAdicional, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 592 - A + nContador, 55, 10, 500, rsInfVisita!cColateGarantia, "F1", 10, hjustify
'    nContador = nContador + nFilaAdicional
'
'    'cuadro de Destino
'    nFilaAdicional = CalculaFilaAdicional(rsInfVisita!cFormalNegocio)
'    oDoc.WTextBox 640 - A + nContador, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 642 - A + nContador, 55, 15, 500, "Sobre el destino y el impacto del mismo:", "F2", 10, hLeft
'    'cuadro de Sustento de Venta : CONTENIDO
'    oDoc.WTextBox 655 - A + nContador, 50, 50 + nFilaAdicional, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 657 - A + nContador, 55, 10, 500, rsInfVisita!cDestino, "F1", 10, hjustify
'    nContador = nContador + nFilaAdicional
'
'    'cuadro de VERIFICACION DE INMUEBLE
'    oDoc.WTextBox 705 - A + nContador, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'    oDoc.WTextBox 707 - A + nContador, 55, 15, 500, "Verificacion de Inmueble :", "F2", 10, hLeft
'    'cuadro de VERIFICACION DE INMUEBLE:coNTENIDO
'    oDoc.WTextBox 720 - A + nContador, 50, 95, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1.2, vbBlack
'
'    oDoc.WTextBox 722 - A + nContador, 55, 15, 500, "Direccion :", "F1", 10, hLeft
'    oDoc.WTextBox 735 - A + nContador, 55, 15, 500, "Referencia de Ubicacion :", "F1", 10, hLeft
'    oDoc.WTextBox 747 - A + nContador, 55, 15, 500, "Zona :", "F1", 10, hLeft
'    oDoc.WTextBox 755 - A + nContador, 200, 15, 500, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 755 - A + nContador, 220, 50, 500, "Urbana", "F1", 10, hjustify
'    oDoc.WTextBox 755 - A + nContador, 280, 60, 500, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 755 - A + nContador, 300, 70, 500, "Rural", "F1", 10, hjustify
'    oDoc.WTextBox 767 - A + nContador, 55, 15, 500, "Tipo de Construccion :", "F1", 10, hLeft
'    oDoc.WTextBox 780 - A + nContador, 100, 15, 500, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 780 - A + nContador, 120, 15, 500, "Material Noble", "F1", 10, hjustify
'    oDoc.WTextBox 780 - A + nContador, 200, 15, 500, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 780 - A + nContador, 220, 15, 500, "Madera", "F1", 10, hjustify
'    oDoc.WTextBox 780 - A + nContador, 280, 15, 500, "( )", "F1", 10, hjustify
'    oDoc.WTextBox 780 - A + nContador, 300, 15, 500, "Otros", "F1", 10, hjustify
'    oDoc.WTextBox 795 - A + nContador, 55, 15, 500, "Estado de la Vivienda :", "F1", 10, hLeft
'
'    'cuadro de VISTO BUENO
'    oDoc.WTextBox 815 - A + nContador, 50, 15, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'
'    oDoc.WTextBox 817 - A + nContador, 55, 15, 500, "Analista de Creditos :", "F2", 10, hjustify
'    oDoc.WTextBox 817 - A + nContador, 320, 15, 500, "Jefe de Grupo :", "F2", 10, hjustify
'
'    'cuadro de VISTO BUENO:Contenido
'    oDoc.WTextBox 780 + nContador, 50, 40, 500, "", "F1", 12, hCenter, vTop, vbBlack, 1, vbBlack
'    oDoc.PDFClose
'    oDoc.Show
'
'    Else
'    MsgBox "Los Datos de la propuesta del Credito no han sido Registrados Correctamente", vbInformation, "Aviso"
'    End If
'End Sub

Private Function CalculaFilaAdicional(ByVal psTexto As String) As Integer
    Dim nTotalAceptable As Integer, nTotalText As Integer, nTotalTextoExceso As Integer, nTotalTextoFila As Integer, nTotalFilasAdicional As Integer, nTotalEspacioLinea As Integer, nTotalEspacioAdicional As Integer
    
    nTotalAceptable = 330
    nTotalTextoFila = 105
    nTotalEspacioLinea = 1
    CalculaFilaAdicional = 0
    nTotalText = Len(psTexto)
        If nTotalText > nTotalAceptable And nTotalText < 400 Then
           ' nTotalTextoExceso = nTotalText - nTotalAceptable
           ' nTotalFilasAdicional = nTotalTextoExceso / nTotalTextoFila
            'nTotalEspacioAdicional = nTotalFilasAdicional * nTotalEspacioLinea
            nTotalEspacioAdicional = 1 * nTotalEspacioLinea
            CalculaFilaAdicional = nTotalEspacioAdicional
        ElseIf nTotalText > 400 Then
            nTotalEspacioAdicional = 2 * nTotalEspacioLinea
            CalculaFilaAdicional = nTotalEspacioAdicional
        End If
    
End Function

