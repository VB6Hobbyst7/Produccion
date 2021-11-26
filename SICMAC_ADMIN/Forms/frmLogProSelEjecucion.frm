VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLogProSelEjecucion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5670
   ClientLeft      =   555
   ClientTop       =   2100
   ClientWidth     =   10950
   Icon            =   "frmLogProSelEjecucion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   10950
   Visible         =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9300
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Proceso de Seleccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1530
      Left            =   120
      TabIndex        =   38
      Top             =   75
      Width           =   10695
      Begin VB.TextBox txtObjeto 
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
         Height          =   315
         Left            =   8640
         Locked          =   -1  'True
         TabIndex        =   124
         Top             =   300
         Width           =   1860
      End
      Begin VB.CommandButton CmdConsultarProceso 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2670
         TabIndex        =   2
         Top             =   310
         Width           =   350
      End
      Begin VB.TextBox txtanio 
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
         Height          =   315
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   300
         Width           =   3255
      End
      Begin VB.TextBox TxtProSelNro 
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
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox TxtTipo 
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
         Height          =   300
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   630
         Width           =   6255
      End
      Begin VB.TextBox TxtMonto 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   9240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   1260
      End
      Begin VB.TextBox TxtDescripcion 
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
         Height          =   495
         Left            =   1680
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   950
         Width           =   8835
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Objeto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8040
         TabIndex        =   125
         Top             =   360
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ejecución"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   69
         Top             =   360
         Width           =   1425
      End
      Begin VB.Label LblMoneda 
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
         Height          =   300
         Left            =   8640
         TabIndex        =   4
         Top             =   630
         Width           =   580
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nº Proceso"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   42
         Top             =   345
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Proceso Selección"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   41
         Top             =   700
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Monto"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8040
         TabIndex        =   40
         Top             =   690
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   1020
         Width           =   840
      End
   End
   Begin VB.Frame FrPostores 
      Caption         =   "Lista de Postores del Porceso de Seleccion"
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
      Height          =   3840
      Left            =   120
      TabIndex        =   70
      Top             =   1680
      Visible         =   0   'False
      Width           =   10680
      Begin VB.CommandButton cmdQuitarPostor 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1380
         TabIndex        =   72
         Top             =   3360
         Width           =   1155
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSItemPostores 
         Height          =   2925
         Left            =   120
         TabIndex        =   73
         Top             =   330
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5159
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   -2147483647
         ForeColorSel    =   -2147483624
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         Enabled         =   0   'False
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.CommandButton CmdRegistrar 
         Caption         =   "Registrar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   3360
         Width           =   1215
      End
      Begin VB.CommandButton cmdAgragarPostor 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   71
         Top             =   3360
         Width           =   1215
      End
   End
   Begin VB.Frame FrameConObs 
      Caption         =   "Lista de Consultas del Proceso de Selección "
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
      Height          =   3810
      Left            =   120
      TabIndex        =   24
      Top             =   1740
      Visible         =   0   'False
      Width           =   10695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSFObsCon 
         Height          =   2925
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5159
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483630
         Cols            =   6
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   -2147483647
         ForeColorSel    =   -2147483624
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         WordWrap        =   -1  'True
         FocusRect       =   0
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.CommandButton cmdModificarConOns 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2520
         TabIndex        =   59
         Top             =   3300
         Width           =   1155
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1320
         TabIndex        =   25
         Top             =   3300
         Width           =   1155
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   3300
         Width           =   1155
      End
      Begin VB.CommandButton cmdResponderConOns 
         Caption         =   "Responder"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   3300
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1320
         TabIndex        =   61
         Top             =   3300
         Visible         =   0   'False
         Width           =   1155
      End
   End
   Begin VB.Frame FrActoPublico 
      Caption         =   "Datos del Acto Publico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3810
      Left            =   120
      TabIndex        =   75
      Top             =   1680
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CommandButton cmdCnsArchOcurrencia 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10035
         TabIndex        =   100
         Top             =   1470
         Width           =   330
      End
      Begin VB.TextBox txtarchOcurrencia 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   101
         Top             =   1440
         Width           =   8775
      End
      Begin VB.TextBox txtNotario 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         TabIndex        =   82
         Top             =   330
         Width           =   8775
      End
      Begin VB.TextBox txtLugar 
         Height          =   315
         Left            =   1620
         MaxLength       =   255
         TabIndex        =   78
         Top             =   1020
         Width           =   8775
      End
      Begin VB.CommandButton cmdGuardarDatosAP 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   6420
         TabIndex        =   77
         Top             =   3345
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelarDatosAP 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7800
         TabIndex        =   76
         Top             =   3345
         Width           =   1275
      End
      Begin VB.TextBox txtveedor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1620
         TabIndex        =   79
         Top             =   675
         Width           =   8775
      End
      Begin RichTextLib.RichTextBox txtocurrencia 
         Height          =   1425
         Left            =   1620
         TabIndex        =   102
         Top             =   1800
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2514
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmLogProSelEjecucion.frx":08CA
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Ocurrencia"
         Height          =   195
         Left            =   120
         TabIndex        =   104
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Archivo"
         Height          =   195
         Left            =   120
         TabIndex        =   103
         Top             =   1500
         Width           =   540
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Notario Publico"
         Height          =   195
         Left            =   180
         TabIndex        =   83
         Top             =   465
         Width           =   1080
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Lugar de Ejecucion"
         Height          =   195
         Left            =   120
         TabIndex        =   81
         Top             =   1140
         Width           =   1380
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Veedor"
         Height          =   195
         Left            =   180
         TabIndex        =   80
         Top             =   780
         Width           =   510
      End
   End
   Begin VB.Frame fraObs 
      Caption         =   "Respuesta a Consultas/Observaciones del Proceso de Selección "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   10695
      Begin MSComDlg.CommonDialog CDlgConsultas 
         Left            =   120
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "*.txt"
         DialogTitle     =   "Abrir Archivo de Consultas"
         FileName        =   "*.txt"
         Filter          =   "*.txt"
      End
      Begin VB.CommandButton cmdCnsArchconsultasResp 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10035
         TabIndex        =   88
         Top             =   2100
         Width           =   315
      End
      Begin VB.CommandButton cmdCnsArchconsultas 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10035
         TabIndex        =   84
         Top             =   870
         Width           =   315
      End
      Begin VB.CommandButton CmdCancelarConsObs 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7800
         TabIndex        =   66
         Top             =   3360
         Width           =   1275
      End
      Begin VB.CommandButton cmdPersObs 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2235
         TabIndex        =   32
         Top             =   510
         Width           =   315
      End
      Begin VB.TextBox txtPersObs 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   480
         Width           =   7815
      End
      Begin VB.CommandButton cmdGrabarObs 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   6420
         TabIndex        =   30
         Top             =   3360
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancelaObs 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   9180
         TabIndex        =   29
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtPersCodObs 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         TabIndex        =   33
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtArchivoConsulta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   840
         Width           =   9255
      End
      Begin RichTextLib.RichTextBox rTxtResp 
         Height          =   855
         Left            =   1140
         TabIndex        =   87
         Top             =   2400
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   0   'False
         ScrollBars      =   2
         TextRTF         =   $"frmLogProSelEjecucion.frx":094D
      End
      Begin VB.TextBox txtArchivoConsultaResp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   2070
         Width           =   9255
      End
      Begin RichTextLib.RichTextBox txtconsulta 
         Height          =   855
         Left            =   1140
         TabIndex        =   91
         Top             =   1200
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   1508
         _Version        =   393217
         ScrollBars      =   2
         TextRTF         =   $"frmLogProSelEjecucion.frx":09D0
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Archivo"
         Height          =   195
         Left            =   300
         TabIndex        =   90
         Top             =   2120
         Width           =   540
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Archivo"
         Height          =   195
         Left            =   300
         TabIndex        =   86
         Top             =   900
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Persona"
         Height          =   195
         Left            =   300
         TabIndex        =   36
         Top             =   540
         Width           =   585
      End
      Begin VB.Label lblDesc 
         AutoSize        =   -1  'True
         Caption         =   "Consulta"
         Height          =   195
         Left            =   300
         TabIndex        =   35
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Respuesta"
         Height          =   195
         Left            =   300
         TabIndex        =   34
         Top             =   2400
         Width           =   765
      End
   End
   Begin VB.Frame fraPos 
      Caption         =   "Registro de Participantes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   43
      Top             =   1680
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CommandButton cmdimprimirpostores 
         Caption         =   "Imprimir Registro de Participante"
         Height          =   360
         Left            =   1530
         TabIndex        =   123
         Top             =   3360
         Width           =   2835
      End
      Begin VB.TextBox txtDNI 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9270
         Locked          =   -1  'True
         TabIndex        =   121
         Top             =   1845
         Width           =   1080
      End
      Begin VB.TextBox txtRUC 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9270
         Locked          =   -1  'True
         TabIndex        =   119
         Top             =   750
         Width           =   1080
      End
      Begin VB.TextBox txtEmail 
         Height          =   315
         Left            =   1530
         MaxLength       =   80
         TabIndex        =   117
         Top             =   1485
         Width           =   6795
      End
      Begin VB.TextBox txtTelefono 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9270
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   1110
         Width           =   1080
      End
      Begin VB.TextBox txtDomicilio 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   1110
         Width           =   6795
      End
      Begin VB.CommandButton cmdPersRep 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2760
         TabIndex        =   110
         Top             =   1875
         Width           =   330
      End
      Begin VB.TextBox txtRepresenta 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   1845
         Width           =   5205
      End
      Begin VB.ComboBox cboMonedaRecibo 
         Height          =   315
         ItemData        =   "frmLogProSelEjecucion.frx":0A53
         Left            =   8220
         List            =   "frmLogProSelEjecucion.frx":0A55
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   2475
         Width           =   750
      End
      Begin VB.TextBox txtDescRecibo 
         Height          =   315
         Left            =   1530
         TabIndex        =   106
         Top             =   2835
         Width           =   8850
      End
      Begin VB.TextBox txtimpoteRecibo 
         Height          =   315
         Left            =   9000
         TabIndex        =   105
         Top             =   2475
         Width           =   1380
      End
      Begin VB.CommandButton cmdCancelarPostor 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7815
         TabIndex        =   67
         Top             =   3345
         Width           =   1275
      End
      Begin VB.TextBox TxtSerie 
         Height          =   315
         Left            =   1530
         TabIndex        =   60
         Top             =   2475
         Width           =   495
      End
      Begin VB.TextBox txtnrorecibo 
         Height          =   315
         Left            =   2010
         TabIndex        =   50
         Top             =   2475
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelaPos 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   9180
         TabIndex        =   47
         Top             =   3345
         Width           =   1335
      End
      Begin VB.CommandButton cmdGrabarPos 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   6315
         TabIndex        =   46
         Top             =   3345
         Width           =   1395
      End
      Begin VB.TextBox txtPersona 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   750
         Width           =   5190
      End
      Begin VB.CommandButton cmdPersona 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2760
         TabIndex        =   44
         Top             =   780
         Width           =   330
      End
      Begin VB.TextBox txtPersCod 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1530
         TabIndex        =   48
         Top             =   750
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtFechaRecibo 
         Height          =   315
         Left            =   1530
         TabIndex        =   108
         Top             =   375
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtPersRep 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1530
         TabIndex        =   111
         Top             =   1845
         Width           =   1575
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "D.N.I."
         Height          =   195
         Left            =   8505
         TabIndex        =   122
         Top             =   1905
         Width           =   420
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C."
         Height          =   195
         Left            =   8505
         TabIndex        =   120
         Top             =   810
         Width           =   480
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "e-mail"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   390
         TabIndex        =   118
         Top             =   1560
         Width           =   405
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
         Height          =   195
         Left            =   8505
         TabIndex        =   116
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
         Height          =   195
         Left            =   375
         TabIndex        =   115
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Representante"
         Height          =   195
         Left            =   390
         TabIndex        =   112
         Top             =   1905
         Width           =   1050
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Importe"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7575
         TabIndex        =   54
         Top             =   2535
         Width           =   525
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   375
         TabIndex        =   53
         Top             =   2895
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   390
         TabIndex        =   52
         Top             =   450
         Width           =   495
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nro de Recibo"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   375
         TabIndex        =   51
         Top             =   2535
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Razón Social"
         Height          =   195
         Left            =   375
         TabIndex        =   49
         Top             =   795
         Width           =   945
      End
   End
   Begin VB.Frame FrameApelacionesLista 
      Caption         =   "Lista de Apelaciones por Item "
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
      Height          =   3900
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CommandButton cmdRespuesta 
         Caption         =   "Responder"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   3345
         Width           =   1275
      End
      Begin VB.CommandButton cmdquitarAp 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   3345
         Width           =   1275
      End
      Begin VB.CommandButton cmdregistrarAp 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3345
         Width           =   1275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSApe 
         Height          =   1575
         Left            =   90
         TabIndex        =   12
         Top             =   1695
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   2778
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   14677503
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         WordWrap        =   -1  'True
         FocusRect       =   0
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSItem 
         Height          =   1410
         Left            =   90
         TabIndex        =   55
         Top             =   240
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   2487
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   7
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   16773857
         ForeColorSel    =   -2147483635
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
   End
   Begin VB.Frame frameApelacionesD 
      Caption         =   "Registro de Apelación "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3855
      Left            =   120
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   10695
      Begin VB.CommandButton cmdCnsArchApeRespuesta 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10155
         TabIndex        =   95
         Top             =   2145
         Width           =   315
      End
      Begin VB.CommandButton cmdCnsArchApelacion 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   10155
         TabIndex        =   92
         Top             =   1090
         Width           =   315
      End
      Begin VB.CommandButton cmdCancelarApelaciones 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   7800
         TabIndex        =   68
         Top             =   3360
         Width           =   1275
      End
      Begin VB.CommandButton cmdGrabarApelacion 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   6420
         TabIndex        =   18
         Top             =   3360
         Width           =   1275
      End
      Begin VB.CommandButton cmdPersApe 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Left            =   2565
         TabIndex        =   17
         Top             =   750
         Width           =   315
      End
      Begin VB.TextBox txtPersApe 
         Appearance      =   0  'Flat
         Height          =   290
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   720
         Width           =   7575
      End
      Begin VB.CommandButton cmdCancelarApelacion 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   9180
         TabIndex        =   15
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox txtnro 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox txtPersCodApe 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   1755
      End
      Begin VB.TextBox txtArchApelaciones 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   93
         Top             =   1060
         Width           =   9375
      End
      Begin VB.TextBox txtArchApeResp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   96
         Top             =   2115
         Width           =   9375
      End
      Begin RichTextLib.RichTextBox txtRespuesta 
         Height          =   855
         Left            =   1140
         TabIndex        =   98
         Top             =   2475
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmLogProSelEjecucion.frx":0A57
      End
      Begin RichTextLib.RichTextBox txtApelacion 
         Height          =   615
         Left            =   1140
         TabIndex        =   99
         Top             =   1440
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   1085
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"frmLogProSelEjecucion.frx":0ADA
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Archivo"
         Height          =   195
         Left            =   480
         TabIndex        =   97
         Top             =   2175
         Width           =   540
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Archivo"
         Height          =   195
         Left            =   480
         TabIndex        =   94
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Apelacion"
         Height          =   195
         Left            =   360
         TabIndex        =   23
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         Height          =   195
         Left            =   480
         TabIndex        =   22
         Top             =   780
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Respuesta"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2475
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nro"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   420
         Width           =   255
      End
   End
   Begin VB.Frame FrameComite 
      BackColor       =   &H8000000A&
      Height          =   3900
      Left            =   120
      TabIndex        =   56
      Top             =   1680
      Visible         =   0   'False
      Width           =   10695
      Begin VB.ComboBox cboEtapas 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   360
         Width           =   8415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSComite 
         Height          =   2205
         Left            =   120
         TabIndex        =   57
         Top             =   1080
         Width           =   10395
         _ExtentX        =   18336
         _ExtentY        =   3889
         _Version        =   393216
         BackColor       =   16777215
         Cols            =   5
         FixedCols       =   0
         ForeColorFixed  =   -2147483646
         BackColorSel    =   16773857
         ForeColorSel    =   -2147483635
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         GridColorUnpopulated=   -2147483633
         FocusRect       =   0
         ScrollBars      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin VB.CommandButton cmdComite 
         Caption         =   "Asignar"
         Height          =   375
         Left            =   7800
         TabIndex        =   58
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   6060
         Picture         =   "frmLogProSelEjecucion.frx":0B5D
         Top             =   3480
         Width           =   240
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Los miembros del Comité Responsable de la Etapa están marcados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   3450
         Width           =   5685
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Miembro del Comité del Responsable Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   840
         Width           =   3885
      End
      Begin VB.Label Label16 
         Caption         =   "Etapas del Proceso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   420
         Width           =   1815
      End
   End
   Begin VB.Image imgOK 
      Height          =   240
      Left            =   180
      Picture         =   "frmLogProSelEjecucion.frx":0E9F
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgNN 
      Height          =   240
      Left            =   480
      Picture         =   "frmLogProSelEjecucion.frx":11E1
      Top             =   240
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmLogProSelEjecucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSQL As String
Dim xEjecutar As Boolean
Dim nMes As Integer, nAnio As Integer, Rspta As Boolean
Dim nTipo As Integer, cTitulo As String
Dim gnProSelNro As Integer, gnBSGrupoCod As String, gnAnio As Integer

Public Sub TipoFuncion(ByVal pnTipo As Integer, psTitulo As String)
nTipo = pnTipo
cTitulo = psTitulo
Me.Show 1
End Sub

Private Sub Form_Load()
Me.Height = 6075 '5780
CentraForm Me
gnProSelNro = 0
gnBSGrupoCod = ""
FormaFlexEta
txtanio.Text = Year(gdFecSis)
Me.Caption = cTitulo
Select Case nTipo
    Case 1 'consultas
        FormaMSFObsCon
        FrameConObs.Visible = True
'        cmdConsultas.Visible = True
'        Caption = "Registro de Consultas"
        FrameConObs.Caption = "Lista de Consultas del Proceso de Selección "
        cmdImprimir.Visible = False
        cmdResponderConOns.Visible = False
        rTxtResp.Enabled = False
        txtArchivoConsultaResp.Enabled = False
        cmdCnsArchconsultasResp.Enabled = False
    Case 2 'observaciones
        FormaMSFObsCon
        FrameConObs.Visible = True
 '       Caption = "Registro de Observaciones"
        FrameConObs.Caption = "Lista de Observaciones del Proceso de Selección "
        cmdImprimir.Visible = False
        cmdResponderConOns.Visible = False
        cmdQuitar.Visible = True
        cmdAgregar.Visible = True
        rTxtResp.Enabled = False
        txtArchivoConsultaResp.Enabled = False
        cmdCnsArchconsultasResp.Enabled = False
    Case 3 'Postores
        FormaFlexItemPostor
        FrPostores.Visible = True
 '       Caption = "Registro de Participantes y Venta de Bases"
        CargarMonedaRecibo
        cmdAgragarPostor.Visible = True
        cmdQuitarPostor.Visible = True
        CmdRegistrar.Visible = False
    Case 4 'Apelaciones
        FormaMSApe
        FormaFlexItem
        FrameApelacionesLista.Visible = True
        Caption = "Registro de Apelaciones"
        'cmdSalir.Top = cmdSalir.Top + 2000
        'Height = Height + 2000
    Case 5
        FrameComite.Visible = True
        FormaMSComite
'        Caption = "Asignacion del Comite por Etapa del Proceso de Seleccion"
    Case 6
        FormaMSFObsCon
        FrameConObs.Visible = True
'        cmdConsultas.Visible = True
'        Caption = "Absolucion de Consultas"
        FrameConObs.Caption = "Lista de Consultas del Proceso de Selección "
        cmdQuitar.Visible = False
        cmdAgregar.Visible = False
        cmdModificarConOns.Visible = False
        cmdResponderConOns.Visible = True
        cmdImprimir.Visible = True
        rTxtResp.Enabled = True
        txtArchivoConsultaResp.Enabled = True
        cmdCnsArchconsultasResp.Enabled = True
    Case 7
        FormaMSFObsCon
        FrameConObs.Visible = True
'        Caption = "Absolucion de Observaciones"
        FrameConObs.Caption = "Lista de Observaciones del Proceso de Selección "
        cmdQuitar.Visible = False
        cmdAgregar.Visible = False
        cmdModificarConOns.Visible = False
        cmdResponderConOns.Visible = True
        cmdImprimir.Visible = True
        rTxtResp.Enabled = True
        txtArchivoConsultaResp.Enabled = True
        cmdCnsArchconsultasResp.Enabled = True
    Case 8
        FormaFlexItemPostor
        FrPostores.Visible = True
'        Caption = "Registrar Entrega de Propuestas"
        cmdAgragarPostor.Visible = False
        cmdQuitarPostor.Visible = False
        cmdimprimirpostores.Left = cmdQuitarPostor.Left + 100
        CmdRegistrar.Visible = True
    Case 9
'        Caption = "Registro de Datos del Acto Publico"
        FrActoPublico.Visible = True
End Select
CentraForm Me
End Sub


Private Sub cboEtapas_Click()
    LimpiaMiembros
    CargarComiteEtapa gnProSelNro, cboEtapas.ItemData(cboEtapas.ListIndex)
End Sub

Private Sub cmdAgragarPostor_Click()
If Len(TxtProSelNro) = 0 Then
   MsgBox "Debe seleccionar un proceso de selección..." + Space(10), vbInformation
End If

If gnProSelNro = 0 Then Exit Sub
cmdCancelaPos_Click
FrPostores.Visible = False
fraPos.Visible = True
End Sub

Private Sub cmdAgregar_Click()
    
    If gnProSelNro = 0 Then Exit Sub
    
    cmdCancelaObs_Click
    fraObs.Visible = True
    FrameConObs.Visible = False
    cmdPersObs.Enabled = True
    txtPersCodObs.Enabled = True
    txtPersObs.Enabled = True
    txtconsulta.Enabled = True
    Rspta = False
End Sub

Private Sub cmdCancelaObs_Click()
fraObs.Visible = False
FrameConObs.Visible = True
txtPersCodObs.Text = ""
txtPersObs.Text = ""
txtconsulta.Text = ""
rTxtResp.Text = ""
End Sub

Private Sub cmdCancelaPos_Click()
    fraPos.Visible = False
    FrPostores.Visible = True
    txtPersCod.Text = ""
    txtPersona.Text = ""
    TxtSerie.Text = ""
    txtnrorecibo.Text = ""
    txtDescRecibo.Text = ""
    txtimpoteRecibo.Text = ""
    txtFechaRecibo.Text = "__/__/____"
End Sub

Private Sub cmdCancelarApelacion_Click()
    frameApelacionesD.Visible = False
    FrameApelacionesLista.Visible = True
    txtPersCodApe.Text = ""
    txtPersApe.Text = ""
    txtApelacion.Text = ""
    txtRespuesta.Text = ""
End Sub

Private Sub cmdCancelarApelaciones_Click()
    FrameApelacionesLista.Visible = True
    frameApelacionesD.Visible = False
    txtArchApelaciones.Text = ""
    txtArchApeResp.Text = ""
    txtArchApeResp.Text = ""
    txtRespuesta.Text = ""
    txtApelacion.Text = ""
    txtArchApelaciones.Text = ""
    txtnro.Text = ""
    txtPersCodApe.Text = ""
    txtPersApe.Text = ""
End Sub

Private Sub CmdCancelarConsObs_Click()
    fraObs.Visible = False
    FrameConObs.Visible = True
    txtArchivoConsulta.Text = ""
    txtArchivoConsultaResp.Text = ""
    txtconsulta.Text = ""
    rTxtResp.Text = ""
End Sub

Private Sub cmdCancelarDatosAP_Click()
    txtNotario.Text = ""
    txtveedor.Text = ""
    txtLugar.Text = ""
    txtocurrencia.Text = ""
    TxtProSelNro = ""
    gnProSelNro = 0
    gnBSGrupoCod = ""
    TxtTipo.Text = ""
    TxtMonto.Text = ""
    LblMoneda.Caption = ""
    TxtDescripcion.Text = ""
    txtarchOcurrencia.Text = ""
    txtanio.Text = Year(gdFecSis)
End Sub

Private Sub cmdCancelarPostor_Click()
    fraPos.Visible = False
    FrPostores.Visible = True
End Sub

Private Sub cmdCnsArchApelacion_Click()
On Error GoTo cmdCnsArchApelacionErr
    Dim sArchivo As String
    CDlgConsultas.FileName = "*.txt"
    CDlgConsultas.InitDir = App.path
    CDlgConsultas.ShowOpen
    txtArchApelaciones.Text = CDlgConsultas.FileName
    sArchivo = Right(txtArchApelaciones.Text, 3)
    
    If sArchivo = "txt" Or sArchivo = "doc" Then
        txtApelacion.FileName = txtArchApelaciones.Text
        txtApelacion.Text = LimpiaString(txtApelacion.Text)
    Else
        txtArchApelaciones.Text = ""
        txtApelacion.Text = ""
        MsgBox "Archivo Incompatible", vbInformation, "Aviso"
    End If
    Exit Sub
cmdCnsArchApelacionErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCnsArchApeRespuesta_Click()
On Error GoTo cmdCnsArchApeRespuestaErr
    Dim sArchivo As String
    CDlgConsultas.FileName = "*.txt"
    CDlgConsultas.InitDir = App.path
    CDlgConsultas.ShowOpen
    txtArchApeResp.Text = CDlgConsultas.FileName
    sArchivo = Right(txtArchApeResp.Text, 3)
    
    If sArchivo = "txt" Or sArchivo = "doc" Then
        txtRespuesta.FileName = txtArchApeResp.Text
        txtRespuesta.Text = LimpiaString(txtRespuesta.Text)
    Else
        txtRespuesta.Text = ""
        txtArchApeResp.Text = ""
        MsgBox "Archivo Incompatible", vbInformation, "Aviso"
    End If
    Exit Sub
cmdCnsArchApeRespuestaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCnsArchconsultas_Click()
On Error GoTo cmdCnsArchconsultasErr
    Dim sArchivo As String
    CDlgConsultas.FileName = "*.txt"
    CDlgConsultas.InitDir = App.path
    CDlgConsultas.ShowOpen
    txtArchivoConsulta.Text = CDlgConsultas.FileName
    sArchivo = Right(txtArchivoConsulta.Text, 3)
    
    If sArchivo = "txt" Or sArchivo = "doc" Then
        txtconsulta.FileName = txtArchivoConsulta.Text
        txtconsulta.Text = LimpiaString(txtconsulta.Text)
    Else
        txtArchivoConsulta.Text = ""
        txtconsulta.Text = ""
        MsgBox "Archivo Incompatible", vbInformation, "Aviso"
    End If
    
    Exit Sub
cmdCnsArchconsultasErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCnsArchconsultasResp_Click()
On Error GoTo cmdCnsArchconsultasRespErr
    Dim sArchivo As String
    CDlgConsultas.FileName = "*.txt"
    CDlgConsultas.ShowOpen
    txtArchivoConsultaResp.Text = CDlgConsultas.FileName
    sArchivo = Right(txtArchivoConsultaResp.Text, 3)
    If sArchivo = "txt" Or sArchivo = "doc" Then
        rTxtResp.FileName = txtArchivoConsultaResp.Text
        rTxtResp.Text = LimpiaString(rTxtResp.Text)
    Else
        txtArchivoConsultaResp.Text = ""
        rTxtResp.Text = ""
        MsgBox "Archivo Incompatible", vbInformation, "Aviso"
    End If
    Exit Sub
cmdCnsArchconsultasRespErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdCnsArchOcurrencia_Click()
On Error GoTo cmdCnsArchOcurrenciaErr
    Dim sArchivo As String
    CDlgConsultas.FileName = "*.txt"
    CDlgConsultas.ShowOpen
    txtarchOcurrencia.Text = CDlgConsultas.FileName
    sArchivo = Right(txtarchOcurrencia.Text, 3)
    If sArchivo = "txt" Or sArchivo = "doc" Then
        txtocurrencia.FileName = txtarchOcurrencia.Text
        txtocurrencia.Text = LimpiaString(txtocurrencia.Text)
    Else
        txtArchivoConsultaResp.Text = ""
        rTxtResp.Text = ""
        MsgBox "Archivo Incompatible", vbInformation, "Aviso"
    End If
    Exit Sub
cmdCnsArchOcurrenciaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdComite_Click()
On Error GoTo cmdComiteErr
    Dim oCon As DConecta, i As Integer, valor As Integer
    If gnProSelNro = 0 Then Exit Sub
    If MSComite.TextMatrix(1, 2) = "" Then Exit Sub
    If Not ValidaComite Then
        MsgBox "Debe Seleccionar por lo menos a un Miembro del Comite", vbInformation, "Aviso"
        Exit Sub
    End If
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        With MSComite
            i = 1
            sSQL = "delete LogProSelEtapaComite where nProSelNro= " & gnProSelNro & " and nEtapaCod=" & cboEtapas.ItemData(cboEtapas.ListIndex)
            oCon.Ejecutar sSQL
            Do While i < .Rows
                .Col = 0
                .row = i
                If .CellPicture = imgOK Then
                    sSQL = "insert into LogProSelEtapaComite (nProSelNro,nEtapaCod,cPersCod)" & _
                       " values(" & gnProSelNro & "," & cboEtapas.ItemData(cboEtapas.ListIndex) & ",'" & .TextMatrix(i, 2) & "')"
                    oCon.Ejecutar sSQL
                End If
                i = i + 1
            Loop
            .ColSel = .Cols - 1
        End With
        MsgBox "Comite para la Etapa " & cboEtapas.Text & " Asignado", vbInformation
        oCon.CierraConexion
    End If
'    cargarComiteItemProceso gnProSelNro
    Exit Sub
cmdComiteErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Function ValidaComite() As Boolean
On Error GoTo ValidaComiteErr
    Dim i As Integer
    With MSComite
        i = 1
        Do While i < .Rows
            .Col = 0
            .row = i
            If .CellPicture = imgOK Then
                ValidaComite = True
                .ColSel = .Cols - 1
                Exit Function
            End If
            i = i + 1
        Loop
        .ColSel = .Cols - 1
        ValidaComite = False
    End With
    Exit Function
ValidaComiteErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Sub CmdConsultarProceso_Click()
On Error GoTo msflex_clckErr
    frmLogProSelCnsProcesoSeleccion.Inicio 2
    With frmLogProSelCnsProcesoSeleccion
        'If Not .gbBandera Then Exit Sub
        TxtProSelNro = .gvnProSelNro '.gvnNro
        gnProSelNro = .gvnProSelNro
        gnBSGrupoCod = .gvcBSGrupoCod
        txtObjeto = .gvcObjeto
        'txtnro.Text = .gvnNro
        TxtTipo.Text = .gvcTipo
        txtanio.Text = .gvcMes + " - " + CStr(.gvnAnio)
        TxtMonto.Text = Format(.gvnMonto, "###,###.00")
        LblMoneda.Caption = .gvcMoneda
        TxtDescripcion.Text = .gvcDescripcion
        gnAnio = .gvnAnio
        
    End With
    Select Case nTipo
        Case 1, 6
            lblDesc.Caption = "Consulta"
'            CargaObservaciones
            If VerificaEtapa(gnProSelNro, cnAbsolucionConsultas) Then
                If Not VerificaEtapaCerrada(gnProSelNro, cnObservaciones) Then
                    CargaConsultas
                    cmdModificarConOns.Visible = True
                    cmdQuitar.Visible = True
                    cmdAgregar.Visible = True
                    cmdImprimir.Visible = True
                    cmdResponderConOns.Visible = True
                Else
                    MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                    cmdModificarConOns.Visible = False
                    cmdQuitar.Visible = False
                    cmdAgregar.Visible = False
                    cmdImprimir.Visible = False
                    cmdResponderConOns.Visible = False
                    Exit Sub
                End If
            Else
               MsgBox "No se Especificado esta Etapa para este proceso", vbInformation, "Aviso"
               cmdModificarConOns.Visible = False
               cmdQuitar.Visible = False
               cmdAgregar.Visible = False
               cmdImprimir.Visible = False
               cmdResponderConOns.Visible = False
               Exit Sub
            End If
        Case 2, 7
            lblDesc.Caption = "Observación"
            'CargaObservaciones
            If VerificaEtapa(gnProSelNro, cnObservaciones) Then
                If Not VerificaEtapaCerrada(gnProSelNro, cnObservaciones) Then
                    CargaObservaciones
                Else
                    MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                    cmdModificarConOns.Visible = False
                    cmdQuitar.Visible = False
                    cmdAgregar.Visible = False
                    cmdImprimir.Visible = False
                    cmdResponderConOns.Visible = False
                    Exit Sub
                End If
            Else
               MsgBox "No se Especificado esta Etapa para este proceso", vbInformation, "Aviso"
               cmdModificarConOns.Visible = False
               cmdQuitar.Visible = False
               cmdAgregar.Visible = False
               cmdImprimir.Visible = False
               cmdResponderConOns.Visible = False
               Exit Sub
            End If
        Case 3
            If VerificaEtapa(gnProSelNro, cnRegistroParticipantes) Then
                If Not VerificaEtapaCerrada(gnProSelNro, cnRegistroParticipantes) Then
                    CargarPostores gnProSelNro
                    cmdAgragarPostor.Enabled = True
                    cmdQuitarPostor.Enabled = True
                    cmdimprimirpostores.Enabled = True
                Else
                    MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                    cmdAgragarPostor.Enabled = False
                    cmdQuitarPostor.Enabled = False
                    cmdimprimirpostores.Enabled = False
                    Exit Sub
                End If
            Else
                MsgBox "Etapa no Esta Configurada para este Proceso", vbInformation, "Aviso"
                cmdAgragarPostor.Enabled = False
                cmdQuitarPostor.Enabled = False
                cmdimprimirpostores.Enabled = False
                Exit Sub
            End If
        Case 8
            If VerificaEtapa(gnProSelNro, cnPresentacionPropuestas) Then
                CargarPostores gnProSelNro
            Else
                MsgBox "Etapa no Esta Configurada para este Proceso", vbInformation, "Aviso"
                Exit Sub
            End If
        Case 4
            If VerificaEtapa(gnProSelNro, cnApelaciones) Then
                If Not VerificaEtapaCerrada(gnProSelNro, cnApelaciones) Then
                    GeneraDetalleItem gnProSelNro
                    cmdregistrarAp.Visible = True
                    cmdquitarAp.Visible = True
                    cmdRespuesta.Visible = True
                Else
                    cmdregistrarAp.Visible = False
                    cmdquitarAp.Visible = False
                    cmdRespuesta.Visible = False
                    MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                    Exit Sub
                End If
            Else
                cmdregistrarAp.Visible = False
                cmdquitarAp.Visible = False
                cmdRespuesta.Visible = False
                MsgBox "Etapa no Esta Configurada para este Proceso", vbInformation, "Aviso"
                Exit Sub
            End If
'            MSItem.SetFocus
        Case 5
            CargarEtapas gnProSelNro
            If cboEtapas.ListCount > 0 Then
                cargarComiteItemProceso gnProSelNro
                CargarComiteEtapa gnProSelNro, cboEtapas.ItemData(cboEtapas.ListIndex)
                cmdComite.Visible = True
            Else
                cmdComite.Visible = False
                Exit Sub
            End If
        Case 9
            CagarDatosActoPublico gnProSelNro
    End Select
Exit Sub
msflex_clckErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CagarDatosActoPublico(ByVal pnProSelNro As Integer)
On Error GoTo CagarDatosActoPublicoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    If TxtTipo.Text <> "CONCURSO PUBLICO" And TxtTipo.Text <> "LICITACION PUBLICA" Then
        cmdCancelarDatosAP.Visible = False
        cmdGuardarDatosAP.Visible = False
        MsgBox "El Proceso Seleccionado no es un Acto Publico", vbInformation, "Aviso"
        Exit Sub
    Else
        cmdCancelarDatosAP.Visible = True
        cmdGuardarDatosAP.Visible = True
    End If
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select cNotario, cVeedor, cLugarEjecucion, tOcurrencias from LogProcesoSeleccion WHERE nProSelNro = " & gnProSelNro
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            txtNotario.Text = Rs!cnotario
            txtveedor.Text = Rs!cveedor
            txtLugar.Text = Rs!clugarejecucion
            txtocurrencia.Text = Rs!tOcurrencias
        Else
            txtNotario.Text = ""
            txtveedor.Text = ""
            txtLugar.Text = ""
            txtocurrencia.Text = ""
        End If
        oCon.CierraConexion
    End If
    Exit Sub
CagarDatosActoPublicoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Sub FormaMSComite()
    With MSComite
        .Clear
        .Rows = 2
        .Cols = 6
        .RowHeight(0) = 360
        .RowHeight(1) = 8
        .ColWidth(0) = 300
        .ColWidth(1) = 0:     .TextMatrix(0, 1) = "Proceso"
        .ColWidth(2) = 0:    .TextMatrix(0, 2) = "cPersCod"
        .ColWidth(3) = 0:    .TextMatrix(0, 3) = "Cargo"
        .ColWidth(4) = 8300:    .TextMatrix(0, 4) = "Nombre"
        .ColWidth(5) = 1500:    .TextMatrix(0, 5) = "Tipo"
        .WordWrap = True
    End With
End Sub

Private Sub cargarComiteItemProceso(ByRef pnProSelNro As Integer)
On Error GoTo cargarComiteItemProcesoErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
    Set oCon = New DConecta
    If oCon.AbreConexion Then
'        sSQL = "select distinct e.*, p.cPersNombre, c.cConsDescripcion, f.cDescripcion from LogProSelEvaluacionComite e " & _
'                "inner join persona p on p.cPersCod=e.cPersCod " & _
'                "inner join constante c on e.nCargo = c.nConsValor and nConsCod=9085 " & _
'                "inner join LogProSelEvaluacionFactor f on e.nFactorNro=f.nFactorNro " & _
'                "where nProSelNro=" & pnProSelNro & " and nProSelItem=" & pnProSelItem
        sSQL = "select s.nProSelNro, s.cPersCod, s.bSuplente, c.cConsDescripcion,p.cPersNombre from LogProSelComite s " & _
                "inner join constante c on s.nCargo=c.nConsValor and c.nConsCod=9085 " & _
                "inner join Persona p on s.cPersCod=p.cPersCod " & _
                "where nProSelNro= " & pnProSelNro & " order by bSuplente"
        Set Rs = oCon.CargaRecordSet(sSQL)
        FormaMSComite
        Do While Not Rs.EOF
            i = i + 1
            InsRow MSComite, i
            MSComite.Col = 0
            MSComite.row = i
'            If Rs!bEvaluador Then
'                Set MSComite.CellPicture = imgOK
'            Else
                Set MSComite.CellPicture = imgNN
'            End If
            MSComite.TextMatrix(i, 1) = Rs!nProselNro
            MSComite.TextMatrix(i, 2) = Rs!cPersCod
            MSComite.TextMatrix(i, 3) = Rs!cConsDescripcion
            MSComite.TextMatrix(i, 4) = Rs!cPersNombre
            MSComite.TextMatrix(i, 5) = IIf(CBool(Rs!bSuplente), "SUPLENTE", "TITULAR")
            Rs.MoveNext
        Loop
        'MSComite.Col = 1
        'MSComite.row = 1
        MSComite.ColSel = MSComite.Cols - 1
        oCon.CierraConexion
    End If
    Exit Sub
cargarComiteItemProcesoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub cmdGrabarApelacion_Click()
    On Error GoTo cmdGrabarErr
    Dim oConn As New DConecta, Rs As ADODB.Recordset, sSQL As String
    Dim nItemApe As Integer

If oConn.AbreConexion Then

    'recuperar ultimo item de apelacion

    Set Rs = oConn.CargaRecordSet("Select nUltItem = coalesce(max(nItemApelacion),0) from LogProSelApelacion WHERE nProSelNro = " & gnProSelNro)
    If Not Rs.EOF Then
       nItemApe = Rs!nUltItem + 1
    Else
       nItemApe = 1
    End If

    If Rspta Then
        sSQL = "update LogProSelApelacion set bResuelto=1, cRespuesta ='" & txtRespuesta.Text & "'" & _
                " where nItemApelacion=" & txtnro
    Else
        sSQL = "insert into LogProSelApelacion(nProSelNro,nProSelItem,nItemApelacion,cPersCod,cApelacion,cRespuesta,bAdmision,bResuelto) " & _
                " values(" & gnProSelNro & "," & MSItem.TextMatrix(MSItem.row, 6) & "," & nItemApe & ",'" & txtPersCodApe.Text & "','" & txtApelacion.Text & "','',0,0)"
    End If
    Set Rs = oConn.CargaRecordSet(sSQL)
    oConn.CierraConexion

    frameApelacionesD.Visible = False
    FrameApelacionesLista.Visible = True

    CargarApelaciones gnProSelNro, MSItem.TextMatrix(MSItem.row, 6)
    MsgBox "Apelacion Registrada con éxito!!" + Space(10), vbInformation, "Aviso"

    cmdRespuesta.Enabled = True
    cmdCancelarApelacion_Click
End If
    Exit Sub
cmdGrabarErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdGrabarObs_Click()
Dim oConn As New DConecta, sSQL As String, xConsulta As String
On Error GoTo cmdGrabarObs_ClickErr

If Len(Trim(txtconsulta.Text)) = 0 Then Exit Sub
If txtPersCodObs.Text = "" Then Exit Sub

If Rspta Then
    Select Case nTipo
        Case 1, 6
            sSQL = "update LogProSelConsultas set nEstado=1, cConsulta='" & txtconsulta.Text & "', cRespuesta='" & rTxtResp.Text & "' where nProSelnro=" & gnProSelNro & " and cPersCod='" & MSFObsCon.TextMatrix(MSFObsCon.row, 1) & "' and cMovNro='" & MSFObsCon.TextMatrix(MSFObsCon.row, 6) & "'"
        Case 2, 7
            sSQL = "update LogProSelObsBases set nEstado=1, cObservacion='" & txtconsulta.Text & "', cRespuesta='" & rTxtResp.Text & "' where nProSelnro=" & gnProSelNro & " and cPersCod='" & MSFObsCon.TextMatrix(MSFObsCon.row, 1) & "' and cMovNro='" & MSFObsCon.TextMatrix(MSFObsCon.row, 6) & "'"
    End Select
Else
    Select Case nTipo
        Case 1
            sSQL = " insert into LogProSelConsultas (nProSelNro,nTipo,cPersCod,cConsulta,cRespuesta,cMovNro) " & _
                    " values (" & gnProSelNro & "," & nTipo & ",'" & txtPersCodObs & "','" & txtconsulta.Text & "','','" & GetLogMovNro & "')"
                   '" values (" & nProSelNro & "," & nTipo & ",'" & txtPersCodObs & "','" & txtConsulta & "','" & txtrespuesta & "' )"
        Case 2
            sSQL = " insert into LogProSelObsBases (nProSelNro,cPersCod,cObservacion,cRespuesta,cMovNro) " & _
                    " values (" & gnProSelNro & ",'" & txtPersCodObs & "','" & txtconsulta.Text & "','','" & GetLogMovNro & "')"
                   '" values (" & nProSelNro & "," & nTipo & ",'" & txtPersCodObs & "','" & txtConsulta & "','" & txtrespuesta & "' )"
    End Select
End If
If oConn.AbreConexion Then
   If MsgBox("¿ Está seguro de agregar la " & lblDesc.Caption & " del Postor ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
      If sSQL <> "" Then
          oConn.Ejecutar sSQL
          oConn.CierraConexion
    '      fraPos.Visible = False
    '      fraVis.Visible = True
        Select Case nTipo
            Case 1, 6
                CargaConsultas
            Case 2, 7
                CargaObservaciones
        End Select
        cmdCancelaObs_Click
      End If
   End If
End If
Exit Sub
cmdGrabarObs_ClickErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdGrabarPos_Click()
Dim oConn As New DConecta, sSQL As String, i As Integer, _
    nItem As Integer, Rs As ADODB.Recordset, nCostoBases As Currency
On Error GoTo GrabarPos

'sSQL = " insert into LogProSelPostorPropuesta (nProSelNro,nProSelItem,cPersCod,nPropEconomica,nPuntaje,bGanador) " & _
'       " values (" & nProSelNro & "," & nProSelItem & ",'" & txtPersCod & "'," & VNumero(txtPropEcon) & "," & VNumero(txtPuntaje) & "," & chkGanador.Value & ")"

If Not VerificaVentaBases(gnProSelNro) And Len(Trim(txtnrorecibo.Text)) > 0 Then
    MsgBox "No Debe Ingresar el Nro de Recibo, las Bases no se Venden", vbInformation, "Aviso"
    Exit Sub
ElseIf VerificaVentaBases(gnProSelNro) And Len(Trim(txtnrorecibo.Text)) = 0 Then
    MsgBox "Debe Ingresar el Nro de Recibo, las Bases se Venden", vbInformation, "Aviso"
    Exit Sub
End If

If Len(Trim(txtPersCod.Text)) = 0 Then
   MsgBox "Debe ingresar un postor válido..." + Space(10), vbInformation, "Aviso"
   txtPersCod.SetFocus
   Exit Sub
End If

'If Len(Trim(txtnrorecibo.Text)) > 0 And CDbl(txtimpoteRecibo.Text) = 0 Then
'    MsgBox "Nro de Recibo no Existe...", vbInformation, "Aviso"
'    TxtSerie.Text = ""
'    txtnrorecibo.Text = ""
'    TxtSerie.SetFocus
'    Exit Sub
'End If

'nCostoBases = CargarCostoBases(gnProSelNro)

If Not VerificarCostoBases(gnProSelNro, Val(txtimpoteRecibo.Text), cboMonedaRecibo.ItemData(cboMonedaRecibo.ListIndex)) Then
    MsgBox "Importe de Recibo Incorrecto ", vbInformation, "Aviso"
    txtimpoteRecibo.SetFocus
    Exit Sub
End If

If MsgBox("¿ Está seguro de agregar un Participante ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   If oConn.AbreConexion Then
        If Val(txtnrorecibo.Text) > 0 And Val(TxtSerie.Text) > 0 Then
            sSQL = "select count(*) from LogProSelPostor where cNroRecibo='" & TxtSerie.Text & txtnrorecibo.Text & "'"
            Set Rs = oConn.CargaRecordSet(sSQL)
            If Rs(0) <> 0 Then
                oConn.CierraConexion
                MsgBox "Error el Nro de Recibo ya fue Registrado...", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        sSQL = " insert into LogProSelPostor (nProSelNro,cPersCod,cNroRecibo,dFechaRecibo,nImporteRecibo,cDescripcionRecibo,nMonedaRecibo, cEmail, cPersRep) " & _
               " values (" & gnProSelNro & ",'" & txtPersCod & "','" & TxtSerie.Text & "-" & txtnrorecibo.Text & "','" & _
               IIf(txtFechaRecibo.Text = "__/__/____", Null, Format(txtFechaRecibo.Text, "yyyymmdd")) & "'," & Val(txtimpoteRecibo.Text) & ",'" & txtDescRecibo.Text & "'," & cboMonedaRecibo.ItemData(cboMonedaRecibo.ListIndex) & ",'" & txtEmail.Text & "','" & txtPersRep.Text & "' )"
        oConn.Ejecutar sSQL
        oConn.CierraConexion
    End If
    fraPos.Visible = False
    FrPostores.Visible = True
    cmdImprimir.Enabled = True
    CargarPostores gnProSelNro
    If MSItemPostores.Enabled Then MSItemPostores.SetFocus
End If

Exit Sub
    
GrabarPos:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Function VerificarCostoBases(ByVal pnProSelNro As Integer, ByVal pnImporte As Currency, ByVal pnMoneda As Integer) As Boolean
On Error GoTo CargarCostoBasesErr
    Dim oCon As DConecta, Rs As ADODB.Recordset, sSQL As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nMonedaCostoBases, nCostoBases from LogProcesoSeleccion where nProSelNro = " & pnProSelNro
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            If pnImporte = Rs!nCostoBases Then
                If pnImporte = 0 Then
                    VerificarCostoBases = True
                ElseIf pnMoneda = Rs!nMonedaCostoBases Then
                    VerificarCostoBases = True
                Else
                    VerificarCostoBases = False
                End If
            Else
                VerificarCostoBases = False
            End If
        End If
        oCon.CierraConexion
    End If
    Exit Function
CargarCostoBasesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Sub cmdGuardarDatosAP_Click()
On Error GoTo GuardarDatosAPErr
    Dim oCon As DConecta, sSQL As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "update LogProcesoSeleccion set cNotario='" & txtNotario.Text & _
                "', cVeedor='" & txtveedor.Text & "', cLugarEjecucion='" & _
                    txtLugar.Text & "', tOcurrencias='" & txtocurrencia.Text & "'" & _
               " where nProSelNro = " & gnProSelNro
        oCon.Ejecutar sSQL
        MsgBox "Datos Registrados Correctamente", vbInformation
        oCon.CierraConexion
    End If
    Exit Sub
GuardarDatosAPErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub cmdImprimir_Click()
    If MSFObsCon.TextMatrix(MSFObsCon.row, 1) = "" Then Exit Sub
    Select Case nTipo
        Case 6
            ImpConsultasWord gnProSelNro
        Case 7
            ImpObservacionesWord gnProSelNro
    End Select
End Sub

Private Sub cmdimprimirpostores_Click()
    If gnProSelNro = 0 Then Exit Sub
    Select Case nTipo
        Case 3
            'ImpRegParticipantesWord gnProSelNro, MSItemPostores.TextMatrix(MSItemPostores.row, 2), TxtTipo.Text & " N°:" & TxtProSelNro & "-" & txtAnio.Text & "-CMAC-T", txtDescripcion.Text
'            ImprimePostores gnProSelNro, TxtTipo & " N° " & TxtProSelNro.Text & "-" & gnAnio & "-CMAC-T", TxtDescripcion.Text, LblMoneda, CDbl(TxtMonto)
             ImpRegParticipantesWord gnProSelNro, TxtDescripcion.Text, txtPersona.Text, txtRUC.Text, _
                         txtDomicilio.Text, txtTelefono.Text, txtEmail.Text, _
                         txtRepresenta.Text, txtDNI.Text, txtFechaRecibo.Text

        Case 8
            ImprimePostores gnProSelNro, TxtTipo & " N° " & TxtProSelNro.Text & "-" & txtanio.Text & "-CMAC-T", TxtDescripcion.Text, LblMoneda, CDbl(TxtMonto)
    End Select
End Sub

Private Sub cmdModificarConOns_Click()
    On Error GoTo cmdResponderConOnsErr
    With MSFObsCon
        If .TextMatrix(.row, 1) = "" Then Exit Sub
        txtPersCodObs.Text = .TextMatrix(.row, 1)
        txtPersObs.Text = .TextMatrix(.row, 3)
        txtconsulta.Text = .TextMatrix(.row, 4)
        rTxtResp.Text = .TextMatrix(.row, 5)
    End With
        cmdPersObs.Enabled = False
        txtPersCodObs.Enabled = False
        txtPersObs.Enabled = False
        txtconsulta.Enabled = True
        fraObs.Visible = True
        FrameConObs.Visible = False
        Rspta = True
        '.Show 1
    'End With
    Exit Sub
cmdResponderConOnsErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdPersApe_Click()
'frmBuscaPersona.Show 1
'If frmBuscaPersona.vpOK Then
'   txtPersApe.Text = frmBuscaPersona.vpPersNom
'   txtPersCodApe = frmBuscaPersona.vpPersCod
'End If
With frmLogProSelCnsPostores
    .Inicio gnProSelNro, ""
    txtPersCodApe.Text = .gcPersCod
    txtPersApe.Text = .gcPersNombre
'    If txtPersCodObs.Text <> "" Then
'        MSItem.Enabled = True
''        MSItem.SetFocus
'    End If
End With
End Sub

Private Sub cmdPersObs_Click()
'Dim X As UPersona
'Set X = frmBuscaPersona.Inicio
'
'If X Is Nothing Then
'    Exit Sub
'End If
'
'If Len(Trim(X.sPersNombre)) > 0 Then
'   txtPersona.Text = X.sPersNombre
'   txtPersCod = X.sPersCod
'End If

'frmBuscaPersona.Show 1
'If frmBuscaPersona.vpOK Then
'   txtPersObs.Text = frmBuscaPersona.vpPersNom
'   txtPersCodObs = frmBuscaPersona.vpPersCod
'End If
'With frmLogCnsPostores
'    .Inicio gnProSelNro, cCadenaPostores
'    txtPersCod.Text = .gcPersCod
'    txtPersona.Text = .gcPersNombre
'    If txtPersCod.Text <> "" Then
'        MSItem.Enabled = True
'        MSItem.SetFocus
'    End If
'End With
With frmLogProSelCnsPostores
    .Inicio gnProSelNro, ""
    txtPersCodObs.Text = .gcPersCod
    txtPersObs.Text = .gcPersNombre
    If txtPersCodObs.Text <> "" Then
        MSItem.Enabled = True
'        MSItem.SetFocus
    End If
End With
End Sub

'Private Sub cmdPostor_Click()
'Dim i As Integer
'i = MSItem.Row
'If Len(MSItem.TextMatrix(i, 0)) > 0 And Len(MSItem.TextMatrix(i, 1)) > 0 Then
'   frmLogRegistroDatosItem.Inicio MSFlex.TextMatrix(MSFlex.Row, 4), MSItem.TextMatrix(i, 1), False, 2
'Else
'   MsgBox "No se halla un Proceso/Item válido..." + Space(10), vbInformation
'End If
'End Sub

Private Sub cmdPersona_Click()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio

If X Is Nothing Then
    Exit Sub
End If

If Len(Trim(X.sPersNombre)) > 0 Then
   txtPersona.Text = X.sPersNombre
   txtDomicilio.Text = X.sPersDireccDomicilio
   txtTelefono.Text = X.sPersTelefono
   txtPersCod = X.sPersCod
   txtRUC.Text = X.sPersIdnroRUC
End If

'frmBuscaPersona.Show 1
'If frmBuscaPersona.vpOK Then
'   txtPersona.Text = frmBuscaPersona.vpPersNom
'   txtPersCod = frmBuscaPersona.vpPersCod
''   valida si el postor ya esta para el item
'End If
End Sub

Private Sub cmdPersRep_Click()
Dim X As UPersona
Set X = frmBuscaPersona.Inicio

If X Is Nothing Then
    Exit Sub
End If

If Len(Trim(X.sPersNombre)) > 0 Then
   txtRepresenta.Text = X.sPersNombre
   txtPersRep = X.sPersCod
   txtDNI.Text = X.sPersIdnroDNI
End If
End Sub

Private Sub cmdQuitar_Click()
On Error GoTo cmdQuitarErr
    
    If gnProSelNro = 0 Then Exit Sub
    If MSFObsCon.TextMatrix(MSFObsCon.row, 1) = "" Then Exit Sub
    
    Dim oCon As DConecta, sSQL As String
    If MsgBox("Seguro que Desea Eliminar...", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Select Case nTipo
        Case 1
            sSQL = "delete LogProSelConsultas where nProSelnro=" & gnProSelNro & " and cPersCod='" & MSFObsCon.TextMatrix(MSFObsCon.row, 1) & "' and cMovNro='" & MSFObsCon.TextMatrix(MSFObsCon.row, 6) & "'"
        Case 2
            sSQL = "delete LogProSelObsBases where nProSelnro=" & gnProSelNro & " and cPersCod='" & MSFObsCon.TextMatrix(MSFObsCon.row, 1) & "' and cMovNro='" & MSFObsCon.TextMatrix(MSFObsCon.row, 6) & "'"
    End Select
    Set oCon = New DConecta
    If sSQL <> "" Then
        If oCon.AbreConexion Then
            oCon.Ejecutar sSQL
            MsgBox "Se Elimino Satisfactoriamente...", vbInformation
            oCon.CierraConexion
        End If
        Select Case nTipo
            Case 1
                CargaConsultas
            Case 2
                CargaObservaciones
        End Select
    End If
Exit Sub
cmdQuitarErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdquitarAp_Click()
Dim i As Integer
Dim k As Integer

If gnProSelNro = 0 Then Exit Sub

i = MSApe.row
If Len(Trim(MSApe.TextMatrix(i, 2))) = 0 Then
   Exit Sub
End If

If MsgBox("¿ está seguro de quitar el elemento ?" + Space(10), vbQuestion + vbYesNo, "Confirme") = vbYes Then
   EliminarApelacion MSApe.TextMatrix(i, 0), gnProSelNro
   If MSApe.Rows - 1 > 1 Then
      MSApe.RemoveItem i
   Else
      'MSFlex.Clear          Quita las cabeceras
      For k = 0 To MSApe.Cols - 1
          MSApe.TextMatrix(i, k) = ""
      Next
      MSApe.RowHeight(i) = 8
      If MSApe.Rows < 2 Then cmdRespuesta.Enabled = False
   End If
End If
End Sub

Private Sub cmdQuitarPostor_Click()
On Error GoTo cmdQuitarErr
    Dim oCon As DConecta, sSQL As String
    
    If Len(TxtProSelNro) = 0 Then
       MsgBox "Debe seleccionar un proceso de selección..." + Space(10), vbInformation
    End If
    
    If gnProSelNro = 0 Then Exit Sub
    
    If MsgBox("Seguro que Desea Eliminar un Postor ...", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "delete LogProSelPostor where nProSelNro=" & gnProSelNro & " and cPersCod='" & MSItemPostores.TextMatrix(MSItemPostores.row, 2) & "'"
        oCon.Ejecutar sSQL
        oCon.CierraConexion
        MsgBox "Se Elimino el Postor...", vbInformation
        CargarPostores gnProSelNro
    End If
    Exit Sub
cmdQuitarErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub CmdRegistrar_Click()
Dim oCon As DConecta, sSQL As String, i As Integer

On Error GoTo CmdRegistrarErr

If MsgBox("¿ Está seguro de registrar la propuesta ?" + Space(10), vbQuestion + vbYesNo, "confirme operación") = vbYes Then
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        oCon.BeginTrans
        i = 1
        With MSItemPostores
            .Col = 0
            Do While i < .Rows
                .row = i
                If .CellPicture = imgOK Then
                    
                    sSQL = "update LogProSelPostor set nPresentoProp = 1 where nProSelNro= " & gnProSelNro & " and cPersCod='" & .TextMatrix(i, 2) & "'"
                    oCon.Ejecutar sSQL
                    
                    sSQL = "insert into LogProSelPostorPropuesta(nProSelNro, nProSelItem, cPersCod) " & _
                           "select " & gnProSelNro & ", nProSelItem, '" & .TextMatrix(i, 2) & "' from LogProSelItem where nProSelNro = " & gnProSelNro
                    oCon.Ejecutar sSQL
                End If
                i = i + 1
            Loop
            oCon.CommitTrans
            .ColSel = .Cols - 1
        End With
        CierraEtapa gnProSelNro, cnAbsolucionConsultas
        CierraEtapa gnProSelNro, cnObservaciones
        CierraEtapa gnProSelNro, cnConvocatoria
        CierraEtapa gnProSelNro, cnRegistroParticipantes
        CierraEtapa gnProSelNro, cnPresentacionPropuestas
        CmdRegistrar.Enabled = False
        MSItemPostores.Enabled = False
        oCon.CierraConexion
    End If
    MsgBox "Registro de Propuestas Terminado", vbInformation
End If

Exit Sub
CmdRegistrarErr:
    oCon.RollbackTrans
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Sub cmdregistrarAp_Click()
    Rspta = False
    If gnProSelNro = 0 Then Exit Sub
    cmdPersApe.Enabled = True
    txtnro.Enabled = True
    txtPersCodApe.Enabled = True
    txtPersApe.Enabled = True
    txtApelacion.Enabled = True
    
    frameApelacionesD.Visible = True
    FrameApelacionesLista.Visible = False
End Sub

Private Sub cmdResponderConOns_Click()
On Error GoTo cmdResponderConOnsErr
'    With frmItemApelaciones
    With MSFObsCon
        If .TextMatrix(.row, 1) = "" Then Exit Sub
        txtPersCodObs.Text = .TextMatrix(.row, 1)
        txtPersObs.Text = .TextMatrix(.row, 3)
        txtconsulta.Text = .TextMatrix(.row, 4)
        rTxtResp.Text = .TextMatrix(.row, 5)
    End With
        rTxtResp.Enabled = True
        cmdPersObs.Enabled = False
        txtPersCodObs.Enabled = False
        txtPersObs.Enabled = False
        txtconsulta.Enabled = False
        fraObs.Visible = True
        FrameConObs.Visible = False
        Rspta = True
        '.Show 1
    'End With
    Exit Sub
cmdResponderConOnsErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdRespuesta_Click()
On Error GoTo cmdRespuestaErr
'    With frmItemApelaciones
        txtnro.Text = MSApe.TextMatrix(MSApe.row, 0)
        txtPersCodApe.Text = MSApe.TextMatrix(MSApe.row, 1)
        txtPersApe.Text = MSApe.TextMatrix(MSApe.row, 2)
        txtApelacion.Text = MSApe.TextMatrix(MSApe.row, 3)
        txtRespuesta.Text = MSApe.TextMatrix(MSApe.row, 4)

        cmdPersApe.Enabled = False
        txtnro.Enabled = False
        txtPersCodApe.Enabled = False
        txtPersApe.Enabled = False
        txtApelacion.Enabled = False
        txtRespuesta.Enabled = True
        frameApelacionesD.Visible = True
        FrameApelacionesLista.Visible = False
        Rspta = True
        '.Show 1
    'End With
    Exit Sub
cmdRespuestaErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub EliminarApelacion(ByRef nIA As Integer, ByRef nPSI As Integer)
    On Error GoTo EliminarApelacionErr
    Dim oConn As New DConecta, sSQL As String
    If oConn.AbreConexion Then
        sSQL = "delete from LogProSelApelacion where nItemApelacion=" & nIA & " and nProSelNro=" & nPSI
        oConn.Ejecutar sSQL
        oConn.CierraConexion
    End If
    Exit Sub
EliminarApelacionErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub Command1_Click()

End Sub


Private Sub CargarMonedaRecibo()
On Error GoTo CargarMonedaBasesErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nConsValor, cConsDescripcion from constante where nConsCod = 1011"
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            cboMonedaRecibo.AddItem IIf(Rs!cConsDescripcion = "SOLES", "S/.", "$"), cboMonedaRecibo.ListCount
            cboMonedaRecibo.ItemData(cboMonedaRecibo.ListCount - 1) = Rs!nConsValor
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    If cboMonedaRecibo.ListCount > 0 Then cboMonedaRecibo.ListIndex = 0
    Exit Sub
CargarMonedaBasesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub


'Sub GeneraMeses()
'Dim oConn As New DConecta, Rs As New ADODB.Recordset
'
'If oConn.AbreConexion Then
'   cboMes.Clear
'   sSql = "select cMes = rtrim(substring(cNomTab,1,12)) from DBComunes..TablaCod where cCodTab like 'EZ%' and len(cCodTab)=4"
'   Set Rs = oConn.CargaRecordSet(sSql)
'   If Not Rs.EOF Then
'      Do While Not Rs.EOF
'         cboMes.AddItem Rs!cMes
'         Rs.MoveNext
'      Loop
'   End If
'   cboMes.ListIndex = 0
'End If
'End Sub

'Private Sub cboMes_Click()
'nMes = cboMes.ListIndex + 1
'FormaMSFObsCon
'FormaMSApe
'ListaProcesosMensual nMes, CInt(txtAnio), cboMes.Text
'End Sub



'Sub ListaProcesosMensual(vMes As Integer, vAnio As Integer, vMesDesc As String)
'Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Currency
'
'sSQL = ""
'nSuma = 0
'FormaFlex
'nAnio = CInt(txtAnio.Text)
'nMes = cboMes.ListIndex + 1
'
'If oConn.AbreConexion Then
'
'   'Equivalente en el PLAN ANUAL
'   'sSQL = "select nPlanAnualNro, cProSel=coalesce(r.cAbreviatura,''), cObjeto=b.cBSDescripcion,  cFuenteFin = coalesce(f.cFuenteFin,''), " & _
'   '"       cAgencias=coalesce(a.cAgeDescripcion,'TODAS'), p.*, b.cBSGrupoCod " & _
'   '"  from LogPlanAnualDetalle p inner join LogProSelBienesServicios b  on p.cObjetoCod = b.cProSelBSCod " & _
'   '"       left outer join (select nProSelTpoCod,nProSelSubTpo,cAbreviatura from LogProSelTpoRangos) r on  r.nProSelTpoCod = p.nProSelTpoCod and r.nProSelSubTpo = p.nProSelSubTpo " & _
'   '"       left outer join (select nConsValor as nFuenteFinCod, cFuenteFin=cConsDescripcion from Constante where nConsCod = " & gcAreaAprobacion & " and nConsCod<>nConsValor) f on p.nFuenteFinCod = f.nFuenteFinCod " & _
'   '"       left outer join Agencias a on p.cAgeCod = a.cAgeCod " & _
'   '" where p.nPlanAnualEstado = 1 and p.nPlanAnualMes = " & vMes & " and p.nPlanAnualAnio = " & vAnio & " "
'
'   sSQL = "select p.*, t.cProSelTpoDescripcion as cProceso" & _
'   "  from LogProcesoSeleccion p " & _
'   "       inner join LogProSelTpo t on p.nProSelTpoCod = t.nProSelTpoCod " & _
'   " Where p.nPlanAnualAnio = " & nAnio & " And p.nPlanAnualMes = " & nMes & " and p.nProSelEstado=1"
'
'   If Len(sSQL) = 0 Then Exit Sub
'
'   Set Rs = oConn.CargaRecordSet(sSQL)
'   If Not Rs.EOF Then
'      Do While Not Rs.EOF
'         i = i + 1
'         InsRow MSFlex, i
'         MSFlex.RowHeight(i) = 500
'         MSFlex.TextMatrix(i, 0) = Rs!nPlanAnualNro
'         MSFlex.TextMatrix(i, 1) = Rs!nPlanAnualItem
'         MSFlex.TextMatrix(i, 2) = Rs!nProSelTpoCod
'         MSFlex.TextMatrix(i, 3) = Rs!nProSelSubTpo
'         MSFlex.TextMatrix(i, 4) = Rs!nProSelNro
'         MSFlex.TextMatrix(i, 5) = Rs!cSintesis
'         MSFlex.TextMatrix(i, 6) = Rs!cProceso
'         MSFlex.TextMatrix(i, 7) = IIf(Rs!nMoneda = 2, "DOLARES", "SOLES")
'         MSFlex.TextMatrix(i, 8) = FNumero(Rs!nMontoRef)
'         MSFlex.TextMatrix(i, 9) = Rs!cBSGrupocod
'         MSFlex.TextMatrix(i, 10) = Rs!cArchivoBases
'         nSuma = nSuma + Rs!nMontoRef
'         Rs.MoveNext
'      Loop
'        MSFlex.SetFocus
'   End If
'End If
'End Sub

'Sub FormaFlex()
'MSFlex.Clear
'MSFlex.Rows = 2
'MSFlex.RowHeight(0) = 360
'MSFlex.RowHeight(1) = 8
'MSFlex.ColWidth(0) = 0
'MSFlex.ColWidth(1) = 0:     MSFlex.ColAlignment(1) = 4
'MSFlex.ColWidth(2) = 0
'MSFlex.ColWidth(3) = 0
'MSFlex.ColWidth(4) = 400:   MSFlex.TextMatrix(0, 4) = "Nro":     MSFlex.ColAlignment(4) = 4
'MSFlex.ColWidth(5) = 4000:  MSFlex.TextMatrix(0, 5) = ""
'MSFlex.ColWidth(6) = 3800:  MSFlex.TextMatrix(0, 6) = "Proceso"
'MSFlex.ColWidth(7) = 1000:   MSFlex.TextMatrix(0, 7) = "      Moneda"
'MSFlex.ColWidth(8) = 1100:  MSFlex.TextMatrix(0, 8) = "           Monto"
'MSFlex.WordWrap = True
'End Sub
'

'Private Sub cmdArchivo_Click()
'Dim oConn As New DConecta, nProSelNro As Integer
'
'dlgArchivo.FileName = "*.txt"
'dlgArchivo.Filter = "*.txt"
'dlgArchivo.ShowOpen
'If Len(Trim(dlgArchivo.FileName)) > 0 Then
'   txtArchivo.Text = dlgArchivo.FileName
'   rtfBases.LoadFile txtArchivo.Text
'   nProSelNro = MSFlex.TextMatrix(MSFlex.Row, 4)
'   If oConn.AbreConexion Then
'      sSQL = "UPDATE LogProcesoSeleccion SET cArchivoBases = '" & txtArchivo & "' " & _
'             " WHERE nProSelNro = " & nProSelNro & " "
'      oConn.Ejecutar sSQL
'      MSFlex.TextMatrix(MSFlex.Row, 10) = txtArchivo
'   End If
'End If
'End Sub

'Private Sub cmdCara_Click()
'Dim i As Integer
'i = MSItem.Row
'If Len(MSItem.TextMatrix(i, 0)) > 0 And Len(MSItem.TextMatrix(i, 1)) > 0 Then
'   frmLogRegistroDatosItem.Inicio MSItem.TextMatrix(i, 0), MSItem.TextMatrix(i, 1), 1
'Else
'   MsgBox "No se halla un Proceso/Item válido..." + Space(10), vbInformation
'End If
'End Sub

'Private Sub cmdConsultas_Click()
'Dim i As Integer
'
'i = MSFlex.Row
'If Len(MSFlex.TextMatrix(i, 4)) > 0 Then
'   'observaciones al proceso y a las bases ----------------------
'   'solo dependen del nro de proceso
'   frmLogRegistroDatosItem.Inicio MSFlex.TextMatrix(i, 4), 0, 3
'Else
'   MsgBox "No se halla un proceso válido..." + Space(10), vbInformation
'End If
'End Sub

'Private Sub cmdEtapas_Click()
'If Len(MSFlex.TextMatrix(MSFlex.Row, 4)) > 0 And Len(MSEta.TextMatrix(MSEta.Row, 0)) > 0 Then
'   frmLogRegistroDatosItem.Inicio MSFlex.TextMatrix(MSFlex.Row, 4), MSEta.TextMatrix(MSEta.Row, 0), 4, MSEta.TextMatrix(MSEta.Row, 2)
'   If frmLogRegistroDatosItem.vpGrabado Then
'      MSFlex_GotFocus
'   End If
'Else
'   MsgBox "No se halla un Proceso/Item válido..." + Space(10), vbInformation
'End If
'End Sub

'Private Sub cmdPostor_Click()
'Dim i As Integer
'i = MSItem.Row
'If Len(MSItem.TextMatrix(i, 0)) > 0 And Len(MSItem.TextMatrix(i, 1)) > 0 Then
'   frmLogRegistroDatosItem.Inicio MSItem.TextMatrix(i, 0), MSItem.TextMatrix(i, 1), 2
'Else
'   MsgBox "No se halla un Proceso/Item válido..." + Space(10), vbInformation
'End If
'End Sub

'Private Sub cmdProSel_Click()
'Dim oConn As New DConecta, rs As New ADODB.Recordset
'Dim rc As New ADODB.Recordset, rn As New ADODB.Recordset
'Dim nProSelNro As Integer, nPlanNro As Integer, nPlanItem As Integer
'Dim nProSelTpo As Integer, nProSelSub As Integer, cBSGrupoCod As String
'Dim nMonto As Currency, nMoneda As Integer, i As Integer, cProSelBSCod As String
'Dim nCant As Integer
'
'If MsgBox("¿ Generar Procesos de Selección para el Mes y Año indicados ?" + Space(10), vbQuestion + vbYesNo, "Confirmación") = vbNo Then
'   Exit Sub
'End If
'
'nAnio = CInt(txtAnio.Text)
'nMes = cboMes.ListIndex + 1
'
'If oConn.AbreConexion Then
'
'   cLogNro = GetLogMovNro
'
'   sSQL = " select nPlanAnualNro,nPlanAnualItem,nProSelTpoCod,nProSelSubTpo,nMoneda,nValorEstimado,cBSGrupoCod " & _
'          "  from LogPlanAnualDetalle where nPlanAnualMes = " & nMes & " and nPlanAnualAnio = " & nAnio & ""
'
'   Set rs = oConn.CargaRecordSet(sSQL)
'   If Not rs.EOF Then
'      Do While Not rs.EOF
'
'         nPlanNro = rs!nPlanAnualNro
'         nPlanItem = rs!nPlanAnualItem
'         nProSelTpo = rs!nProSelTpoCod
'         nProSelSub = rs!nProSelSubTpo
'         cBSGrupoCod = rs!cBSGrupoCod
'         nMoneda = rs!nMoneda
'         nMonto = rs!nValorEstimado
'
'         'ANULA PROCESOS DE SELECCION CREADOS ANTERIORMENTE PARA CADA ITEM
'         sSQL = "UPDATE LogProcesoSeleccion SET nProSelEstado = 0 WHERE nPlanAnualNro = " & nPlanNro & " AND nPlanAnualItem = " & nPlanItem & " "
'         oConn.Ejecutar sSQL
'
'         'CREA UN PROCESO DE SELECCION PARA CADA ITEM DEL PLAN ANUAL
'
'         sSQL = "INSERT INTO LogProcesoSeleccion (nPlanAnualNro,nPlanAnualItem,nProSelTpoCod,nProSelSubTpo,nMoneda,nMontoRef, cLogMovNro) " & _
'                " VALUES ( " & nPlanNro & "," & nPlanItem & "," & nProSelTpo & "," & nProSelSub & "," & nMoneda & "," & nMonto & ",'" & cLogNro & "')"
'         oConn.Ejecutar sSQL
'
'         Set rn = oConn.CargaRecordSet("Select nUlt=@@identity from LogProcesoSeleccion")
'         If Not rn.EOF Then
'            nProSelNro = rn!nUlt
'         End If
'
'         'CREACION DE ETAPAS DE CADA PROCESO DE SELECCION
'         sSQL = "INSERT INTO LogProSelEtapa (nProSelNro,nEtapaCod,nOrden,nEstado) " & _
'                " Select " & nProSelNro & ",nEtapaCod,nOrden,1 from LogProSelTpoEtapa where nProSelTpoCod = 1 order by nOrden "
'         oConn.Ejecutar sSQL
'
'
'         sSQL = "select v.cProSelBSCod,v.nCantidad " & _
'         "  from LogPlanAnualValor v inner join LogProSelBienesServicios b on v.cProSelBSCod = b.cProSelBSCod " & _
'         " where b.cBSGrupoCod = '" & cBSGrupoCod & "' and v.nEstado=1 "
'
'         i = 0
'         Set rc = oConn.CargaRecordSet(sSQL)
'         If Not rc.EOF Then
'            Do While Not rc.EOF
'               i = i + 1
'               nCant = rc!nCantidad
'               cProSelBSCod = rc!cProSelBSCod
'               sSQL = "INSERT INTO LogProSelItem (nProSelNro,nProSelItem,cProSelBSCod,nCantidad) " & _
'                      " VALUES (" & nProSelNro & "," & i & ",'" & cProSelBSCod & "'," & nCant & ") "
'               oConn.Ejecutar sSQL
'               rc.MoveNext
'            Loop
'         End If
'         rs.MoveNext
'      Loop
'   End If
'   MsgBox "Se han generado los procesos de Selección correctamente!!" + Space(10), vbInformation
'End If
'End Sub

'Private Sub MSFlex_Click()
'
'End Sub

Sub CargaConsultas()
Dim oConn As New DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
On Error GoTo CargaProp
   
FormaMSFObsCon
If oConn.AbreConexion Then
          
   sSQL = "select x.nProSelNro, x.nTipo, x.cPersCod, cPersNombre=replace(p.cPersNombre,'/',' '),x.cConsulta, x.cRespuesta, cMovNro  " & _
          " from LogProSelConsultas x inner join Persona p on x.cPersCod = p.cPersCod " & _
          " Where x.nProSelNro = " & gnProSelNro & ""
          
   Set Rs = oConn.CargaRecordSet(sSQL)
   i = 0
   Do While Not Rs.EOF
        i = i + 1
        InsRow MSFObsCon, i
        MSFObsCon.RowHeight(i) = 900
        MSFObsCon.TextMatrix(i, 1) = Rs!cPersCod
        MSFObsCon.TextMatrix(i, 2) = "Consulta"
        MSFObsCon.TextMatrix(i, 3) = Rs!cPersNombre
        MSFObsCon.TextMatrix(i, 4) = Rs!cConsulta
        MSFObsCon.TextMatrix(i, 5) = IIf(IsNull(Rs!cRespuesta), "", Rs!cRespuesta)
        MSFObsCon.TextMatrix(i, 6) = Rs!cMovNro
        MSFObsCon.ScrollBars = flexScrollBarBoth
        Rs.MoveNext
        cmdImprimir.Enabled = True
    Loop
End If
Exit Sub
CargaProp:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Sub CargaObservaciones()
Dim oConn As New DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
On Error GoTo CargaProp
   
FormaMSFObsCon
If oConn.AbreConexion Then
          
   sSQL = "select x.nProSelNro, x.cPersCod, cPersNombre=replace(p.cPersNombre,'/',''),x.cObservacion, x.cRespuesta, cMovNro  " & _
          " from LogProSelObsBases x inner join Persona p on x.cPersCod = p.cPersCod " & _
          " Where x.nProSelNro = " & gnProSelNro & ""
          
   Set Rs = oConn.CargaRecordSet(sSQL)
   i = 0
   Do While Not Rs.EOF
        i = i + 1
        InsRow MSFObsCon, i
        MSFObsCon.RowHeight(i) = 900
        MSFObsCon.TextMatrix(i, 1) = Rs!cPersCod
        MSFObsCon.TextMatrix(i, 2) = "Observacion"
        MSFObsCon.TextMatrix(i, 3) = Rs!cPersNombre
        MSFObsCon.TextMatrix(i, 4) = Rs!cObservacion
        MSFObsCon.TextMatrix(i, 5) = IIf(IsNull(Rs!cRespuesta), "", Rs!cRespuesta)
        MSFObsCon.TextMatrix(i, 6) = Rs!cMovNro
        MSFObsCon.ScrollBars = flexScrollBarBoth
        Rs.MoveNext
        cmdImprimir.Enabled = True
    Loop
End If
Exit Sub
CargaProp:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

'Sub CargaObservaciones()
'Dim oConn As New DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
'On Error GoTo CargaProp
'
'FormaMSFObsCon
'If oConn.AbreConexion Then
'
'   sSQL = "select x.nProSelNro, x.nTipo, x.cPersCod, p.cPersNombre,x.cConsulta, x.cRespuesta  " & _
'          " from LogProSelConsultas x inner join Persona p on x.cPersCod = p.cPersCod " & _
'          " Where x.nProSelNro = " & MSFlex.TextMatrix(MSFlex.Row, 4) & ""
'
'   Set Rs = oConn.CargaRecordSet(sSQL)
'   i = 0
'   Do While Not Rs.EOF
'        If nTipo = 1 And Rs!nTipo = 1 Then
'            i = i + 1
'            InsRow MSFObsCon, i
'            MSFObsCon.RowHeight(i) = 280
'            MSFObsCon.TextMatrix(i, 2) = "Consulta"
'            MSFObsCon.TextMatrix(i, 3) = Rs!cPersNombre
'            MSFObsCon.TextMatrix(i, 4) = Rs!cConsulta
'            'MSFlex.TextMatrix(i, 5) = rs!cRespuesta
'            MSFObsCon.ScrollBars = flexScrollBarBoth
'        ElseIf nTipo = 2 And Rs!nTipo = 2 Then
'            i = i + 1
'            InsRow MSFObsCon, i
'            MSFObsCon.RowHeight(i) = 280
'            MSFObsCon.TextMatrix(i, 2) = "Observacion"
'            MSFObsCon.TextMatrix(i, 3) = Rs!cPersNombre
'            MSFObsCon.TextMatrix(i, 4) = Rs!cConsulta
'            'MSFlex.TextMatrix(i, 5) = rs!cRespuesta
'            MSFObsCon.ScrollBars = flexScrollBarBoth
'        End If
'        Rs.MoveNext
'    Loop
'End If
'Exit Sub
'CargaProp:
'    MsgBox Err.Number & vbCrLf & Err.Description
'End Sub

Sub FormaMSFObsCon()
    Dim i As Integer
MSFObsCon.Clear
MSFObsCon.Rows = 2
MSFObsCon.Cols = 7
MSFObsCon.RowHeight(0) = 320
MSFObsCon.RowHeight(1) = 10
MSFObsCon.ColWidth(0) = 0
MSFObsCon.ColWidth(1) = 0
MSFObsCon.ColWidth(2) = 1000:    MSFObsCon.TextMatrix(0, 2) = "Tipo"
MSFObsCon.ColWidth(3) = 3000:    MSFObsCon.TextMatrix(0, 3) = "Persona"
MSFObsCon.ColWidth(4) = 5000:    MSFObsCon.TextMatrix(0, 4) = "Descripción"
MSFObsCon.ColWidth(5) = 5000:    MSFObsCon.TextMatrix(0, 5) = "Respuesta"
MSFObsCon.ColWidth(6) = 0:       MSFObsCon.TextMatrix(0, 6) = "Mov Nro"
End Sub

'Private Sub MSFlex_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 And Not xEjecutar Then
'   frmLogPlanAnualDocDet.Inicio MSFlex.TextMatrix(MSFlex.Row, 0), MSFlex.TextMatrix(MSFlex.Row, 1), MSFlex.TextMatrix(MSFlex.Row, 2), MSFlex.TextMatrix(MSFlex.Row, 5), MSFlex.TextMatrix(MSFlex.Row, 7), MSFlex.TextMatrix(MSFlex.Row, 3), VNumero(MSFlex.TextMatrix(MSFlex.Row, 10))
'End If
'End Sub

Sub FormaFlexItemPostor()
With MSItemPostores
    .Clear
    .Rows = 2
    .Cols = 5
    .RowHeight(0) = 320
    .RowHeight(1) = 8
    .ColWidth(0) = 800:     .ColAlignment(1) = 4:   .TextMatrix(0, 0) = " Item"
    .ColWidth(1) = 1500:     .ColAlignment(1) = 4:   .TextMatrix(0, 1) = " Fecha"
    .ColWidth(2) = 1500:     .ColAlignment(2) = 4:   .TextMatrix(0, 2) = " Código"
    .ColWidth(3) = 4000:    .TextMatrix(0, 3) = " Nombre"
    .ColWidth(4) = 2000:    .TextMatrix(0, 4) = " Nro de Recibo"
End With
End Sub

Sub FormaFlexItem()
MSItem.Clear
MSItem.Rows = 2
MSItem.RowHeight(0) = 320
MSItem.RowHeight(1) = 8
MSItem.ColWidth(0) = 250
MSItem.ColWidth(1) = 850:   MSItem.ColAlignment(1) = 4:  MSItem.TextMatrix(0, 1) = " Item"
MSItem.ColWidth(2) = 0:   MSItem.ColAlignment(2) = 4:  MSItem.TextMatrix(0, 2) = " Código"
MSItem.ColWidth(3) = 8000:  MSItem.TextMatrix(0, 3) = " Descripción"
MSItem.ColWidth(4) = 800:  MSItem.TextMatrix(0, 4) = " Cantidad"
MSItem.ColWidth(5) = 0:  MSItem.TextMatrix(0, 5) = " nProSelNro"
MSItem.ColWidth(6) = 0:  MSItem.TextMatrix(0, 6) = " nProSelItem"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set frmLogProSelEjecucion = Nothing
End Sub

Private Sub MSItem_DblClick()
On Error GoTo MSItemErr
    Dim i As Integer, bTipo As Boolean
    With MSItem
        If Trim(.TextMatrix(.row, 0)) = "-" Then
           .TextMatrix(.row, 0) = "+"
           i = .row + 1
           bTipo = True
        ElseIf Trim(.TextMatrix(.row, 0)) = "+" Then
           .TextMatrix(.row, 0) = "-"
           i = .row + 1
           bTipo = False
        Else
            Exit Sub
        End If
        
        Do While i < .Rows
            If Trim(.TextMatrix(i, 0)) = "+" Or Trim(.TextMatrix(i, 0)) = "-" Then
                Exit Sub
            End If
            
            If bTipo Then
                .RowHeight(i) = 0
            Else
                .RowHeight(i) = 260
            End If
            i = i + 1
        Loop
    End With
Exit Sub
MSItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Sub GeneraDetalleItem(vProSelNro As Integer)
Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Currency
Dim sSQL As String, sGrupo As String

sSQL = ""
nSuma = 0
FormaFlexItem

If oConn.AbreConexion Then

    sSQL = "select v.nProSelNro, v.nProSelItem, b.cBSGrupoDescripcion, b.cBSGrupoCod,x.cBSCod, y.cBSDescripcion, x.nCantidad, x.nMonto " & _
            "from LogProSelItem v " & _
            "inner join BSGrupos b on v.cBSGrupoCod = b.cBSGrupoCod " & _
            "inner join LogProSelItemBS x on v.nProSelNro = x.nProSelNro and v.nProSelItem = x.nProSelItem " & _
            "inner join LogProSelBienesServicios y on x.cBSCod = y.cProSelBSCod " & _
            "where v.nProSelNro = " & vProSelNro & " order by v.nProSelItem, b.cBSGrupoDescripcion "
    
   Set Rs = oConn.CargaRecordSet(sSQL)
   If Not Rs.EOF Then
      Do While Not Rs.EOF
        If sGrupo <> Rs!nProSelItem Then
         sGrupo = Rs!nProSelItem
         i = i + 1
         InsRow MSItem, i
         MSItem.Col = 0
         MSItem.row = i
         MSItem.CellFontSize = 10
         MSItem.CellFontBold = True
         MSItem.TextMatrix(i, 0) = "+"
         MSItem.TextMatrix(i, 1) = Rs!nProSelItem
         MSItem.TextMatrix(i, 3) = Rs!cBSGrupoDescripcion
         MSItem.TextMatrix(i, 4) = ""
         MSItem.TextMatrix(i, 5) = Rs!nProselNro
         MSItem.TextMatrix(i, 6) = Rs!nProSelItem
        End If
        i = i + 1
        InsRow MSItem, i
        MSItem.RowHeight(i) = 0
        MSItem.TextMatrix(i, 1) = Rs!cBSCod
        MSItem.TextMatrix(i, 3) = Rs!cBSDescripcion
        MSItem.TextMatrix(i, 4) = Rs!nCantidad
        MSItem.TextMatrix(i, 5) = Rs!nProselNro
        MSItem.TextMatrix(i, 6) = Rs!nProSelItem
        Rs.MoveNext
      Loop
      MSItem.ColSel = MSItem.Cols - 1
   End If
   MSItem.SetFocus
End If
End Sub


'Sub GeneraDetalleItem(vProSelNro As Integer, vBSGrupoCod As String)
'Dim oConn As New DConecta, Rs As New ADODB.Recordset, i As Integer, nSuma As Currency
'Dim sSQL As String
'
'sSQL = ""
'nSuma = 0
'FormaFlexItem
'
'If oConn.AbreConexion Then
'
'   'sSQL = "select v.cProSelBSCod,b.cBSDescripcion, v.nCantidad, v.nPrecioUnitario " & _
'   '       "  from LogPlanAnualValor v inner join LogProSelBienesServicios b on v.cProSelBSCod = b.cProSelBSCod " & _
'   '       " where b.cBSGrupoCod = '01' "
'
'   sSQL = "select v.nProSelNro, v.nProSelItem, v.cProSelBSCod,b.cBSDescripcion, v.nCantidad  " & _
'          "  from LogProSelItem v inner join LogProSelBienesServicios b on v.cProSelBSCod = b.cProSelBSCod " & _
'          " where v.nProSelNro = " & vProSelNro & "  and b.cBSGrupoCod = '" & vBSGrupoCod & "' "
'
'   Set Rs = oConn.CargaRecordSet(sSQL)
'   If Not Rs.EOF Then
'      Do While Not Rs.EOF
'         i = i + 1
'         InsRow MSItemPostores, i
'         'MSItem.TextMatrix(i, 0) = "+ "
'         MSItemPostores.TextMatrix(i, 0) = Rs!nProSelNro
'         MSItemPostores.TextMatrix(i, 1) = Rs!nProSelItem
'         MSItemPostores.TextMatrix(i, 2) = Rs!cProSelBSCod
'         MSItemPostores.TextMatrix(i, 3) = Rs!cBSDescripcion
'         MSItemPostores.TextMatrix(i, 4) = Rs!nCantidad
'         Rs.MoveNext
'      Loop
'   End If
'End If
'End Sub


'Private Sub MSFlex_GotFocus()
'If Len(Trim(MSFlex.TextMatrix(MSFlex.Row, 0))) > 0 And Len(Trim(MSFlex.TextMatrix(MSFlex.Row, 1))) > 0 Then
'    Select Case nTipo
'        Case 1
'            lblDesc.Caption = "Consulta"
''            CargaObservaciones
'            CargaConsultas
'        Case 2
'            lblDesc.Caption = "Observación"
'            'CargaObservaciones
'            CargaObservaciones
'        Case 3
'            GeneraDetalleItem MSFlex.TextMatrix(MSFlex.Row, 4), MSFlex.TextMatrix(MSFlex.Row, 9)
'        Case 4
'            CargarApelaciones MSFlex.TextMatrix(MSFlex.Row, 4), MSFlex.TextMatrix(MSFlex.Row, 1)
'    End Select
'End If
'End Sub

'Private Sub MSFlex_RowColChange()
'If Len(Trim(MSFlex.TextMatrix(MSFlex.Row, 0))) > 0 And Len(Trim(MSFlex.TextMatrix(MSFlex.Row, 1))) > 0 Then
'    Select Case nTipo
'        Case 1
'            lblDesc.Caption = "Consulta"
''            CargaObservaciones
'            CargaConsultas
'        Case 2
'            lblDesc.Caption = "Observación"
'            'CargaObservaciones
'            CargaObservaciones
'        Case 3
'            GeneraDetalleItem MSFlex.TextMatrix(MSFlex.Row, 4), MSFlex.TextMatrix(MSFlex.Row, 9)
'        Case 4
'            CargarApelaciones MSFlex.TextMatrix(MSFlex.Row, 4), MSFlex.TextMatrix(MSFlex.Row, 1)
'    End Select
'End If
'End Sub

'Sub CargaArchivoBases(vArchivo As String)
'Dim oConn As New DConecta
'Dim fs
'
'Set fs = CreateObject("Scripting.FileSystemObject")
'
'
''Set a = fs.CreateTextFile("c:\archivoprueba.txt", True)
''a.WriteLine ("Esto es una prueba.")
''a.Close
'
'
'On Error GoTo Salida
'
'rtfBases.Text = ""
'If Len(Trim(vArchivo)) > 0 Then
'   txtArchivo.Text = vArchivo
'   If fs.fileExists(vArchivo) Then
'      rtfBases.LoadFile vArchivo
'   Else
'      MsgBox "El archivo [" & vArchivo & "] no existe..." + Space(10), vbInformation
'   End If
'End If
'Exit Sub
'
'Salida:
'   MsgBox "Error: " + Err.Description
'End Sub

'
'Sub ListaEtapas(nProSelNro As Integer)
''Sub ListaEtapas(nPlanNro As Integer, nPlanItem As Integer)
'Dim oConn As New DConecta, rs As New ADODB.Recordset, i As Integer, nSuma As Currency
'
'nSuma = 0
'FormaFlexEta
'
'If oConn.AbreConexion Then
'
'   'sSQL = "select e.*,e.nEtapaCod,t.cEtapa " & _
'   '" From LogProcesoSeleccion p inner join LogProSelEtapa e on p.nProSelNro = e.nProSelNro " & _
'   '"     inner join (select nConsValor as nEtapaCod, cConsDescripcion as cEtapa from Constante where nConsCod= " & gcEtapasProcesoSel & " and nConsCod<>nConsValor) t on t.nEtapaCod = e.nEtapaCod " & _
'   ' Where P.nPlanAnualNro = " & nPlanNro & " And P.nPlanAnualItem = " & nPlanItem & " "
'
'   sSQL = "select e.*, t.cEtapa " & _
'   " From LogProSelEtapa e  " & _
'   "     inner join (select nConsValor as nEtapaCod, cConsDescripcion as cEtapa from Constante where nConsCod= " & gcEtapasProcesoSel & " and nConsCod<>nConsValor) t on e.nEtapaCod = t.nEtapaCod " & _
'   " Where e.nProSelNro = " & nProSelNro & "  "
'
'   Set rs = oConn.CargaRecordSet(sSQL)
'   If Not rs.EOF Then
'      Do While Not rs.EOF
'         i = i + 1
'         InsRow MSEta, i
'         'MSEta.RowHeight(i) = 290
'         MSEta.TextMatrix(i, 0) = rs!nEtapaCod
'         MSEta.TextMatrix(i, 1) = rs!nOrden
'         MSEta.TextMatrix(i, 2) = rs!cEtapa
'         'MSEta.TextMatrix(i, 3) = rs!cResponsable
'         MSEta.TextMatrix(i, 3) = IIf(IsNull(rs!dFechaInicio), "", rs!dFechaInicio)
'         MSEta.TextMatrix(i, 4) = IIf(IsNull(rs!dFechaTermino), "", rs!dFechaTermino)
'         MSEta.TextMatrix(i, 5) = rs!cObservacion
'         rs.MoveNext
'      Loop
'   End If
'End If
'End Sub

Sub FormaFlexEta()
'MSEta.Clear
'MSEta.Rows = 2
'MSEta.RowHeight(0) = 320
'MSEta.RowHeight(1) = 10
'MSEta.ColWidth(0) = 0
'MSEta.ColWidth(1) = 300:   MSEta.TextMatrix(0, 1) = "Orden":       MSEta.ColAlignment(1) = 4
'MSEta.ColWidth(2) = 5000:  MSEta.TextMatrix(0, 2) = "Etapa"
''MSEta.ColWidth(3) = 2400
'MSEta.ColWidth(3) = 900:   MSEta.TextMatrix(0, 3) = "Inicio":      MSEta.ColAlignment(3) = 4
'MSEta.ColWidth(4) = 900:   MSEta.TextMatrix(0, 4) = "Término":     MSEta.ColAlignment(4) = 4
'MSEta.ColWidth(5) = 4000:  MSEta.TextMatrix(0, 5) = "Observación"
'MSEta.ColWidth(6) = 0
End Sub

Private Sub CargarApelaciones(ByVal pnPSN As Integer, ByVal pnPSI As Integer)
    On Error GoTo CargarApelacionesErr
    Dim oConn As New DConecta, Rs As ADODB.Recordset, sSQL As String, i As Integer
    sSQL = "select nItemApelacion,p.cPersCod,p.cPersNombre,cApelacion,cRespuesta,bAdmision,bResuelto " & _
            "from LogProSelApelacion a inner join Persona p on a.cPersCod = p.cPersCod " & _
            "where nProSelNro=" & pnPSN & " and nProSelItem=" & pnPSI
    oConn.AbreConexion
    FormaMSApe
    Set Rs = oConn.CargaRecordSet(sSQL)
    i = 1
    cmdRespuesta.Enabled = False
    Do While Not Rs.EOF
        InsRow MSApe, i
        MSApe.RowHeight(i) = 1000
        MSApe.TextMatrix(i, 0) = Rs!nItemApelacion
        MSApe.TextMatrix(i, 1) = Rs!cPersCod
        MSApe.TextMatrix(i, 2) = Rs!cPersNombre
        MSApe.TextMatrix(i, 3) = Rs!cApelacion
        MSApe.TextMatrix(i, 4) = Rs!cRespuesta
        MSApe.TextMatrix(i, 5) = CBool(Rs!bAdmision)
        MSApe.TextMatrix(i, 6) = CBool(Rs!bResuelto)
        i = i + 1
        Rs.MoveNext
        cmdRespuesta.Enabled = True
        MSApe.SetFocus
    Loop
    oConn.CierraConexion
    Exit Sub
CargarApelacionesErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Private Sub CargarPostores(nPSN As Integer)
    On Error GoTo CargarPostoresErr
    Dim oConn As New DConecta, Rs As ADODB.Recordset, sSQL As String, i As Integer
    sSQL = "select x.nPresentoProp, x.nMovNroVentaBase, x.dFecha, x.cNroRecibo, p.cPersCod, p.cPersNombre  from LogProSelPostor x " & _
           "    inner join Persona p on x.cPersCod = p.cPersCod " & _
           "    where nProSelNro=" & nPSN & _
           "    order by x.nMovNroVentaBase"
    oConn.AbreConexion
    FormaFlexItemPostor
    Set Rs = oConn.CargaRecordSet(sSQL)
    i = 1
    Do While Not Rs.EOF
        If Not VerificaEtapaCerrada(gnProSelNro, 9) Then
            InsRow MSItemPostores, i
            If nTipo = 8 Then
                MSItemPostores.Col = 0
                MSItemPostores.row = i
                If Rs!nPresentoProp Then
                    Set MSItemPostores.CellPicture = imgOK
                Else
                    Set MSItemPostores.CellPicture = imgNN
                End If
                    CmdRegistrar.Enabled = True
                    MSItemPostores.Enabled = True
            End If
            Else
                CmdRegistrar.Enabled = False
                MSItemPostores.Enabled = False
                cmdimprimirpostores.Enabled = False
                MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                Exit Sub
            End If
        
        MSItemPostores.TextMatrix(i, 0) = Rs!nMovNroVentaBase
        MSItemPostores.TextMatrix(i, 1) = Format(Rs!dFecha, "dd/mm/yyyy")
        MSItemPostores.TextMatrix(i, 2) = Rs!cPersCod
        MSItemPostores.TextMatrix(i, 3) = Rs!cPersNombre
        MSItemPostores.TextMatrix(i, 4) = Rs!cNroRecibo
        MSItemPostores.Enabled = True
        cmdimprimirpostores.Enabled = True
        i = i + 1
        Rs.MoveNext
    Loop
    MSItemPostores.ColSel = MSItemPostores.Cols - 1
    oConn.CierraConexion
    Exit Sub
CargarPostoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Sub

Sub FormaMSApe()
With MSApe
    .Clear
    .Rows = 2
    .RowHeight(0) = 320
    .RowHeight(1) = 8
'    .ColWidth(0) = 300:  .ColAlignment(1) = 4
    .ColWidth(0) = 450:   .TextMatrix(0, 0) = "Item": .ColAlignment(0) = 4
    .ColWidth(1) = 800:  .TextMatrix(0, 1) = "Codigo":        .ColAlignment(1) = 4
    .ColWidth(2) = 2000:  .TextMatrix(0, 2) = "Persona":      '.ColAlignment(2) = 4
    .ColWidth(3) = 3300:  .TextMatrix(0, 3) = "Apelacion":      '  .ColAlignment(3) = 4
    .ColWidth(4) = 3300:  .TextMatrix(0, 4) = "Respuesta":      ' .ColAlignment(4) = 4
    .ColWidth(5) = 0:     .TextMatrix(0, 5) = "Admision":       .ColAlignment(5) = 4
    .ColWidth(6) = 0:     .TextMatrix(0, 6) = "Resuelta":       .ColAlignment(6) = 4
End With
End Sub

Private Sub CargarDatosRecibo(pcNroDoc As String)
    On Error GoTo CargarDatosReciboErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nDocTpo, cDocNro, d.nMovNro, dDocFecha, cMovDesc, nMovImporte, nMoneda = substring(c.cCtaContCod,3,1) " & _
                "  from movdoc d " & _
                "  inner join mov m on d.nMovNro=m.nMovNro " & _
                "  inner join movcta c on d.nMovNro=c.nMovNro " & _
                "  where d.cDocNro='" & pcNroDoc & "' and nMovItem=1"
        Set Rs = oCon.CargaRecordSet(sSQL)
        txtDescRecibo = ""
        txtimpoteRecibo = ""
        txtFechaRecibo.Text = ""
        If Not Rs.EOF Then
            txtDescRecibo = Rs!cMovDesc
            txtimpoteRecibo = FNumero(Rs!nMovImporte)
            txtFechaRecibo.Text = Rs!dDocFecha
            cboMonedaRecibo.ListIndex = Rs!nMoneda - 1
        End If
        oCon.CierraConexion
    End If
    Exit Sub
CargarDatosReciboErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Private Sub MSComite_Click()
    Dim nCol As Integer, nRow As Integer
    With MSComite
        nRow = .row
        nCol = .Col
        .Col = 0
'        If .Rows = 7 Then
'            NroM = 3
'        ElseIf .Rows = 7 Then
'            NroM = 5
'        ElseIf .Rows = 3 Then
'            NroM = 1
'        Else
'            MsgBox "El Comite esta inconmpleto...", vbInformation, "Aviso"
'        End If
        If .CellPicture = imgNN Then
            .row = nRow
            Set .CellPicture = imgOK
        Else
            Set .CellPicture = imgNN
        End If
        '.Col = nCol
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub MSFObsCon_GotFocus()
    Select Case nTipo
        Case 1, 2
            If MSFObsCon.TextMatrix(MSFObsCon.row, 5) = "" Then
                cmdModificarConOns.Enabled = True
            Else
                cmdModificarConOns.Enabled = False
            End If
        Case 6, 7
            cmdResponderConOns.Enabled = True
    End Select
End Sub

Private Sub MSFObsCon_SelChange()
    Select Case nTipo
        Case 1, 2
            If MSFObsCon.TextMatrix(MSFObsCon.row, 5) = "" Then
                cmdModificarConOns.Enabled = True
            Else
                cmdModificarConOns.Enabled = False
            End If
        Case 6, 7
            cmdResponderConOns.Enabled = True
    End Select
End Sub

Private Sub MSItem_GotFocus()
    CargarApelaciones gnProSelNro, Val(MSItem.TextMatrix(MSItem.row, 6))
End Sub

Private Sub MSItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then MSItem_DblClick
End Sub

Private Sub MSItem_SelChange()
    MSItem_GotFocus
End Sub

Private Sub MSItemPostores_DblClick()
    Dim nCol As Integer
    If nTipo <> 8 Then Exit Sub
    With MSItemPostores
        nCol = .Col
        .Col = 0
        If .CellPicture = imgOK Then
            Set .CellPicture = imgNN
        Else
            Set .CellPicture = imgOK
        End If
        '.Col = nCol
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(TxtProSelNro.Text) >= 0 Then
            ConsultarProcesoNro Val(TxtProSelNro.Text), Val(txtanio.Text)
            Exit Sub
        Else
            TxtProSelNro.SetFocus
        End If
    End If
    KeyAscii = DigNumEnt(KeyAscii)
End Sub

Private Sub txtApelacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtConsulta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txtDescRecibo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub


Private Sub txtFechaRecibo_GotFocus()
SelTexto txtFechaRecibo
End Sub

Private Sub txtFechaRecibo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdPersona.SetFocus
End If
End Sub

Private Sub txtimpoteRecibo_KeyPress(KeyAscii As Integer)
    KeyAscii = DigNumDec(txtimpoteRecibo, KeyAscii)
End Sub

Private Sub txtnrorecibo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtnrorecibo = Format(txtnrorecibo, "00000000")
'        CargarDatosRecibo TxtSerie.Text & "-" & txtnrorecibo.Text
    End If
End Sub

Private Sub TxtProSelNro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtanio.Text) > 0 Then
            ConsultarProcesoNro Val(TxtProSelNro.Text), Val(txtanio.Text)
            Exit Sub
        Else
            txtanio.SetFocus
        End If
    End If
    KeyAscii = DigNumEnt(KeyAscii)
End Sub

'Private Sub TxtResp_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
'End Sub

Private Sub txtrespuesta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtSerie = Format(TxtSerie, "000")
        txtnrorecibo.SetFocus
        Exit Sub
    End If
    KeyAscii = FNumero(KeyAscii)
End Sub

Private Function VerificaVentaBases(ByVal pnProSelNro As Integer) As Boolean
    On Error GoTo VerificaVetaBasesErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select nCostoBases from LogProcesoSeleccion where nProSelNro =" & pnProSelNro
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            If Rs!nCostoBases > 0 Then
                VerificaVentaBases = True
            Else
                VerificaVentaBases = False
            End If
        End If
        oCon.CierraConexion
    End If
    Exit Function
VerificaVetaBasesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Private Sub CargarEtapas(ByVal pnProSelNro As Integer)
On Error GoTo CargarEtapasErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSQL = "select e.nEtapaCod, e.dFechaInicio, e.dFechaTermino, c.cDescripcion " & _
               " from LogProSelEtapa e " & _
               " inner join LogEtapa c on e.nEtapaCod = c.nEtapaCod and c.nEstado = 1 " & _
               " where e.nEstado=1 and nProSelNro = " & pnProSelNro & " order by nOrden"
        Set Rs = oCon.CargaRecordSet(sSQL)
        cboEtapas.Clear
        Do While Not Rs.EOF
            cboEtapas.AddItem Rs!cDescripcion, cboEtapas.ListCount
            cboEtapas.ItemData(cboEtapas.ListCount - 1) = Rs!nEtapaCod
            Rs.MoveNext
        Loop
        oCon.CierraConexion
        If cboEtapas.ListCount > 0 Then cboEtapas.ListIndex = 0
    End If
    Exit Sub
CargarEtapasErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Function NroMiembros() As Integer
On Error GoTo nroMiembrosErr
    Dim i As Integer, j As Integer
    With MSComite
        .Col = 0
        Do While i < .Rows
            .row = i
            If .CellPicture = imgOK Then j = j + 1
            i = i + 1
        Loop
        NroMiembros = j
        .ColSel = .Cols - 1
    End With
    Exit Function
nroMiembrosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Private Sub LimpiaMiembros()
On Error GoTo nroMiembrosErr
    Dim i As Integer, j As Integer
    With MSComite
        .Col = 0
        i = 1
        Do While i < .Rows
            .row = i
            Set .CellPicture = imgNN
            i = i + 1
        Loop
        .ColSel = .Cols - 1
    End With
    Exit Sub
nroMiembrosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Sub CargarComiteEtapa(ByVal pnProSelNro As Integer, ByVal pnEtapaCod As Integer)
On Error GoTo CargarComiteEtapaErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset, i As Integer
    Set oCon = New DConecta
    sSQL = "select cPersCod from LogProSelEtapaComite where nProSelNro = " & pnProSelNro & " and nEtapaCod=" & pnEtapaCod
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        Do While Not Rs.EOF
            i = 1
            With MSComite
                .Col = 0
                Do While i < .Rows
                    .row = i
                    If .TextMatrix(i, 2) = Rs!cPersCod Then
                        Set .CellPicture = imgOK
                        Exit Do
                    End If
                    i = i + 1
                Loop
                .ColSel = .Cols - 1
            End With
            Rs.MoveNext
        Loop
        oCon.CierraConexion
    End If
    Exit Sub
CargarComiteEtapaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Sub ConsultarProcesoNro(ByVal pnNro As Integer, ByVal pnAnio As Integer)
    On Error GoTo ConsultarProcesoNroErr
    Dim oCon As DConecta, sSQL As String, Rs As ADODB.Recordset
    Set oCon = New DConecta
    sSQL = "select t.cProSelTpoDescripcion, s.nProSelNro, s.nPlanAnualNro, s.nPlanAnualAnio, " & _
            "s.nPlanAnualMes, s.nProSelTpoCod, s.nProSelSubTpo, nNroProceso, c.cConsDescripcion, " & _
            "s.nObjetoCod , s.nMoneda, s.nProSelMonto, s.nProSelEstado, cSintesis, nModalidadCompra " & _
            "from LogProcesoSeleccion s " & _
            "inner join LogProSelTpo t on s.nProSelTpoCod = t.nProSelTpoCod " & _
            "left outer join constante c on s.nObjetoCod=c.nConsValor and c.nConsCod = 9048 " & _
            "where s.nProSelEstado > -1 and s.nNroProceso=" & pnNro & " and nPlanAnualAnio = " & pnAnio
    If oCon.AbreConexion Then
        Set Rs = oCon.CargaRecordSet(sSQL)
        If Not Rs.EOF Then
            gnProSelNro = Rs!nProselNro
'            gnBSGrupoCod = rs!cBSGrupoCod
            TxtTipo.Text = Rs!cProSelTpoDescripcion
            TxtMonto.Text = FNumero(Rs!nProSelMonto)
            LblMoneda.Caption = IIf(Rs!nMoneda = 1, "S/.", "$")
            TxtDescripcion.Text = Rs!cSintesis
        Else
            gnProSelNro = 0
            TxtTipo.Text = ""
            TxtMonto.Text = ""
            LblMoneda.Caption = ""
            TxtDescripcion.Text = ""
            MsgBox "Proceso no Existe...", vbInformation, "Aviso"
            Exit Sub
        End If
        oCon.CierraConexion
    End If
    Select Case nTipo
        Case 1, 6
            lblDesc.Caption = "Consulta"
'            CargaObservaciones
            If VerificaEtapa(gnProSelNro, cnAbsolucionConsultas) Then
                If Not VerificaEtapaCerrada(gnProSelNro, cnObservaciones) Then
                    CargaConsultas
                    If nTipo = 1 Then
                        cmdModificarConOns.Visible = True
                        cmdQuitar.Visible = True
                        cmdAgregar.Visible = True
                    ElseIf nTipo = 6 Then
                        cmdImprimir.Visible = True
                        cmdResponderConOns.Visible = True
                        Exit Sub
                    End If
                Else
                    FormaMSFObsCon
                    MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                    cmdModificarConOns.Visible = False
                    cmdQuitar.Visible = False
                    cmdAgregar.Visible = False
                    cmdImprimir.Visible = False
                    cmdResponderConOns.Visible = False
                    Exit Sub
                End If
            Else
               FormaMSFObsCon
               MsgBox "No se Especificado esta Etapa para este proceso", vbInformation, "Aviso"
               cmdModificarConOns.Visible = False
               cmdQuitar.Visible = False
               cmdAgregar.Visible = False
               cmdImprimir.Visible = False
               cmdResponderConOns.Visible = False
               Exit Sub
            End If
        Case 2, 7
            lblDesc.Caption = "Observación"
            'CargaObservaciones
            If VerificaEtapa(gnProSelNro, cnObservaciones) Then
                If Not VerificaEtapaCerrada(gnProSelNro, cnObservaciones) Then
                    CargaObservaciones
                    If nTipo = 2 Then
                        cmdModificarConOns.Visible = True
                        cmdQuitar.Visible = True
                        cmdAgregar.Visible = True
                    ElseIf nTipo = 7 Then
                        cmdImprimir.Visible = True
                        cmdResponderConOns.Visible = True
                        Exit Sub
                    End If
                Else
                    FormaMSFObsCon
                    MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                    cmdModificarConOns.Visible = False
                    cmdQuitar.Visible = False
                    cmdAgregar.Visible = False
                    cmdImprimir.Visible = False
                    cmdResponderConOns.Visible = False
                    Exit Sub
                End If
            Else
               FormaMSFObsCon
               MsgBox "No se Especificado esta Etapa para este proceso", vbInformation, "Aviso"
               cmdModificarConOns.Visible = False
               cmdQuitar.Visible = False
               cmdAgregar.Visible = False
               cmdImprimir.Visible = False
               cmdResponderConOns.Visible = False
               Exit Sub
            End If
        Case 3
            If VerificaEtapa(gnProSelNro, cnRegistroParticipantes) Then
'                CierraEtapa gnProSelNro, 1
'                CierraEtapa gnProSelNro, 2
                If Not VerificaEtapaCerrada(gnProSelNro, cnRegistroParticipantes) Then
                    CargarPostores gnProSelNro
                    cmdAgragarPostor.Enabled = True
                    cmdQuitarPostor.Enabled = True
                    cmdimprimirpostores.Enabled = True
                Else
                    MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                    cmdAgragarPostor.Enabled = False
                    cmdQuitarPostor.Enabled = False
                    cmdimprimirpostores.Enabled = False
                    Exit Sub
                End If
            Else
                FormaFlexItemPostor
                MsgBox "Etapa no Esta Configurada para este Proceso", vbInformation, "Aviso"
                cmdAgragarPostor.Enabled = False
                cmdQuitarPostor.Enabled = False
                cmdimprimirpostores.Enabled = False
                Exit Sub
            End If
        Case 8
            If VerificaEtapa(gnProSelNro, cnPresentacionPropuestas) Then
                CargarPostores gnProSelNro
            Else
                FormaFlexItemPostor
                MsgBox "Etapa no Esta Configurada para este Proceso", vbInformation, "Aviso"
                Exit Sub
            End If
        Case 4
            If VerificaEtapa(gnProSelNro, cnApelaciones) Then
                If Not VerificaEtapaCerrada(gnProSelNro, cnApelaciones) Then
                    GeneraDetalleItem gnProSelNro
                    cmdregistrarAp.Visible = True
                    cmdquitarAp.Visible = True
                    cmdRespuesta.Visible = True
                Else
                    FormaFlexItem
                    cmdregistrarAp.Visible = False
                    cmdquitarAp.Visible = False
                    cmdRespuesta.Visible = False
                    MsgBox "Etapa Cerrada", vbInformation, "Aviso"
                    Exit Sub
                End If
            Else
                FormaFlexItem
                cmdregistrarAp.Visible = False
                cmdquitarAp.Visible = False
                cmdRespuesta.Visible = False
                MsgBox "Etapa no Esta Configurada para este Proceso", vbInformation, "Aviso"
                Exit Sub
            End If
'            MSItem.SetFocus
        Case 5
            CargarEtapas gnProSelNro
            If cboEtapas.ListCount > 0 Then
                cargarComiteItemProceso gnProSelNro
                CargarComiteEtapa gnProSelNro, cboEtapas.ItemData(cboEtapas.ListIndex)
                cmdComite.Visible = True
            Else
                cmdComite.Visible = False
                Exit Sub
            End If
'            CargarEtapas gnProSelNro
'            cargarComiteItemProceso gnProSelNro
'            If cboEtapas.ListIndex > -1 Then CargarComiteEtapa gnProSelNro, cboEtapas.ItemData(cboEtapas.ListIndex)
        Case 9
            CagarDatosActoPublico gnProSelNro
    End Select
    Exit Sub
ConsultarProcesoNroErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

'Private Function VerificarEtapa(ByVal pnProSelNro As Integer, ByVal nEtapa As Integer) As Boolean
'On Error GoTo VerificarEtapaErr
'    Dim oCon As DConecta, sSQL As String, rs As New ADODB.Recordset
'    Set oCon = New DConecta
'    sSQL = "select nNro=count(*)  From LogProSelEtapa e " & _
'           "      inner join (select nConsValor as nEtapaCod, " & _
'           "      cConsDescripcion as cEtapa from Constante where " & _
'           "      nConsCod= 9041 and nConsCod<>nConsValor) t on " & _
'           "    e.nEtapaCod = t.nEtapaCod  Where e.nProSelNro = " & pnProSelNro & " and e.nEtapaCod = " & nEtapa
'    If oCon.AbreConexion Then
'       Set rs = oCon.CargaRecordSet(sSQL)
'       If Not rs.EOF Then
'          If rs!nNro > 0 Then
'            VerificarEtapa = True
'          Else
'            VerificarEtapa = False
'          End If
'       End If
'       oCon.CierraConexion
'    End If
'    Exit Function
'VerificarEtapaErr:
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
'End Function
