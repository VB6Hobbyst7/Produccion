VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCredRegistrarParametroBPPR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de BPPR"
   ClientHeight    =   5955
   ClientLeft      =   3120
   ClientTop       =   3150
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredRegistrarParametroBPPR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9855
   Begin VB.Frame Frame2 
      Caption         =   "Metas Minimas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   69
      Top             =   1320
      Width           =   9615
      Begin VB.TextBox txtSaldo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6240
         TabIndex        =   74
         Top             =   345
         Width           =   1485
      End
      Begin VB.TextBox txtOpe 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3960
         TabIndex        =   72
         Top             =   345
         Width           =   650
      End
      Begin VB.TextBox txtNroCliente 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         TabIndex        =   70
         Top             =   345
         Width           =   650
      End
      Begin VB.Label Label8 
         Caption         =   "Saldo Cartera:"
         Height          =   255
         Left            =   4920
         TabIndex        =   75
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Nro Operaciones:"
         Height          =   255
         Left            =   2400
         TabIndex        =   73
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Nro Clientes:"
         Height          =   255
         Left            =   360
         TabIndex        =   71
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   65
      Top             =   4920
      Width           =   9615
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   8040
         TabIndex        =   68
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdRegistrar 
         Caption         =   "Guardar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   46
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   67
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2880
         TabIndex        =   66
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCN 
      Height          =   285
      Left            =   5520
      TabIndex        =   8
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtMC 
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Top             =   8040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Cartera"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9615
      Begin MSDataListLib.DataCombo dcTipoCartera 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcTipoClasificacion 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label4 
         Caption         =   "Clasificación:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cartera:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Parametros de cumplimiento de Mora"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   9615
      Begin VB.TextBox txtCMDeAceptable 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   50
         Top             =   240
         Width           =   650
      End
      Begin VB.TextBox txtCMDeSinDescuento 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   49
         Top             =   600
         Width           =   650
      End
      Begin VB.TextBox txtCMHastaAceptable 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   48
         Top             =   240
         Width           =   650
      End
      Begin VB.TextBox txtCMHastaSinDescuento 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3840
         TabIndex        =   47
         Top             =   600
         Width           =   650
      End
      Begin VB.TextBox txtDescuento11 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7920
         TabIndex        =   42
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta11 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7920
         TabIndex        =   41
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe11 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7920
         TabIndex        =   40
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento10 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   39
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta10 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   38
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe10 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         TabIndex        =   37
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento9 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   36
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta9 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   35
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe9 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6480
         TabIndex        =   34
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento8 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   33
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta8 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   32
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe8 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   31
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   30
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   29
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         TabIndex        =   28
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   27
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   26
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4320
         TabIndex        =   25
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   24
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   23
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         TabIndex        =   22
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   21
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   20
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   19
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   17
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe3 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   16
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe2 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtDescuento1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   2040
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   10
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoDe12 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         TabIndex        =   43
         Top             =   1320
         Width           =   650
      End
      Begin VB.TextBox txtConDescuentoHasta12 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         TabIndex        =   44
         Top             =   1680
         Width           =   650
      End
      Begin VB.TextBox txtDescuento12 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         TabIndex        =   45
         Top             =   2040
         Width           =   650
      End
      Begin VB.Label Label28 
         Caption         =   "%"
         Height          =   255
         Left            =   4560
         TabIndex        =   64
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label27 
         Caption         =   "%"
         Height          =   255
         Left            =   4560
         TabIndex        =   63
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label25 
         Caption         =   "%"
         Height          =   255
         Left            =   2880
         TabIndex        =   62
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label24 
         Caption         =   "%"
         Height          =   255
         Left            =   2880
         TabIndex        =   61
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label22 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   3240
         TabIndex        =   60
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   3240
         TabIndex        =   59
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "De:"
         Height          =   255
         Left            =   1800
         TabIndex        =   58
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label18 
         Caption         =   "De:"
         Height          =   255
         Left            =   1800
         TabIndex        =   57
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label16 
         Caption         =   "Sin Descuento:"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label15 
         Caption         =   "Aceptable:"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   1095
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   9480
         Y1              =   1980
         Y2              =   1980
      End
      Begin VB.Label Label5 
         Caption         =   "Desc:"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label17 
         Caption         =   "Con Descuento:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label20 
         Caption         =   "De:"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label Label23 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1680
         Width           =   615
      End
   End
   Begin VB.Label Label2 
      Caption         =   "CLIENTES NUEVOS:"
      Height          =   255
      Left            =   3480
      TabIndex        =   9
      Top             =   8040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "META POR CRECIMIENTO:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   8040
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmCredRegistrarParametroBPPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCOMDCredito As COMDCredito.DCOMBPPR
Dim objCOMNCredito As COMNCredito.NCOMBPPR
Dim lSaldo As String

Private Sub cmdCancelar_Click()
    DeshabilitarText
    cmdCancelar.Enabled = False
    cmdRegistrar.Enabled = False
    'cmdEditar.Enabled = True
    Me.dcTipoClasificacion.BoundText = 0
    LimpiarText
End Sub

Private Sub cmdEditar_Click()
    Me.cmdRegistrar.Enabled = True
    Me.cmdCancelar.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.Frame2.Enabled = True
    Me.Frame3.Enabled = True
    'Me.txtNroCliente.Enabled = True
    HabilitarText
    
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub dcTipoCartera_Change()
    'CargarTipoClasificacion
    If dcTipoCartera.BoundText <> 0 Then
        CargarTipoClasificacion
    Else
        LimpiarText
        Set dcTipoClasificacion.RowSource = Nothing
    End If
End Sub

Private Sub dcTipoClasificacion_Change()
    If dcTipoClasificacion.BoundText <> 0 Then
        LimpiarText
        CargarParametrosTipoClasificacion
    Else
        LimpiarText
    End If
End Sub



Private Sub Form_Load()
    CargarTipoCartera
    Me.Frame2.Enabled = False
    Me.Frame3.Enabled = False
End Sub

Private Sub cmdRegistrar_Click()
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    
    If MsgBox("Se van a Regsitrar los Parametros", vbInformation + vbYesNo) = vbYes Then
            sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            Set objCOMNCredito = New COMNCredito.NCOMBPPR
            
            objCOMNCredito.actualizarParametrosBPP dcTipoCartera.BoundText, dcTipoClasificacion.BoundText
            
            objCOMNCredito.InsertarParametrosBPP dcTipoCartera.BoundText, dcTipoClasificacion.BoundText, txtNroCliente.Text, txtOpe.Text, txtSaldo.Text, IIf(txtCMDeAceptable.Text <> "", txtCMDeAceptable.Text, 0), IIf(txtCMHastaAceptable.Text <> "", txtCMHastaAceptable.Text, 0), IIf(txtCMDeSinDescuento.Text <> "", txtCMDeSinDescuento.Text, 0), IIf(txtCMHastaSinDescuento.Text <> "", txtCMHastaSinDescuento.Text, 0), _
            IIf(txtConDescuentoDe1.Text <> "", txtConDescuentoDe1.Text, 0), IIf(txtConDescuentoHasta1.Text <> "", txtConDescuentoHasta1.Text, 0), IIf(txtDescuento1.Text <> "", txtDescuento1.Text, 0), _
            IIf(txtConDescuentoDe2.Text <> "", txtConDescuentoDe2.Text, 0), IIf(txtConDescuentoHasta2.Text <> "", txtConDescuentoHasta2.Text, 0), IIf(txtDescuento2.Text <> "", txtDescuento2.Text, 0), _
            IIf(txtConDescuentoDe3.Text <> "", txtConDescuentoDe3.Text, 0), IIf(txtConDescuentoHasta3.Text <> "", txtConDescuentoHasta3.Text, 0), IIf(txtDescuento3.Text <> "", txtDescuento3.Text, 0), _
            IIf(txtConDescuentoDe4.Text <> "", txtConDescuentoDe4.Text, 0), IIf(txtConDescuentoHasta4.Text <> "", txtConDescuentoHasta4.Text, 0), IIf(txtDescuento4.Text <> "", txtDescuento4.Text, 0), _
            IIf(txtConDescuentoDe5.Text <> "", txtConDescuentoDe5.Text, 0), IIf(txtConDescuentoHasta5.Text <> "", txtConDescuentoHasta5.Text, 0), IIf(txtDescuento5.Text <> "", txtDescuento5.Text, 0), _
            IIf(txtConDescuentoDe6.Text <> "", txtConDescuentoDe6.Text, 0), IIf(txtConDescuentoHasta6.Text <> "", txtConDescuentoHasta6.Text, 0), IIf(txtDescuento6.Text <> "", txtDescuento6.Text, 0), _
            IIf(txtConDescuentoDe7.Text <> "", txtConDescuentoDe7.Text, 0), IIf(txtConDescuentoHasta7.Text <> "", txtConDescuentoHasta7.Text, 0), IIf(txtDescuento7.Text <> "", txtDescuento7.Text, 0), _
            IIf(txtConDescuentoDe8.Text <> "", txtConDescuentoDe8.Text, 0), IIf(txtConDescuentoHasta8.Text <> "", txtConDescuentoHasta8.Text, 0), IIf(txtDescuento8.Text <> "", txtDescuento8.Text, 0), _
            IIf(txtConDescuentoDe9.Text <> "", txtConDescuentoDe9.Text, 0), IIf(txtConDescuentoHasta9.Text <> "", txtConDescuentoHasta9.Text, 0), IIf(txtDescuento9.Text <> "", txtDescuento9.Text, 0), _
            IIf(txtConDescuentoDe10.Text <> "", txtConDescuentoDe10.Text, 0), IIf(txtConDescuentoHasta10.Text <> "", txtConDescuentoHasta10.Text, 0), IIf(txtDescuento10.Text <> "", txtDescuento10.Text, 0), _
            IIf(txtConDescuentoDe11.Text <> "", txtConDescuentoDe11.Text, 0), IIf(txtConDescuentoHasta11.Text <> "", txtConDescuentoHasta11.Text, 0), IIf(txtDescuento11.Text <> "", txtDescuento11.Text, 0), _
            IIf(txtConDescuentoDe12.Text <> "", txtConDescuentoDe12.Text, 0), IIf(txtConDescuentoHasta12.Text <> "", txtConDescuentoHasta12.Text, 0), IIf(txtDescuento12.Text <> "", txtDescuento12.Text, 0), _
            sMovNro
            
            MsgBox "Se han Registrados los Parametros", vbInformation
            LimpiarText
            Me.dcTipoClasificacion.BoundText = 0
    End If
    
End Sub

Private Sub LimpiarText()
    Dim ctlControl As Object
    For Each ctlControl In Me.Controls
        If TypeOf ctlControl Is TextBox Then
                ctlControl.Text = ""
        End If
    Next
End Sub
Private Sub DeshabilitarText()
    Dim ctlControl As Object
    For Each ctlControl In Me.Controls
        If TypeOf ctlControl Is TextBox Then
                ctlControl.Enabled = False
        End If
    Next
End Sub
Private Sub HabilitarText()
    Dim ctlControl As Object
    For Each ctlControl In Me.Controls
        If TypeOf ctlControl Is TextBox Then
                ctlControl.Enabled = True
        End If
    Next
End Sub

Private Sub CargarTipoCartera()
    Dim rsTipoCartera As New ADODB.Recordset
    Set objCOMNCredito = New COMNCredito.NCOMBPPR
    Set rsTipoCartera.DataSource = objCOMNCredito.getCargarTipoCartera
    dcTipoCartera.BoundColumn = "iTipoCarteraId"
    dcTipoCartera.DataField = "iTipoCarteraId"
    Set dcTipoCartera.RowSource = rsTipoCartera
    dcTipoCartera.ListField = "vTipoCartera"
    dcTipoCartera.BoundText = 0
End Sub

Private Sub CargarTipoClasificacion()
    Dim rsTipoClasificacion As New ADODB.Recordset
   Set objCOMNCredito = New COMNCredito.NCOMBPPR
    Set rsTipoClasificacion.DataSource = objCOMNCredito.getCargarTipoClasificacion(dcTipoCartera.BoundText)
    dcTipoClasificacion.BoundColumn = "iTipoClasificacionId"
    dcTipoClasificacion.DataField = "iTipoClasificacionId"
    Set dcTipoClasificacion.RowSource = rsTipoClasificacion
    dcTipoClasificacion.ListField = "vTipoClasificacion"
    dcTipoClasificacion.BoundText = 0
End Sub
'JACA 20110308
Private Sub CargarParametrosTipoClasificacion()
    Dim rsTipoClasificacion As New ADODB.Recordset
      Set objCOMNCredito = New COMNCredito.NCOMBPPR
    
       
    Set rsTipoClasificacion.DataSource = objCOMNCredito.getCargarParametrosTipoClasificacion(dcTipoCartera.BoundText, dcTipoClasificacion.BoundText)
    
    
    If rsTipoClasificacion.RecordCount > 0 Then
        
         Me.cmdCancelar.Enabled = False
         Me.cmdRegistrar.Enabled = False
         Me.Frame2.Enabled = False
         Me.Frame3.Enabled = False
         
            Me.txtNroCliente.Text = rsTipoClasificacion!iMetaClieNue
            Me.txtOpe.Text = rsTipoClasificacion!iMetaNroOpe
            Me.txtSaldo.Text = Format(rsTipoClasificacion!iMetaSaldo, "##,##0.00")
            
            Me.txtCMDeAceptable.Text = Format(rsTipoClasificacion!nMAceptableDe, "#0.00")
            Me.txtCMHastaAceptable.Text = Format(rsTipoClasificacion!nMAceptableHasta, "#0.00")
            Me.txtCMDeSinDescuento.Text = Format(rsTipoClasificacion!nMSinDescuentoDe, "#0.00")
            Me.txtCMHastaSinDescuento.Text = Format(rsTipoClasificacion!nMSinDescuentoHasta, "#0.00")
            
            Me.txtConDescuentoDe1.Text = Format(rsTipoClasificacion!nMConDescuento1, "#0.00")
            Me.txtConDescuentoHasta1.Text = Format(rsTipoClasificacion!nMConDescuentoHasta1, "#0.00")
            Me.txtDescuento1.Text = Format(rsTipoClasificacion!nMConDescuentoDesc1, "#0.00")
            
            
            Me.txtConDescuentoDe2.Text = Format(rsTipoClasificacion!nMConDescuento2, "#0.00")
            Me.txtConDescuentoHasta2.Text = Format(rsTipoClasificacion!nMConDescuentoHasta2, "#0.00")
            Me.txtDescuento2.Text = Format(rsTipoClasificacion!nMConDescuentoDesc2, "#0.00")
            
            
            Me.txtConDescuentoDe3.Text = Format(rsTipoClasificacion!nMConDescuento3, "#0.00")
            Me.txtConDescuentoHasta3.Text = Format(rsTipoClasificacion!nMConDescuentoHasta3, "#0.00")
            Me.txtDescuento3.Text = Format(rsTipoClasificacion!nMConDescuentoDesc3, "#0.00")
            
            
            Me.txtConDescuentoDe4.Text = Format(rsTipoClasificacion!nMConDescuento4, "#0.00")
            Me.txtConDescuentoHasta4.Text = Format(rsTipoClasificacion!nMConDescuentoHasta4, "#0.00")
            Me.txtDescuento4.Text = Format(rsTipoClasificacion!nMConDescuentoDesc4, "#0.00")
            
            
            Me.txtConDescuentoDe5.Text = Format(rsTipoClasificacion!nMConDescuento5, "#0.00")
            Me.txtConDescuentoHasta5.Text = Format(rsTipoClasificacion!nMConDescuentoHasta5, "#0.00")
            Me.txtDescuento5.Text = Format(rsTipoClasificacion!nMConDescuentoDesc5, "#0.00")
            
            
            Me.txtConDescuentoDe6.Text = Format(rsTipoClasificacion!nMConDescuento6, "#0.00")
            Me.txtConDescuentoHasta6.Text = Format(rsTipoClasificacion!nMConDescuentoHasta6, "#0.00")
            Me.txtDescuento6.Text = Format(rsTipoClasificacion!nMConDescuentoDesc6, "#0.00")
            
            
            Me.txtConDescuentoDe7.Text = Format(rsTipoClasificacion!nMConDescuento7, "#0.00")
            Me.txtConDescuentoHasta7.Text = Format(rsTipoClasificacion!nMConDescuentoHasta7, "#0.00")
            Me.txtDescuento7.Text = Format(rsTipoClasificacion!nMConDescuentoDesc7, "#0.00")
            
            
            Me.txtConDescuentoDe8.Text = Format(rsTipoClasificacion!nMConDescuento8, "#0.00")
            Me.txtConDescuentoHasta8.Text = Format(rsTipoClasificacion!nMConDescuentoHasta8, "#0.00")
            Me.txtDescuento8.Text = Format(rsTipoClasificacion!nMConDescuentoDesc8, "#0.00")
            
            
            Me.txtConDescuentoDe9.Text = Format(rsTipoClasificacion!nMConDescuento9, "#0.00")
            Me.txtConDescuentoHasta9.Text = Format(rsTipoClasificacion!nMConDescuentoHasta9, "#0.00")
            Me.txtDescuento9.Text = Format(rsTipoClasificacion!nMConDescuentoDesc9, "#0.00")
            
            
            Me.txtConDescuentoDe10.Text = Format(rsTipoClasificacion!nMConDescuento10, "#0.00")
            Me.txtConDescuentoHasta10.Text = Format(rsTipoClasificacion!nMConDescuentoHasta10, "#0.00")
            Me.txtDescuento10.Text = Format(rsTipoClasificacion!nMConDescuentoDesc10, "#0.00")
            
            
            Me.txtConDescuentoDe11.Text = Format(rsTipoClasificacion!nMConDescuento11, "#0.00")
            Me.txtConDescuentoHasta11.Text = Format(rsTipoClasificacion!nMConDescuentoHasta11, "#0.00")
            Me.txtDescuento11.Text = Format(rsTipoClasificacion!nMConDescuentoDesc11, "#0.00")
            
            
            Me.txtConDescuentoDe12.Text = Format(rsTipoClasificacion!nMConDescuento12, "#0.00")
            Me.txtConDescuentoHasta12.Text = Format(rsTipoClasificacion!nMConDescuentoHasta12, "#0.00")
            Me.txtDescuento12.Text = Format(rsTipoClasificacion!nMConDescuentoDesc12, "#0.00")
        
            Me.cmdEditar.Enabled = True
    Else
            
            Me.cmdCancelar.Enabled = True
            Me.cmdEditar.Enabled = False
            Me.Frame2.Enabled = True
            Me.Frame3.Enabled = True
            Me.txtNroCliente.Enabled = True
            Me.txtNroCliente.SetFocus
            
    End If
End Sub
'END JACA
'Private Sub CargarComite()
'    Dim rsComite As New ADODB.Recordset
'    Set objCOMDCredito = New COMDCredito.DCOMBPPR
'    Set rsComite.DataSource = objCOMDCredito.CargarComite
'    dcComite.BoundColumn = "iComiteId"
'    dcComite.DataField = "iComiteId"
'    Set dcComite.RowSource = rsComite
'    dcComite.ListField = "vComite"
'    dcComite.BoundText = 0
'End Sub
'
'Private Sub chkComite_Click()
'    If chkComite.value = 1 Then
'        lblComite.Visible = True
'        dcComite.Visible = True
'        CargarComite
'    Else
'        lblComite.Visible = False
'        dcComite.Visible = False
'    End If
'End Sub

'Private Sub txtFRegistro_LostFocus()
'    Dim sCad As String
'    sCad = ValidaFecha(txtFRegistro.Text)
'        If Not Trim(sCad) = "" Then
'            MsgBox sCad, vbInformation, "Aviso"
'            Exit Sub
'        End If
'        If CDate(txtFRegistro.Text) > gdFecSis Then
'            MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
'            txtFRegistro.SetFocus
'            Exit Sub
'        End If
'End Sub

Private Sub txtCN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMC.SetFocus
End If
End Sub
Private Sub txtConDescuentoDe1_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoDe1.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoHasta1.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta1_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoHasta1.Text <> "" Then
        If KeyAscii = 13 Then
            txtDescuento1.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtMC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMC.Text = Format(txtMC.Text, "#,##0.00")
        txtCMDeAceptable.SetFocus
    End If
End Sub

Private Sub txtMC_LostFocus()
    txtMC.Text = Format(txtMC.Text, "#,##0.00")
End Sub

Private Sub txtCMDeAceptable_KeyPress(KeyAscii As Integer)
    If Me.txtCMDeAceptable.Text <> "" Then
        If KeyAscii = 13 Then
            txtCMDeAceptable.Text = Format(txtCMDeAceptable.Text, "#,##0.00")
            txtCMHastaAceptable.Enabled = True
            txtCMHastaAceptable.SetFocus
        End If
    End If
End Sub

Private Sub txtCMDeAceptable_LostFocus()
    txtCMDeAceptable.Text = Format(txtCMDeAceptable.Text, "#,##0.00")
    'txtConDescuentoDe1.Text = Format(txtCMDeAceptable.Text, "#,##0.00")
End Sub

Private Sub txtCMHastaAceptable_KeyPress(KeyAscii As Integer)
    If Me.txtCMHastaAceptable.Text <> "" Then
        If KeyAscii = 13 Then
            txtCMHastaAceptable.Text = Format(txtCMHastaAceptable.Text, "#,##0.00")
            txtCMDeSinDescuento.Enabled = True
            txtCMDeSinDescuento.SetFocus
        End If
    End If
End Sub

Private Sub txtCMHastaAceptable_LostFocus()
    txtCMHastaAceptable.Text = Format(txtCMHastaAceptable.Text, "#,##0.00")
End Sub

Private Sub txtCMDeSinDescuento_KeyPress(KeyAscii As Integer)
    If Me.txtCMDeSinDescuento.Text <> "" Then
        If KeyAscii = 13 Then
            txtCMDeSinDescuento.Text = Format(txtCMDeSinDescuento.Text, "#,##0.00")
            txtCMHastaSinDescuento.Enabled = True
             txtCMHastaSinDescuento.SetFocus
        End If
    End If
End Sub

Private Sub txtCMDeSinDescuento_LostFocus()
    txtCMDeSinDescuento.Text = Format(txtCMDeSinDescuento.Text, "#,##0.00")
End Sub

Private Sub txtCMHastaSinDescuento_KeyPress(KeyAscii As Integer)
    If Me.txtCMHastaSinDescuento.Text <> "" Then
        If KeyAscii = 13 Then
            txtCMHastaSinDescuento.Text = Format(txtCMHastaSinDescuento.Text, "#,##0.00")
            txtConDescuentoHasta1.Text = Format(txtCMHastaSinDescuento.Text, "#,##0.00")
            txtConDescuentoDe1.Enabled = True
            txtConDescuentoDe1.SetFocus
        End If
    End If
End Sub

Private Sub txtCMHastaSinDescuento_LostFocus()
    txtCMHastaSinDescuento.Text = Format(txtCMHastaSinDescuento.Text, "#,##0.00")
    txtConDescuentoHasta1.Text = Format(txtCMHastaSinDescuento.Text, "#,##0.00")
End Sub

Private Sub txtDescuento1_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento1.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoDe2.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento1_LostFocus()
    txtDescuento1.Text = Format(txtDescuento1.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoDe2_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoDe2.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoHasta2.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta2_KeyPress(KeyAscii As Integer)
   If Me.txtConDescuentoHasta2.Text <> "" Then
        If KeyAscii = 13 Then
            txtDescuento2.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento2_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento2.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoDe3.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe2_LostFocus()
    txtConDescuentoDe2.Text = Format(txtConDescuentoDe2.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta2_LostFocus()
    txtConDescuentoHasta2.Text = Format(txtConDescuentoHasta2.Text, "#,##0.00")
End Sub

Private Sub txtDescuento2_LostFocus()
    txtDescuento2.Text = Format(txtDescuento2.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoDe3_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoDe3.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoHasta3.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta3_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoHasta3.Text <> "" Then
        If KeyAscii = 13 Then
            txtDescuento3.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento3_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento3.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoDe4.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe3_LostFocus()
    txtConDescuentoDe3.Text = Format(txtConDescuentoDe3.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta3_LostFocus()
    txtConDescuentoHasta3.Text = Format(txtConDescuentoHasta3.Text, "#,##0.00")
End Sub

Private Sub txtDescuento3_LostFocus()
    txtDescuento3.Text = Format(txtDescuento3.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoDe4_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoDe4.Text <> "" Then
        If KeyAscii = 13 Then
           txtConDescuentoHasta4.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta4_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoHasta4.Text <> "" Then
        If KeyAscii = 13 Then
           txtDescuento4.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento4_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento4.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoDe5.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe4_LostFocus()
    txtConDescuentoDe4.Text = Format(txtConDescuentoDe4.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta4_LostFocus()
    txtConDescuentoHasta4.Text = Format(txtConDescuentoHasta4.Text, "#,##0.00")
End Sub

Private Sub txtDescuento4_LostFocus()
    txtDescuento4.Text = Format(txtDescuento4.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoDe5_KeyPress(KeyAscii As Integer)
   If Me.txtConDescuentoDe5.Text <> "" Then
        If KeyAscii = 13 Then
           txtConDescuentoHasta5.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta5_KeyPress(KeyAscii As Integer)
   If Me.txtConDescuentoHasta5.Text <> "" Then
        If KeyAscii = 13 Then
            txtDescuento5.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento5_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento5.Text <> "" Then
        If KeyAscii = 13 Then
           txtConDescuentoDe6.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe5_LostFocus()
    txtConDescuentoDe5.Text = Format(txtConDescuentoDe5.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta5_LostFocus()
    txtConDescuentoHasta5.Text = Format(txtConDescuentoHasta5.Text, "#,##0.00")
End Sub

Private Sub txtDescuento5_LostFocus()
    txtDescuento5.Text = Format(txtDescuento5.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoDe6_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoDe6.Text <> "" Then
        If KeyAscii = 13 Then
           txtConDescuentoHasta6.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta6_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoHasta6.Text <> "" Then
        If KeyAscii = 13 Then
            txtDescuento6.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento6_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento6.Text <> "" Then
        If KeyAscii = 13 Then
           txtConDescuentoDe7.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe6_LostFocus()
    txtConDescuentoDe6.Text = Format(txtConDescuentoDe6.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta6_LostFocus()
    txtConDescuentoHasta6.Text = Format(txtConDescuentoHasta6.Text, "#,##0.00")
End Sub

Private Sub txtDescuento6_LostFocus()
    txtDescuento6.Text = Format(txtDescuento6.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoDe7_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoDe7.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoHasta7.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta7_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoHasta7.Text <> "" Then
        If KeyAscii = 13 Then
           txtDescuento7.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento7_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento7.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoDe8.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe7_LostFocus()
    txtConDescuentoDe7.Text = Format(txtConDescuentoDe7.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta7_LostFocus()
    txtConDescuentoHasta7.Text = Format(txtConDescuentoHasta7.Text, "#,##0.00")
End Sub

Private Sub txtDescuento7_LostFocus()
    txtDescuento7.Text = Format(txtDescuento7.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoDe8_KeyPress(KeyAscii As Integer)
   If Me.txtConDescuentoDe8.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoHasta8.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta8_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoHasta8.Text <> "" Then
        If KeyAscii = 13 Then
           txtDescuento8.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento8_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento8.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoDe9.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe8_LostFocus()
    txtConDescuentoDe8.Text = Format(txtConDescuentoDe8.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta8_LostFocus()
    txtConDescuentoHasta8.Text = Format(txtConDescuentoHasta8.Text, "#,##0.00")
End Sub

Private Sub txtDescuento8_LostFocus()
    txtDescuento8.Text = Format(txtDescuento8.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoDe9_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoDe9.Text <> "" Then
        If KeyAscii = 13 Then
           txtConDescuentoHasta9.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta9_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoHasta9.Text <> "" Then
        If KeyAscii = 13 Then
           txtDescuento9.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento9_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento9.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoDe10.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe9_LostFocus()
    txtConDescuentoDe9.Text = Format(txtConDescuentoDe9.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta9_LostFocus()
    txtConDescuentoHasta9.Text = Format(txtConDescuentoHasta9.Text, "#,##0.00")
End Sub

Private Sub txtDescuento9_LostFocus()
    txtDescuento9.Text = Format(txtDescuento9.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoDe11_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoDe11.Text <> "" Then
        If KeyAscii = 13 Then
           txtConDescuentoHasta11.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta11_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoHasta11.Text <> "" Then
        If KeyAscii = 13 Then
          txtDescuento11.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento11_KeyPress(KeyAscii As Integer)
    If Me.txtDescuento11.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoDe12.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe11_LostFocus()
    txtConDescuentoDe11.Text = Format(txtConDescuentoDe11.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta11_LostFocus()
    txtConDescuentoHasta11.Text = Format(txtConDescuentoHasta11.Text, "#,##0.00")
End Sub

Private Sub txtDescuento11_LostFocus()
    txtDescuento11.Text = Format(txtDescuento11.Text, "#,##0.00")
End Sub


'Private Sub txtRentabilidad_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        txtRentabilidad.Text = Format(txtRentabilidad.Text, "#,##0.00")
'    End If
'End Sub

Private Sub txtConDescuentoDe10_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoDe10.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoHasta10.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta10_KeyPress(KeyAscii As Integer)
    If Me.txtConDescuentoHasta10.Text <> "" Then
        If KeyAscii = 13 Then
            txtDescuento10.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento10_KeyPress(KeyAscii As Integer)
   If Me.txtDescuento10.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoDe11.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoDe10_LostFocus()
    txtConDescuentoDe10.Text = Format(txtConDescuentoDe10.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta10_LostFocus()
txtConDescuentoHasta10.Text = Format(txtConDescuentoHasta10.Text, "#,##0.00")
End Sub

Private Sub txtDescuento10_LostFocus()
txtDescuento10.Text = Format(txtDescuento10.Text, "#,##0.00")
End Sub


Private Sub txtConDescuentoDe12_KeyPress(KeyAscii As Integer)
   If Me.txtConDescuentoDe12.Text <> "" Then
        If KeyAscii = 13 Then
            txtConDescuentoHasta12.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtConDescuentoHasta12_KeyPress(KeyAscii As Integer)
   If Me.txtConDescuentoHasta12.Text <> "" Then
        If KeyAscii = 13 Then
            txtDescuento12.Enabled = True
            SendKeys "{tab}"
        End If
    End If
End Sub

Private Sub txtDescuento12_KeyPress(KeyAscii As Integer)
   If Me.txtDescuento12.Text <> "" Then
        If KeyAscii = 13 Then
            Me.cmdRegistrar.Enabled = True
            cmdRegistrar.SetFocus
        End If
    End If
End Sub

Private Sub txtConDescuentoDe12_LostFocus()
    txtConDescuentoDe12.Text = Format(txtConDescuentoDe12.Text, "#,##0.00")
End Sub

Private Sub txtConDescuentoHasta12_LostFocus()
txtConDescuentoHasta12.Text = Format(txtConDescuentoHasta12.Text, "#,##0.00")
End Sub

Private Sub txtDescuento12_LostFocus()
txtDescuento12.Text = Format(txtDescuento12.Text, "#,##0.00")
End Sub

'Private Sub txtRentabilidad_LostFocus()
'    txtRentabilidad.Text = Format(txtRentabilidad.Text, "#,##0.00")
'End Sub

Private Sub txtNroCliente_KeyPress(KeyAscii As Integer)
    If Me.txtNroCliente.Text <> "" Then
        If KeyAscii = 13 Then
            Me.txtOpe.Enabled = True
            Me.txtOpe.SetFocus
        End If
    End If
End Sub
Private Sub txtOpe_KeyPress(KeyAscii As Integer)
    If Me.txtOpe.Text <> "" Then
        If KeyAscii = 13 Then
            Me.txtSaldo.Enabled = True
            Me.txtSaldo.SetFocus
        
        End If
    End If
End Sub

Private Sub txtSaldo_Change()
'    txtSaldo.Text = Format(lSaldo, "S/. ##,##0.00")
End Sub

Private Sub txtSaldo_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case 46
'            Me.txtSaldo.Text = ""
'            lSaldo = ""
'    End Select
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
    
        If KeyAscii = 13 Then
            If txtSaldo.Text <> "" Then
                Me.txtCMDeAceptable.Enabled = True
                Me.txtCMDeAceptable.SetFocus
            End If
'        ElseIf KeyAscii = 8 Then 'si es retroceso
'            If Len(lSaldo) > 1 Then
'                lSaldo = Mid(lSaldo, 1, Len(lSaldo) - 1) 'retrocede 1 digito
'            Else
'                lSaldo = "0"
'            End If
'                txtSaldo.Text = Format(lSaldo, "S/. ##,##0.00")
'        Else
'            If InStr("0123456789.", Chr(KeyAscii)) = 0 Then
'                KeyAscii = 0
'            Else
'                lSaldo = lSaldo + Chr(KeyAscii)
'                txtSaldo.Text = Format(lSaldo, "S/. ##,##0.00")
'            End If
        End If
  
End Sub

Private Sub txtSaldo_LostFocus()
     txtSaldo.Text = Format(txtSaldo.Text, "##,##0.00")
End Sub
