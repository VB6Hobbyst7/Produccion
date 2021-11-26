VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersEstadosFinancierosDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados Financieros"
   ClientHeight    =   10455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12810
   Icon            =   "frmPersEstadosFinancierosDetalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10455
   ScaleWidth      =   12810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   320
      Left            =   11700
      TabIndex        =   92
      Top             =   20
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   320
      Left            =   10580
      TabIndex        =   91
      Top             =   20
      Width           =   1095
   End
   Begin VB.CommandButton cmdSolicitarEdicionEF 
      Caption         =   "Solicitar Edición"
      Height          =   320
      Left            =   7560
      TabIndex        =   84
      Top             =   20
      Width           =   1335
   End
   Begin VB.CommandButton cmdAutorizarEF 
      Caption         =   "Autorizar Edición"
      Height          =   320
      Left            =   8940
      TabIndex        =   83
      Top             =   20
      Width           =   1335
   End
   Begin VB.CheckBox chkEEFFAuditado 
      Caption         =   "Auditado"
      Height          =   255
      Left            =   3600
      TabIndex        =   73
      Top             =   50
      Width           =   975
   End
   Begin TabDlg.SSTab SSTabDatosEEFF 
      Height          =   10155
      Left            =   0
      TabIndex        =   0
      Top             =   300
      Width           =   12795
      _ExtentX        =   22569
      _ExtentY        =   17912
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      ForeColor       =   -2147483635
      TabCaption(0)   =   "Balance General"
      TabPicture(0)   =   "frmPersEstadosFinancierosDetalle.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "feActivos"
      Tab(0).Control(1)=   "fePasivos"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Estado Resultado"
      TabPicture(1)   =   "frmPersEstadosFinancierosDetalle.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblExpresado"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblEstadoGP"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "feEstadoResultado"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Flujo Efectivo"
      TabPicture(2)   =   "frmPersEstadosFinancierosDetalle.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fefujoEfectivo"
      Tab(2).Control(1)=   "Line4"
      Tab(2).Control(2)=   "Label6"
      Tab(2).Control(3)=   "Label5"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Cambios en Patrimonio"
      TabPicture(3)   =   "frmPersEstadosFinancierosDetalle.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtPatrimonioIni"
      Tab(3).Control(1)=   "txtPatrimonioFin"
      Tab(3).Control(2)=   "feCambiosPatrimonio"
      Tab(3).Control(3)=   "Line1"
      Tab(3).Control(4)=   "Label10"
      Tab(3).Control(5)=   "Label8"
      Tab(3).Control(6)=   "Label4"
      Tab(3).Control(7)=   "Label3"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "EEFF Anterior"
      TabPicture(4)   =   "frmPersEstadosFinancierosDetalle.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraNotaEEFFAnterior"
      Tab(4).Control(1)=   "chkAudit"
      Tab(4).Control(2)=   "fraBalGeneral"
      Tab(4).Control(3)=   "fraFlujo"
      Tab(4).Control(4)=   "fraIndicadores"
      Tab(4).Control(5)=   "fraEstResutado"
      Tab(4).Control(6)=   "cmdGrabar"
      Tab(4).Control(7)=   "cmdCerrar"
      Tab(4).Control(8)=   "Label63"
      Tab(4).ControlCount=   9
      Begin VB.Frame fraNotaEEFFAnterior 
         Caption         =   "Nota:"
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
         Left            =   -74640
         TabIndex        =   74
         Top             =   480
         Width           =   11775
         Begin VB.Label lblNota 
            AutoSize        =   -1  'True
            Caption         =   "La información mostrada es de carácter informativo y en modo sólo lectura (representan EEFF registrados con anterioridad.). "
            Height          =   195
            Left            =   600
            TabIndex        =   75
            Top             =   360
            Width           =   8775
         End
      End
      Begin VB.CheckBox chkAudit 
         Caption         =   "Auditado"
         Height          =   255
         Left            =   -63840
         TabIndex        =   68
         Top             =   1500
         Width           =   975
      End
      Begin VB.Frame fraBalGeneral 
         Caption         =   "Balance General"
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
         Height          =   2415
         Left            =   -74760
         TabIndex        =   39
         Top             =   1680
         Width           =   11895
         Begin VB.TextBox txtEFTotalPat 
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
            Height          =   285
            Left            =   9960
            TabIndex        =   53
            Text            =   "0.00"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtEFResulAcum 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9960
            TabIndex        =   52
            Text            =   "0.00"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtEFTotalPas 
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
            Height          =   285
            Left            =   5880
            TabIndex        =   51
            Text            =   "0.00"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtEFDeudaFinanL 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5880
            TabIndex        =   50
            Text            =   "0.00"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtEFPasCorriente 
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
            Height          =   285
            Left            =   5880
            TabIndex        =   49
            Text            =   "0.00"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtEFDeudaFinanC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5880
            TabIndex        =   48
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtEFTotalAct 
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
            Height          =   285
            Left            =   2040
            TabIndex        =   47
            Text            =   "0.00"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtEFActFijo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   46
            Text            =   "0.00"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtEFActCorriente 
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
            Height          =   285
            Left            =   2040
            TabIndex        =   45
            Text            =   "0.00"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtEFExiste 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   44
            Text            =   "0.00"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtEFCuentaCobrar 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   43
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtEFCapSocial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9960
            TabIndex        =   42
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtEFProveedor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5880
            TabIndex        =   41
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtEFCajaBanco 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2040
            TabIndex        =   40
            Text            =   "0.00"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label62 
            Caption         =   "Capital Social + Adicional"
            Height          =   195
            Left            =   8040
            TabIndex        =   67
            Top             =   240
            Width           =   1905
         End
         Begin VB.Label Label61 
            Caption         =   "Resultados Acumulados"
            Height          =   195
            Left            =   8160
            TabIndex        =   66
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label60 
            Caption         =   "Total Patrimonio"
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
            Left            =   8400
            TabIndex        =   65
            Top             =   2040
            Width           =   1515
         End
         Begin VB.Label Label59 
            Caption         =   "Cuentas x Cobrar Comerc."
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   600
            Width           =   1845
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            Caption         =   "Existencias"
            Height          =   195
            Left            =   1200
            TabIndex        =   63
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label57 
            AutoSize        =   -1  'True
            Caption         =   "Activo Corriente"
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
            Left            =   600
            TabIndex        =   62
            Top             =   1320
            Width           =   1380
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "Activo Fijo"
            Height          =   195
            Left            =   1200
            TabIndex        =   61
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "Total Activo"
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
            Left            =   840
            TabIndex        =   60
            Top             =   2040
            Width           =   1050
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "Proveedores"
            Height          =   195
            Left            =   4800
            TabIndex        =   59
            Top             =   240
            Width           =   900
         End
         Begin VB.Label Label53 
            Caption         =   "Deuda Financiera Cte."
            Height          =   315
            Left            =   4200
            TabIndex        =   58
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label52 
            AutoSize        =   -1  'True
            Caption         =   "Pasivo Corriente"
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
            Left            =   4320
            TabIndex        =   57
            Top             =   960
            Width           =   1410
         End
         Begin VB.Label Label51 
            Caption         =   "Deuda Financiera a L.P."
            Height          =   315
            Left            =   4080
            TabIndex        =   56
            Top             =   1320
            Width           =   1725
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            Caption         =   "Total Pasivo"
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
            Left            =   4680
            TabIndex        =   55
            Top             =   2040
            Width           =   1080
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            Caption         =   "Caja-Bancos"
            Height          =   195
            Left            =   1080
            TabIndex        =   54
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame fraFlujo 
         Caption         =   "Estado de Flujo Efectivo"
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
         Height          =   2415
         Left            =   -68880
         TabIndex        =   29
         Top             =   4560
         Width           =   6015
         Begin VB.TextBox txtEFFlujoEfec 
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
            Height          =   285
            Left            =   3480
            TabIndex        =   33
            Text            =   "0.00"
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtEFFlujoFinan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3480
            TabIndex        =   32
            Text            =   "0.00"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtEFFujoInv 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3480
            TabIndex        =   31
            Text            =   "0.00"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtEFFlujoOpe 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3480
            TabIndex        =   30
            Text            =   "0.00"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label78 
            Caption         =   "Flujo de Efectivo por Act. de Inversión"
            Height          =   195
            Left            =   720
            TabIndex        =   37
            Top             =   840
            Width           =   2805
         End
         Begin VB.Label Label76 
            Caption         =   "Flujo de Efectivo por Act. de Financiamiento"
            Height          =   195
            Left            =   240
            TabIndex        =   36
            Top             =   1320
            Width           =   3285
         End
         Begin VB.Label Label75 
            Caption         =   "Flujo Efectivo Total"
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
            Left            =   1680
            TabIndex        =   35
            Top             =   1800
            Width           =   1785
         End
         Begin VB.Label Label64 
            Caption         =   "Flujo de Efectivo por Act. de Operación"
            Height          =   195
            Left            =   600
            TabIndex        =   34
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame fraIndicadores 
         Caption         =   "Indicadores de Riesgo Cambiario Créditicio"
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
         Height          =   975
         Left            =   -74880
         TabIndex        =   20
         Top             =   8280
         Width           =   12495
         Begin VB.TextBox txtEFPosCambios 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   600
            TabIndex        =   24
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtEFIngresoME 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7080
            TabIndex        =   23
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtEFDeudaFinan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3840
            TabIndex        =   22
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtEFGastosME 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9960
            TabIndex        =   21
            Text            =   "0.00"
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label73 
            Caption         =   "% Gastos en M.E."
            Height          =   195
            Left            =   9960
            TabIndex        =   28
            Top             =   360
            Width           =   1365
         End
         Begin VB.Label Label79 
            Caption         =   "Posición de Cambios (S/. 000)"
            Height          =   195
            Left            =   600
            TabIndex        =   27
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label69 
            Caption         =   "Deuda Financiera en M.E."
            Height          =   255
            Left            =   3840
            TabIndex        =   26
            Top             =   360
            Width           =   1965
         End
         Begin VB.Label Label65 
            Caption         =   "% Ingreso en M.E."
            Height          =   315
            Left            =   7080
            TabIndex        =   25
            Top             =   360
            Width           =   1365
         End
      End
      Begin VB.Frame fraEstResutado 
         Caption         =   "Estado de Resultados"
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
         Height          =   3615
         Left            =   -74760
         TabIndex        =   3
         Top             =   4560
         Width           =   5655
         Begin VB.TextBox txtEFIngresoFinan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3360
            TabIndex        =   11
            Text            =   "0.00"
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox txtEFUtNeta 
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
            Height          =   285
            Left            =   3360
            TabIndex        =   10
            Text            =   "0.00"
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox txtEFVentas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3360
            TabIndex        =   9
            Text            =   "0.00"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtEFCostVentas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3360
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtEFUtBruta 
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
            Height          =   285
            Left            =   3360
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtEFGastosOpe 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3360
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtEFUtOpe 
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
            Height          =   285
            Left            =   3360
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   1920
            Width           =   1695
         End
         Begin VB.TextBox txtEFGastoFinan 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3360
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "Ingresos Financieros"
            Height          =   195
            Left            =   1800
            TabIndex        =   19
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label Label77 
            AutoSize        =   -1  'True
            Caption         =   "Ventas"
            Height          =   195
            Left            =   2640
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label74 
            Caption         =   "Utilidad/ (Pérdida)Neta"
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
            Left            =   1200
            TabIndex        =   17
            Top             =   3120
            Width           =   1980
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            Caption         =   "Gastos Financieros"
            Height          =   195
            Left            =   1920
            TabIndex        =   16
            Top             =   2760
            Width           =   1350
         End
         Begin VB.Label Label70 
            Caption         =   "Utilidad/ (Pérdida)Operativa"
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
            Left            =   840
            TabIndex        =   15
            Top             =   1995
            Width           =   2460
         End
         Begin VB.Label Label68 
            Caption         =   "Utilidad/ (Pérdida)Bruta"
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
            Left            =   1200
            TabIndex        =   14
            Top             =   1080
            Width           =   2025
         End
         Begin VB.Label Label67 
            Caption         =   "Gastos Operativos (Adm+Srv+Vtas)"
            Height          =   195
            Left            =   720
            TabIndex        =   13
            Top             =   1560
            Width           =   2685
         End
         Begin VB.Label Label66 
            Caption         =   "Costo de Ventas"
            Height          =   315
            Left            =   2040
            TabIndex        =   12
            Top             =   720
            Width           =   1245
         End
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   -64800
         TabIndex        =   2
         Top             =   9405
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   -63600
         TabIndex        =   1
         Top             =   9405
         Visible         =   0   'False
         Width           =   1095
      End
      Begin SICMACT.FlexEdit feActivos 
         Height          =   9795
         Left            =   -74880
         TabIndex        =   72
         ToolTipText     =   "Favor ingresar los montos en soles. (TC - del día)"
         Top             =   330
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   17277
         Cols0           =   9
         ScrollBars      =   0
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "AuxInicial-Concepto-Totales-%-nConsCod-nConsValor-nConsValorGrupo-nEstFinanTipo-AuxFinal"
         EncabezadosAnchos=   "0-4000-1500-700-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-2-X-X-X-X-X-X"
         ListaControles  =   "0-0-1-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-C-C-C-C-C"
         FormatosEdit    =   "0-0-2-2-2-2-2-2-2"
         CantEntero      =   12
         TextArray0      =   "AuxInicial"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit fefujoEfectivo 
         Height          =   4935
         Left            =   -74685
         TabIndex        =   78
         Top             =   1680
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   8705
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "AuxInicial-Concepto-Totales-nConsCod-nConsValor-nEstFinanTipo-AuxFinal"
         EncabezadosAnchos=   "0-10000-2000-0-0-0-0"
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
         ColumnasAEditar =   "X-X-2-3-X-X-X"
         ListaControles  =   "0-0-1-1-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-C-C-C-C"
         FormatosEdit    =   "0-0-2-2-2-2-2"
         CantEntero      =   12
         TextArray0      =   "AuxInicial"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit fePasivos 
         Height          =   9795
         Left            =   -68560
         TabIndex        =   85
         Top             =   330
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   17277
         Cols0           =   9
         ScrollBars      =   0
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "AuxInicial-Concepto-Totales-%-nConsCod-nConsValor-nConsValorGrupo-nEstFinanTipo-AuxFinal"
         EncabezadosAnchos=   "0-4000-1500-700-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-2-X-X-X-X-X-X"
         ListaControles  =   "0-0-1-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-C-C-C-C-C"
         FormatosEdit    =   "0-0-2-2-2-2-2-2-2"
         CantEntero      =   12
         TextArray0      =   "AuxInicial"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   6
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin MSMask.MaskEdBox txtPatrimonioIni 
         Height          =   315
         Left            =   -70080
         TabIndex        =   86
         Top             =   1560
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtPatrimonioFin 
         Height          =   315
         Left            =   -68040
         TabIndex        =   87
         Top             =   1560
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.FlexEdit feEstadoResultado 
         Height          =   5535
         Left            =   1320
         TabIndex        =   69
         Top             =   1440
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   9763
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Concepto-Totales-Porcentaje-nConsCod-nConsValor-nEstFinanTipo-AuxFinal"
         EncabezadosAnchos=   "500-6200-1900-1200-0-0-0-0"
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
         ColumnasAEditar =   "X-X-2-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-C-C-C-C"
         FormatosEdit    =   "0-0-2-2-2-2-2-2"
         CantEntero      =   12
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   495
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit feCambiosPatrimonio 
         Height          =   3735
         Left            =   -74950
         TabIndex        =   90
         Top             =   2400
         Width           =   12660
         _ExtentX        =   22331
         _ExtentY        =   6588
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmPersEstadosFinancierosDetalle.frx":0396
         EncabezadosAnchos=   "0-3200-1550-1550-1550-1550-1550-1650-0-0-0"
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
         ColumnasAEditar =   "X-X-2-3-4-5-6-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-R-R-R-R-R-C-C-C"
         FormatosEdit    =   "0-0-2-2-2-2-2-2-2-2-2"
         CantEntero      =   12
         TextArray0      =   "AuxInicial"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Line Line4 
         X1              =   -72840
         X2              =   -63360
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line3 
         X1              =   1560
         X2              =   11040
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         X1              =   -73680
         X2              =   -63360
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label10 
         Caption         =   "Al :"
         Height          =   255
         Left            =   -68280
         TabIndex        =   89
         Top             =   1620
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "Del:"
         Height          =   255
         Left            =   -70440
         TabIndex        =   88
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "(Expresado en Soles)"
         Height          =   255
         Left            =   -69720
         TabIndex        =   80
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "FLUJO DE EFECTIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69840
         TabIndex        =   79
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "(Expresado en Soles)"
         Height          =   255
         Left            =   -69480
         TabIndex        =   77
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "ESTADO DE CAMBIO EN EL PATRIMONIO NETO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -70920
         TabIndex        =   76
         Top             =   960
         Width           =   4695
      End
      Begin VB.Label lblEstadoGP 
         Caption         =   "Estado de Ganancias y Pérdidas"
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
         Left            =   5040
         TabIndex        =   71
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblExpresado 
         Caption         =   "(Expresado en Soles)"
         Height          =   255
         Left            =   5520
         TabIndex        =   70
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label63 
         Caption         =   $"frmPersEstadosFinancierosDetalle.frx":0422
         Height          =   795
         Left            =   -68880
         TabIndex        =   38
         Top             =   7320
         Width           =   6015
      End
   End
   Begin MSMask.MaskEdBox txtFechaEF 
      Height          =   315
      Left            =   2280
      TabIndex        =   81
      Top             =   5
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label lblNumEdicion 
      Caption         =   "-"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6720
      TabIndex        =   95
      Top             =   60
      Width           =   135
   End
   Begin VB.Line Line6 
      X1              =   7320
      X2              =   7320
      Y1              =   45
      Y2              =   300
   End
   Begin VB.Label Label2 
      Caption         =   "]"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6850
      TabIndex        =   94
      Top             =   60
      Width           =   135
   End
   Begin VB.Line Line5 
      X1              =   4800
      X2              =   4800
      Y1              =   45
      Y2              =   300
   End
   Begin VB.Label Label1 
      Caption         =   "[Número de ediciones:"
      Enabled         =   0   'False
      Height          =   210
      Left            =   5040
      TabIndex        =   93
      Top             =   60
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   10440
      X2              =   10440
      Y1              =   45
      Y2              =   300
   End
   Begin VB.Label lblFechaEEFF 
      Caption         =   "Fecha del Estado Financiero:"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   82
      Top             =   40
      Width           =   2175
   End
End
Attribute VB_Name = "frmPersEstadosFinancierosDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************
'** DESARROLLADO POR: FRHU
'** REQUERIMIENTO: ERS013-2015
'** FECHA DESARROLLO: 20150326
'** MODIFICADO POR: LUCV20170915 / EAAS20170915 - ERS051-2017
'******************************
Option Explicit
Public Enum TEFCambios
    EFFilaSinCambios = 0
    EFFilaNueva = 1
    EFFilaModificada = 2
    EFFilaConsulta = 3
End Enum
Dim nTipoAccion As Integer  '1: Nuevo, 2:Editar, 3: Consultar, 4: Eliminar
Dim nCodEF As Long
Dim sPersCod As String
Dim sFechaEF As String
'*****-> LUCV20170915-Agregó, según: ERS051-2017 *****->
Dim fnTipoPermiso As Integer
Dim rsFeActivosEEFF As ADODB.Recordset
Dim rsFePasivosEEFF As ADODB.Recordset
Dim rsFeEstadoResultadoEEFF As ADODB.Recordset
Dim rsFeFlujoEfectivoEEFF As ADODB.Recordset
Dim rsFeCambiosPatrimonioEEFF As ADODB.Recordset
Dim rsDatosPersEstadosFinancieros As ADODB.Recordset
Dim ofrmPersEstadosFinancierosDetalleDet As frmPersEstadosFinancierosDetalleDet

Dim lvPrincipalActivos() As tfrmPersEstadosFinancieros             'Principal-> BalanceActivos
Dim lvPrincipalPasivos() As tfrmPersEstadosFinancieros             'Principal-> BalancePasivos
Dim lvPrincipalFlujoEfectivo() As tfrmPersEstadosFinancieros       'Principal-> FlujoEfectivo
Dim lvDetalleActivos() As tfrmPersEstadosFinancierosDetalle        'Detalle-> BalanceActivos
Dim lvDetallePasivos() As tfrmPersEstadosFinancierosDetalle        'Detalle-> BalancePasivos
Dim lvDetalleFlujoEfectivo() As tfrmPersEstadosFinancierosDetalle  'Detalle-> FlujoEfectivo

Dim i As Integer, j As Integer, K As Integer
Dim bActivoDetTotal(33) As Boolean
Dim bPasivoDetTotal(33) As Boolean
Dim bFlujoEfectivoDetTotal(16) As Boolean

'->***** LUCV20171015
Dim nMaximo As Integer
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
'<-***** Fin LUCV20171015

'<-*****Fin LUCV20170915 <-*****
Dim cEstadoEF As String 'EAAS20170915
Public Sub Inicio(ByVal pnCodEF As Integer, ByVal psPersCod As String, ByVal pnTipoAccion As TEFCambios, ByVal pnAuditado As Integer, Optional ByVal psFechaEF As String = "")
'LUCV20170915, Comentó según ERS051-2017
'    nTipoAccion = pnTipoAccion
'    nCodEF = pnCodEF
'    sPersCod = psPersCod
'    chkAudit.Enabled = True
'    If nTipoAccion = 2 Or nTipoAccion = 3 Then
'
'        chkAudit.value = IIf(pnAuditado = 1, 1, 0)
'        'txtFechaEF.Text = Format(psFechaEF, "yyyymmdd")
'        txtFechaEF.Text = Format(psFechaEF, "dd/mm/yyyy") 'FRHU 20160401
'        txtFechaEF.Enabled = False
'        chkAudit.Enabled = False
'        Call MostrarEstadosFinancieros(pnCodEF)
'
'
'        If nTipoAccion = 3 Then
'            Call HabilitaEstFinancieros(False)
'        End If
'    Else
'        txtFechaEF.Text = "__/__/____"
'    End If
'    Me.Show 1
'Fin LUCV20170915
    
    '**********->LUCV20170915, Agregó según ERS051-2017->**********
    '->********** EAAS20170926 ->**********
    sFechaEF = psFechaEF
    Dim bTipoPermiso As String
    Dim nEstado As Integer
    Dim nNumEditEF As Integer
    Dim nNumAut As Integer
    Dim bAutoriza As Boolean
    Dim sUser As String
    Dim sCargo As String
    Dim bCargoAutorizado As Boolean
    Dim nConsSisCod As Integer
    Dim oRS As ADODB.Recordset
    Dim oRsVerificarExisteAutorizacionEF As ADODB.Recordset
    Dim oRsObtieneTipoPermiso As ADODB.Recordset
    Dim oRsVerificaMigracion As ADODB.Recordset
    Dim oDFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDFormatosEval = New COMDCredito.DCOMFormatosEval
    Set oRsVerificarExisteAutorizacionEF = New ADODB.Recordset
    sCargo = gsCodCargo
    nConsSisCod = 702
    '<-********* fin EAAS20170915 <-**********
    
    nTipoAccion = pnTipoAccion
    nCodEF = pnCodEF
    sPersCod = psPersCod
    chkAudit.Enabled = True
    
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim psTipoAccion As String
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    If nTipoAccion = 1 Then
        psTipoAccion = "Nuevo"
    ElseIf nTipoAccion = 2 Then
        psTipoAccion = "Editar"
    ElseIf nTipoAccion = 3 Then
        psTipoAccion = "Consultar"
    ElseIf nTipoAccion Then
        psTipoAccion = "Eliminar"
    Else
        psTipoAccion = ""
    End If
    
    frmPersEstadosFinancierosDetalle.Caption = "Estados Financieros [" & psTipoAccion & "]"
    Call CargaControlesInicio
    
    '->********** EAAS20170915, Agregó según ERS051-2017 ->**********
    sUser = gsCodUser
    cEstadoEF = -1
    Set oRS = oDFormatosEval.RecuperarNumeroEdicionesEF(sPersCod, nCodEF)
    Set oRsObtieneTipoPermiso = oDFormatosEval.ObtieneTipoPermiso(gsCodCargo)
    Set oRsVerificaMigracion = oDFormatosEval.VerificaMigracionEstadoFinanciero(nCodEF)
    bCargoAutorizado = oDFormatosEval.VerificaPermisoEdicionEF(nConsSisCod, gsCodCargo)
    bTipoPermiso = oRsObtieneTipoPermiso!bExiste
    nNumEditEF = oRS.RecordCount
    '<-********** EAAS20170915 <-**********
    
    lblNumEdicion.Caption = nNumEditEF
    
    If nTipoAccion = 1 Then '=>Nuevo**
        txtFechaEF.Text = "__/__/____"
        Call HabilitaEstFinancieros(True, , , , True, True)
    ElseIf nTipoAccion = 2 Or nTipoAccion = 3 Then '=>Editar/Consultar **
        chkEEFFAuditado.value = IIf(pnAuditado = 1, 1, 0)
        txtFechaEF.Text = Format(psFechaEF, "dd/mm/yyyy")
        If CargaDatos Then
            Call MostrarEstadosFinancieros(pnCodEF)
            Call MostrarEstadosFinancierosNew(pnCodEF)
            Call HabilitaEstFinancieros(True, True, True, True, True, True)
            If nTipoAccion = 3 Then
                Call HabilitaEstFinancieros
            End If
        Else
            MsgBox "No existen datos actuales en la fecha seleccionada del EEFF del cliente", vbOKOnly, "Atención"
            Call MostrarEstadosFinancieros(pnCodEF)
            Call CargaControlesInicio
        End If
        
        '->***** EAAS20170915, Agregó según ERS051-2017 ->*****
        If nTipoAccion = 2 And bCargoAutorizado = False Then
            MsgBox "No tiene los permisos para la edición del EEFF. ", vbInformation, "Aviso"
            Exit Sub
        End If
        If nTipoAccion = 2 And oRsVerificaMigracion!bMigrado = 1 Then
            Set oRsVerificarExisteAutorizacionEF = oDFormatosEval.VerificarExisteAutorizacionEF(sPersCod, nCodEF, sUser)
            If oRsVerificarExisteAutorizacionEF.RecordCount = 1 Then
                cEstadoEF = oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto
                If oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto <> 2 Then
                    If oRsVerificarExisteAutorizacionEF!cPersUsuarioSolicitante = sUser Then
                        If oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto = 1 Then
                             MsgBox "Se autoriza la edicion del EEFF. ", vbInformation, "Aviso"
                             Call HabilitaEstFinancieros(True, False, False, True, True, True)
                        Else
                             If oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto = 3 Then
                                 MsgBox "Su solicitud de edición del EEFF ha sido Anulada, vuelva a solicitar la edición ", vbInformation, "Aviso"
                                 Call HabilitaEstFinancieros(True, True, False, True, False, True)
                             Else
                                MsgBox "Su solicitud de edición del EEFF aún no ha sido autorizada. ", vbInformation, "Aviso"
                                If bTipoPermiso = 1 Then
                                 Call HabilitaEstFinancieros(True, False, True, True, False, True)
                                Else
                                Call HabilitaEstFinancieros(True, False, False, True, False, True)
                                End If
                             End If
                        End If
                Else
                       If oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto = 3 Then
                             MsgBox "Se alcanzó las 3 veces permitidas de edición del EEFF,solicite la edición del EEFF . ", vbInformation, "Aviso"
                             Call HabilitaEstFinancieros(True, True, False, True, False, True)
                        Else
                            MsgBox "Ya hay una solicitud de edición pendiente del Usuario " & oRsVerificarExisteAutorizacionEF!cPersUsuarioSolicitante & " para el EEFF. ", vbInformation, "Aviso"
                            If bTipoPermiso = 1 Then
                            Call HabilitaEstFinancieros(True, False, True, True, False, True)
                            Else
                            Call HabilitaEstFinancieros(True, False, False, True, False, True)
                            End If
                        End If
                End If
                Else
                        MsgBox "Se alcanzó el límite permitido de Edición(3 veces), solicite su autorización al Dpto. Riesgos. ", vbInformation, "Aviso"
                        Call HabilitaEstFinancieros(True, True, False, True, False, True)
                End If
            Else
                        MsgBox "Se alcanzó el límite permitido de Edición(3 veces), solicite su autorización al Dpto. Riesgos.", vbInformation, "Aviso"
                        If bTipoPermiso = 1 Then
                        Call HabilitaEstFinancieros(True, True, False, True, False, True)
                        Else
                        Call HabilitaEstFinancieros(True, True, False, True, False, True)
                        End If
            End If
        ElseIf nTipoAccion = 2 And nNumEditEF < 3 Then
            Call HabilitaEstFinancieros(True, False, False, True, True, True)
        ElseIf nTipoAccion = 2 And nNumEditEF >= 3 Then
            Set oRsVerificarExisteAutorizacionEF = oDFormatosEval.VerificarExisteAutorizacionEF(sPersCod, nCodEF, sUser) 'EAAS
            If oRsVerificarExisteAutorizacionEF.RecordCount = 1 Then
                cEstadoEF = oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto
                If oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto <> 2 Then
                    If oRsVerificarExisteAutorizacionEF!cPersUsuarioSolicitante = sUser Then
                        If oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto = 1 Then
                             MsgBox "Se autoriza la edicion del EEFF. ", vbInformation, "Aviso"
                             Call HabilitaEstFinancieros(True, False, False, True, True, True)
                        Else
                             If oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto = 3 Then
                                 MsgBox "Su solicitud de edición del EEFF ha sido Anulada, vuelva a solicitar la edición ", vbInformation, "Aviso"
                                 Call HabilitaEstFinancieros(True, False, False, True, True, True)
                             Else
                                MsgBox "Su solicitud de edición del EEFF aún no ha sido autorizada. ", vbInformation, "Aviso"
                                If bTipoPermiso = 1 Then
                                    Call HabilitaEstFinancieros(True, False, False, True, False, True)
                                Else
                                    Call HabilitaEstFinancieros(True, False, False, True, False, True)
                                End If
                             End If
                        End If
                Else
                       If oRsVerificarExisteAutorizacionEF!nEstadoSoliAuto = 3 Then
                             MsgBox "Se alcanzó las 3 veces permitidas de edición del EEFF,solicite la edición del EEFF . ", vbInformation, "Aviso"
                             Call HabilitaEstFinancieros(True, True, False, True, False, True)
                        Else
                            MsgBox "Ya hay una solicitud de edición pendiente del Usuario " & oRsVerificarExisteAutorizacionEF!cPersUsuarioSolicitante & " para el EEFF. ", vbInformation, "Aviso"
                            If bTipoPermiso = 1 Then
                                Call HabilitaEstFinancieros(True, False, True, True, False, True)
                            Else
                                Call HabilitaEstFinancieros(True, False, False, True, False, True)
                            End If
                        End If
                End If
                Else
                        MsgBox "Se alcanzó el límite permitido de Edición(3 veces), solicite su autorización al Dpto. Riesgos. ", vbInformation, "Aviso"
                        Call HabilitaEstFinancieros(True, True, False, True, False, True)
                End If
            Else
                        MsgBox "Se alcanzó el límite permitido de Edición(3 veces), solicite su autorización al Dpto. Riesgos. ", vbInformation, "Aviso"
                        Call HabilitaEstFinancieros(True, True, False, True, False, True)
            End If
        End If
        '<-***** Fin EAAS20170915 <-*****
    End If
    Me.SSTabDatosEEFF.Tab = 0
        Me.Show 1
    '<-********** Fin LUCV20170915 <-**********
End Sub
Private Sub cmdGrabar_Click()
    Dim oDPersona As COMDPersona.DCOMPersonas
    '->***** LUCV20170915, Agregó segun ERS051-2017
    Dim dFechaPatriIni, dFechaPatriFin As Date
    dFechaPatriIni = txtPatrimonioIni.Text
    dFechaPatriFin = txtPatrimonioFin.Text
    '<-***** Fin LUCV20170915
    If Not HayDatosEstFinan Then
        MsgBox "Verificar si se registraron todos los datos"
        Exit Sub
    End If
    If Right(Format(txtFechaEF.Text, "yyyymmdd"), 4) = "0630" Or Right(Format(txtFechaEF.Text, "yyyymmdd"), 4) = "1231" Then
        If nTipoAccion = 1 Then
            If ExisteEstadoFinanciero Then
                MsgBox "Ya existe la Fecha de Estado Financiero."
                Exit Sub
            End If
            Set oDPersona = New COMDPersona.DCOMPersonas
            nCodEF = oDPersona.GrabarEstFinan(sPersCod, txtFechaEF.Text, gsCodUser, IIf(chkAudit.value = 1, 1, 0), Format(dFechaPatriIni, "yyyymmdd"), Format(dFechaPatriFin, "yyyymmdd"))
            If nCodEF = 0 Then
                MsgBox ("Los datos no se guardaron correctamente, vuelve a grabar")
                Exit Sub
            End If
            Set oDPersona = Nothing
        End If
    Else
        MsgBox "Solo se pueden ingresar fechas Semestrales 30/06/Año o 31/12/Año"
        Exit Sub
    End If
    Set oDPersona = New COMDPersona.DCOMPersonas
    Call oDPersona.GrabarEstFinanDet(CargaRSEstFinanDet, nCodEF)
    Set oDPersona = Nothing
    
    MsgBox ("Los datos se guardaron correctamente")
    Call LimpiarEstFinan
    Unload Me
End Sub
Private Sub cmdCerrar_Click()
    Call LimpiarEstFinan
    Unload Me
End Sub
Private Sub HabilitaEstFinancieros(Optional ByVal pbHabEstFinanFecha As Boolean = 0, _
                                  Optional ByVal pbHabEstFinanEdicion As Boolean = 0, _
                                  Optional ByVal pbHabEstFinanAutoriza As Boolean = 0, _
                                  Optional ByVal pbHabEstFinanComparativo As Boolean = 0, _
                                  Optional ByVal pbHabEstFinanGuardar As Boolean = 0, _
                                  Optional ByVal pbHabEstFinanAuditado As Boolean = 0, _
                                  Optional ByVal pbHabEstFinanAnt As Boolean = 0) 'LUCV20170915 -Agregó: HabEstFinanNew, según: ERS051 -2017
    'Controles anteriores
    Me.fraBalGeneral.Enabled = pbHabEstFinanAnt
    Me.fraEstResutado.Enabled = pbHabEstFinanAnt
    Me.fraFlujo.Enabled = pbHabEstFinanAnt
    Me.fraIndicadores.Enabled = pbHabEstFinanAnt
    
    'Controles nuevos
    txtFechaEF.Enabled = pbHabEstFinanFecha
    cmdSolicitarEdicionEF.Enabled = pbHabEstFinanEdicion
    cmdAutorizarEF.Enabled = pbHabEstFinanAutoriza
    
    chkEEFFAuditado.Enabled = pbHabEstFinanAuditado
    cmdGuardar.Enabled = pbHabEstFinanGuardar
End Sub
Private Function ExisteEstadoFinanciero() As Boolean
    Dim oDPersona As COMDPersona.DCOMPersonas
    Dim RS As ADODB.Recordset
    
    Set oDPersona = New COMDPersona.DCOMPersonas
    Set RS = oDPersona.ExistePersonaEstadoFinanciero(sPersCod, txtFechaEF.Text)
    Set oDPersona = Nothing
    
    If Not RS.BOF And Not RS.EOF Then
        ExisteEstadoFinanciero = True
    Else
        ExisteEstadoFinanciero = False
    End If
    Exit Function
End Function
Private Sub CalculoTotal(ByVal pnTipo As Integer)
On Error GoTo ErrorCalculo
Select Case pnTipo
    Case 1:
            txtEFActCorriente.Text = Format(CDbl(txtEFCajaBanco.Text) + CDbl(txtEFCuentaCobrar.Text) + CDbl(txtEFExiste.Text), "###," & String(15, "#") & "#0.00")
            Call CalculoTotal(2)
    Case 2:
            txtEFTotalAct.Text = Format(CDbl(txtEFActFijo.Text) + CDbl(txtEFActCorriente.Text), "###," & String(15, "#") & "#0.00")
    Case 3:
            txtEFPasCorriente.Text = Format(CDbl(txtEFProveedor.Text) + CDbl(txtEFDeudaFinanC.Text), "###," & String(15, "#") & "#0.00")
            Call CalculoTotal(4)
    Case 4:
            txtEFTotalPas.Text = Format(CDbl(txtEFPasCorriente.Text) + CDbl(txtEFDeudaFinanL.Text), "###," & String(15, "#") & "#0.00")
    Case 5:
            txtEFTotalPat.Text = Format(CDbl(txtEFCapSocial.Text) + CDbl(txtEFResulAcum.Text), "###," & String(15, "#") & "#0.00")
    Case 6:
            txtEFUtBruta.Text = Format(CDbl(txtEFVentas.Text) - CDbl(txtEFCostVentas.Text), "###," & String(15, "#") & "#0.00")
            Call CalculoTotal(7)
    Case 7:
            txtEFUtOpe.Text = Format(CDbl(txtEFUtBruta.Text) - CDbl(txtEFGastosOpe.Text), "###," & String(15, "#") & "#0.00")
            Call CalculoTotal(8)
    Case 8:
            txtEFUtNeta.Text = Format(CDbl(txtEFUtOpe.Text) + CDbl(txtEFIngresoFinan.Text) - CDbl(txtEFGastoFinan.Text), "###," & String(15, "#") & "#0.00")
    Case 9:
            txtEFFlujoEfec.Text = Format(CDbl(txtEFFlujoOpe.Text) + CDbl(txtEFFujoInv.Text) + CDbl(txtEFFlujoFinan.Text), "###," & String(15, "#") & "#0.00")
End Select
Exit Sub

ErrorCalculo:
MsgBox "Error: Ingrese los datos Correctamente." & Chr(13) & "Detalles de error: " & Err.Description, vbCritical, "Error"
Select Case pnTipo
    Case 1:
            txtEFCajaBanco.Text = "0.00"
            txtEFCuentaCobrar.Text = "0.00"
            txtEFExiste.Text = "0.00"
    Case 2:
            txtEFActFijo.Text = "0.00"
    Case 3:
            txtEFProveedor.Text = "0.00"
            txtEFDeudaFinanC.Text = "0.00"
    Case 4:
            txtEFDeudaFinanL.Text = "0.00"
    Case 5:
            txtEFCapSocial.Text = "0.00"
            txtEFResulAcum.Text = "0.00"
    Case 6:
            txtEFVentas.Text = "0.00"
            txtEFCostVentas.Text = "0.00"
    Case 7:
            txtEFGastosOpe.Text = "0.00"
    Case 8:
            txtEFIngresoFinan.Text = "0.00"
            txtEFGastoFinan.Text = "0.00"
    Case 9:
            txtEFFlujoEfec.Text = "0.00"
            txtEFFlujoOpe.Text = "0.00"
            txtEFFujoInv.Text = "0.00"
            txtEFFlujoFinan.Text = "0.00"
End Select
    Call CalculoTotal(pnTipo)
End Sub
Private Sub MostrarEstadosFinancieros(ByVal pnCodEstFinan As Integer)
    Dim oDPersona  As COMDPersona.DCOMPersonas
    Dim rsDatos  As ADODB.Recordset
    Dim i As Integer
    Dim sCadena As String
    Set oDPersona = New COMDPersona.DCOMPersonas
    Set rsDatos = Nothing
    Dim fbEditarEF As Boolean
    sCadena = "###," & String(15, "#") & "#0.00"
    Call LimpiarEstFinan
    Set rsDatos = oDPersona.RecuperaEstFinanDetalle(pnCodEstFinan)
     
        If Not (rsDatos.EOF And rsDatos.BOF) Then
            fbEditarEF = True
            For i = 1 To rsDatos.RecordCount
                Select Case CInt(rsDatos!nEstFinanTipo)
                    Case 1:
                        Select Case CInt(rsDatos!nEstFinanSubTipo)
                            Case 1: txtEFCajaBanco.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 2: txtEFCuentaCobrar.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 3: txtEFExiste.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 4: txtEFActCorriente.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 5: txtEFActFijo.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 6: txtEFTotalAct.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 7: txtEFProveedor.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 8: txtEFDeudaFinanC.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 9: txtEFPasCorriente.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 10: txtEFDeudaFinanL.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 11: txtEFTotalPas.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 12: txtEFCapSocial.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 13: txtEFResulAcum.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 14: txtEFTotalPat.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        End Select
                    Case 2:
                        Select Case CInt(rsDatos!nEstFinanSubTipo)
                            Case 1: txtEFVentas.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 2: txtEFCostVentas.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 3: txtEFUtBruta.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 4: txtEFGastosOpe.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 5: txtEFUtOpe.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 6: txtEFIngresoFinan.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 7: txtEFGastoFinan.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 8: txtEFUtNeta.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        End Select
                    Case 3:
                          Select Case CInt(rsDatos!nEstFinanSubTipo)
                            Case 1: txtEFFlujoOpe.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 2: txtEFFujoInv.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 3: txtEFFlujoFinan.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 4: txtEFFlujoEfec.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            End Select
                    Case 4:
                         Select Case CInt(rsDatos!nEstFinanSubTipo)
                            Case 1: txtEFPosCambios.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 2: txtEFDeudaFinan.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 3: txtEFIngresoME.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                            Case 4: txtEFGastosME.Text = Format(CDbl(rsDatos!nMonto), sCadena)
                        End Select
                End Select
                rsDatos.MoveNext
            Next i
        Else
            fbEditarEF = False
        End If
End Sub

Private Function CargaRSEstFinanDet() As ADODB.Recordset
    Dim rsEF As ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    Set rsEF = New ADODB.Recordset
    With rsEF
        'Crear RecordSet
        .Fields.Append "nCodEstFinan", adInteger           '1
        .Fields.Append "nEstFinanTipo", adInteger          '2
        .Fields.Append "nEstFinanSubTipo", adInteger       '3
        .Fields.Append "nMonto", adCurrency                '4
        .Open
        
        'Llenar Recordset
        If HayDatosEstFinan Then
            For i = 1 To 4
                For j = 1 To 14
                    .AddNew
                    .Fields("nCodEstFinan") = nCodEF
                    .Fields("nEstFinanTipo") = i
                    .Fields("nEstFinanSubTipo") = j
                    
                    Select Case i
                        Case 1:
                            Select Case j
                                Case 1: .Fields("nMonto") = CDbl(txtEFCajaBanco.Text)
                                Case 2: .Fields("nMonto") = CDbl(txtEFCuentaCobrar.Text)
                                Case 3: .Fields("nMonto") = CDbl(txtEFExiste.Text)
                                Case 4: .Fields("nMonto") = CDbl(txtEFActCorriente.Text)
                                Case 5: .Fields("nMonto") = CDbl(txtEFActFijo.Text)
                                Case 6: .Fields("nMonto") = CDbl(txtEFTotalAct.Text)
                                Case 7: .Fields("nMonto") = CDbl(txtEFProveedor.Text)
                                Case 8: .Fields("nMonto") = CDbl(txtEFDeudaFinanC.Text)
                                Case 9: .Fields("nMonto") = CDbl(txtEFPasCorriente.Text)
                                Case 10: .Fields("nMonto") = CDbl(txtEFDeudaFinanL.Text)
                                Case 11: .Fields("nMonto") = CDbl(txtEFTotalPas.Text)
                                Case 12: .Fields("nMonto") = CDbl(txtEFCapSocial.Text)
                                Case 13: .Fields("nMonto") = CDbl(txtEFResulAcum.Text)
                                Case 14: .Fields("nMonto") = CDbl(txtEFTotalPat.Text)
                                        Exit For
                            End Select
                        Case 2:
                            Select Case j
                                Case 1: .Fields("nMonto") = CDbl(txtEFVentas.Text)
                                Case 2: .Fields("nMonto") = CDbl(txtEFCostVentas.Text)
                                Case 3: .Fields("nMonto") = CDbl(txtEFUtBruta.Text)
                                Case 4: .Fields("nMonto") = CDbl(txtEFGastosOpe.Text)
                                Case 5: .Fields("nMonto") = CDbl(txtEFUtOpe.Text)
                                Case 6: .Fields("nMonto") = CDbl(txtEFIngresoFinan.Text)
                                Case 7: .Fields("nMonto") = CDbl(txtEFGastoFinan.Text)
                                Case 8: .Fields("nMonto") = CDbl(txtEFUtNeta.Text)
                                        Exit For
                            End Select
                        Case 3:
                              Select Case j
                                Case 1: .Fields("nMonto") = CDbl(txtEFFlujoOpe.Text)
                                Case 2: .Fields("nMonto") = CDbl(txtEFFujoInv.Text)
                                Case 3: .Fields("nMonto") = CDbl(txtEFFlujoFinan.Text)
                                Case 4: .Fields("nMonto") = CDbl(txtEFFlujoEfec.Text)
                                        Exit For
                                End Select
                        Case 4:
                             Select Case j
                                Case 1: .Fields("nMonto") = CDbl(txtEFPosCambios.Text)
                                Case 2: .Fields("nMonto") = CDbl(txtEFDeudaFinan.Text)
                                Case 3: .Fields("nMonto") = CDbl(txtEFIngresoME.Text)
                                Case 4: .Fields("nMonto") = CDbl(txtEFGastosME.Text)
                                        Exit For
                            End Select
                    End Select
                    
                Next j
            Next i
        End If
        If Not .EOF Then .MoveFirst
    End With
    Set CargaRSEstFinanDet = rsEF
End Function
Private Function HayDatosEstFinan() As Boolean
    HayDatosEstFinan = False
    If Not (Trim(txtEFCajaBanco.Text) = "0.00" And _
        Trim(txtEFCuentaCobrar.Text) = "0.00" And _
        Trim(txtEFExiste.Text) = "0.00" And _
        Trim(txtEFActCorriente.Text) = "0.00" And _
        Trim(txtEFActFijo.Text) = "0.00" And _
        Trim(txtEFTotalAct.Text) = "0.00" And _
        Trim(txtEFProveedor.Text) = "0.00" And _
        Trim(txtEFDeudaFinanC.Text) = "0.00" And _
        Trim(txtEFPasCorriente.Text) = "0.00" And _
        Trim(txtEFDeudaFinanL.Text) = "0.00" And _
        Trim(txtEFTotalPas.Text) = "0.00" And _
        Trim(txtEFCapSocial.Text) = "0.00" And _
        Trim(txtEFResulAcum.Text) = "0.00" And _
        Trim(txtEFTotalPat.Text) = "0.00") Then
       HayDatosEstFinan = True
    End If
       
    If Not (Trim(txtEFVentas.Text) = "0.00" And _
       Trim(txtEFCostVentas.Text) = "0.00" And _
       Trim(txtEFUtBruta.Text) = "0.00" And _
       Trim(txtEFGastosOpe.Text) = "0.00" And _
       Trim(txtEFUtOpe.Text) = "0.00" And _
       Trim(txtEFIngresoFinan.Text) = "0.00" And _
       Trim(txtEFGastoFinan.Text) = "0.00" And _
       Trim(txtEFUtNeta.Text) = "0.00") Then
         HayDatosEstFinan = True
    End If
    
    If Not (Trim(txtEFFlujoOpe.Text) = "0.00" And _
       Trim(txtEFFujoInv.Text) = "0.00" And _
       Trim(txtEFFlujoFinan.Text) = "0.00" And _
       Trim(txtEFFlujoEfec.Text) = "0.00") Then
        HayDatosEstFinan = True
    End If
    
    If Not (Trim(txtEFPosCambios.Text) = "0.00" And _
       Trim(txtEFDeudaFinan.Text) = "0.00" And _
       Trim(txtEFIngresoME.Text) = "0.00" And _
       Trim(txtEFGastosME.Text) = "0.00") Then
         HayDatosEstFinan = True
    End If
End Function
Private Sub LimpiarEstFinan()
    txtEFCajaBanco.Text = "0.00"
    txtEFCuentaCobrar.Text = "0.00"
    txtEFExiste.Text = "0.00"
    txtEFActFijo.Text = "0.00"
    txtEFActCorriente.Text = "0.00"
    txtEFTotalAct.Text = "0.00"
    txtEFProveedor.Text = "0.00"
    txtEFDeudaFinanC.Text = "0.00"
    txtEFPasCorriente.Text = "0.00"
    txtEFDeudaFinanL.Text = "0.00"
    txtEFTotalPas.Text = "0.00"
    txtEFResulAcum.Text = "0.00"
    txtEFCapSocial.Text = "0.00"
    txtEFTotalPat.Text = "0.00"
    txtEFCostVentas.Text = "0.00"
    txtEFVentas.Text = "0.00"
    txtEFUtBruta.Text = "0.00"
    txtEFGastosOpe.Text = "0.00"
    txtEFUtOpe.Text = "0.00"
    txtEFIngresoFinan.Text = "0.00"
    txtEFGastoFinan.Text = "0.00"
    txtEFUtNeta.Text = "0.00"
    txtEFFlujoOpe.Text = "0.00"
    txtEFFujoInv.Text = "0.00"
    txtEFFlujoFinan.Text = "0.00"
    txtEFFlujoEfec.Text = "0.00"
    txtEFPosCambios.Text = "0.00"
    txtEFDeudaFinan.Text = "0.00"
    txtEFIngresoME.Text = "0.00"
    txtEFGastosME.Text = "0.00"
End Sub

Private Sub fefujoEfectivo_OnCellChange(pnRow As Long, pnCol As Long)
    Select Case CInt(fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, 0)) 'Activa TextBuscar
        Case 1, 4, 5, 8, 9, 12, 13, 14, 15
            Me.fefujoEfectivo.ListaControles = "0-0-0-0-0-0-0"
        Case Else
            Me.fefujoEfectivo.ListaControles = "0-0-1-0-0-0-0"
        End Select

    Select Case CInt(fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, 0)) 'Celdas que pueden Editarse
        Case 1, 4, 5, 8, 9, 12, 13, 15
            Me.fefujoEfectivo.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.fefujoEfectivo.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
    Call CalculoCeldasFlujoEfectivo
End Sub

Private Sub fePasivos_OnCellChange(pnRow As Long, pnCol As Long)
    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0)) 'Activa TextBuscar
        Case 1, 23, 30, 31, 32
            Me.fePasivos.ListaControles = "0-0-0-0-0-0-0"
        Case Else
            Me.fePasivos.ListaControles = "0-0-1-0-0-0-0"
    End Select

    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0)) 'Celdas que pueden Editarse
        Case 1, 23, 31, 32, 24
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.fePasivos.ColumnasAEditar = "X-X-2-X-X-X-X"
    End Select
    Call CalculoCeldasBalanceGeneral(1)
    Call CalculoCeldasBalanceGeneral(2)
End Sub
'**********-> LUCV20170915-Agregó, según: ERS051-2017 *****->
Private Sub Form_Load()
    CentraForm Me
    EnfocaControl feActivos

    bActivoDetTotal(1) = False: bActivoDetTotal(2) = False: bActivoDetTotal(3) = False: bActivoDetTotal(4) = False: bActivoDetTotal(5) = False: bActivoDetTotal(6) = False
    bActivoDetTotal(7) = False: bActivoDetTotal(8) = False: bActivoDetTotal(9) = False: bActivoDetTotal(10) = False: bActivoDetTotal(11) = False: bActivoDetTotal(12) = False
    bActivoDetTotal(13) = False: bActivoDetTotal(14) = False: bActivoDetTotal(15) = False: bActivoDetTotal(16) = False: bActivoDetTotal(17) = False: bActivoDetTotal(18) = False
    bActivoDetTotal(19) = False: bActivoDetTotal(20) = False: bActivoDetTotal(21) = False: bActivoDetTotal(22) = False: bActivoDetTotal(23) = False: bActivoDetTotal(24) = False
    bActivoDetTotal(25) = False: bActivoDetTotal(26) = False: bActivoDetTotal(27) = False: bActivoDetTotal(28) = False: bActivoDetTotal(29) = False: bActivoDetTotal(30) = False
    bActivoDetTotal(31) = False: bActivoDetTotal(32) = False
    
    bPasivoDetTotal(1) = False: bPasivoDetTotal(2) = False: bPasivoDetTotal(3) = False: bPasivoDetTotal(4) = False: bPasivoDetTotal(5) = False: bPasivoDetTotal(6) = False
    bPasivoDetTotal(7) = False: bPasivoDetTotal(8) = False: bPasivoDetTotal(9) = False: bPasivoDetTotal(10) = False: bPasivoDetTotal(11) = False: bPasivoDetTotal(12) = False
    bPasivoDetTotal(13) = False: bPasivoDetTotal(14) = False: bPasivoDetTotal(15) = False: bPasivoDetTotal(16) = False: bPasivoDetTotal(17) = False: bPasivoDetTotal(18) = False
    bPasivoDetTotal(19) = False: bPasivoDetTotal(20) = False: bPasivoDetTotal(21) = False: bPasivoDetTotal(22) = False: bPasivoDetTotal(23) = False: bPasivoDetTotal(24) = False
    bPasivoDetTotal(25) = False: bPasivoDetTotal(26) = False: bPasivoDetTotal(27) = False: bPasivoDetTotal(28) = False: bPasivoDetTotal(29) = False: bPasivoDetTotal(30) = False
    bPasivoDetTotal(31) = False: bPasivoDetTotal(32) = False
    
    bFlujoEfectivoDetTotal(1) = False: bFlujoEfectivoDetTotal(2) = False: bFlujoEfectivoDetTotal(3) = False: bFlujoEfectivoDetTotal(4) = False: bFlujoEfectivoDetTotal(5) = False: bFlujoEfectivoDetTotal(6) = False
    bFlujoEfectivoDetTotal(7) = False: bFlujoEfectivoDetTotal(8) = False: bFlujoEfectivoDetTotal(9) = False: bFlujoEfectivoDetTotal(10) = False: bFlujoEfectivoDetTotal(11) = False: bFlujoEfectivoDetTotal(12) = False
    bFlujoEfectivoDetTotal(13) = False: bFlujoEfectivoDetTotal(14) = False: bFlujoEfectivoDetTotal(15) = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
    If KeyCode = 113 And Shift = 0 Then
        KeyCode = 10
    End If
    If KeyCode = 27 And Shift = 0 Then
        Unload Me
    End If
End Sub
Private Sub Cmdguardar_Click()
    'LUCV20170915
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Dim oDCOMFormatosEval As COMDCredito.DCOMFormatosEval
    Dim oDPersona As COMDPersona.DCOMPersonas
    Set oDCOMFormatosEval = New COMDCredito.DCOMFormatosEval
    Dim rsEstadoResultado As ADODB.Recordset
    Dim rsFlujoEfectivo As ADODB.Recordset
    Dim rsCambiosPatrimonio As ADODB.Recordset
    Dim GrabarDatos As Boolean
    
    Dim nEstFinanOrdenItem As Integer
    Dim nEstFinanTipo As Integer
    Dim nEstFinanSubTipoCod As Integer
    Dim nEstFinanSubTipoValor As Integer
    Dim nEstFinanSubTipoValorGrupo As Integer
    Dim nEstFinanCapitalSocial As Currency
    Dim nEstFinanReservaLegal As Currency
    Dim nEstFinanResultadoAcumulados As Currency
    Dim nMonto As Currency
    Dim nPorcentaje As Double
    Dim dFechaPatriIni, dFechaPatriFin As Date
    
    nEstFinanOrdenItem = 0: nEstFinanTipo = 0: nEstFinanSubTipoCod = 0: nEstFinanSubTipoValor = 0: nEstFinanSubTipoValorGrupo = 0:
    nEstFinanCapitalSocial = 0: nEstFinanReservaLegal = 0: nEstFinanResultadoAcumulados = 0: nMonto = 0: nPorcentaje = 0
    Set rsEstadoResultado = IIf(feEstadoResultado.rows - 1 > 0, feEstadoResultado.GetRsNew(), Nothing)
    Set rsFlujoEfectivo = IIf(fefujoEfectivo.rows - 1 > 0, fefujoEfectivo.GetRsNew(), Nothing)
    Set rsCambiosPatrimonio = IIf(feCambiosPatrimonio.rows - 1 > 0, feCambiosPatrimonio.GetRsNew(), Nothing)
    'Fin LUCV20170915
    
    'EAAS20170915
    Dim lcMovNroEF As String
    Dim nNumEditEF As Integer
    Dim oRS As ADODB.Recordset
    Dim oDFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDFormatosEval = New COMDCredito.DCOMFormatosEval
    Set oRS = oDFormatosEval.RecuperarNumeroEdicionesEF(sPersCod, nCodEF)
    nNumEditEF = oRS.RecordCount
    'Fin EAAS20170915
    
    If ValidaDatos Then
    If MsgBox("Se guardarán los datos ingresados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
        dFechaPatriIni = txtPatrimonioIni.Text
        dFechaPatriFin = txtPatrimonioFin.Text
        '***** Registra datos de identificacion/ Cabecera EEFF ****
        If Right(Format(txtFechaEF.Text, "yyyymmdd"), 4) = "0630" Or Right(Format(txtFechaEF.Text, "yyyymmdd"), 4) = "1231" Then
            If nTipoAccion = 1 Then
                If ExisteEstadoFinanciero Then
                    MsgBox "Ya existe la Fecha de Estado Financiero."
                    Exit Sub
                End If
                Set oDPersona = New COMDPersona.DCOMPersonas
                nCodEF = oDPersona.GrabarEstFinan(sPersCod, txtFechaEF.Text, gsCodUser, IIf(chkEEFFAuditado.value = 1, 1, 0), Format(dFechaPatriIni, "yyyymmdd"), Format(dFechaPatriFin, "yyyymmdd"))
                If nCodEF = 0 Then
                    MsgBox ("Los datos no se guardaron correctamente, vuelve a grabar")
                    Exit Sub
                End If
                Set oDPersona = Nothing
                
            ElseIf nTipoAccion = 2 Then
                Call oDCOMFormatosEval.ActualizaPersEstadoFinanciero(nCodEF, IIf(chkEEFFAuditado.value = 1, 1, 0), Format(dFechaPatriIni, "yyyymmdd"), Format(dFechaPatriFin, "yyyymmdd"))
            End If
        Else
            MsgBox "Solo se pueden ingresar fechas Semestrales 30/06/Año o 31/12/Año"
            Exit Sub
        End If

            '************************************************************************************
            '******* Carga Subitems de detalle en matices, para el caso de Mantenimiento ********
            '************************************************************************************
            '****->Detalles SubItem FlujoEfectivo
            If UBound(lvPrincipalFlujoEfectivo) > 0 Then
                For i = 1 To UBound(lvPrincipalFlujoEfectivo)
                    If IsArray(lvPrincipalFlujoEfectivo(i).vTotal) Then
                        If UBound(lvPrincipalFlujoEfectivo(i).vTotal) <= 0 And bFlujoEfectivoDetTotal(i) = False Then
                            Call CargarMatrizDatosMantenimientoSubItems(lvDetalleFlujoEfectivo, nCodEF, lvPrincipalFlujoEfectivo(i).nConsCod, lvPrincipalFlujoEfectivo(i).nConsValor)
                            lvPrincipalFlujoEfectivo(i).vTotal = lvDetalleFlujoEfectivo
                        End If
                        ReDim lvDetalleFlujoEfectivo(0)
                    End If
                Next i
            End If
            '****->Detalle SubItem Activo
            If UBound(lvPrincipalActivos) > 0 Then
                For i = 1 To UBound(lvPrincipalActivos)
                    If IsArray(lvPrincipalActivos(i).vTotal) Then
                        If UBound(lvPrincipalActivos(i).vTotal) <= 0 And bActivoDetTotal(i) = False Then
                            Call CargarMatrizDatosMantenimientoSubItems(lvDetalleActivos, nCodEF, lvPrincipalActivos(i).nConsCod, lvPrincipalActivos(i).nConsValor)
                            lvPrincipalActivos(i).vTotal = lvDetalleActivos
                        End If
                        ReDim lvDetalleActivos(0)
                    End If
                Next i
            End If
            '****->Detalle SubItem Pasivo
            If UBound(lvPrincipalPasivos) > 0 Then
                For i = 1 To UBound(lvPrincipalPasivos)
                    If IsArray(lvPrincipalPasivos(i).vTotal) Then
                        If UBound(lvPrincipalPasivos(i).vTotal) <= 0 And bPasivoDetTotal(i) = False Then
                            Call CargarMatrizDatosMantenimientoSubItems(lvDetallePasivos, nCodEF, lvPrincipalPasivos(i).nConsCod, lvPrincipalPasivos(i).nConsValor, lvPrincipalPasivos(i).nConsValorGrupo)
                            lvPrincipalPasivos(i).vTotal = lvDetallePasivos
                        End If
                        ReDim lvDetalleActivos(0)
                    End If
                Next i
            End If
            '*********************************************************************************************
            '***** Registro de Items y SubItems de EEFF (Balance, EERR, FLujo Caja, Cambios Patrimonio)***
            '*********************************************************************************************
            GrabarDatos = oNCOMFormatosEval.GrabarPersEstadosFinancierosDetNew(nCodEF, rsEstadoResultado, rsFlujoEfectivo, rsCambiosPatrimonio)
            '***** Activos *****
            If UBound(lvPrincipalActivos) > 0 Or UBound(lvPrincipalPasivos) > 0 Then 'Eliminamos Activos Registrados
                Call oDCOMFormatosEval.EliminaPersEstadoFinancieroDetNew(nCodEF, 7041)
            End If
            If UBound(lvPrincipalActivos) > 0 Then
                For i = 1 To UBound(lvPrincipalActivos)
                    nEstFinanOrdenItem = CInt(Me.feActivos.TextMatrix(i, 0))
                    nEstFinanTipo = CInt(Me.feActivos.TextMatrix(i, 7))
                    nEstFinanSubTipoCod = CInt(Me.feActivos.TextMatrix(i, 4))
                    nEstFinanSubTipoValor = CDbl(Me.feActivos.TextMatrix(i, 5))
                    nEstFinanSubTipoValorGrupo = CDbl(Me.feActivos.TextMatrix(i, 6))
                    nMonto = CDbl(Me.feActivos.TextMatrix(i, 2))
                    nPorcentaje = CDbl(Me.feActivos.TextMatrix(i, 3))
                    Call oDCOMFormatosEval.InsertaPersEstadoFinancieroDetNew(nCodEF, nEstFinanOrdenItem, nEstFinanTipo, nEstFinanSubTipoCod, nEstFinanSubTipoValor, nEstFinanSubTipoValorGrupo, nMonto, nPorcentaje)
                
                    If IsArray(lvPrincipalActivos(i).vTotal) Then 'Para subitems de Activos
                        For j = 1 To UBound(lvPrincipalActivos(i).vTotal)
                            Call oDCOMFormatosEval.InsertaPersEstadoFinancieroDetNew(nCodEF, j, nEstFinanTipo, nEstFinanSubTipoCod, Trim(Right(lvPrincipalActivos(i).vTotal(j).cDescripcion, 4)), nEstFinanSubTipoValorGrupo, lvPrincipalActivos(i).vTotal(j).nImporte)
                        Next j
                    End If
                    nEstFinanOrdenItem = 0: nEstFinanSubTipoValor = 0: nMonto = 0: nPorcentaje = 0: nEstFinanTipo = 0: nEstFinanSubTipoCod = 0: nEstFinanSubTipoValorGrupo = 0
                Next i
            End If
            '***** Pasivos *****
            If UBound(lvPrincipalPasivos) > 0 Or UBound(lvPrincipalPasivos) > 0 Then 'Eliminamos Pasivos Registrados
                Call oDCOMFormatosEval.EliminaPersEstadoFinancieroDetNew(nCodEF, 7042)
            End If
            If UBound(lvPrincipalPasivos) > 0 Then
                For i = 1 To UBound(lvPrincipalPasivos)
                    nEstFinanOrdenItem = CInt(Me.fePasivos.TextMatrix(i, 0))
                    nEstFinanTipo = CInt(Me.fePasivos.TextMatrix(i, 7))
                    nEstFinanSubTipoCod = CInt(Me.fePasivos.TextMatrix(i, 4))
                    nEstFinanSubTipoValor = CDbl(Me.fePasivos.TextMatrix(i, 5))
                    nEstFinanSubTipoValorGrupo = CDbl(Me.fePasivos.TextMatrix(i, 6))
                    nMonto = CDbl(Me.fePasivos.TextMatrix(i, 2))
                    nPorcentaje = CDbl(Me.fePasivos.TextMatrix(i, 3))
                    
                    Call oDCOMFormatosEval.InsertaPersEstadoFinancieroDetNew(nCodEF, nEstFinanOrdenItem, nEstFinanTipo, nEstFinanSubTipoCod, nEstFinanSubTipoValor, nEstFinanSubTipoValorGrupo, nMonto, nPorcentaje)
                    If IsArray(lvPrincipalPasivos(i).vTotal) Then 'Para subitem Pasivos
                        If i = 7 Or i = 18 Then
                            For j = 1 To UBound(lvPrincipalPasivos(i).vTotal)
                                Call oDCOMFormatosEval.InsertaPersEstadoFinancieroDetNewDet(nCodEF, nEstFinanSubTipoCod, nEstFinanSubTipoValor, nEstFinanSubTipoValorGrupo, Trim(Right(lvPrincipalPasivos(i).vTotal(j).cDescripcion, 13)), lvPrincipalPasivos(i).vTotal(j).nImporte)
                            Next j
                        Else
                            For j = 1 To UBound(lvPrincipalPasivos(i).vTotal)
                                Call oDCOMFormatosEval.InsertaPersEstadoFinancieroDetNew(nCodEF, j, nEstFinanTipo, nEstFinanSubTipoCod, Trim(Right(lvPrincipalPasivos(i).vTotal(j).cDescripcion, 4)), nEstFinanSubTipoValorGrupo, lvPrincipalPasivos(i).vTotal(j).nImporte)
                            Next j
                        End If
                    End If
                    nEstFinanOrdenItem = 0: nEstFinanTipo = 0: nEstFinanSubTipoCod = 0: nEstFinanSubTipoValor = 0: nEstFinanSubTipoValorGrupo = 0: nMonto = 0: nPorcentaje = 0
                    
                Next i
            End If
            '***** Flujo de Efectivo *****
            If UBound(lvPrincipalFlujoEfectivo) > 0 Then
                For i = 1 To UBound(lvPrincipalFlujoEfectivo)
                    nEstFinanOrdenItem = CInt(Me.fefujoEfectivo.TextMatrix(i, 0))
                    nEstFinanTipo = CInt(Me.fefujoEfectivo.TextMatrix(i, 5))
                    nEstFinanSubTipoCod = CInt(Me.fefujoEfectivo.TextMatrix(i, 3))
                    nEstFinanSubTipoValor = CDbl(Me.fefujoEfectivo.TextMatrix(i, 4))
                    nEstFinanSubTipoValorGrupo = 0
                    nMonto = CDbl(Me.fefujoEfectivo.TextMatrix(i, 2))
                    nPorcentaje = "0.00"
                    
                    If IsArray(lvPrincipalFlujoEfectivo(i).vTotal) Then 'Para SubItem FlujoEfectivo
                        For j = 1 To UBound(lvPrincipalFlujoEfectivo(i).vTotal)
                            Call oDCOMFormatosEval.InsertaPersEstadoFinancieroDetNew(nCodEF, j, nEstFinanTipo, nEstFinanSubTipoCod, Trim(Right(lvPrincipalFlujoEfectivo(i).vTotal(j).cDescripcion, 3)), nEstFinanSubTipoValorGrupo, lvPrincipalFlujoEfectivo(i).vTotal(j).nImporte)
                        Next j
                    End If
                    nEstFinanOrdenItem = 0: nEstFinanSubTipoValor = 0: nMonto = 0: nPorcentaje = 0: nEstFinanTipo = 0: nEstFinanSubTipoCod = 0: nEstFinanSubTipoValorGrupo = 0
                Next i
            End If
   
            Set oDPersona = New COMDPersona.DCOMPersonas
            Set oDPersona = Nothing
            MsgBox ("Los datos se guardaron correctamente")
            
            'EAAS20170926
            lcMovNroEF = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            Call oDFormatosEval.GrabarEstadoFinancieroEdicion(gsCodPersUser, nCodEF, gdFecSis, nTipoAccion, cEstadoEF, lcMovNroEF, GetMaquinaUsuario) 'EAAS 20170926
            Set oDFormatosEval = Nothing
            'Fin EAAS20170915
            
            Call LimpiarEstFinan
            Call LimpiarEstFinanNew
            Unload Me
    Else
        MsgBox "Verificar si se registraron todos los datos"
        Exit Sub
    End If
End Sub
Private Sub cmdSalir_Click()
   Call LimpiarEstFinan
   Call LimpiarEstFinanNew
    Unload Me
End Sub
Private Sub feActivos_EnterCell()
    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0)) 'Activa TextBuscar
        Case 1, 21, 32
            Me.feActivos.ListaControles = "0-0-0-0-0-0-0"
        Case Else
            Me.feActivos.ListaControles = "0-0-1-0-0-0-0"
        End Select

    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0)) 'Celdas que pueden Editarse
        Case 1, 21, 32
            Me.feActivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.feActivos.ColumnasAEditar = "X-X-2-X-X-X-X"
        End Select
    Call CalculoCeldasBalanceGeneral(1)
    Call CalculoCeldasBalanceGeneral(2)
End Sub
Private Sub feActivos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim pnMonto As Double
    Dim Index As Integer
    Dim nTotal As Double
    Set ofrmPersEstadosFinancierosDetalleDet = New frmPersEstadosFinancierosDetalleDet
    
    If txtFechaEF.Text = "__/__/____" Then
        fEnfoque txtFechaEF
        MsgBox "Ingrese una fecha de Estados Financieros.", vbOKOnly, "Atención"
        Call CalculoCeldasBalanceGeneral(1)
        Call CalculoCeldasBalanceGeneral(2)
        Exit Sub
    End If
      
    If feActivos.TextMatrix(1, 0) = "" Then Exit Sub
    Index = CInt(feActivos.TextMatrix(feActivos.row, 0))
    Select Case CInt(feActivos.TextMatrix(Me.feActivos.row, 0))
        Case 1, 21, 32
            MsgBox "No se pueden editar los registros de totales.", vbInformation, "Aviso"
        Case Else
            If feActivos.Col = 2 Then 'Columna Total
                If IsArray(lvPrincipalActivos(Index).vTotal) Then
                    lvDetalleActivos = lvPrincipalActivos(Index).vTotal
                Else
                    ReDim lvDetalleActivos(0)
                End If
            End If
            
            If ofrmPersEstadosFinancierosDetalleDet.Inicio(lvDetalleActivos, nTotal, CInt(feActivos.TextMatrix(Me.feActivos.row, 4)), CInt(feActivos.TextMatrix(Me.feActivos.row, 5)), feActivos.TextMatrix(Me.feActivos.row, 1), nCodEF, feActivos.TextMatrix(Me.feActivos.row, 6)) Then
                If feActivos.Col = 2 Then 'column Total
                    lvPrincipalActivos(Index).vTotal = lvDetalleActivos
                    If IsArray(lvPrincipalActivos(Index).vTotal) Then
                        lvDetalleActivos = lvPrincipalActivos(Index).vTotal
                        bActivoDetTotal(Index) = True
                    Else
                        ReDim lvDetalleActivos(0)
                    End If
                End If
            End If
            If feActivos.Col = 2 Then
                Me.feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.Col) = Format(nTotal, "#,#0.00")
                  'Pinta de color celda, según monto
                If CDbl(feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.Col)) > 0 Then
                    Me.feActivos.CellForeColor = &H80000002
                    Me.feActivos.CellBackColor = &HC0FFFF
                ElseIf CDbl(feActivos.TextMatrix(Me.feActivos.row, Me.feActivos.Col)) < 0 Then
                    Me.feActivos.CellForeColor = &HFF&
                    Me.feActivos.CellBackColor = &HC0FFFF
                Else
                    Me.feActivos.CellForeColor = &H80000012
                    Me.feActivos.CellBackColor = &H80000005
                End If
            End If
    End Select
    Call CalculoCeldasBalanceGeneral(1)
    Call CalculoCeldasBalanceGeneral(2)
End Sub
Private Sub fePasivos_EnterCell()
    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0)) 'Activa TextBuscar
        Case 1, 23, 30, 31, 32
            Me.fePasivos.ListaControles = "0-0-0-0-0-0-0"
        Case Else
            Me.fePasivos.ListaControles = "0-0-1-0-0-0-0"
    End Select

    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0)) 'Celdas que pueden Editarse
        Case 1, 23, 31, 32, 24
            Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.fePasivos.ColumnasAEditar = "X-X-2-X-X-X-X"
    End Select
    Call CalculoCeldasBalanceGeneral(1)
    Call CalculoCeldasBalanceGeneral(2)
End Sub
Private Sub fePasivos_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim pnMonto As Double
    Dim Index As Integer
    Dim nTotal As Double
    
    Dim lnConsCod As Integer
    Dim lnConsValor As Integer
    Dim lsTitulo As String
    Dim lnValorGrupo As Integer
    
    Set ofrmPersEstadosFinancierosDetalleDet = New frmPersEstadosFinancierosDetalleDet
    
    If txtFechaEF.Text = "__/__/____" Then
        fEnfoque txtFechaEF
        MsgBox "Ingrese una fecha de Estados Financieros.", vbOKOnly, "Atención"
        Call CalculoCeldasBalanceGeneral(1)
        Call CalculoCeldasBalanceGeneral(2)
        Exit Sub
    End If
      
    If fePasivos.TextMatrix(1, 0) = "" Then Exit Sub
    Index = CInt(fePasivos.TextMatrix(fePasivos.row, 0))
    Select Case CInt(fePasivos.TextMatrix(Me.fePasivos.row, 0))
        Case 1, 23, 32
            MsgBox "No se pueden editar los registros de totales.", vbInformation, "Aviso"
        Case Else
            If fePasivos.Col = 2 Then 'Columna Total
                If IsArray(lvPrincipalPasivos(Index).vTotal) Then
                    lvDetallePasivos = lvPrincipalPasivos(Index).vTotal
                Else
                    ReDim lvDetallePasivos(0)
                End If
            End If
            
            lnConsCod = CInt(fePasivos.TextMatrix(Me.fePasivos.row, 4))
            lnConsValor = CInt(fePasivos.TextMatrix(Me.fePasivos.row, 5))
            lsTitulo = fePasivos.TextMatrix(Me.fePasivos.row, 1)
            lnValorGrupo = fePasivos.TextMatrix(Me.fePasivos.row, 6)
            
            If ofrmPersEstadosFinancierosDetalleDet.Inicio(lvDetallePasivos, nTotal, lnConsCod, lnConsValor, lsTitulo, nCodEF, lnValorGrupo) Then
                If fePasivos.Col = 2 Then 'column Total
                    lvPrincipalPasivos(Index).vTotal = lvDetallePasivos
                    If IsArray(lvPrincipalPasivos(Index).vTotal) Then
                        lvDetallePasivos = lvPrincipalPasivos(Index).vTotal
                        bPasivoDetTotal(Index) = True
                    Else
                        ReDim lvDetallePasivos(0)
                    End If
                    
                End If
            End If
            If fePasivos.Col = 2 Then
                Me.fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col) = Format(nTotal, "#,#0.00")
                'Pinta de color celda, según monto
                If CDbl(fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col)) > 0 Then
                    Me.fePasivos.CellForeColor = &H80000002
                    Me.fePasivos.CellBackColor = &HC0FFFF
                ElseIf CDbl(fePasivos.TextMatrix(Me.fePasivos.row, Me.fePasivos.Col)) < 0 Then
                    Me.fePasivos.CellForeColor = &HFF&
                    Me.fePasivos.CellBackColor = &HC0FFFF
                Else
                    Me.fePasivos.CellForeColor = &H80000012
                    Me.fePasivos.CellBackColor = &H80000005
                End If
            End If
    End Select
    Call CalculoCeldasBalanceGeneral(1)
    Call CalculoCeldasBalanceGeneral(2)
End Sub
Private Sub feEstadoResultado_EnterCell()
    Dim i As Integer
    Select Case CInt(feEstadoResultado.TextMatrix(Me.feEstadoResultado.row, 5)) 'Celdas que no pueden editarse
        Case 10, 20, 30, 50, 60
            Me.feEstadoResultado.ColumnasAEditar = "X-X-X-X-X-X"
            'MsgBox "No se pueden editar los registros de totales.", vbInformation, "Aviso"
        Case 60
            MsgBox "El total se mostrará en el pasivo, cuenta 51-Resultado del ejercicio", vbInformation, "Aviso"
        Case Else
            Me.feEstadoResultado.ColumnasAEditar = "X-X-2-X-X-X"
        End Select
        Call CalculoCeldaResultadoEjercicio
        Call CalculoCeldasBalanceGeneral(1)
        Call CalculoCeldasBalanceGeneral(2)
End Sub
Private Sub feEstadoResultado_Validate(Cancel As Boolean)
    Call CalculoCeldaResultadoEjercicio
    Call CalculoCeldasBalanceGeneral(1)
    Call CalculoCeldasBalanceGeneral(2)
End Sub
Private Sub feEstadoResultado_OnCellChange(pnRow As Long, pnCol As Long)
'Carga Detalle de activos - Estados de resultado
    Call CalculoCeldasEstadoResultado
    Call CalculoCeldaResultadoEjercicio
    Call CalculoCeldasBalanceGeneral(1)
    Call CalculoCeldasBalanceGeneral(2)
    
    'Sólo valores positivos
      If Me.feEstadoResultado.TextMatrix(pnRow, pnCol) < 0 Then
        Me.feEstadoResultado.TextMatrix(pnRow, pnCol) = Format(Abs(CCur(feEstadoResultado.TextMatrix(pnRow, pnCol))), "#,#0.00")   '"0.00"
      End If
    'Pinta de color celda, según monto
    If CDbl(feEstadoResultado.TextMatrix(Me.feEstadoResultado.row, Me.feEstadoResultado.Col)) > 0 Then
        Me.feEstadoResultado.CellForeColor = &H80000002
        Me.feEstadoResultado.CellBackColor = &HC0FFFF
    ElseIf CDbl(feEstadoResultado.TextMatrix(Me.feEstadoResultado.row, Me.feEstadoResultado.Col)) < 0 Then
        Me.feEstadoResultado.CellForeColor = &HFF&
        Me.feEstadoResultado.CellBackColor = &HC0FFFF
    Else
        Me.feEstadoResultado.CellForeColor = &H80000012
        Me.feEstadoResultado.CellBackColor = &H80000005
    End If
End Sub

Private Sub CalculoCeldaResultadoEjercicio()
    Dim nTotal As Double
    nTotal = 0
    If feEstadoResultado.TextMatrix(1, 0) = "" Then Exit Sub
    nTotal = feEstadoResultado.TextMatrix(17, 2)
    
    If nTotal <> 0 Then
        Me.fePasivos.TextMatrix(31, 2) = Format(nTotal, "#,#0.00")
    Else
        Me.fePasivos.TextMatrix(31, 2) = Format(0, "#,#0.00")
    End If
End Sub

Private Sub fefujoEfectivo_EnterCell()
    Select Case CInt(fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, 0)) 'Activa TextBuscar
        Case 1, 4, 5, 8, 9, 12, 13, 14, 15
            Me.fefujoEfectivo.ListaControles = "0-0-0-0-0-0-0"
        Case Else
            Me.fefujoEfectivo.ListaControles = "0-0-1-0-0-0-0"
        End Select

    Select Case CInt(fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, 0)) 'Celdas que pueden Editarse
        Case 1, 4, 5, 8, 9, 12, 13, 15
            Me.fefujoEfectivo.ColumnasAEditar = "X-X-X-X-X-X-X"
        Case Else
            Me.fefujoEfectivo.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
        Call CalculoCeldasFlujoEfectivo
End Sub

Private Sub txtPatrimonioIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPatrimonioFin.SetFocus
        If Not IsDate(txtPatrimonioIni.Text) Then
            MsgBox "Verifique Dia,Mes,Año de la fecha inicial, Fecha Incorrecta", vbInformation, "Aviso"
            txtPatrimonioIni.SetFocus
        Else
            feCambiosPatrimonio.TextMatrix(1, 1) = "Saldo al: " & txtPatrimonioIni.Text
        End If
    End If
End Sub
Private Sub txtPatrimonioFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        feCambiosPatrimonio.SetFocus
        If Not IsDate(txtPatrimonioIni.Text) Then
            MsgBox "Verifique Dia,Mes,Año de la fecha inicial, Fecha Incorrecta", vbInformation, "Aviso"
            txtPatrimonioIni.SetFocus
        ElseIf Not IsDate(txtPatrimonioFin.Text) Then
            MsgBox "Verifique Dia,Mes,Año de la fecha final, Fecha Incorrecta", vbInformation, "Aviso"
            txtPatrimonioFin.SetFocus
        Else
            feCambiosPatrimonio.TextMatrix(1, 1) = "Saldo al: " & txtPatrimonioIni.Text
            feCambiosPatrimonio.TextMatrix(7, 1) = "Saldo al: " & txtPatrimonioFin.Text
        End If
    End If
End Sub
Private Sub fefujoEfectivo_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim pnMonto As Double
    Dim Index As Integer
    Dim nTotal As Double
    Set ofrmPersEstadosFinancierosDetalleDet = New frmPersEstadosFinancierosDetalleDet
    
    If fefujoEfectivo.TextMatrix(1, 0) = "" Then Exit Sub
    Index = CInt(fefujoEfectivo.TextMatrix(fefujoEfectivo.row, 0))
    Select Case CInt(fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, 0))
        Case 1, 4, 5, 8, 9, 12, 13, 14, 15
            MsgBox "No se pueden editar los registros de totales.", vbInformation, "Aviso"
        Case Else
            If fefujoEfectivo.Col = 2 Then 'Columna Total
                If IsArray(lvPrincipalFlujoEfectivo(Index).vTotal) Then
                    lvDetalleFlujoEfectivo = lvPrincipalFlujoEfectivo(Index).vTotal
                Else
                    ReDim lvDetalleFlujoEfectivo(0)
                End If
            End If
            
            If ofrmPersEstadosFinancierosDetalleDet.Inicio(lvDetalleFlujoEfectivo, nTotal, CInt(fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, 3)), CInt(fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, 4)), fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, 1), nCodEF) Then
                If fefujoEfectivo.Col = 2 Then 'column Total
                    lvPrincipalFlujoEfectivo(Index).vTotal = lvDetalleFlujoEfectivo
                    If IsArray(lvPrincipalFlujoEfectivo(Index).vTotal) Then
                        lvDetalleFlujoEfectivo = lvPrincipalFlujoEfectivo(Index).vTotal
                        bFlujoEfectivoDetTotal(Index) = True
                    Else
                        ReDim lvDetalleFlujoEfectivo(0)
                    End If
                End If
            End If
            If fefujoEfectivo.Col = 2 Then
                Me.fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, Me.fefujoEfectivo.Col) = Format(nTotal, "#,#0.00")
                
                'Pinta de color celda, según monto
                If CDbl(fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, Me.fefujoEfectivo.Col)) > 0 Then
                    Me.fefujoEfectivo.CellForeColor = &H80000002
                    Me.fefujoEfectivo.CellBackColor = &HC0FFFF
                ElseIf CDbl(fefujoEfectivo.TextMatrix(Me.fefujoEfectivo.row, Me.fefujoEfectivo.Col)) < 0 Then
                    Me.fefujoEfectivo.CellForeColor = &HFF&
                    Me.fefujoEfectivo.CellBackColor = &HC0FFFF
                Else
                    Me.fefujoEfectivo.CellForeColor = &H80000012
                    Me.fefujoEfectivo.CellBackColor = &H80000005
                End If
            End If
    End Select
    CalculoCeldasFlujoEfectivo
End Sub
Private Sub feCambiosPatrimonio_EnterCell()
    Select Case CInt(feCambiosPatrimonio.TextMatrix(Me.feCambiosPatrimonio.row, 0)) 'Celdas que pueden Editarse
        Case 7
            Me.feCambiosPatrimonio.ColumnasAEditar = "X-X-X-X-X-X-X-X-X-X-X"
        Case Else
            Me.feCambiosPatrimonio.ColumnasAEditar = "X-X-2-3-4-5-6-X-X-X-X"
    End Select
End Sub
Private Sub feCambiosPatrimonio_OnCellChange(pnRow As Long, pnCol As Long)
    CalculoCeldasCambioPatrimonio
    
    'Pinta de color celda, según monto
    If CDbl(feCambiosPatrimonio.TextMatrix(Me.feCambiosPatrimonio.row, Me.feCambiosPatrimonio.Col)) > 0 Then
        Me.feCambiosPatrimonio.CellForeColor = &H80000002
        Me.feCambiosPatrimonio.CellBackColor = &HC0FFFF
    ElseIf CDbl(feCambiosPatrimonio.TextMatrix(Me.feCambiosPatrimonio.row, Me.feCambiosPatrimonio.Col)) < 0 Then
        Me.feCambiosPatrimonio.CellForeColor = &HFF&
        Me.feCambiosPatrimonio.CellBackColor = &HC0FFFF
    Else
        Me.feCambiosPatrimonio.CellForeColor = &H80000012
        Me.feCambiosPatrimonio.CellBackColor = &H80000005
    End If
End Sub
Private Function CargaDatos() As Boolean 'Mantenimiento EEFF
    On Error GoTo ErrorCargaDatos
      Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
      Dim i As Integer
      Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
      CargaDatos = True
      CargaDatos = oNCOMFormatosEval.CargaDatosPersEstadosFinancierosDetNew(nCodEF, rsFeActivosEEFF, rsFePasivosEEFF, rsFeEstadoResultadoEEFF, rsFeFlujoEfectivoEEFF, rsFeCambiosPatrimonioEEFF, rsDatosPersEstadosFinancieros)
      If (rsFeActivosEEFF.EOF) Or (rsFePasivosEEFF.EOF) Or (rsFeEstadoResultadoEEFF.EOF) Or (rsFeFlujoEfectivoEEFF.EOF) Or (rsFeCambiosPatrimonioEEFF.EOF) Then 'Sino tiene datos
        CargaDatos = False
      End If
      Exit Function
ErrorCargaDatos:
    CargaDatos = False
    MsgBox Err.Description + ": Error al carga datos", vbInformation, "Error"
End Function
Private Sub MostrarEstadosFinancierosNew(ByVal pnCodEstFinan As Integer)
    Dim lnFila As Integer
    Dim dPatFechaInicial  As Date
    Dim dPatFechaFinal  As Date
    '-->Carga Datos=> feActivos ***
    Call LimpiaFlex(feActivos)
        Do While Not rsFeActivosEEFF.EOF
            feActivos.AdicionaFila
            lnFila = feActivos.row
            feActivos.TextMatrix(lnFila, 1) = rsFeActivosEEFF!Concepto
            feActivos.TextMatrix(lnFila, 2) = Format(rsFeActivosEEFF!nMonto, "#,#0.00")
            feActivos.TextMatrix(lnFila, 3) = Format(rsFeActivosEEFF!nPorcentaje, "#,#0.00")
            feActivos.TextMatrix(lnFila, 4) = rsFeActivosEEFF!nConsCod
            feActivos.TextMatrix(lnFila, 5) = rsFeActivosEEFF!nConsValor
            feActivos.TextMatrix(lnFila, 6) = rsFeActivosEEFF!nConsValorGrupo
            feActivos.TextMatrix(lnFila, 7) = rsFeActivosEEFF!nEstFinanTipo
            
            Select Case CInt(feActivos.TextMatrix(feActivos.row, 5)) 'celda que o se puede editar
            Case 100, 200, 300, 1000
                Me.feActivos.BackColorRow (&H80000000)
                Me.feActivos.ForeColorRow vbBlack, True
            Case 29, 36, 39
            Me.feActivos.ForeColorRow vbRed
            Case Else
                Me.feActivos.BackColorRow (&HFFFFFF)
                Me.feActivos.ColumnasAEditar = "X-X-2-X-X-X-X"
            End Select
            rsFeActivosEEFF.MoveNext
        Loop
    rsFeActivosEEFF.Close
    Set rsFeActivosEEFF = Nothing

     '-->Carga Datos=> fePasivos ***
        Call LimpiaFlex(fePasivos)
        Do While Not rsFePasivosEEFF.EOF
            fePasivos.AdicionaFila
            lnFila = fePasivos.row
            fePasivos.TextMatrix(lnFila, 1) = rsFePasivosEEFF!Concepto
            fePasivos.TextMatrix(lnFila, 2) = Format(rsFePasivosEEFF!nMonto, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 3) = Format(rsFePasivosEEFF!nPorcentaje, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 4) = rsFePasivosEEFF!nConsCod
            fePasivos.TextMatrix(lnFila, 5) = rsFePasivosEEFF!nConsValor
            fePasivos.TextMatrix(lnFila, 6) = rsFePasivosEEFF!nConsValorGrupo
            fePasivos.TextMatrix(lnFila, 7) = rsFePasivosEEFF!nEstFinanTipo
            
        Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 5)) 'celda que  o se puede editar
            Case 100, 200, 300, 1000, 1002
                Me.fePasivos.BackColorRow (&H80000000)
                Me.fePasivos.ForeColorRow vbBlack, True
            Case 45, 59
                Me.fePasivos.ForeColorRow vbBlack, True
            Case 50, 5912
                Me.fePasivos.ForeColorRow vbGrayText, True
            Case Else
                Me.fePasivos.BackColorRow (&HFFFFFF)
                Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
            rsFePasivosEEFF.MoveNext
        Loop
    rsFePasivosEEFF.Close
    Set rsFePasivosEEFF = Nothing

    'Carga Datos=> Estado Resultado ***
    Call LimpiaFlex(feEstadoResultado)
        lnFila = 0
        Do While Not rsFeEstadoResultadoEEFF.EOF
            feEstadoResultado.AdicionaFila
            lnFila = feEstadoResultado.row
            feEstadoResultado.TextMatrix(lnFila, 1) = rsFeEstadoResultadoEEFF!Concepto
            feEstadoResultado.TextMatrix(lnFila, 2) = Format(rsFeEstadoResultadoEEFF!nMonto, "#,#0.00")
            feEstadoResultado.TextMatrix(lnFila, 3) = Format(rsFeEstadoResultadoEEFF!nPorcentaje, "#,#0.00")
            feEstadoResultado.TextMatrix(lnFila, 4) = rsFeEstadoResultadoEEFF!nConsCod
            feEstadoResultado.TextMatrix(lnFila, 5) = rsFeEstadoResultadoEEFF!nConsValor
            feEstadoResultado.TextMatrix(lnFila, 6) = rsFeEstadoResultadoEEFF!nEstFinanTipo
            
        Select Case CInt(feEstadoResultado.TextMatrix(feEstadoResultado.row, 5)) 'celda que  o se puede editar
            Case 1, 11
                Me.feEstadoResultado.ForeColorRow vbBlack, True
            Case 10, 20, 30, 50, 60
                Me.feEstadoResultado.BackColorRow (&H80000000)
                Me.feEstadoResultado.ForeColorRow vbBlack, True
                Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
            Case Else
                Me.feEstadoResultado.BackColorRow (&HFFFFFF)
                Me.fePasivos.ColumnasAEditar = "X-X-2-X-X-X-X"
        End Select
        rsFeEstadoResultadoEEFF.MoveNext
        Loop
    rsFeEstadoResultadoEEFF.Close
    Set rsFeEstadoResultadoEEFF = Nothing

    'Carga Datos =>Flujo Efectivo
    Call LimpiaFlex(fefujoEfectivo)
    lnFila = 0
    Do While Not rsFeFlujoEfectivoEEFF.EOF
        fefujoEfectivo.AdicionaFila
        lnFila = fefujoEfectivo.row
        

        fefujoEfectivo.TextMatrix(lnFila, 1) = rsFeFlujoEfectivoEEFF!Concepto
        fefujoEfectivo.TextMatrix(lnFila, 2) = Format(rsFeFlujoEfectivoEEFF!nMonto, "#,#0.00")
        fefujoEfectivo.TextMatrix(lnFila, 3) = rsFeFlujoEfectivoEEFF!nConsCod
        fefujoEfectivo.TextMatrix(lnFila, 4) = rsFeFlujoEfectivoEEFF!nConsValor
        fefujoEfectivo.TextMatrix(lnFila, 5) = rsFeFlujoEfectivoEEFF!nEstFinanTipo
        
    Select Case CInt(fefujoEfectivo.TextMatrix(fefujoEfectivo.row, 4)) 'celda que se puede editar
        Case 1, 2, 3, 4, 5, 6
            Me.fefujoEfectivo.BackColorRow (&H80000000)
            Me.fefujoEfectivo.ForeColorRow vbBlack, True
        Case 130, 230, 330
            Me.fefujoEfectivo.ForeColorRow vbBlack, True
        Case Else
            Me.fefujoEfectivo.BackColorRow (&HFFFFFF)
            Me.fefujoEfectivo.ColumnasAEditar = "X-X-2-3-X-X-X"
    End Select
    
    rsFeFlujoEfectivoEEFF.MoveNext
    Loop
    rsFeFlujoEfectivoEEFF.Close
    Set rsFeFlujoEfectivoEEFF = Nothing
    
    'Carga Datos=> Cambios en el Patrimonio
    Call LimpiaFlex(feCambiosPatrimonio)
        lnFila = 0
        Do While Not rsFeCambiosPatrimonioEEFF.EOF
            feCambiosPatrimonio.AdicionaFila
            lnFila = feCambiosPatrimonio.row
            
            If feCambiosPatrimonio.TextMatrix(lnFila, 0) = 1 Then  'Concepto con->Fecha Inicio
                feCambiosPatrimonio.TextMatrix(lnFila, 1) = rsFeCambiosPatrimonioEEFF!Concepto & rsDatosPersEstadosFinancieros!dFechaPatrimonioIni
            ElseIf feCambiosPatrimonio.TextMatrix(lnFila, 0) = 7 Then 'Concepto con->Fecha Final
                feCambiosPatrimonio.TextMatrix(lnFila, 1) = rsFeCambiosPatrimonioEEFF!Concepto & rsDatosPersEstadosFinancieros!dFechaPatrimonioFin
            Else 'Concepto de constante
                feCambiosPatrimonio.TextMatrix(lnFila, 1) = rsFeCambiosPatrimonioEEFF!Concepto
            End If
            
            feCambiosPatrimonio.TextMatrix(lnFila, 2) = Format(rsFeCambiosPatrimonioEEFF!nCapSocial, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 3) = Format(rsFeCambiosPatrimonioEEFF!nCapSocialAdicional, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 4) = Format(rsFeCambiosPatrimonioEEFF!nReservaLegal, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 5) = Format(rsFeCambiosPatrimonioEEFF!nResultadoAcumulado, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 6) = Format(rsFeCambiosPatrimonioEEFF!nUtilidadEjercicio, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 7) = Format(rsFeCambiosPatrimonioEEFF!nTotalPatrimonio, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 8) = rsFeCambiosPatrimonioEEFF!nConsCod
            feCambiosPatrimonio.TextMatrix(lnFila, 9) = rsFeCambiosPatrimonioEEFF!nConsValor
            feCambiosPatrimonio.TextMatrix(lnFila, 10) = rsFeCambiosPatrimonioEEFF!nEstFinanTipo
            
        Select Case CInt(feCambiosPatrimonio.TextMatrix(feCambiosPatrimonio.row, 9)) 'celda que  o se puede editar
            Case 1, 7
                Me.feCambiosPatrimonio.BackColorRow (&H80000000)
                Me.feCambiosPatrimonio.ForeColorRow vbBlack, True
        End Select
        rsFeCambiosPatrimonioEEFF.MoveNext
        Loop
    rsFeCambiosPatrimonioEEFF.Close
    Set rsFeCambiosPatrimonioEEFF = Nothing
        'Carga Fecha registradas
        dPatFechaInicial = IIf(rsDatosPersEstadosFinancieros!dFechaPatrimonioIni = "", "__/__/____", rsDatosPersEstadosFinancieros!dFechaPatrimonioIni)
        dPatFechaFinal = IIf(rsDatosPersEstadosFinancieros!dFechaPatrimonioFin = "", "__/__/____", rsDatosPersEstadosFinancieros!dFechaPatrimonioFin)
        txtPatrimonioIni.Text = Format(dPatFechaInicial, "dd/mm/yyyy")
        txtPatrimonioFin.Text = Format(dPatFechaFinal, "dd/mm/yyyy")
        rsDatosPersEstadosFinancieros.Close
        Set rsDatosPersEstadosFinancieros = Nothing
    
End Sub
Public Function ValidaDatos() As Boolean
    Dim nIndice As Integer
    Dim i As Integer
    ValidaDatos = False
    
    If Trim(txtFechaEF.Text) = "__/__/____" Or Not IsDate(Trim(txtFechaEF.Text)) Then
            MsgBox "Falta ingresar la fecha del estado financiero." & Chr(10) & " Formato: DD/MM/YYYY", vbInformation, "Aviso"
            SSTabDatosEEFF.Tab = 0
            ValidaDatos = False
            Exit Function
    End If
    '********** Para TAB:0 -> Balance General
    If CCur(Me.feActivos.TextMatrix(32, 2)) <> CCur(Me.fePasivos.TextMatrix(33, 2)) Then
        MsgBox "El monto de [Total Activo] no coincide con el monto de [Total Pasivo y Patrimonio], favor verificarlo.", vbInformation, "Aviso"
            SSTabDatosEEFF.Tab = 0
            Exit Function
    End If

    '********** Para TAB:1 -> Estado Resultados
    If CCur(Me.feEstadoResultado.TextMatrix(15, 2)) = 0 Or CCur(Me.feEstadoResultado.TextMatrix(16, 2)) = 0 Then
        MsgBox "Se recomienda actualizar la pestaña de [Estados de Resultados]. Impuesto a la renta. " & Chr(10) & "Utilidad (o pérdida) neta antes del IR", vbInformation, "Aviso"
            SSTabDatosEEFF.Tab = 1
            Exit Function
    End If
    '********** Para TAB:2 -> Flujo Efectivo
    If CCur(Me.fefujoEfectivo.TextMatrix(14, 2)) <> CCur(Me.feActivos.TextMatrix(2, 2)) Then
        MsgBox "El monto ingresado en la fila [E], no coincide con el monto de [10-Efectivo y equivalentes de efectivo], favor verificarlo.", vbInformation, "Aviso"
             SSTabDatosEEFF.Tab = 2
            'Exit Function
    End If
    '********** Para TAB:3 -> Cambios en el Patrimonio
    If Right(Format(txtFechaEF.Text, "yyyymmdd"), 4) = "1231" Then
        If txtPatrimonioIni.Text = "__/__/____" Or Not IsDate(Trim(txtPatrimonioIni.Text)) Then
                MsgBox "Falta ingresar del periodo inicial de cambios en el patrimonio." & Chr(10) & " Formato: DD/MM/YYYYY", vbInformation, "Aviso"
                SSTabDatosEEFF.Tab = 3
                ValidaDatos = False
                Exit Function
        End If
        If txtPatrimonioFin.Text = "__/__/____" Or Not IsDate(Trim(txtPatrimonioFin.Text)) Then
                MsgBox "Falta ingresar del periodo final de cambios en el patrimonio." & Chr(10) & " Formato: DD/MM/YYYY", vbInformation, "Aviso"
                SSTabDatosEEFF.Tab = 3
                ValidaDatos = False
                Exit Function
        End If
    End If
    'Valida fecha de cambios en el patrimonio
    If txtPatrimonioIni.Text = "__/__/____" Then
        txtPatrimonioIni.Text = "01/01/1900"
    End If
    If txtPatrimonioFin.Text = "__/__/____" Then
        txtPatrimonioFin.Text = "01/01/1900"
    End If
    ValidaDatos = True
End Function
Private Sub CalculoCeldasBalanceGeneral(pnActPas As Integer)
    Dim i As Integer
    Dim nTotalActPas, nTotalActPasCte, nTotalActPasNoCte, nTotalActPatrimonio, nTotalPasPatrimonio As Currency
    Dim nPorcActPas, nPorcActPasCte, nPorcActPasNoCte, nPorcActPatrimonio As Currency
    Dim nPorcTotalActPas, nPorcTotalActPasCte, nPorcTotalActPasNoCte, nPorcTotalActPatrimonio As Currency
    Dim nPatCapital, nPatAccionesInversion, nPatCapitalAdicional, nPatResultadoNoRealizado, nPatExcedenteRevaluacion, nPatReservas, nPatResultadoAcumulado, nPatUtilidadEjercicio As Currency
    Dim nActDesvalorizacionExistencia, nDesvalorizacionInmovilizado, nDepreciacionAmortizacion As Currency
    '*******************************************************
    '*******************-> Activos <-***********************
    '*******************************************************
    If pnActPas = 1 Then
        'Para Totales
        nTotalActPas = 0
        nTotalActPasCte = 0
        nTotalActPasNoCte = 0
        nTotalActPatrimonio = 0
        nActDesvalorizacionExistencia = 0: nDesvalorizacionInmovilizado = 0: nDepreciacionAmortizacion = 0
        'Para Porcentaje
        nPorcActPas = 0: nPorcTotalActPas = 0
        nPorcActPasCte = 0: nPorcTotalActPasCte = 0
        nPorcActPasNoCte = 0: nPorcTotalActPasNoCte = 0
        nPorcActPatrimonio = 0: nPorcTotalActPatrimonio = 0
        nTotalPasPatrimonio = 0
    '******************** Columna Totales ************************
        'Valida que todos los registros sean números
        For i = 1 To feActivos.rows - 1
            If Not IsNumeric(Me.feActivos.TextMatrix(i, 2)) Then Me.feActivos.TextMatrix(i, 2) = "0.00"
            If Not IsNumeric(Me.feActivos.TextMatrix(i, 3)) Then Me.feActivos.TextMatrix(i, 3) = "0.00"
        Next i
        'Suma:Activo Corriente
        For i = 2 To 19
            nTotalActPasCte = nTotalActPasCte + CCur(Me.feActivos.TextMatrix(i, 2))
        Next i
        nActDesvalorizacionExistencia = CCur(Me.feActivos.TextMatrix(20, 2))
        Me.feActivos.TextMatrix(1, 2) = Format(nTotalActPasCte - nActDesvalorizacionExistencia, "#,#0.00")
        'Suma: Activo No Corriente
        For i = 22 To 31
            nTotalActPasNoCte = nTotalActPasNoCte + CCur(Me.feActivos.TextMatrix(i, 2))
        Next i
        nDesvalorizacionInmovilizado = CCur(Me.feActivos.TextMatrix(28, 2))
        nDepreciacionAmortizacion = CCur(Me.feActivos.TextMatrix(31, 2))
        
        nTotalActPasNoCte = nTotalActPasNoCte - nDesvalorizacionInmovilizado - nDepreciacionAmortizacion
        Me.feActivos.TextMatrix(21, 2) = Format(nTotalActPasNoCte - nDesvalorizacionInmovilizado - nDepreciacionAmortizacion, "#,#0.00")
        'Total Activo
        Me.feActivos.TextMatrix(32, 2) = Format((nTotalActPasCte - nActDesvalorizacionExistencia) + (nTotalActPasNoCte - nDesvalorizacionInmovilizado - nDepreciacionAmortizacion), "#,#0.00")
        nTotalActPas = Me.feActivos.TextMatrix(32, 2)
    '******************** Columna Porcentaje ********************
        'Calcula Activo Corriente
        For i = 2 To 20
            nPorcActPas = 0
            nPorcActPas = Round((CCur(Me.feActivos.TextMatrix(i, 2)) / IIf(nTotalActPas = 0, 1, nTotalActPas)), 4) * 100
            Me.feActivos.TextMatrix(i, 3) = Format(nPorcActPas, "#,#0.00")
            nPorcTotalActPasCte = nPorcTotalActPasCte + nPorcActPas
        Next i
        Me.feActivos.TextMatrix(1, 3) = Format(nPorcTotalActPasCte, "#,#0.00")
        'Calcula Activo No Corriente
        For i = 22 To 31
            nPorcActPasNoCte = 0
            nPorcActPasNoCte = Round((CCur(Me.feActivos.TextMatrix(i, 2)) / IIf(nTotalActPas = 0, 1, nTotalActPas)), 4) * 100
            Me.feActivos.TextMatrix(i, 3) = Format(nPorcActPasNoCte, "#,#0.00")
            nPorcTotalActPasNoCte = nPorcTotalActPasNoCte + nPorcActPasNoCte
        Next i
        Me.feActivos.TextMatrix(21, 3) = Format(nPorcTotalActPasNoCte, "#,#0.00")
        'Total Activo Porcentaje
        Me.feActivos.TextMatrix(32, 3) = Format(nPorcTotalActPasCte + nPorcTotalActPasNoCte, "#,#0.00")
    End If
    '*******************************************************
    '*******************-> Pasivos <-***********************
    '*******************************************************
    If pnActPas = 2 Then
      'Para Totales
        nTotalActPas = 0
        nTotalActPasCte = 0
        nTotalActPasNoCte = 0
        nTotalActPatrimonio = 0
       'Para Montos de Patrimonio
        nPatCapital = 0
        nPatAccionesInversion = 0
        nPatCapitalAdicional = 0
        nPatResultadoNoRealizado = 0
        nPatExcedenteRevaluacion = 0
        nPatReservas = 0
        nPatResultadoAcumulado = 0
        nPatUtilidadEjercicio = 0
        'Para Porcentaje
        nPorcActPas = 0: nPorcTotalActPas = 0
        nPorcActPasCte = 0: nPorcTotalActPasCte = 0
        nPorcActPasNoCte = 0: nPorcTotalActPasNoCte = 0
        nPorcActPatrimonio = 0: nPorcTotalActPatrimonio = 0
        nTotalPasPatrimonio = 0
    '******************** Columna Totales ************************
        'Valida que todos los registros sean números
        For i = 1 To fePasivos.rows - 1
            If Not IsNumeric(Me.fePasivos.TextMatrix(i, 2)) Then Me.fePasivos.TextMatrix(i, 2) = "0.00"
            If Not IsNumeric(Me.fePasivos.TextMatrix(i, 3)) Then Me.fePasivos.TextMatrix(i, 3) = "0.00"
        Next i
        'Suma:Pasivo Corriente
        For i = 2 To 11
            nTotalActPasCte = nTotalActPasCte + CCur(Me.fePasivos.TextMatrix(i, 2))
        Next i
        Me.fePasivos.TextMatrix(1, 2) = Format(nTotalActPasCte, "#,#0.00")
        'Suma: Pasivo No Corriente
        For i = 13 To 22
            nTotalActPasNoCte = nTotalActPasNoCte + CCur(Me.fePasivos.TextMatrix(i, 2))
        Next i
        Me.fePasivos.TextMatrix(12, 2) = Format(nTotalActPasNoCte, "#,#0.00")
        
        'Resultado del ejercicio [Estado de Resultado]
        Call CalculoCeldaResultadoEjercicio
        
        'Suma: Patrimonio
        nPatAccionesInversion = CCur(Me.fePasivos.TextMatrix(25, 2))
        nPatCapitalAdicional = CCur(Me.fePasivos.TextMatrix(26, 2))
        nPatResultadoNoRealizado = CCur(Me.fePasivos.TextMatrix(27, 2))
        nPatExcedenteRevaluacion = CCur(Me.fePasivos.TextMatrix(28, 2))
        nPatReservas = CCur(Me.fePasivos.TextMatrix(29, 2))
        nPatResultadoAcumulado = CCur(Me.fePasivos.TextMatrix(30, 2))
        nPatUtilidadEjercicio = CCur(Me.fePasivos.TextMatrix(31, 2))
        
        nPatCapital = CCur(Me.feActivos.TextMatrix(32, 2)) - (nTotalActPasCte + nTotalActPasNoCte) - (nPatAccionesInversion + nPatCapitalAdicional + nPatResultadoNoRealizado + nPatExcedenteRevaluacion + nPatReservas + nPatResultadoAcumulado + nPatUtilidadEjercicio)
        Me.fePasivos.TextMatrix(24, 2) = Format(nPatCapital, "#,#0.00")
        
        'Me.feCambiosPatrimonio.TextMatrix(1, 2) = Format(nPatCapital, "#,#0.00") '[CambiosPatrominio]
        'Me.feCambiosPatrimonio.TextMatrix(3, 5) = Format(fePasivos.TextMatrix(30, 2), "#,#0.00") '[CambiosPatrominio- ResultadoAcumulado]
        Me.feCambiosPatrimonio.TextMatrix(4, 6) = Format(fePasivos.TextMatrix(31, 2), "#,#0.00") '[CambiosPatrominio- ResultadoEjercicio]
        For i = 24 To 31
            nTotalActPatrimonio = nTotalActPatrimonio + CCur(Me.fePasivos.TextMatrix(i, 2))
        Next i
        Me.fePasivos.TextMatrix(23, 2) = Format(nTotalActPatrimonio, "#,#0.00")
 
        'Total Pasivo
        Me.fePasivos.TextMatrix(32, 2) = Format(nTotalActPasCte + nTotalActPasNoCte, "#,#0.00")
        'Total Pasivo y Patrimonio
        Me.fePasivos.TextMatrix(33, 2) = Format(nTotalActPasCte + nTotalActPasNoCte + nTotalActPatrimonio, "#,#0.00")
        
        nTotalPasPatrimonio = Me.fePasivos.TextMatrix(33, 2)
        nTotalActPas = Me.fePasivos.TextMatrix(32, 2)
    '******************** Columna Porcentaje ********************
        'Calcula Activo Corriente
        For i = 2 To 11
            nPorcActPas = 0
            nPorcActPas = Round((CCur(Me.fePasivos.TextMatrix(i, 2)) / IIf(nTotalPasPatrimonio = 0, 1, nTotalPasPatrimonio)), 4) * 100
            Me.fePasivos.TextMatrix(i, 3) = Format(nPorcActPas, "#,#0.00")
            nPorcTotalActPasCte = nPorcTotalActPasCte + nPorcActPas
        Next i
        Me.fePasivos.TextMatrix(1, 3) = Format(nPorcTotalActPasCte, "#,#0.00")
        'Calcula Activo No Corriente
        For i = 13 To 22
            nPorcActPasNoCte = 0
            nPorcActPasNoCte = Round((CCur(Me.fePasivos.TextMatrix(i, 2)) / IIf(nTotalPasPatrimonio = 0, 1, nTotalPasPatrimonio)), 4) * 100
            Me.fePasivos.TextMatrix(i, 3) = Format(nPorcActPasNoCte, "#,#0.00")
            nPorcTotalActPasNoCte = nPorcTotalActPasNoCte + nPorcActPasNoCte
        Next i
        Me.fePasivos.TextMatrix(12, 3) = Format(nPorcTotalActPasNoCte, "#,#0.00")
        'Calcula Patrimonio
        For i = 24 To 31
            nPorcActPatrimonio = 0
            nPorcActPatrimonio = Round((CCur(Me.fePasivos.TextMatrix(i, 2)) / IIf(nTotalPasPatrimonio = 0, 1, nTotalPasPatrimonio)), 4) * 100
            Me.fePasivos.TextMatrix(i, 3) = Format(nPorcActPatrimonio, "#,#0.00")
            nPorcTotalActPatrimonio = nPorcTotalActPatrimonio + nPorcActPatrimonio
        Next i
            Me.fePasivos.TextMatrix(23, 3) = Format(Round(nTotalActPatrimonio / IIf(nTotalPasPatrimonio = 0, 1, nTotalPasPatrimonio), 4) * 100, "#,#0.00")
        'Total Pasivo Porcentaje
        Me.fePasivos.TextMatrix(32, 3) = Format(nPorcTotalActPasCte + nPorcTotalActPasNoCte, "#,#0.00")
        'Total Pasivo y Patrimonio
        Me.fePasivos.TextMatrix(33, 3) = Format(nPorcTotalActPasCte + nPorcTotalActPasNoCte + nPorcTotalActPatrimonio, "#,#0.00")
    End If
End Sub
Private Sub CalculoCeldasEstadoResultado()
    Dim i As Integer
    Dim nTotalIngresos, nTotalUtilidaBruta, nTotalUtilidadOperativa, nTotalUtilidadNetaIR As Currency, nTotalUtilidadNeta As Currency
    Dim nPorcTotalIngresos As Currency
    Dim nPorcCostoVentas As Currency
    Dim nPorcUtilidadBruta As Currency
    Dim nPorcGastosAdministrativos As Currency
    Dim nPorcGastosVentas As Currency
    Dim nPorcUtilidadOperativa  As Currency
    Dim nPorcDepreciacion As Currency
    Dim nPorcIngresoFinanciero As Currency
    Dim nPorcGastosFinancieros  As Currency
    Dim nPorcOtrosIngresos As Currency
    Dim nPorcOtrosEgresos As Currency
    Dim nPorcUtilidadNetaIR  As Currency
    Dim nPorcImpuestoRenta As Currency
    Dim nPorcUtilidadNeta As Currency
    
    Dim nDepreciacion As Currency
    Dim nIngresoFinanciero As Currency
    Dim nGastosFinanciero As Currency
    Dim nOtrosIngresos As Currency
    Dim nOtrosEgresos As Currency
    
    '*******************************************************
    'Para Totales
    nTotalIngresos = 0
    nTotalUtilidaBruta = 0
    nTotalUtilidadOperativa = 0
    nTotalUtilidadNetaIR = 0
    nTotalUtilidadNeta = 0
    
    nDepreciacion = 0
    nIngresoFinanciero = 0
    nGastosFinanciero = 0
    nOtrosIngresos = 0
    nOtrosEgresos = 0
    'Para Porcentaje
    nPorcTotalIngresos = 0
    nPorcCostoVentas = 0
    nPorcUtilidadBruta = 0
    nPorcGastosVentas = 0
    nPorcGastosAdministrativos = 0
    nPorcUtilidadOperativa = 0
    nPorcDepreciacion = 0
    nPorcIngresoFinanciero = 0
    nPorcGastosFinancieros = 0
    nPorcOtrosIngresos = 0
    nPorcOtrosEgresos = 0
    nPorcUtilidadNetaIR = 0
    nPorcImpuestoRenta = 0
    nPorcUtilidadNeta = 0
    '******************** Columna Totales ************************
    'Valida que todos los registros sean números
    For i = 1 To feEstadoResultado.rows - 1
        If Not IsNumeric(Me.feEstadoResultado.TextMatrix(i, 2)) Then Me.feEstadoResultado.TextMatrix(i, 2) = "0.00"
        If Not IsNumeric(Me.feEstadoResultado.TextMatrix(i, 3)) Then Me.feEstadoResultado.TextMatrix(i, 3) = "0.00"
    Next i
    'Suma: Total Ingresos
    For i = 1 To 2
        nTotalIngresos = nTotalIngresos + CCur(Me.feEstadoResultado.TextMatrix(i, 2))
    Next i
    nTotalIngresos = nTotalIngresos - Me.feEstadoResultado.TextMatrix(3, 2)
    Me.feEstadoResultado.TextMatrix(4, 2) = Format(nTotalIngresos, "#,#0.00")
    If nTotalIngresos = 0 Then
        MsgBox "Se recomienda que el monto de: (+)TOTAL INGRESO, no sea cero", vbInformation, "Aviso"
        Exit Sub
    End If
    'Suma: Utilidad Bruta
    nTotalUtilidaBruta = nTotalIngresos - CCur(Me.feEstadoResultado.TextMatrix(5, 2))
    Me.feEstadoResultado.TextMatrix(6, 2) = Format(nTotalUtilidaBruta, "#,#0.00")
    'Suma:Utilidad Operativa
    nTotalUtilidadOperativa = nTotalUtilidaBruta - CCur(Me.feEstadoResultado.TextMatrix(7, 2)) - CCur(Me.feEstadoResultado.TextMatrix(8, 2))
    Me.feEstadoResultado.TextMatrix(9, 2) = Format(nTotalUtilidadOperativa, "#,#0.00")
    'Suma:Utilidad (o pérdida) neta antes del IR
    nDepreciacion = CCur(Me.feEstadoResultado.TextMatrix(10, 2))
    nIngresoFinanciero = CCur(Me.feEstadoResultado.TextMatrix(11, 2))
    nGastosFinanciero = CCur(Me.feEstadoResultado.TextMatrix(12, 2))
    nOtrosIngresos = CCur(Me.feEstadoResultado.TextMatrix(13, 2))
    nOtrosEgresos = CCur(Me.feEstadoResultado.TextMatrix(14, 2))
    
    nTotalUtilidadNetaIR = CCur(nTotalUtilidadOperativa - nDepreciacion + nIngresoFinanciero - nGastosFinanciero + nOtrosIngresos - nOtrosEgresos)
     Me.feEstadoResultado.TextMatrix(15, 2) = Format(nTotalUtilidadNetaIR, "#,#0.00")
    'Suma:Utilidad (o pérdida) neta
    nTotalUtilidadNeta = CCur(Me.feEstadoResultado.TextMatrix(15, 2)) - CCur(Me.feEstadoResultado.TextMatrix(16, 2))
    Me.feEstadoResultado.TextMatrix(17, 2) = Format(nTotalUtilidadNeta, "#,#0.00")
    '******************** Columna Porcentaje ********************
    '4.-Porcentaje: Total Ingreso
        nPorcTotalIngresos = Round((nTotalIngresos / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(4, 3) = Format(nPorcTotalIngresos, "#0.00")
    '5.-Porcentaje: Costos de Ventas
        nPorcCostoVentas = Round((CCur(Me.feEstadoResultado.TextMatrix(5, 2)) / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(5, 3) = Format(nPorcCostoVentas, "#0.00")
    '6.-Porcentaje: Utilidad Bruta
        nPorcUtilidadBruta = Round((nTotalUtilidaBruta / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(6, 3) = Format(nPorcUtilidadBruta, "#0.00")
    '7.-Porcentaje: Gastos de Ventas
        nPorcGastosVentas = Round((CCur(Me.feEstadoResultado.TextMatrix(7, 2)) / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(7, 3) = Format(nPorcGastosVentas, "#0.00")
    '8.-Porcentaje: Gastos Administrativos
        nPorcGastosAdministrativos = Round((CCur(Me.feEstadoResultado.TextMatrix(8, 2)) / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(8, 3) = Format(nPorcGastosAdministrativos, "#0.00")
    '9.-Porcentaje: Utilidad Operativo
        nPorcUtilidadOperativa = Round((nTotalUtilidadOperativa / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(9, 3) = Format(nPorcUtilidadOperativa, "#0.00")
    '10.-Porcentaje: Depreciacion
        nPorcDepreciacion = Round((CCur(Me.feEstadoResultado.TextMatrix(10, 2)) / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(10, 3) = Format(nPorcDepreciacion, "#0.00")
    '11.-Porcentaje: Ingresos Financieros
        nPorcIngresoFinanciero = Round((CCur(Me.feEstadoResultado.TextMatrix(11, 2)) / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(11, 3) = Format(nPorcIngresoFinanciero, "#0.00")
    '12.-Porcentaje: Gastos Financieros
        nPorcGastosFinancieros = Round((CCur(Me.feEstadoResultado.TextMatrix(12, 2)) / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(12, 3) = Format(nPorcGastosFinancieros, "#0.00")
    '13.-Porcentaje: Otros Ingresos
        nPorcOtrosIngresos = Round((CCur(Me.feEstadoResultado.TextMatrix(13, 2)) / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(13, 3) = Format(nPorcOtrosIngresos, "#0.00")
    '14.-Porcentaje: Otros Egresos
        nPorcOtrosEgresos = Round((CCur(Me.feEstadoResultado.TextMatrix(14, 2)) / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(14, 3) = Format(nPorcOtrosEgresos, "#0.00")
    '15.-Porcentaje: Utilidad(o perdida) netas antes del IR
        nPorcUtilidadNetaIR = Round((nTotalUtilidadNetaIR / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(15, 3) = Format(nPorcUtilidadNetaIR, "#0.00")
    '16.-Porcentaje: Impuesto a la renta
        nPorcUtilidadNetaIR = Round((CCur(Me.feEstadoResultado.TextMatrix(16, 2)) / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(16, 3) = Format(nPorcUtilidadNetaIR, "#0.00")
    '17.-Porcentaje: Utilidad Neta
        nPorcUtilidadNeta = Round((nTotalUtilidadNeta / nTotalIngresos) * 100, 2)
        Me.feEstadoResultado.TextMatrix(17, 3) = Format(nPorcUtilidadNeta, "#0.00")
End Sub
Private Sub CalculoCeldasFlujoEfectivo()
    Dim i As Integer
    Dim nTotalEfectivoA, nTotalEfectivoB, nTotalEfectivoC, nTotalEfectivoD, nTotalEfectivoE, nTotalEfectivoF As Currency
    '*******************************************************
    'Para Totales
    nTotalEfectivoA = 0
    nTotalEfectivoB = 0
    nTotalEfectivoC = 0
    nTotalEfectivoD = 0
    nTotalEfectivoE = 0
    nTotalEfectivoF = 0
    '******************** Columna Totales ************************
    'Valida que todos los registros sean números
    For i = 1 To 15
        If Not IsNumeric(Me.fefujoEfectivo.TextMatrix(i, 2)) Then Me.fefujoEfectivo.TextMatrix(i, 2) = "0.00"
        If Not IsNumeric(Me.fefujoEfectivo.TextMatrix(i, 3)) Then Me.fefujoEfectivo.TextMatrix(i, 3) = "0.00"
    Next i
    'Suma: Actividad de la Operación (A)
        Me.fefujoEfectivo.TextMatrix(1, 2) = Format(CCur(Me.fefujoEfectivo.TextMatrix(2, 2)), "#,#0.00")
        nTotalEfectivoA = CCur(Me.fefujoEfectivo.TextMatrix(2, 2)) - CCur(Me.fefujoEfectivo.TextMatrix(3, 2))
        Me.fefujoEfectivo.TextMatrix(4, 2) = Format(nTotalEfectivoA, "#,#0.00")
    'Suma: Actividad de Inversión (B)
        Me.fefujoEfectivo.TextMatrix(5, 2) = Format(CCur(Me.fefujoEfectivo.TextMatrix(6, 2)), "#,#0.00")
        nTotalEfectivoB = CCur(Me.fefujoEfectivo.TextMatrix(6, 2)) - CCur(Me.fefujoEfectivo.TextMatrix(7, 2))
        Me.fefujoEfectivo.TextMatrix(8, 2) = Format(nTotalEfectivoB, "#,#0.00")
    'Suma:Actividad de Financiamiento(C)
        Me.fefujoEfectivo.TextMatrix(9, 2) = Format(CCur(Me.fefujoEfectivo.TextMatrix(10, 2)), "#,#0.00")
        nTotalEfectivoC = CCur(Me.fefujoEfectivo.TextMatrix(10, 2)) - CCur(Me.fefujoEfectivo.TextMatrix(11, 2))
        Me.fefujoEfectivo.TextMatrix(12, 2) = Format(nTotalEfectivoC, "#,#0.00")
    'Suma:Actividad de Financiamiento(D)
        nTotalEfectivoD = nTotalEfectivoA + nTotalEfectivoB + nTotalEfectivoC
        Me.fefujoEfectivo.TextMatrix(13, 2) = Format(nTotalEfectivoD, "#,#0.00")
    'Suma:Actividad de Financiamiento(E)
        'Me.fefujoEfectivo.TextMatrix(14, 2) = Format(100, "#,#0.00")
    'Suma:Actividad de Financiamiento(F)
        nTotalEfectivoF = nTotalEfectivoD + Me.fefujoEfectivo.TextMatrix(14, 2)
        Me.fefujoEfectivo.TextMatrix(15, 2) = Format(nTotalEfectivoF, "#,#0.00")
End Sub
Private Sub CalculoCeldasCambioPatrimonio()
    Dim i As Integer
    Dim nTotalCapital As Currency
    Dim nTotalCapitalSocial As Currency
    Dim nTotalReserva As Currency
    Dim nTotalResultados As Currency
    Dim nTotalUtilidadEjercicio As Currency
    Dim nTotalPatrimonio1 As Currency
    Dim nTotalPatrimonio2 As Currency
    Dim nTotalPatrimonio3 As Currency
    Dim nTotalPatrimonio4 As Currency
    Dim nTotalPatrimonio5 As Currency
    Dim nTotalPatrimonio6 As Currency
    Dim nTotalPatrimonio7 As Currency

    nTotalCapital = 0
    nTotalCapitalSocial = 0
    nTotalReserva = 0
    nTotalResultados = 0
    nTotalUtilidadEjercicio = 0
    nTotalPatrimonio1 = 0: nTotalPatrimonio2 = 0: nTotalPatrimonio3 = 0: nTotalPatrimonio4 = 0:  nTotalPatrimonio5 = 0: nTotalPatrimonio6 = 0: nTotalPatrimonio7 = 0
    Call CalculoCeldasBalanceGeneral(2)
'******************** Columna Totales ************************
    'Valida que todos los registros sean números
        For i = 1 To 7
            If Not IsNumeric(Me.feCambiosPatrimonio.TextMatrix(i, 2)) Then Me.feCambiosPatrimonio.TextMatrix(i, 2) = "0.00"
            If Not IsNumeric(Me.feCambiosPatrimonio.TextMatrix(i, 3)) Then Me.feCambiosPatrimonio.TextMatrix(i, 3) = "0.00"
            If Not IsNumeric(Me.feCambiosPatrimonio.TextMatrix(i, 4)) Then Me.feCambiosPatrimonio.TextMatrix(i, 4) = "0.00"
            If Not IsNumeric(Me.feCambiosPatrimonio.TextMatrix(i, 5)) Then Me.feCambiosPatrimonio.TextMatrix(i, 5) = "0.00"
        Next i
    'Suma: Capital Social
        For i = 1 To 6
            nTotalCapital = nTotalCapital + CCur(Me.feCambiosPatrimonio.TextMatrix(i, 2))
            Me.feCambiosPatrimonio.TextMatrix(7, 2) = Format(nTotalCapital, "#,#0.00")
        Next i
    'Suma: Capital Social Adicional
          For i = 1 To 6
            nTotalCapitalSocial = nTotalCapitalSocial + CCur(Me.feCambiosPatrimonio.TextMatrix(i, 3))
            Me.feCambiosPatrimonio.TextMatrix(7, 3) = Format(nTotalCapitalSocial, "#,#0.00")
        Next i
    'Suma: Reserva Legal
        For i = 1 To 6
            nTotalReserva = nTotalReserva + CCur(Me.feCambiosPatrimonio.TextMatrix(i, 4))
            Me.feCambiosPatrimonio.TextMatrix(7, 4) = Format(nTotalReserva, "#,#0.00")
        Next i
    'Suma: Resultados Acumulados
        For i = 1 To 6
            nTotalResultados = nTotalResultados + CCur(Me.feCambiosPatrimonio.TextMatrix(i, 5))
            Me.feCambiosPatrimonio.TextMatrix(7, 5) = Format(nTotalResultados, "#,#0.00")
        Next i
    'Suma: Utilidad Ejercicio
        For i = 1 To 6
            nTotalUtilidadEjercicio = nTotalUtilidadEjercicio + CCur(Me.feCambiosPatrimonio.TextMatrix(i, 6))
            Me.feCambiosPatrimonio.TextMatrix(7, 6) = Format(nTotalUtilidadEjercicio, "#,#0.00")
        Next i
    'Suma Fila - Total Patrimonio
        nTotalPatrimonio1 = CCur(Me.feCambiosPatrimonio.TextMatrix(1, 2)) + CCur(Me.feCambiosPatrimonio.TextMatrix(1, 3)) + CCur(Me.feCambiosPatrimonio.TextMatrix(1, 4)) + CCur(Me.feCambiosPatrimonio.TextMatrix(1, 5)) + CCur(Me.feCambiosPatrimonio.TextMatrix(1, 6))
        nTotalPatrimonio2 = CCur(Me.feCambiosPatrimonio.TextMatrix(2, 2)) + CCur(Me.feCambiosPatrimonio.TextMatrix(2, 3)) + CCur(Me.feCambiosPatrimonio.TextMatrix(2, 4)) + CCur(Me.feCambiosPatrimonio.TextMatrix(2, 5)) + CCur(Me.feCambiosPatrimonio.TextMatrix(2, 6))
        nTotalPatrimonio3 = CCur(Me.feCambiosPatrimonio.TextMatrix(3, 2)) + CCur(Me.feCambiosPatrimonio.TextMatrix(3, 3)) + CCur(Me.feCambiosPatrimonio.TextMatrix(3, 4)) + CCur(Me.feCambiosPatrimonio.TextMatrix(3, 5)) + CCur(Me.feCambiosPatrimonio.TextMatrix(3, 6))
        nTotalPatrimonio4 = CCur(Me.feCambiosPatrimonio.TextMatrix(4, 2)) + CCur(Me.feCambiosPatrimonio.TextMatrix(4, 3)) + CCur(Me.feCambiosPatrimonio.TextMatrix(4, 4)) + CCur(Me.feCambiosPatrimonio.TextMatrix(4, 5)) + CCur(Me.feCambiosPatrimonio.TextMatrix(4, 6))
        nTotalPatrimonio5 = CCur(Me.feCambiosPatrimonio.TextMatrix(5, 2)) + CCur(Me.feCambiosPatrimonio.TextMatrix(5, 3)) + CCur(Me.feCambiosPatrimonio.TextMatrix(5, 4)) + CCur(Me.feCambiosPatrimonio.TextMatrix(5, 5)) + CCur(Me.feCambiosPatrimonio.TextMatrix(5, 6))
        nTotalPatrimonio6 = CCur(Me.feCambiosPatrimonio.TextMatrix(6, 2)) + CCur(Me.feCambiosPatrimonio.TextMatrix(6, 3)) + CCur(Me.feCambiosPatrimonio.TextMatrix(6, 4)) + CCur(Me.feCambiosPatrimonio.TextMatrix(6, 5)) + CCur(Me.feCambiosPatrimonio.TextMatrix(6, 6))
        
        Me.feCambiosPatrimonio.TextMatrix(1, 7) = Format(nTotalPatrimonio1, "#,#0.00")
        Me.feCambiosPatrimonio.TextMatrix(2, 7) = Format(nTotalPatrimonio2, "#,#0.00")
        Me.feCambiosPatrimonio.TextMatrix(3, 7) = Format(nTotalPatrimonio3, "#,#0.00")
        Me.feCambiosPatrimonio.TextMatrix(4, 7) = Format(nTotalPatrimonio4, "#,#0.00")
        Me.feCambiosPatrimonio.TextMatrix(5, 7) = Format(nTotalPatrimonio5, "#,#0.00")
        Me.feCambiosPatrimonio.TextMatrix(6, 7) = Format(nTotalPatrimonio6, "#,#0.00")
        
        For i = 1 To 6
            nTotalPatrimonio7 = nTotalPatrimonio7 + CCur(Me.feCambiosPatrimonio.TextMatrix(i, 7))
            Me.feCambiosPatrimonio.TextMatrix(7, 7) = Format(nTotalPatrimonio7, "#,#0.00")
        Next i
End Sub
Private Sub CargaControlesInicio()
    Call CargarFlexEEFF
End Sub
Private Sub CargarFlexEEFF()
    Dim lnFila As Integer
    Dim lnNumRegRS As Integer
    Dim CargarFlexEEFF As Boolean
    Dim oNCOMFormatosEval As COMNCredito.NCOMFormatosEval
    Set oNCOMFormatosEval = New COMNCredito.NCOMFormatosEval
    
       CargarFlexEEFF = oNCOMFormatosEval.CargaDatosFlexEEFF(sPersCod, _
                                                             rsFeActivosEEFF, _
                                                             rsFePasivosEEFF, _
                                                             rsFeEstadoResultadoEEFF, _
                                                             rsFeFlujoEfectivoEEFF, _
                                                             rsFeCambiosPatrimonioEEFF)
    'Flex Activos ->CargaInicial
    feActivos.Clear
    feActivos.FormaCabecera
    feActivos.rows = 2
    Call LimpiaFlex(feActivos)
    lnFila = 0
    lnNumRegRS = 0
    lnNumRegRS = rsFeActivosEEFF.RecordCount
    ReDim lvPrincipalActivos(lnNumRegRS)
    Do While Not rsFeActivosEEFF.EOF
        feActivos.AdicionaFila
        lnFila = feActivos.row
        feActivos.TextMatrix(lnFila, 1) = rsFeActivosEEFF!Concepto
        feActivos.TextMatrix(lnFila, 2) = Format(rsFeActivosEEFF!nMonto, "#,#0.00")
        feActivos.TextMatrix(lnFila, 3) = Format(rsFeActivosEEFF!nPorcentaje, "#,#0.00")
        feActivos.TextMatrix(lnFila, 4) = rsFeActivosEEFF!nConsCod
        feActivos.TextMatrix(lnFila, 5) = rsFeActivosEEFF!nConsValor
        feActivos.TextMatrix(lnFila, 6) = rsFeActivosEEFF!nConsValorGrupo
        feActivos.TextMatrix(lnFila, 7) = rsFeActivosEEFF!nEstFinanTipo
        
        'Lena datos de Registro en Matrix "lvPrincipalActivos"
        lvPrincipalActivos(lnFila).cConcepto = rsFeActivosEEFF!Concepto
        lvPrincipalActivos(lnFila).nImporte = rsFeActivosEEFF!nMonto
        lvPrincipalActivos(lnFila).nConsCod = rsFeActivosEEFF!nConsCod
        lvPrincipalActivos(lnFila).nConsValor = rsFeActivosEEFF!nConsValor
        lvPrincipalActivos(lnFila).nConsValorGrupo = rsFeActivosEEFF!nConsValorGrupo
        lvPrincipalActivos(lnFila).nEstFinanTipo = rsFeActivosEEFF!nEstFinanTipo
        
        
    Select Case CInt(feActivos.TextMatrix(feActivos.row, 5)) 'celda que o se puede editar
        Case 100, 200, 300, 1000
            Me.feActivos.BackColorRow (&H80000000)
            Me.feActivos.ForeColorRow vbBlack, True
        Case 29, 36, 39
            Me.feActivos.ForeColorRow vbRed
        Case Else
            Me.feActivos.BackColorRow (&HFFFFFF)
            Me.feActivos.ColumnasAEditar = "X-X-2-X-X-X-X"
    End Select
    rsFeActivosEEFF.MoveNext
    Loop
    rsFeActivosEEFF.Close
    Set rsFeActivosEEFF = Nothing

'Flex Pasivos ->CargaInicial
    fePasivos.Clear
    fePasivos.FormaCabecera
    fePasivos.rows = 2
    Call LimpiaFlex(fePasivos)
        lnFila = 0
        lnNumRegRS = 0
        lnNumRegRS = rsFePasivosEEFF.RecordCount
        ReDim lvPrincipalPasivos(lnNumRegRS)
        Do While Not rsFePasivosEEFF.EOF
            fePasivos.AdicionaFila
            lnFila = fePasivos.row
            fePasivos.TextMatrix(lnFila, 1) = rsFePasivosEEFF!Concepto
            fePasivos.TextMatrix(lnFila, 2) = Format(rsFePasivosEEFF!nMonto, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 3) = Format(rsFePasivosEEFF!nPorcentaje, "#,#0.00")
            fePasivos.TextMatrix(lnFila, 4) = rsFePasivosEEFF!nConsCod
            fePasivos.TextMatrix(lnFila, 5) = rsFePasivosEEFF!nConsValor
            fePasivos.TextMatrix(lnFila, 6) = rsFePasivosEEFF!nConsValorGrupo
            fePasivos.TextMatrix(lnFila, 7) = rsFePasivosEEFF!nEstFinanTipo
            
            'Lena datos de Registro en Matrix "lvPrincipalPasivos"
            lvPrincipalPasivos(lnFila).cConcepto = rsFePasivosEEFF!Concepto
            lvPrincipalPasivos(lnFila).nImporte = rsFePasivosEEFF!nMonto
            lvPrincipalPasivos(lnFila).nConsCod = rsFePasivosEEFF!nConsCod
            lvPrincipalPasivos(lnFila).nConsValor = rsFePasivosEEFF!nConsValor
            lvPrincipalPasivos(lnFila).nConsValorGrupo = rsFePasivosEEFF!nConsValorGrupo
            lvPrincipalPasivos(lnFila).nEstFinanTipo = rsFePasivosEEFF!nEstFinanTipo
            
        Select Case CInt(fePasivos.TextMatrix(fePasivos.row, 5)) 'celda que  o se puede editar
            Case 100, 200, 300, 1000, 1002
                Me.fePasivos.BackColorRow (&H80000000)
                Me.fePasivos.ForeColorRow vbBlack, True
            Case 45, 59
                Me.fePasivos.ForeColorRow vbBlack, True
            Case 50, 5912
                Me.fePasivos.ForeColorRow vbGrayText, True
            Case Else
                Me.fePasivos.BackColorRow (&HFFFFFF)
                Me.fePasivos.ColumnasAEditar = "X-X-2-3-X-X-X"
        End Select
        rsFePasivosEEFF.MoveNext
        Loop
    rsFePasivosEEFF.Close
    Set rsFePasivosEEFF = Nothing
    
   'Flex Estado Resultado ->CargaInicial
    feEstadoResultado.Clear
    feEstadoResultado.FormaCabecera
    feEstadoResultado.rows = 2
    Call LimpiaFlex(feEstadoResultado)
        lnFila = 0
        Do While Not rsFeEstadoResultadoEEFF.EOF
            feEstadoResultado.AdicionaFila
            lnFila = feEstadoResultado.row
            feEstadoResultado.TextMatrix(lnFila, 1) = rsFeEstadoResultadoEEFF!Concepto
            feEstadoResultado.TextMatrix(lnFila, 2) = Format(rsFeEstadoResultadoEEFF!nMonto, "#,#0.00")
            feEstadoResultado.TextMatrix(lnFila, 3) = Format(rsFeEstadoResultadoEEFF!nPorcentaje, "#,#0.00")
            feEstadoResultado.TextMatrix(lnFila, 4) = rsFeEstadoResultadoEEFF!nConsCod
            feEstadoResultado.TextMatrix(lnFila, 5) = rsFeEstadoResultadoEEFF!nConsValor
            feEstadoResultado.TextMatrix(lnFila, 6) = rsFeEstadoResultadoEEFF!nEstFinanTipo
            
        Select Case CInt(feEstadoResultado.TextMatrix(feEstadoResultado.row, 5)) 'celda que  o se puede editar
            Case 1, 11
                Me.feEstadoResultado.ForeColorRow vbBlack, True
            Case 10, 20, 30, 50, 60
                Me.feEstadoResultado.BackColorRow (&H80000000)
                Me.feEstadoResultado.ForeColorRow vbBlack, True
                Me.fePasivos.ColumnasAEditar = "X-X-X-X-X-X-X"
            Case Else
                Me.feEstadoResultado.BackColorRow (&HFFFFFF)
                Me.fePasivos.ColumnasAEditar = "X-X-2-X-X-X-X"
        End Select
        rsFeEstadoResultadoEEFF.MoveNext
        Loop
    rsFeEstadoResultadoEEFF.Close
    Set rsFeEstadoResultadoEEFF = Nothing
    
    'Flex Flujo Efectivo ->CargaInicial
    fefujoEfectivo.Clear
    fefujoEfectivo.FormaCabecera
    fefujoEfectivo.rows = 2
    Call LimpiaFlex(fefujoEfectivo)
    lnFila = 0
    lnNumRegRS = 0
    lnNumRegRS = rsFeFlujoEfectivoEEFF.RecordCount
    ReDim lvPrincipalFlujoEfectivo(lnNumRegRS)
    Do While Not rsFeFlujoEfectivoEEFF.EOF
        fefujoEfectivo.AdicionaFila
        lnFila = fefujoEfectivo.row
        fefujoEfectivo.TextMatrix(lnFila, 1) = rsFeFlujoEfectivoEEFF!Concepto
        fefujoEfectivo.TextMatrix(lnFila, 2) = Format(rsFeFlujoEfectivoEEFF!nMonto, "#,#0.00")
        fefujoEfectivo.TextMatrix(lnFila, 3) = rsFeFlujoEfectivoEEFF!nConsCod
        fefujoEfectivo.TextMatrix(lnFila, 4) = rsFeFlujoEfectivoEEFF!nConsValor
        fefujoEfectivo.TextMatrix(lnFila, 5) = rsFeFlujoEfectivoEEFF!nEstFinanTipo
        
        'Lena datos de Registro en Matrix "lvPrincipalActivosPasivos"
        lvPrincipalFlujoEfectivo(lnFila).cConcepto = rsFeFlujoEfectivoEEFF!Concepto
        lvPrincipalFlujoEfectivo(lnFila).nImporte = rsFeFlujoEfectivoEEFF!nMonto
        lvPrincipalFlujoEfectivo(lnFila).nConsCod = rsFeFlujoEfectivoEEFF!nConsCod
        lvPrincipalFlujoEfectivo(lnFila).nConsValor = rsFeFlujoEfectivoEEFF!nConsValor
        lvPrincipalFlujoEfectivo(lnFila).nEstFinanTipo = rsFeFlujoEfectivoEEFF!nEstFinanTipo
        
    Select Case CInt(fefujoEfectivo.TextMatrix(fefujoEfectivo.row, 4)) 'celda que  o se puede editar
        Case 1, 2, 3, 4, 5, 6
            Me.fefujoEfectivo.BackColorRow (&H80000000)
            Me.fefujoEfectivo.ForeColorRow vbBlack, True
        Case 130, 230, 330
            Me.fefujoEfectivo.ForeColorRow vbBlack, True
        Case Else
            Me.fefujoEfectivo.BackColorRow (&HFFFFFF)
            Me.fefujoEfectivo.ColumnasAEditar = "X-X-2-3-X-X-X"
    End Select
    rsFeFlujoEfectivoEEFF.MoveNext
    Loop
    rsFeFlujoEfectivoEEFF.Close
    Set rsFeFlujoEfectivoEEFF = Nothing
    
    'Flex Cambios en el Patrimonio ->CargaInicial
    feCambiosPatrimonio.Clear
    feCambiosPatrimonio.FormaCabecera
    feCambiosPatrimonio.rows = 2
    Call LimpiaFlex(feCambiosPatrimonio)
        lnFila = 0
        Do While Not rsFeCambiosPatrimonioEEFF.EOF
            feCambiosPatrimonio.AdicionaFila
            lnFila = feCambiosPatrimonio.row
            feCambiosPatrimonio.TextMatrix(lnFila, 1) = rsFeCambiosPatrimonioEEFF!Concepto
            feCambiosPatrimonio.TextMatrix(lnFila, 2) = Format(rsFeCambiosPatrimonioEEFF!nCapSocial, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 3) = Format(rsFeCambiosPatrimonioEEFF!nCapSocialAdicional, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 4) = Format(rsFeCambiosPatrimonioEEFF!nReservaLegal, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 5) = Format(rsFeCambiosPatrimonioEEFF!nResultadoAcumulado, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 6) = Format(rsFeCambiosPatrimonioEEFF!nUtilidadEjercicio, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 7) = Format(rsFeCambiosPatrimonioEEFF!nTotalPatrimonio, "#,#0.00")
            feCambiosPatrimonio.TextMatrix(lnFila, 8) = rsFeCambiosPatrimonioEEFF!nConsCod
            feCambiosPatrimonio.TextMatrix(lnFila, 9) = rsFeCambiosPatrimonioEEFF!nConsValor
            feCambiosPatrimonio.TextMatrix(lnFila, 10) = rsFeCambiosPatrimonioEEFF!nEstFinanTipo
            
        Select Case CInt(feCambiosPatrimonio.TextMatrix(feCambiosPatrimonio.row, 9)) 'celda que  o se puede editar
            Case 1, 7
                Me.feCambiosPatrimonio.BackColorRow (&H80000000)
                Me.feCambiosPatrimonio.ForeColorRow vbBlack, True
        End Select
        rsFeCambiosPatrimonioEEFF.MoveNext
        Loop
    rsFeCambiosPatrimonioEEFF.Close
    Set rsFeCambiosPatrimonioEEFF = Nothing
End Sub
Public Function CargarMatrizDatosMantenimientoSubItems(ByRef pvDetalleItemFlex() As tfrmPersEstadosFinancierosDetalle, ByVal pnCodEF As Integer, ByVal pnConsCod As Integer, ByVal pnConsValor As Integer, Optional ByVal pnConsValorGrupo As Integer = -1) As Integer
    Dim oNCOMFormatosEval As New COMNCredito.NCOMFormatosEval
    Dim rsSubItemsRegistrados As ADODB.Recordset
    Dim i As Integer
    
    Set rsSubItemsRegistrados = oNCOMFormatosEval.RecuperaDatosSubItemDetalle(pnCodEF, pnConsCod, pnConsValor, pnConsValorGrupo)
    If Not (rsSubItemsRegistrados.EOF And rsSubItemsRegistrados.BOF) Then
        CargarMatrizDatosMantenimientoSubItems = rsSubItemsRegistrados.RecordCount

        ReDim Preserve pvDetalleItemFlex(CargarMatrizDatosMantenimientoSubItems)
        For i = 1 To CargarMatrizDatosMantenimientoSubItems
            pvDetalleItemFlex(i).cDescripcion = rsSubItemsRegistrados!Concepto
            pvDetalleItemFlex(i).nImporte = rsSubItemsRegistrados!nMonto
            rsSubItemsRegistrados.MoveNext
        Next i
    Else
        ReDim Preserve pvDetalleItemFlex(0)
    End If
End Function
Private Sub LimpiarEstFinanNew()
    Me.txtFechaEF.Text = "__/__/____"
    Me.txtPatrimonioIni.Text = "__/__/____"
    Me.txtPatrimonioFin.Text = "__/__/____"
        
    Call LimpiaFlex(feActivos)
    Call LimpiaFlex(fePasivos)
    Call LimpiaFlex(feEstadoResultado)
    Call LimpiaFlex(fefujoEfectivo)
    Call LimpiaFlex(feCambiosPatrimonio)
End Sub
'<-*************** Fin LUCV20170915 **********<-
'->********** EAAS20170915 ->**********
Private Sub cmdSolicitarEdicionEF_Click()
    Dim nOp As Integer
    nOp = 0
    frmPersEstadosFinancierosSolicitaAutoriza.Inicio nCodEF, sPersCod, sFechaEF, nOp
    cmdSolicitarEdicionEF.Enabled = False
End Sub
Private Sub cmdAutorizarEF_Click()
    If cEstadoEF = 1 Then
        MsgBox "Ya hay una Autorización de edición para el EEFF. ", vbInformation, "Aviso"
        cmdAutorizarEF.Enabled = False
        ElseIf cEstadoEF = -1 Then
        MsgBox "No hay solicitud de edición para el EEFF. ", vbInformation, "Aviso"
        cmdAutorizarEF.Enabled = False
    Else
        frmPersEstadosFinancierosSolicitaAutoriza.Inicio nCodEF, sPersCod, sFechaEF
    End If
    cmdSolicitarEdicionEF.Enabled = False
    cmdAutorizarEF.Enabled = False
End Sub
'<-********** Fin EAAS20170915 <-**********
Private Sub txtFechaEF_GotFocus()
    fEnfoque txtFechaEF
End Sub
Private Sub txtFechaEF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    feActivos.SetFocus
    If Not IsDate(txtFechaEF) Then
        MsgBox "Verifique Dia,Mes,Año , Fecha Incorrecta", vbInformation, "Aviso"
        txtFechaEF.SetFocus
    End If
    'SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFActFijo_GotFocus()
    fEnfoque txtEFActFijo
End Sub
Private Sub txtEFActFijo_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFActFijo, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFActFijo, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFActFijo_LostFocus()
    If Trim(txtEFActFijo.Text) = "" Then
        txtEFActFijo.Text = "0.00"
    Else
        txtEFActFijo.Text = Format(txtEFActFijo.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(2)
End Sub
Private Sub txtEFCajaBanco_GotFocus()
    fEnfoque txtEFCajaBanco
End Sub
Private Sub txtEFCajaBanco_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFCajaBanco, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFCajaBanco, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFCajaBanco_LostFocus()
    If Trim(txtEFCajaBanco.Text) = "" Then
        txtEFCajaBanco.Text = "0.00"
    Else
        txtEFCajaBanco.Text = Format(txtEFCajaBanco.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub
Private Sub txtEFCapSocial_GotFocus()
    fEnfoque txtEFCapSocial
End Sub
Private Sub txtEFCapSocial_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFCapSocial, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFCapSocial, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFCapSocial_LostFocus()
    If Trim(txtEFCapSocial.Text) = "" Then
        txtEFCapSocial.Text = "0.00"
    Else
        txtEFCapSocial.Text = Format(txtEFCapSocial.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(5)
End Sub
Private Sub txtEFCostVentas_GotFocus()
    fEnfoque txtEFCostVentas
End Sub
Private Sub txtEFCostVentas_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFCostVentas, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFCostVentas, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFCostVentas_LostFocus()
    If Trim(txtEFCostVentas.Text) = "" Then
        txtEFCostVentas.Text = "0.00"
    Else
        txtEFCostVentas.Text = Format(txtEFCostVentas.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(6)
End Sub
Private Sub txtEFCuentaCobrar_GotFocus()
    fEnfoque txtEFCuentaCobrar
End Sub
Private Sub txtEFCuentaCobrar_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFCuentaCobrar, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFCuentaCobrar, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFCuentaCobrar_LostFocus()
    If Trim(txtEFCuentaCobrar.Text) = "" Then
        txtEFCuentaCobrar.Text = "0.00"
    Else
        txtEFCuentaCobrar.Text = Format(txtEFCuentaCobrar.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub
Private Sub txtEFDeudaFinan_GotFocus()
    fEnfoque txtEFDeudaFinan
End Sub
Private Sub txtEFDeudaFinan_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFDeudaFinan, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFDeudaFinan, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFDeudaFinan_LostFocus()
    If Trim(txtEFDeudaFinan.Text) = "" Then
        txtEFDeudaFinan.Text = "0.00"
    Else
        txtEFDeudaFinan.Text = Format(txtEFDeudaFinan.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub
Private Sub txtEFDeudaFinanC_GotFocus()
    fEnfoque txtEFDeudaFinanC
End Sub
Private Sub txtEFDeudaFinanC_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFDeudaFinanC, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFDeudaFinanC, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFDeudaFinanC_LostFocus()
    If Trim(txtEFDeudaFinanC.Text) = "" Then
        txtEFDeudaFinanC.Text = "0.00"
    Else
        txtEFDeudaFinanC.Text = Format(txtEFDeudaFinanC.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(3)
End Sub
Private Sub txtEFDeudaFinanL_GotFocus()
    fEnfoque txtEFDeudaFinanL
End Sub
Private Sub txtEFDeudaFinanL_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFDeudaFinanL, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFDeudaFinanL, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFDeudaFinanL_LostFocus()
    If Trim(txtEFDeudaFinanL.Text) = "" Then
        txtEFDeudaFinanL.Text = "0.00"
    Else
        txtEFDeudaFinanL.Text = Format(txtEFDeudaFinanL.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(3)
End Sub
Private Sub txtEFExiste_GotFocus()
    fEnfoque txtEFExiste
End Sub
Private Sub txtEFExiste_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFExiste, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFExiste, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFExiste_LostFocus()
    If Trim(txtEFExiste.Text) = "" Then
        txtEFExiste.Text = "0.00"
    Else
        txtEFExiste.Text = Format(txtEFExiste.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(1)
End Sub
Private Sub txtEFFlujoFinan_GotFocus()
    fEnfoque txtEFFlujoFinan
End Sub
Private Sub txtEFFlujoFinan_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFFlujoFinan, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFFlujoFinan, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFFlujoFinan_LostFocus()
    If Trim(txtEFFlujoFinan.Text) = "" Then
        txtEFFlujoFinan.Text = "0.00"
    Else
        txtEFFlujoFinan.Text = Format(txtEFFlujoFinan.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(9)
End Sub
Private Sub txtEFFlujoOpe_GotFocus()
    fEnfoque txtEFFlujoOpe
End Sub
Private Sub txtEFFlujoOpe_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFFlujoOpe, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFFlujoOpe, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFFlujoOpe_LostFocus()
    If Trim(txtEFFlujoOpe.Text) = "" Then
        txtEFFlujoOpe.Text = "0.00"
    Else
        txtEFFlujoOpe.Text = Format(txtEFFlujoOpe.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(9)
End Sub
Private Sub txtEFFujoInv_GotFocus()
    fEnfoque txtEFFujoInv
End Sub
Private Sub txtEFFujoInv_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFFujoInv, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFFujoInv, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFFujoInv_LostFocus()
    If Trim(txtEFFujoInv.Text) = "" Then
        txtEFFujoInv.Text = "0.00"
    Else
        txtEFFujoInv.Text = Format(txtEFFujoInv.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(9)
End Sub
Private Sub txtEFGastoFinan_GotFocus()
    fEnfoque txtEFGastoFinan
End Sub
Private Sub txtEFGastoFinan_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFGastoFinan, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFGastoFinan, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFGastoFinan_LostFocus()
    If Trim(txtEFGastoFinan.Text) = "" Then
        txtEFGastoFinan.Text = "0.00"
    Else
        txtEFGastoFinan.Text = Format(txtEFGastoFinan.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(8)
End Sub
Private Sub txtEFGastosME_GotFocus()
    fEnfoque txtEFGastosME
End Sub
Private Sub txtEFGastosME_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFGastosME, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFGastosME, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFGastosME_LostFocus()
    If Trim(txtEFGastosME.Text) = "" Then
        txtEFGastosME.Text = "0.00"
    Else
        txtEFGastosME.Text = Format(txtEFGastosME.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub
Private Sub txtEFGastosOpe_GotFocus()
    fEnfoque txtEFGastosOpe
End Sub
Private Sub txtEFGastosOpe_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFGastosOpe, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFGastosOpe, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFGastosOpe_LostFocus()
    If Trim(txtEFGastosOpe.Text) = "" Then
        txtEFGastosOpe.Text = "0.00"
    Else
        txtEFGastosOpe.Text = Format(txtEFGastosOpe.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(7)
End Sub
Private Sub txtEFIngresoFinan_GotFocus()
    fEnfoque txtEFIngresoFinan
End Sub
Private Sub txtEFIngresoFinan_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFIngresoFinan, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFIngresoFinan, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFIngresoFinan_LostFocus()
   If Trim(txtEFIngresoFinan.Text) = "" Then
        txtEFIngresoFinan.Text = "0.00"
    Else
        txtEFIngresoFinan.Text = Format(txtEFIngresoFinan.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(8)
End Sub
Private Sub txtEFIngresoME_GotFocus()
    fEnfoque txtEFIngresoME
End Sub
Private Sub txtEFIngresoME_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFIngresoME, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFIngresoME, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFIngresoME_LostFocus()
    If Trim(txtEFIngresoME.Text) = "" Then
        txtEFIngresoME.Text = "0.00"
    Else
        txtEFIngresoME.Text = Format(txtEFIngresoME.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub
Private Sub txtEFPosCambios_GotFocus()
    fEnfoque txtEFPosCambios
End Sub
Private Sub txtEFPosCambios_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFPosCambios, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFPosCambios, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFPosCambios_LostFocus()
    If Trim(txtEFPosCambios.Text) = "" Then
        txtEFPosCambios.Text = "0.00"
    Else
        txtEFPosCambios.Text = Format(txtEFPosCambios.Text, "###," & String(15, "#") & "#0.00")
    End If
End Sub
Private Sub txtEFProveedor_GotFocus()
    fEnfoque txtEFProveedor
End Sub
Private Sub txtEFProveedor_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFProveedor, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFProveedor, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFProveedor_LostFocus()
    If Trim(txtEFProveedor.Text) = "" Then
        txtEFProveedor.Text = "0.00"
    Else
        txtEFProveedor.Text = Format(txtEFProveedor.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(3)
End Sub
Private Sub txtEFResulAcum_GotFocus()
    fEnfoque txtEFResulAcum
End Sub
Private Sub txtEFResulAcum_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFResulAcum, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFResulAcum, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFResulAcum_LostFocus()
    If Trim(txtEFResulAcum.Text) = "" Then
        txtEFResulAcum.Text = "0.00"
    Else
        txtEFResulAcum.Text = Format(txtEFResulAcum.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(5)
End Sub
Private Sub txtEFVentas_GotFocus()
    fEnfoque txtEFVentas
End Sub
Private Sub txtEFVentas_KeyPress(KeyAscii As Integer)
    'KeyAscii = NumerosDecimales(txtEFVentas, KeyAscii, , , True)
    KeyAscii = NumerosDecimales(txtEFVentas, KeyAscii, 10, , True) 'FRHU 20150611
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
    End If
End Sub
Private Sub txtEFVentas_LostFocus()
    If Trim(txtEFVentas.Text) = "" Then
        txtEFVentas.Text = "0.00"
    Else
        txtEFVentas.Text = Format(txtEFVentas.Text, "###," & String(15, "#") & "#0.00")
    End If
    Call CalculoTotal(6)
End Sub



