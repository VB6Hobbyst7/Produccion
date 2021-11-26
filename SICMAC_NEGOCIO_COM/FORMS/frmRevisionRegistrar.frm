VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRevisionRegistrar 
   Caption         =   "Registrar Revisión de la Calificación"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRevisionRegistrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Formato de Revisión de Clasificación "
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      Begin VB.Frame Frame5 
         Caption         =   "Datos"
         Height          =   1815
         Left            =   120
         TabIndex        =   66
         Top             =   1320
         Width           =   11295
         Begin VB.ComboBox cboTMonedaSF 
            Height          =   315
            ItemData        =   "frmRevisionRegistrar.frx":030A
            Left            =   4320
            List            =   "frmRevisionRegistrar.frx":0314
            TabIndex        =   14
            Top             =   1320
            Width           =   670
         End
         Begin VB.ComboBox cboTMonedaCMAC 
            Height          =   315
            ItemData        =   "frmRevisionRegistrar.frx":0320
            Left            =   10320
            List            =   "frmRevisionRegistrar.frx":032A
            TabIndex        =   11
            Top             =   960
            Width           =   670
         End
         Begin VB.TextBox txtMontoSF 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   13
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtMontoCMAC 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8880
            TabIndex        =   10
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtDireccion 
            Height          =   285
            Left            =   7320
            TabIndex        =   7
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox txtDeudor 
            Height          =   285
            Left            =   1200
            TabIndex        =   6
            Top             =   600
            Width           =   3735
         End
         Begin VB.TextBox txtGiro 
            Height          =   285
            Left            =   1920
            TabIndex        =   8
            Top             =   960
            Width           =   3015
         End
         Begin VB.TextBox txtCAnalista 
            Height          =   285
            Left            =   9600
            TabIndex        =   5
            Top             =   240
            Width           =   1335
         End
         Begin MSMask.MaskEdBox txtFSDCmac 
            Height          =   300
            Left            =   7560
            TabIndex        =   9
            Top             =   960
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFRegistro 
            Height          =   300
            Left            =   1200
            TabIndex        =   4
            Top             =   240
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFSDSF 
            Height          =   300
            Left            =   1920
            TabIndex        =   12
            Top             =   1320
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            Caption         =   "F. Reg:"
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Dirección:"
            Height          =   255
            Left            =   6120
            TabIndex        =   72
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Deudor:"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "Giro del Negocio:"
            Height          =   255
            Left            =   240
            TabIndex        =   70
            Top             =   915
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Cód del Analista:"
            Height          =   255
            Left            =   6120
            TabIndex        =   69
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label6 
            Caption         =   "S.D.Cmac al:"
            Height          =   255
            Left            =   6120
            TabIndex        =   68
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "SD-Sist. Fin. al:"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   1320
            Width           =   1455
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Parámetros de Búsqueda"
         Height          =   1095
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   11295
         Begin VB.TextBox txtCodCta 
            Height          =   285
            Left            =   7200
            TabIndex        =   2
            Top             =   285
            Width           =   2055
         End
         Begin VB.TextBox txtTCambio 
            Height          =   285
            Left            =   10320
            TabIndex        =   3
            Top             =   285
            Width           =   735
         End
         Begin SICMACT.TxtBuscar txtCodigo 
            Height          =   285
            Left            =   960
            TabIndex        =   1
            Top             =   280
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   503
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
            TipoBusqueda    =   3
            sTitulo         =   ""
         End
         Begin MSMask.MaskEdBox mskPeriodo1Del 
            Height          =   300
            Left            =   4440
            TabIndex        =   81
            Top             =   285
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   12648447
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Cta"
            Height          =   195
            Left            =   6360
            TabIndex        =   80
            Top             =   285
            Width           =   765
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "T. C:"
            Height          =   195
            Left            =   9840
            TabIndex        =   74
            Top             =   285
            Width           =   435
         End
         Begin VB.Label lblPersCod 
            AutoSize        =   -1  'True
            Caption         =   "Cod Per:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   280
            Width           =   765
         End
         Begin VB.Label Label24 
            Caption         =   "F. Cierre:"
            Height          =   255
            Left            =   3600
            TabIndex        =   64
            Top             =   285
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   62
         Top             =   8160
         Width           =   11295
         Begin VB.CommandButton Command1 
            Caption         =   "&Aceptar"
            Height          =   375
            Left            =   4440
            TabIndex        =   42
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton Command3 
            Caption         =   "&Actualizar"
            Height          =   375
            Left            =   4440
            TabIndex        =   44
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "&Salir"
            Height          =   375
            Left            =   6000
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Revisión"
         Height          =   5055
         Left            =   120
         TabIndex        =   45
         Top             =   3120
         Width           =   11295
         Begin VB.TextBox txtPPerd 
            Height          =   285
            Left            =   10560
            TabIndex        =   29
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtPDud 
            Height          =   285
            Left            =   9000
            TabIndex        =   27
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtPCPP 
            Height          =   285
            Left            =   5880
            TabIndex        =   23
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtPDefic 
            Height          =   285
            Left            =   7440
            TabIndex        =   25
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtPNorm 
            Height          =   285
            Left            =   4320
            TabIndex        =   21
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtOCI 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   9840
            TabIndex        =   34
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtOCI 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   8280
            TabIndex        =   33
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtOCI 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   6720
            TabIndex        =   32
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtOCI 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   5160
            TabIndex        =   31
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtSF 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   9840
            TabIndex        =   28
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtSF 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   8280
            TabIndex        =   26
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtSF 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   6720
            TabIndex        =   24
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtSF 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   5160
            TabIndex        =   22
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtOCI 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3600
            TabIndex        =   30
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtSF 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3600
            TabIndex        =   20
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox txtN 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   9840
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtN 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   8280
            TabIndex        =   18
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtN 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   6720
            TabIndex        =   17
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtN 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   5160
            TabIndex        =   16
            Top             =   720
            Width           =   495
         End
         Begin VB.TextBox txtConclusion 
            Height          =   495
            Left            =   6120
            MultiLine       =   -1  'True
            TabIndex        =   41
            Top             =   3480
            Width           =   4815
         End
         Begin VB.TextBox txtComentario 
            Height          =   495
            Left            =   6120
            MultiLine       =   -1  'True
            TabIndex        =   40
            Top             =   2640
            Width           =   4815
         End
         Begin VB.TextBox txtEvaluacion 
            Height          =   495
            Left            =   6120
            MultiLine       =   -1  'True
            TabIndex        =   39
            Top             =   1800
            Width           =   4815
         End
         Begin VB.TextBox txtInformacion 
            Height          =   495
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   4440
            Width           =   4815
         End
         Begin VB.TextBox txtGarantia 
            Height          =   495
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   3600
            Width           =   4815
         End
         Begin VB.TextBox txtDesarrollo 
            Height          =   495
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   2760
            Width           =   4815
         End
         Begin VB.TextBox txtSituacion 
            Height          =   495
            Left            =   480
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   1920
            Width           =   4815
         End
         Begin VB.TextBox txtN 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3600
            TabIndex        =   15
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label30 
            Caption         =   "%"
            Height          =   255
            Left            =   11040
            TabIndex        =   79
            Top             =   960
            Width           =   165
         End
         Begin VB.Label Label29 
            Caption         =   "%"
            Height          =   255
            Left            =   9480
            TabIndex        =   78
            Top             =   960
            Width           =   165
         End
         Begin VB.Label Label28 
            Caption         =   "%"
            Height          =   255
            Left            =   7920
            TabIndex        =   77
            Top             =   960
            Width           =   165
         End
         Begin VB.Label Label27 
            Caption         =   "%"
            Height          =   255
            Left            =   6360
            TabIndex        =   76
            Top             =   960
            Width           =   165
         End
         Begin VB.Label Label26 
            Caption         =   "%"
            Height          =   255
            Left            =   4800
            TabIndex        =   75
            Top             =   960
            Width           =   165
         End
         Begin VB.Line Line15 
            X1              =   9720
            X2              =   10440
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line14 
            X1              =   9720
            X2              =   10440
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line13 
            X1              =   9720
            X2              =   10440
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Shape Shape5 
            Height          =   1095
            Left            =   9720
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label23 
            Caption         =   "Perd."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9840
            TabIndex        =   61
            Top             =   480
            Width           =   615
         End
         Begin VB.Line Line12 
            X1              =   8160
            X2              =   8880
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line11 
            X1              =   8160
            X2              =   8880
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line10 
            X1              =   8160
            X2              =   8880
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Shape Shape4 
            Height          =   1095
            Left            =   8160
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label22 
            Caption         =   "Dud."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8280
            TabIndex        =   60
            Top             =   480
            Width           =   615
         End
         Begin VB.Line Line9 
            X1              =   6600
            X2              =   7320
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line8 
            X1              =   6600
            X2              =   7320
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line7 
            X1              =   6600
            X2              =   7320
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Shape Shape3 
            Height          =   1095
            Left            =   6600
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "Defic."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6720
            TabIndex        =   59
            Top             =   480
            Width           =   615
         End
         Begin VB.Line Line6 
            X1              =   5040
            X2              =   5760
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line5 
            X1              =   5040
            X2              =   5760
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line4 
            X1              =   5040
            X2              =   5760
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Shape Shape2 
            Height          =   1095
            Left            =   5040
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label20 
            Caption         =   "CPP"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   58
            Top             =   480
            Width           =   615
         End
         Begin VB.Line Line3 
            X1              =   3480
            X2              =   4200
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Line Line2 
            X1              =   3480
            X2              =   4200
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Line Line1 
            X1              =   3480
            X2              =   4200
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Shape Shape1 
            Height          =   1095
            Left            =   3480
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label19 
            Caption         =   "Norm."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3600
            TabIndex        =   57
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "c). Calificación UAI"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "b). Calificación Sistema Financiero"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   960
            Width           =   3015
         End
         Begin VB.Label Label16 
            Caption         =   "a). Calificación CMAC Maynas-Riesgos"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   720
            Width           =   3375
         End
         Begin VB.Label Label15 
            Caption         =   "8.- Conclusión:"
            Height          =   255
            Left            =   6120
            TabIndex        =   53
            Top             =   3240
            Width           =   1335
         End
         Begin VB.Label Label14 
            Caption         =   "7.- Comentario de la Clasificación Asignada:"
            Height          =   255
            Left            =   6120
            TabIndex        =   52
            Top             =   2400
            Width           =   3855
         End
         Begin VB.Label Label13 
            Caption         =   "6.- Evaluación Capacidad de Pago:"
            Height          =   255
            Left            =   6120
            TabIndex        =   51
            Top             =   1560
            Width           =   3015
         End
         Begin VB.Label Label12 
            Caption         =   "5.- Información Financiera:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   4200
            Width           =   2415
         End
         Begin VB.Label Label11 
            Caption         =   "4.- Garantías:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label Label10 
            Caption         =   "3.- Desarrollo Crediticio:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   2520
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "2.- Situación del Deudor en el Sistema Financiero:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1680
            Width           =   4335
         End
         Begin VB.Label Label8 
            Caption         =   "1.- Calificación Crediticia:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   2295
         End
      End
   End
End
Attribute VB_Name = "frmRevisionRegistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmRevisionRegistrar
'** Descripción : formulario que permitirá el registro del formato de la revision de la calificacion.
'** Creación : MAVM, 20080807 10:00:00 AM
'** Modificación:
'********************************************************************

Option Explicit
Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
Dim FechaFinMes As Date
Public lsmensaje As String

Private Sub Command1_Click()
    Dim varCalificacionCMAC As String
    Dim varCalificacionSF As String
    Dim varCalificacionOCI As String
    Dim i, j, K As Integer
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    For i = 0 To 4
        If txtN(i).Text = "x" Then
            varCalificacionCMAC = i
        End If
    Next i
    For j = 0 To 4
        If txtSF(j).Text = "x" Then
            varCalificacionSF = j
        End If
    Next j
    For K = 0 To 4
        If txtOCI(K).Text = "x" Then
            varCalificacionOCI = K
        End If
    Next K
    If Len(Trim(txtCodigo.Text)) <> "0" And txtFRegistro.Text <> "__/__/____" And txtFSDSF.Text <> "__/__/____" And txtFSDCmac.Text <> "__/__/____" Then
        If MsgBox("Esta Seguro de Registar los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            gRevisionId = CInt(objCOMNAuditoria.InsertarRevision(txtCodigo.Text, txtCodCta.Text, mskPeriodo1Del.Text, txtTCambio.Text, txtFRegistro.Text, txtCAnalista.Text, txtGiro.Text, txtFSDCmac.Text, txtMontoCMAC.Text, cboTMonedaCMAC.Text, txtFSDSF.Text, txtMontoSF.Text, cboTMonedaSF.Text, txtPNorm.Text, txtPCPP.Text, txtPDefic.Text, txtPDud.Text, txtPPerd.Text, varCalificacionCMAC, varCalificacionSF, varCalificacionOCI, txtSituacion.Text, txtDesarrollo.Text, txtGarantia.Text, txtInformacion.Text, txtEvaluacion.Text, txtComentario.Text, txtConclusion.Text))
            MsgBox "Los Datos se Registraron Correctamente", vbInformation, Me.Caption
            If MsgBox("Desea Imprimir el Formato?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                Call ImprimeFormatoRevision
            End If
            cboTMonedaSF.Text = "MN"
            cboTMonedaCMAC.Text = "MN"
            LimpiarControles Me, True, False, False
            gRevisionId = 0
            txtCodigo.Text = ""
            txtFRegistro.Text = "__/__/____"
            txtFSDCmac.Text = "__/__/____"
            txtFSDSF.Text = "__/__/____"
            CargarDatosLoad 2
        End If
    Else
        MsgBox "Debe Completar los Campos Vacios", vbCritical, "Aviso"
    End If
    Set objCOMNAuditoria = Nothing
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Dim varCalificacionCMAC As String
    Dim varCalificacionSF As String
    Dim varCalificacionOCI As String
    Dim i, j, K As Integer
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    For i = 0 To 4
        If txtN(i).Text = "x" Then
            varCalificacionCMAC = i
        End If
    Next i
    For j = 0 To 4
        If txtSF(j).Text = "x" Then
            varCalificacionSF = j
        End If
    Next j
    For K = 0 To 4
        If txtOCI(K).Text = "x" Then
            varCalificacionOCI = K
        End If
    Next K
    If Len(Trim(txtCodigo.Text)) <> "0" Then
        If MsgBox("Esta Seguro de Actualizar los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
            objCOMNAuditoria.ActualizarRevisionCalificacion gRevisionId, txtCodigo.Text, txtCodCta.Text, mskPeriodo1Del.Text, txtTCambio.Text, txtFRegistro.Text, txtCAnalista.Text, txtGiro.Text, txtFSDCmac.Text, txtMontoCMAC.Text, cboTMonedaCMAC.Text, txtFSDSF.Text, txtMontoSF.Text, cboTMonedaSF.Text, txtPNorm.Text, txtPCPP.Text, txtPDefic.Text, txtPDud.Text, txtPPerd.Text, varCalificacionCMAC, varCalificacionSF, varCalificacionOCI, txtSituacion.Text, txtDesarrollo.Text, txtGarantia.Text, txtInformacion.Text, txtEvaluacion.Text, txtComentario.Text, txtConclusion.Text
            MsgBox "Los Datos se Actualizaron Correctamente", vbInformation, Me.Caption
            If MsgBox("Desea Imprimir el Formato?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                Call ImprimeFormatoRevision
            End If
            cboTMonedaSF.Text = "MN"
            cboTMonedaCMAC.Text = "MN"
            LimpiarControles Me, True, False, False
            Command1.Visible = True
            Command3.Visible = False
            gRevisionId = 0
            txtCodigo.Text = ""
            txtFRegistro.Text = "__/__/____"
            txtFSDCmac.Text = "__/__/____"
            txtFSDSF.Text = "__/__/____"
            CargarDatosLoad 2
        End If
    Else
    MsgBox "Debe Completar los Campos Vacios", vbCritical, "Aviso"
    End If
    Set objCOMNAuditoria = Nothing
End Sub

Private Sub Form_Load()
    cboTMonedaSF.Text = "MN"
    cboTMonedaCMAC.Text = "MN"
    If gRevisionId <> 0 Then
        CargarDatosModificar (gRevisionId)
        Command1.Visible = False
        Command3.Visible = True
    Else
        CargarDatosLoad 1
    End If
End Sub

Public Sub CargarDatosModificar(ByVal RevisionId As Integer)
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim rs As ADODB.Recordset
    Set rs = objCOMNAuditoria.ObtenerRevisionXId(gRevisionId)
    Dim i, j, K As Integer
    txtCodigo.Text = rs("cPersCod")
    txtCodCta.Text = rs("vCodCta")
    mskPeriodo1Del = rs("cFCierre")
    txtTCambio.Text = rs("mTCambio")
    txtFRegistro.Text = rs("cFRegistro")
    txtCAnalista.Text = rs("vCAnalista")
    txtDeudor.Text = rs("cPersNombre")
    txtDireccion.Text = rs("cPersDireccDomicilio")
    txtGiro.Text = rs("vGiro")
    txtFSDCmac.Text = rs("cFSDCMAC")
    txtMontoCMAC.Text = Format(rs("mMontoCMAC"), "#,##0.00")
    cboTMonedaCMAC.Text = rs("cTMonedaCMAC")
    txtFSDSF.Text = rs("cFSDSF")
    txtMontoSF.Text = Format(rs("mMontoSF"), "#,##0.00")
    cboTMonedaSF.Text = rs("cTMonedaSF")
    txtPNorm.Text = rs("cPNorm")
    txtPCPP.Text = rs("cPCPP")
    txtPDefic.Text = rs("cPDefic")
    txtPDud.Text = rs("cPDud")
    txtPPerd.Text = rs("cPPerd")
        If rs("vCalificacionCMAC") <> "" Then
            For i = 0 To 4
                If i = CInt(rs("vCalificacionCMAC")) Then
                    txtN(i).Text = "x"
                End If
            Next i
        End If
        
        If rs("vCalificacionSF") <> "" Then
            For j = 0 To 4
                If j = CInt(rs("vCalificacionSF")) Then
                    txtSF(j).Text = "x"
                End If
            Next j
        End If
        
        If rs("vCalificacionOCI") <> "" Then
            For K = 0 To 4
                If K = CInt(rs("vCalificacionOCI")) Then
                    txtOCI(K).Text = "x"
                End If
            Next K
        End If
    txtSituacion.Text = rs("vSituacion")
    txtDesarrollo.Text = rs("vDesarrollo")
    txtGarantia.Text = rs("vGarantia")
    txtInformacion.Text = rs("vInformacion")
    txtEvaluacion.Text = rs("vEvaluacion")
    txtComentario.Text = rs("vComentario")
    txtConclusion.Text = rs("vConclusion")
    Set objCOMNAuditoria = Nothing
End Sub

Public Sub CargarDatosLoad(ByVal sAccion As String)
    Dim oTipCambio As nTipoCambio
    If sAccion = "1" Then 'MAVM NR 20090818
        Me.mskPeriodo1Del = gdFecData 'MAVM NR 20090818
    End If 'MAVM NR 20090818
    txtFRegistro.Text = gdFecSis 'MAVM NR 20090818
    FechaFinMes = gdFecData
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Set oTipCambio = New nTipoCambio
        txtTCambio.Text = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#0.000")
    Set oTipCambio = Nothing
    txtPNorm.Text = "0"
    txtPCPP.Text = "0"
    txtPDefic.Text = "0"
    txtPDud.Text = "0"
    txtPPerd.Text = "0"
End Sub

Public Function ImprimeFormatoRevision() As String
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim rs1 As ADODB.Recordset
    Dim oWord As Word.Application
    Dim oDoc As Word.Document
    Dim sArchivo As String
    Set rs1 = objCOMNAuditoria.ObtenerRevisionXId(gRevisionId)
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = False
    Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\FormatoRevisionClasificacion.doc")
    sArchivo = App.path & "\FormatoCarta\FR_" & gRevisionId & "_" & Replace(Left(Time, 5), ":", "") & ".doc"
    oDoc.SaveAs (sArchivo)
    With oWord.Selection.Find
            .Text = "<<cFRegistro>>"
            .Replacement.Text = rs1.Fields("cFRegistro")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<cPersDireccDomicilio>>"
            .Replacement.Text = rs1.Fields("cPersDireccDomicilio")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<cPersNombre>>"
            .Replacement.Text = rs1.Fields("cPersNombre")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<vGiro>>"
            .Replacement.Text = rs1.Fields("vGiro")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<vCA>>"
            .Replacement.Text = rs1.Fields("vCAnalista")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<cFSDCMAC>>"
            .Replacement.Text = rs1.Fields("cFSDCMAC")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<MontoCMAC>>"
            .Replacement.Text = Format(rs1.Fields("mMontoCMAC"), "#,##0.00") & " " & rs1.Fields("cTMonedaCMAC")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<cFSDSF>>"
            .Replacement.Text = rs1.Fields("cFSDSF")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<MontoSF>>"
            .Replacement.Text = Format(rs1.Fields("mMontoSF"), "#,##0.00") & " " & rs1.Fields("cTMonedaSF")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    Dim i, j, K As Integer
    Dim var As String
    var = " "
    If rs1.Fields("vCalificacionCMAC") <> "" Then
        For i = 0 To 4
            If CInt(rs1.Fields("vCalificacionCMAC")) = i Then
                    With oWord.Selection.Find
                        .Text = "[" & i & "]"
                        .Replacement.Text = "x"
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
               Else
                With oWord.Selection.Find
                        .Text = "[" & i & "]"
                        .Replacement.Text = "     "
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
            End If
        Next i
    Else
                For i = 0 To 4
                    With oWord.Selection.Find
                            .Text = "(" & i & ")"
                            .Replacement.Text = "     "
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                    End With
                Next i
    End If
    If rs1.Fields("vCalificacionSF") <> "" Then
        For j = 0 To 4
            If CInt(rs1.Fields("vCalificacionSF")) = j Then
                With oWord.Selection.Find
                    .Text = "(" & j + 5 & ")"
                    .Replacement.Text = "x"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            Else
                With oWord.Selection.Find
                        .Text = "(" & j + 5 & ")"
                        .Replacement.Text = "    "
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
            End If
        Next j
        Else
                For j = 0 To 4
                    With oWord.Selection.Find
                            .Text = "(" & j + 5 & ")"
                            .Replacement.Text = "     "
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                    End With
                Next j
    End If
    If rs1.Fields("vCalificacionOCI") <> "" Then
        For K = 0 To 4
            If CInt(rs1.Fields("vCalificacionOCI")) = K Then
                With oWord.Selection.Find
                    .Text = "(" & K + 10 & ")"
                    .Replacement.Text = "x"
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            Else
                With oWord.Selection.Find
                        .Text = "(" & K + 10 & ")"
                        .Replacement.Text = "      "
                        .Forward = True
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                End With
            End If
        Next K
        Else
                For K = 0 To 4
                    With oWord.Selection.Find
                            .Text = "(" & K + 10 & ")"
                            .Replacement.Text = "     "
                            .Forward = True
                            .Wrap = wdFindContinue
                            .Format = False
                            .Execute Replace:=wdReplaceAll
                    End With
                Next K
    End If
    With oWord.Selection.Find
            .Text = "<<vSituacion>>"
            .Replacement.Text = rs1.Fields("vSituacion")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<vDesarrollo>>"
            .Replacement.Text = rs1.Fields("vDesarrollo")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<vGarantia>>"
            .Replacement.Text = rs1.Fields("vGarantia")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<vInformacion>>"
            .Replacement.Text = rs1.Fields("vInformacion")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<vEvaluacion>>"
            .Replacement.Text = rs1.Fields("vEvaluacion")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<vComentario>>"
            .Replacement.Text = rs1.Fields("vComentario")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "<<vConclusion>>"
            .Replacement.Text = rs1.Fields("vConclusion")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "[N]"
            .Replacement.Text = rs1.Fields("cPNorm") & "%"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "[C]"
            .Replacement.Text = rs1.Fields("cPCPP") & "%"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "[D]"
            .Replacement.Text = rs1.Fields("cPDefic") & "%"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "[U]"
            .Replacement.Text = rs1.Fields("cPDud") & "%"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    With oWord.Selection.Find
            .Text = "[P]"
            .Replacement.Text = rs1.Fields("cPPerd") & "%"
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
    End With
    oDoc.Close
    Set oDoc = Nothing
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing
    Set objCOMNAuditoria = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    gRevisionId = 0
End Sub

Private Sub txtCodigo_EmiteDatos()
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim rs As ADODB.Recordset
    If mskPeriodo1Del.Text <> "__/__/____" Then
        If txtCodigo.Text <> "" Then
            Set rs = objCOMNAuditoria.ValidarCierre(Format(mskPeriodo1Del.Text, "yyyymmdd"))
            If rs.RecordCount <> 0 Then
                Dim svarbol As Boolean
                svarbol = False
                frmAuditListarCreditos.CargarCreditos svarbol, txtCodigo.Text
                If svarbol = True Then frmAuditListarCreditos.Show 1
            Else
                MsgBox "El Proceso de Calificación aún no ha sido culminado", vbInformation, "Aviso"
                txtCodigo.Text = ""
            End If
        End If
    Else
        MsgBox ("Por Favor Elegir la Fecha de Cierre del Mes"), vbCritical
        txtCodigo.Text = ""
    End If
    Set objCOMNAuditoria = Nothing
End Sub

Public Sub BuscarValores()
    CargarDatosLoad 2
    Call CargarClasificacion(txtCodigo.Text, txtCodCta.Text, mskPeriodo1Del, txtTCambio.Text)
    txtFRegistro.SetFocus
End Sub

Public Sub CargarClasificacion(ByVal lsCodPer As String, ByVal lsCodCta As String, pdFechaProc As Date, pnTipCamb As Double)
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim rs As ADODB.Recordset
    Dim lsCalificacion As String
    Dim i As Integer
    lsmensaje = ""
    Set rs = objCOMNAuditoria.ObtenerDatosCalificacionXCliente(lsCodPer, lsCodCta, pdFechaProc, pnTipCamb, lsmensaje)
    If lsmensaje = "" Then
        txtCodCta.Text = rs("cCtaCod")
        txtGiro.Text = Trim(rs("cActiGiro"))
        txtDeudor.Text = rs("cliente")
        txtCAnalista.Text = rs("cAnalista")
        txtDireccion.Text = rs("cPersDireccDomicilio")
        txtMontoCMAC.Text = Format(rs("nSaldoCap"), "#,##0.00")
        cboTMonedaCMAC.Text = rs("cMoneda")
        txtFSDCmac.Text = mskPeriodo1Del
        lsCalificacion = Mid(rs("cCalifActual"), 1, 1)
        For i = 0 To 4
            If CStr(i) = lsCalificacion Then
                txtN(i).Text = "x"
                txtN(i).Enabled = True
            Else
            txtN(i).Enabled = False
            End If
        Next i
        CargarPosicionClienteSBS rs("cCodDoc")
        CargarGarantia (txtCodCta.Text)
    Else
        MsgBox lsmensaje, vbCritical, "Aviso"
    End If
    Set objCOMNAuditoria = Nothing
    Set rs = Nothing
End Sub

'NR MAVM 20090922
Private Sub CargarPosicionClienteSBS(ByVal sPersId As String)
    Dim rsCargarDatos As ADODB.Recordset
    Dim rsCargarValores As ADODB.Recordset
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set rsCargarDatos = objCOMNAuditoria.ObtenerPosicionClienteSBS(sPersId)
    If rsCargarDatos.RecordCount <> 0 Then
        If rsCargarDatos!nNormal <> 0 Then
            txtPNorm.Text = rsCargarDatos!nNormal
            If rsCargarDatos!nNormal = 100 Then
                txtSF(0).Text = "x"
            End If
        End If
        
        If rsCargarDatos!nPotencial <> 0 Then
            txtPCPP.Text = rsCargarDatos!nPotencial
            If rsCargarDatos!nPotencial = 100 Then
                txtSF(1).Text = "x"
            End If
        End If
        
        If rsCargarDatos!nDeficiente <> 0 Then
            txtPDefic.Text = rsCargarDatos!nDeficiente
            If rsCargarDatos!nDeficiente = 100 Then
                txtSF(2).Text = "x"
            End If
        End If
        
        If rsCargarDatos!nDudoso <> 0 Then
            txtPDud.Text = rsCargarDatos!nDudoso
            If rsCargarDatos!nDudoso = 100 Then
                txtSF(3).Text = "x"
            End If
        End If
        
        If rsCargarDatos!nPerdido <> 0 Then
            txtPPerd.Text = rsCargarDatos!nPerdido
            If rsCargarDatos!nPerdido = 100 Then
                txtSF(4).Text = "x"
            End If
        End If
        
        txtFSDSF.Text = Format(rsCargarDatos!Fec_Rep, "dd/mm/yyyy")
        Set rsCargarValores = objCOMNAuditoria.ObtenerEndeudamientoSBS(rsCargarDatos!Cod_Edu)
        txtMontoSF.Text = Format(rsCargarValores!DDirSoles + rsCargarValores!DDirDolar, "#,##0.00")
    End If
    Set objCOMNAuditoria = Nothing
End Sub

Public Sub CargarGarantia(ByVal lsCodCta As String)
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim rs As ADODB.Recordset
    Dim lsmensaje As String
    Dim lsCadena As String
    Dim i As Integer
    lsmensaje = ""
    Set rs = objCOMNAuditoria.ObtenerGarantia(lsCodCta, lsmensaje)
    If lsmensaje = "" Then
    For i = 0 To rs.RecordCount - 1
        If i = rs.RecordCount - 1 Then
            lsCadena = lsCadena & rs.Fields("cDescripcion") & " " & Format(rs.Fields("nGravado"), "#,##0.00") & " " & rs.Fields("cMoneda")
        Else
        lsCadena = lsCadena & rs.Fields("cDescripcion") & " " & Format(rs.Fields("nGravado"), "#,##0.00") & " " & rs.Fields("cMoneda") & ","
        End If
    rs.MoveNext
    Next i
    End If
    txtGarantia.Text = lsCadena
    Set objCOMNAuditoria = Nothing
End Sub

Private Sub txtFRegistro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtDeudor.SetFocus
    End If
End Sub

Private Sub txtFRegistro_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFRegistro.Text)
        If Not Trim(sCad) = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        If CDate(txtFRegistro.Text) > gdFecSis Then
            MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
            txtFRegistro.SetFocus
            Exit Sub
        End If
End Sub

Private Sub txtFSDCmac_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFSDCmac.Text)
        If Not Trim(sCad) = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        If CDate(txtFSDCmac.Text) >= gdFecSis Then
            MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
            txtFSDCmac.SetFocus
            Exit Sub
        End If
End Sub

Public Sub LimpiarControles(frmForm As Form, Optional cText As Boolean, Optional cFG As Boolean, Optional cCombo As Boolean, Optional cDataCombo As Boolean, Optional cLabel As Boolean)
      Dim ctlControl As Object
      On Error Resume Next
      For Each ctlControl In frmForm.Controls
        If cText = True Then
            If TypeOf ctlControl Is TextBox Then
                ctlControl.Text = ""
            End If
        End If
        If cCombo = True Then
            If TypeOf ctlControl Is ComboBox Then
                ctlControl.ListIndex = 0
            End If
        End If
        If cDataCombo = True Then
            If TypeOf ctlControl Is DataCombo Then
                ctlControl.Text = ""
            End If
        End If
        If cLabel = True Then
            If TypeOf ctlControl Is Label Then
                ctlControl.Caption = ""
                ctlControl.BorderStyle = 0
            End If
        End If
         DoEvents
      Next ctlControl
End Sub

Private Sub txtFSDSF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtMontoSF.SetFocus
    End If
End Sub

Private Sub txtFSDSF_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFSDSF.Text)
        If Not Trim(sCad) = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        If CDate(txtFSDSF.Text) > gdFecSis Then
            MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
            txtFSDSF.SetFocus
            Exit Sub
        End If
End Sub

Private Sub txtMontoCMAC_LostFocus()
    txtMontoCMAC.Text = Format(txtMontoCMAC.Text, "#,##0.00")
End Sub

Private Sub txtMontoSF_LostFocus()
    txtMontoSF.Text = Format(txtMontoSF.Text, "#,##0.00")
End Sub

Private Sub txtPNorm_LostFocus()
    ValidarPorcentaje
End Sub

Private Sub txtPCPP_LostFocus()
    ValidarPorcentaje
End Sub

Private Sub txtPDefic_LostFocus()
    ValidarPorcentaje
End Sub

Private Sub txtPDud_LostFocus()
    ValidarPorcentaje
End Sub

Private Sub txtPPerd_LostFocus()
    ValidarPorcentaje
End Sub

Private Sub ValidarPorcentaje()
    If val(txtPNorm.Text) + val(txtPCPP.Text) + val(txtPDefic.Text) + val(txtPDud.Text) + val(txtPPerd.Text) > 100 Then
        MsgBox "Solo Puede Sumar el 100%", vbCritical, "Aviso"
        txtPNorm.Text = "0"
        txtPCPP.Text = "0"
        txtPDefic.Text = "0"
        txtPDud.Text = "0"
        txtPPerd.Text = "0"
    End If
End Sub

Private Sub txtSF_LostFocus(Index As Integer)
    Dim i As Integer
    Dim Cont As Integer
    For i = 0 To 4
        If txtSF(i).Text = "x" Or txtSF(i).Text = "" Then
            If txtSF(i).Text <> "" Then
            Cont = Cont + 1
            End If
            If Cont > 1 Then
            MsgBox "Solo Puede elegir una Calificación", vbCritical, "Aviso"
            txtSF(i).Text = ""
                If Cont > 1 Then
                    txtOCI(0).SetFocus
                    Else
                    txtSF(i).SetFocus
                End If
            Exit Sub
            End If
        Else
        MsgBox "Solo Puede Ingresar la letra x", vbCritical, "Aviso"
        txtSF(i).Text = ""
        If Cont = 1 Then
                    txtOCI(0).SetFocus
                    Else
                    txtSF(i).SetFocus
                End If
        Exit Sub
        End If
        
    Next i
End Sub

Private Sub txtOCI_LostFocus(Index As Integer)
    Dim i As Integer
    Dim Cont As Integer
    For i = 0 To 4
    If txtOCI(i).Text = "x" Or txtOCI(i).Text = "" Then
            If txtOCI(i).Text <> "" Then
            Cont = Cont + 1
            End If
            If Cont > 1 Then
            MsgBox "Solo Puede elegir una Calificación", vbCritical, "Aviso"
            txtOCI(i).Text = ""
                If Cont = 1 Then
                    txtSituacion.SetFocus
                    Else
                    txtOCI(i).SetFocus
                End If
            Exit Sub
            End If
        Else
        MsgBox "Solo Puede Ingresar la letra x", vbCritical, "Aviso"
        txtOCI(i).Text = ""
        If Cont = 1 Then
                    txtSituacion.SetFocus
                    Else
                    txtOCI(i).SetFocus
                End If
        Exit Sub
        End If
    Next i
End Sub
