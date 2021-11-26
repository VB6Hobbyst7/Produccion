VERSION 5.00
Begin VB.Form frmConfirmaRetiro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6135
   ClientLeft      =   1290
   ClientTop       =   1800
   ClientWidth     =   8310
   Icon            =   "frmConfirmaRetiro.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frageneral 
      Caption         =   "Datos de Retiro"
      Height          =   2280
      Left            =   210
      TabIndex        =   1
      Top             =   3630
      Width           =   7800
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   390
         Left            =   4710
         TabIndex        =   14
         Top             =   1770
         Width           =   1485
      End
      Begin Sicmact.TxtBuscar TxtBuscar1 
         Height          =   360
         Left            =   4785
         TabIndex        =   13
         Top             =   225
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
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
         sTitulo         =   ""
      End
      Begin VB.PictureBox Picture1 
         Height          =   1890
         Left            =   120
         ScaleHeight     =   1830
         ScaleWidth      =   3765
         TabIndex        =   3
         Top             =   255
         Width           =   3825
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
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
            Height          =   315
            Left            =   2400
            TabIndex        =   11
            Top             =   1425
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Monto :"
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
            Index           =   2
            Left            =   1695
            TabIndex        =   10
            Top             =   1470
            Width           =   660
         End
         Begin VB.Label lblcuenta 
            Caption         =   "Cuenta"
            ForeColor       =   &H00004080&
            Height          =   225
            Left            =   270
            TabIndex        =   9
            Top             =   975
            Width           =   3465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuenta :"
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
            Index           =   1
            Left            =   60
            TabIndex        =   8
            Top             =   735
            Width           =   735
         End
         Begin VB.Label lblBanco 
            AutoSize        =   -1  'True
            Caption         =   "Banco :"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   270
            TabIndex        =   7
            Top             =   510
            Width           =   3315
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Banco :"
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
            Left            =   60
            TabIndex        =   6
            Top             =   270
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha :"
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
            Index           =   0
            Left            =   1890
            TabIndex        =   5
            Top             =   45
            Width           =   660
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            Caption         =   "dd/mm/yyyy"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2625
            TabIndex        =   4
            Top             =   45
            Width           =   1035
         End
      End
      Begin Sicmact.TxtBuscar TxtBuscar2 
         Height          =   360
         Left            =   4845
         TabIndex        =   15
         Top             =   1005
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   635
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
         sTitulo         =   ""
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Destino:"
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
         Height          =   195
         Left            =   4095
         TabIndex        =   12
         Top             =   1065
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Origen :"
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
         Height          =   195
         Left            =   4080
         TabIndex        =   2
         Top             =   285
         Width           =   690
      End
   End
   Begin Sicmact.FlexEdit fgRetiros 
      Height          =   3420
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   6033
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorControl=   -2147483643
      BackColorControl=   -2147483628
      BackColorControl=   -2147483643
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   -1
      RowHeight0      =   240
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmConfirmaRetiro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
