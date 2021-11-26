VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmColocEvalConsultaCailf 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   Icon            =   "FrmColocEvalConsultaCailf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   6015
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1845
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   3254
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Evaluacion"
      TabPicture(0)   =   "FrmColocEvalConsultaCailf.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TxtComentario"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Calificaiones"
      TabPicture(1)   =   "FrmColocEvalConsultaCailf.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblEval"
      Tab(1).Control(1)=   "LblCmac"
      Tab(1).Control(2)=   "LblSist"
      Tab(1).Control(3)=   "LblDiaAtr"
      Tab(1).Control(4)=   "Label9"
      Tab(1).Control(5)=   "Label8"
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(7)=   "Label6"
      Tab(1).ControlCount=   8
      Begin VB.TextBox TxtComentario 
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label LblEval 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -69960
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label LblCmac 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -69960
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
      Begin VB.Label LblSist 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -72840
         TabIndex        =   16
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label LblDiaAtr 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -72840
         TabIndex        =   15
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Calif x CMAC"
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
         Left            =   -71400
         TabIndex        =   14
         Top             =   720
         Width           =   1110
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Calif x Evaluac."
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
         Left            =   -71400
         TabIndex        =   13
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Calif x Sistem. Finac."
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
         Left            =   -74760
         TabIndex        =   12
         Top             =   1320
         Width           =   1800
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Calif  x Dias Atraso"
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
         Left            =   -74760
         TabIndex        =   11
         Top             =   720
         Width           =   1635
      End
   End
   Begin VB.TextBox TxtRuc 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox TxtCal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3960
      TabIndex        =   6
      Text            =   "1"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox TxtDni 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox TxtCliente 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "DNI"
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
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "RUC"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CALIF GENERAL"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "FrmColocEvalConsultaCailf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdsalir_Click()
Unload Me
End Sub

Public Sub Inicio(ByVal psCodCta As String)
Dim Pers As UPersona
Dim rs As ADODB.Recordset
Set Pers = New UPersona


FrmColocEvalConsultaCailf.Show vbModal
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
