VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredNewNivAutorizaConfEscalona 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Escalonamiento"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmCredNewNivAutorizaConfEscalona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2400
      TabIndex        =   5
      ToolTipText     =   "Aceptar"
      Top             =   5120
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3480
      TabIndex        =   6
      ToolTipText     =   "Cancelar"
      Top             =   5120
      Width           =   1000
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Escalonamiento"
      TabPicture(0)   =   "frmCredNewNivAutorizaConfEscalona.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblMoraPorcentaje"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label9"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label10"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label12"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label14"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label15"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label17"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label18"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label19"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label20"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label21"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label22"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label23"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label24"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label25"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label26"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label27"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label28"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label29"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label30"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label31"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label32"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label33"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtMoraPorcentaje"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtMonCuoMoraMayorA"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtMonCuoMoraMenorIgual"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtMonCreMoraMayorA"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtMonCreMoraMenorIgual"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      Begin VB.TextBox txtMonCreMoraMenorIgual 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5265
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   3600
         Width           =   780
      End
      Begin VB.TextBox txtMonCreMoraMayorA 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2925
         MaxLength       =   15
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   3585
         Width           =   780
      End
      Begin VB.TextBox txtMonCuoMoraMenorIgual 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   5280
         MaxLength       =   15
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   2115
         Width           =   780
      End
      Begin VB.TextBox txtMonCuoMoraMayorA 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   2115
         Width           =   780
      End
      Begin VB.TextBox txtMoraPorcentaje 
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   3000
         MaxLength       =   15
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6150
         TabIndex        =   41
         Top             =   3645
         Width           =   165
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hasta"
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
         Height          =   195
         Left            =   4560
         TabIndex        =   40
         Top             =   3645
         Width           =   525
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "adicional de la exposición máxima anterior del cliente otorgado hasta la fecha."
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
         Height          =   870
         Left            =   4320
         TabIndex        =   39
         Top             =   3885
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3810
         TabIndex        =   38
         Top             =   3630
         Width           =   165
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hasta"
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
         Height          =   195
         Left            =   2220
         TabIndex        =   37
         Top             =   3630
         Width           =   525
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "adicional de la exposición máxima anterior del cliente otorgado hasta la fecha."
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
         Height          =   870
         Left            =   1980
         TabIndex        =   36
         Top             =   3870
         Width           =   2175
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Monto del Crédito"
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
         Height          =   435
         Left            =   360
         TabIndex        =   35
         Top             =   3840
         Width           =   1305
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   4245
         TabIndex        =   34
         Top             =   3390
         Width           =   2355
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   1860
         TabIndex        =   33
         Top             =   3390
         Width           =   2355
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   195
         TabIndex        =   32
         Top             =   3390
         Width           =   1635
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Caption         =   "% de escalonamiento"
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
         Height          =   195
         Left            =   2175
         TabIndex        =   31
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Caption         =   "% de escalonamiento"
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
         Height          =   195
         Left            =   4455
         TabIndex        =   30
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   6165
         TabIndex        =   29
         Top             =   2160
         Width           =   165
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hasta el"
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
         Height          =   195
         Left            =   4455
         TabIndex        =   28
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "sobre la cuota o cuotas máxima pagadas satisfactoriamente del calendario aprobado"
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
         Left            =   4440
         TabIndex        =   27
         Top             =   2400
         Width           =   2085
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3765
         TabIndex        =   26
         Top             =   2160
         Width           =   165
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hasta el"
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
         Height          =   195
         Left            =   2055
         TabIndex        =   25
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "sobre la cuota o cuotas máxima pagadas satisfactoriamente del calendario aprobado"
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
         Left            =   2040
         TabIndex        =   24
         Top             =   2400
         Width           =   2085
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Monto de la cuota a pagar"
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
         Height          =   435
         Left            =   480
         TabIndex        =   23
         Top             =   2400
         Width           =   1245
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   4245
         TabIndex        =   22
         Top             =   1890
         Width           =   2355
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   1860
         TabIndex        =   21
         Top             =   1890
         Width           =   2355
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   195
         TabIndex        =   20
         Top             =   1890
         Width           =   1635
      End
      Begin VB.Label Label9 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   4245
         TabIndex        =   19
         Top             =   1245
         Width           =   2355
      End
      Begin VB.Label lblMoraPorcentaje 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Caption         =   "0.00 %"
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
         Height          =   195
         Left            =   5745
         TabIndex        =   18
         Top             =   870
         Width           =   585
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Caption         =   "menor o igual"
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
         Height          =   195
         Left            =   4515
         TabIndex        =   17
         Top             =   870
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Caption         =   "Agencias con Mora"
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
         Height          =   195
         Left            =   4530
         TabIndex        =   16
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label6 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   4245
         TabIndex        =   15
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   3840
         TabIndex        =   14
         Top             =   870
         Width           =   150
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Caption         =   "mayor al "
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
         Height          =   195
         Left            =   2145
         TabIndex        =   13
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Caption         =   "Agencias con Mora"
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
         Height          =   195
         Left            =   2160
         TabIndex        =   12
         Top             =   600
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   1860
         TabIndex        =   11
         Top             =   1245
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Left            =   1860
         TabIndex        =   10
         Top             =   480
         Width           =   2355
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00F0FFFF&
         Caption         =   "Criterio"
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
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   1440
         Width           =   645
      End
      Begin VB.Label Label16 
         BackColor       =   &H00F0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   200
         TabIndex        =   8
         Top             =   1240
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmCredNewNivAutorizaConfEscalona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************************
'** Nombre : frmCredNewNivAutorizaConfEscalona
'** Descripción : Formulario para configurar autorizaciones de Escalonamiento de créditos según ERS002-2016
'** Creación : EJVG, 20160205 09:18:00 AM
'**********************************************************************************************************
Option Explicit
Dim fvEscalonaConf As TEscalonamientoConf

Dim fbAceptar As Boolean

Private Sub cmdAceptar_Click()
    If Not ValidaDatos Then Exit Sub
    
    fvEscalonaConf.nMoraPorcentaje = txtMoraPorcentaje.Text
    fvEscalonaConf.nMontoCuoMayorA = txtMonCuoMoraMayorA.Text
    fvEscalonaConf.nMontoCuoMenorIgual = txtMonCuoMoraMenorIgual.Text
    fvEscalonaConf.nMontoCreMayorA = txtMonCreMoraMayorA.Text
    fvEscalonaConf.nMontoCreMenorIgual = txtMonCreMoraMenorIgual.Text
    
    fbAceptar = True
    Unload Me
End Sub
Private Function ValidaDatos() As Boolean
    If Not IsNumeric(txtMoraPorcentaje.Text) Then
       MsgBox "Ud. debe ingresar el porcentaje de la Mora.", vbInformation, "Aviso"
       EnfocaControl txtMoraPorcentaje
       Exit Function
    Else
        If CDbl(txtMoraPorcentaje.Text) <= 0 Or CDbl(txtMoraPorcentaje.Text) > 100 Then
            MsgBox "Ud. debe ingresar el porcentaje de la mora a validar." & Chr(13) & "Debe ser mayor a cero (0.00) y menor o igual a cien (100.00).", vbInformation, "Aviso"
            EnfocaControl txtMoraPorcentaje
            Exit Function
        End If
    End If
    If Not IsNumeric(txtMonCuoMoraMayorA.Text) Then
       MsgBox "Ud. debe ingresar el porcentaje máximo sobre la cuota si la mora es mayor a " & Format(txtMoraPorcentaje.Text, "#,##0.00") & " %.", vbInformation, "Aviso"
       EnfocaControl txtMonCuoMoraMayorA
       Exit Function
    Else
        If CDbl(txtMonCuoMoraMayorA.Text) <= 0 Or CDbl(txtMonCuoMoraMayorA.Text) > 100 Then
            MsgBox "Ud. debe ingresar el porcentaje máximo sobre la cuota si la mora es mayor a " & Format(txtMoraPorcentaje.Text, "#,##0.00") & " %." & Chr(13) & "Debe ser mayor a cero (0.00) y menor o igual a cien (100.00).", vbInformation, "Aviso"
            EnfocaControl txtMonCuoMoraMayorA
            Exit Function
        End If
    End If
    If Not IsNumeric(txtMonCuoMoraMenorIgual.Text) Then
       MsgBox "Ud. debe ingresar el porcentaje máximo sobre la cuota si la mora es menor o igual a " & Format(txtMoraPorcentaje.Text, "#,##0.00") & " %.", vbInformation, "Aviso"
       EnfocaControl txtMonCuoMoraMenorIgual
       Exit Function
    Else
        If CDbl(txtMonCuoMoraMenorIgual.Text) <= 0 Or CDbl(txtMonCuoMoraMenorIgual.Text) > 100 Then
            MsgBox "Ud. debe ingresar el porcentaje máximo sobre la cuota si la mora es menor o igual a " & Format(txtMoraPorcentaje.Text, "#,##0.00") & " %." & Chr(13) & "Debe ser mayor a cero (0.00) y menor o igual a cien (100.00).", vbInformation, "Aviso"
            EnfocaControl txtMonCuoMoraMenorIgual
            Exit Function
        End If
    End If
    If Not IsNumeric(txtMonCreMoraMayorA.Text) Then
       MsgBox "Ud. debe ingresar el porcentaje máximo sobre la exposición anterior si la mora es mayor a " & Format(txtMoraPorcentaje.Text, "#,##0.00") & " %.", vbInformation, "Aviso"
       EnfocaControl txtMonCreMoraMayorA
       Exit Function
    Else
        If CDbl(txtMonCreMoraMayorA.Text) <= 0 Or CDbl(txtMonCreMoraMayorA.Text) > 100 Then
            MsgBox "Ud. debe ingresar el porcentaje máximo sobre la exposición anterior si la mora es mayor a " & Format(txtMoraPorcentaje.Text, "#,##0.00") & " %." & Chr(13) & "Debe ser mayor a cero (0.00) y menor o igual a cien (100.00).", vbInformation, "Aviso"
            EnfocaControl txtMonCreMoraMayorA
            Exit Function
        End If
    End If
    If Not IsNumeric(txtMonCreMoraMenorIgual.Text) Then
       MsgBox "Ud. debe ingresar el porcentaje máximo sobre la exposición anterior si la mora es menor o igual a " & Format(txtMoraPorcentaje.Text, "#,##0.00") & " %.", vbInformation, "Aviso"
       EnfocaControl txtMonCreMoraMenorIgual
       Exit Function
    Else
        If CDbl(txtMonCreMoraMenorIgual.Text) <= 0 Or CDbl(txtMonCreMoraMenorIgual.Text) > 100 Then
            MsgBox "Ud. debe ingresar el porcentaje máximo sobre la exposición anterior si la mora es menor o igual a " & Format(txtMoraPorcentaje.Text, "#,##0.00") & " %." & Chr(13) & "Debe ser mayor a cero (0.00) y menor o igual a cien (100.00).", vbInformation, "Aviso"
            EnfocaControl txtMonCreMoraMenorIgual
            Exit Function
        End If
    End If
    
    ValidaDatos = True
End Function
Private Sub Form_Load()
    fbAceptar = False
    
    txtMoraPorcentaje.Text = Format(fvEscalonaConf.nMoraPorcentaje, "#,##0.00")
    txtMonCuoMoraMayorA.Text = Format(fvEscalonaConf.nMontoCuoMayorA, "#,##0.00")
    txtMonCuoMoraMenorIgual.Text = Format(fvEscalonaConf.nMontoCuoMenorIgual, "#,##0.00")
    txtMonCreMoraMayorA.Text = Format(fvEscalonaConf.nMontoCreMayorA, "#,##0.00")
    txtMonCreMoraMenorIgual.Text = Format(fvEscalonaConf.nMontoCreMenorIgual, "#,##0.00")
    cmdAceptar.Enabled = False
End Sub
Public Function Inicio(ByRef pvEscalonaConf As TEscalonamientoConf) As Boolean
    fvEscalonaConf = pvEscalonaConf
    Show 1
    pvEscalonaConf = fvEscalonaConf
    Inicio = fbAceptar
End Function
Private Sub cmdCancelar_Click()
    fbAceptar = False
    Unload Me
End Sub
Private Sub txtMonCreMoraMayorA_Change()
    cmdAceptar.Enabled = True
End Sub
Private Sub txtMonCreMoraMayorA_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonCreMoraMayorA, KeyAscii, 6)
    
    If KeyAscii = 13 Then
        EnfocaControl txtMonCreMoraMenorIgual
    End If
End Sub
Private Sub txtMonCreMoraMayorA_LostFocus()
    txtMonCreMoraMayorA.Text = Format(txtMonCreMoraMayorA.Text, "#,##0.00")
End Sub
Private Sub txtMonCreMoraMenorIgual_Change()
    cmdAceptar.Enabled = True
End Sub
Private Sub txtMonCreMoraMenorIgual_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonCreMoraMenorIgual, KeyAscii, 6)
    
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
Private Sub txtMonCreMoraMenorIgual_LostFocus()
    txtMonCreMoraMenorIgual.Text = Format(txtMonCreMoraMenorIgual.Text, "#,##0.00")
End Sub
Private Sub txtMonCuoMoraMayorA_Change()
    cmdAceptar.Enabled = True
End Sub
Private Sub txtMonCuoMoraMayorA_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonCuoMoraMayorA, KeyAscii, 6)
    
    If KeyAscii = 13 Then
        EnfocaControl txtMonCuoMoraMenorIgual
    End If
End Sub
Private Sub txtMonCuoMoraMayorA_LostFocus()
    txtMonCuoMoraMayorA.Text = Format(txtMonCuoMoraMayorA.Text, "#,##0.00")
End Sub
Private Sub txtMonCuoMoraMenorIgual_Change()
    cmdAceptar.Enabled = True
End Sub
Private Sub txtMonCuoMoraMenorIgual_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonCuoMoraMenorIgual, KeyAscii, 6)
    
    If KeyAscii = 13 Then
        EnfocaControl txtMonCreMoraMayorA
    End If
End Sub
Private Sub txtMonCuoMoraMenorIgual_LostFocus()
    txtMonCuoMoraMenorIgual.Text = Format(txtMonCuoMoraMenorIgual.Text, "#,##0.00")
End Sub
Private Sub txtMoraPorcentaje_Change()
    cmdAceptar.Enabled = True
    lblMoraPorcentaje.Caption = Format(txtMoraPorcentaje.Text, "#,##0.00") & " %"
End Sub
Private Sub txtMoraPorcentaje_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMoraPorcentaje, KeyAscii, 6)
    
    If KeyAscii = 13 Then
        EnfocaControl txtMonCuoMoraMayorA
    End If
End Sub
Private Sub txtMoraPorcentaje_LostFocus()
    txtMoraPorcentaje.Text = Format(txtMoraPorcentaje.Text, "#,##0.00")
End Sub
