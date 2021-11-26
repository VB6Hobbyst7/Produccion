VERSION 5.00
Begin VB.Form frmPigConsultaDetPagos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle del Pago"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ControlBox      =   0   'False
   Icon            =   "frmPigConsultaDetPagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   360
      Left            =   2145
      TabIndex        =   10
      Top             =   3735
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Height          =   3660
      Left            =   60
      TabIndex        =   0
      Top             =   -15
      Width           =   4905
      Begin VB.Label lblCustDiferida 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2835
         TabIndex        =   19
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label lblComServicio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2835
         TabIndex        =   18
         Top             =   585
         Width           =   1845
      End
      Begin VB.Label lblPenalidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2835
         TabIndex        =   17
         Top             =   930
         Width           =   1845
      End
      Begin VB.Label lblDerRemate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2835
         TabIndex        =   16
         Top             =   1275
         Width           =   1845
      End
      Begin VB.Label lblComVcto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2835
         TabIndex        =   15
         Top             =   1620
         Width           =   1845
      End
      Begin VB.Label lblIntMoratorio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2835
         TabIndex        =   14
         Top             =   1965
         Width           =   1845
      End
      Begin VB.Label lblIntCompensatorio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2835
         TabIndex        =   13
         Top             =   2310
         Width           =   1845
      End
      Begin VB.Label lblCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2835
         TabIndex        =   12
         Top             =   2655
         Width           =   1845
      End
      Begin VB.Label Label9 
         Caption         =   "Custodia Diferida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   11
         Top             =   285
         Width           =   2595
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2835
         TabIndex        =   9
         Top             =   3225
         Width           =   1845
      End
      Begin VB.Label Label8 
         Caption         =   "Total Pagado              S/. "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   195
         TabIndex        =   8
         Top             =   3270
         Width           =   2595
      End
      Begin VB.Label Label7 
         Caption         =   "Comision de Servicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   7
         Top             =   630
         Width           =   2595
      End
      Begin VB.Label Label6 
         Caption         =   "Penalidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   6
         Top             =   975
         Width           =   2595
      End
      Begin VB.Label Label5 
         Caption         =   "Derecho de Remate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   5
         Top             =   1305
         Width           =   2595
      End
      Begin VB.Label Label4 
         Caption         =   "Comision Vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   4
         Top             =   1650
         Width           =   2595
      End
      Begin VB.Label Label3 
         Caption         =   "Interes Moratorio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   3
         Top             =   1995
         Width           =   2595
      End
      Begin VB.Label Label2 
         Caption         =   "Interes Compensatorio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   2
         Top             =   2325
         Width           =   2595
      End
      Begin VB.Label Label1 
         Caption         =   "Capital Amortizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   195
         TabIndex        =   1
         Top             =   2670
         Width           =   2595
      End
   End
End
Attribute VB_Name = "frmPigConsultaDetPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Inicio(ByVal pnTotal As Currency, Optional ByVal pnCapital As Currency = 0, _
                        Optional ByVal pnIntCompensatorio As Currency = 0, Optional ByVal pnIntMoratorio As Currency = 0, _
                        Optional ByVal pnComVcto As Currency = 0, Optional ByVal pnDerRemate As Currency = 0, _
                        Optional ByVal pnPenalidad As Currency = 0, Optional ByVal pnComServicio As Currency = 0, _
                        Optional ByVal pnCustDif As Currency = 0)
     Limpiar
     lblTotal.Caption = Format(pnTotal, "###,##0.00")
     lblCapital.Caption = Format(pnCapital, "###,##0.00")
     lblIntCompensatorio.Caption = Format(pnIntCompensatorio, "###,##0.00")
     lblIntMoratorio.Caption = Format(pnIntMoratorio, "###,##0.00")
     lblComVcto.Caption = Format(pnComVcto, "###,##0.00")
     lblDerRemate.Caption = Format(pnDerRemate, "###,##0.00")
     lblPenalidad.Caption = Format(pnPenalidad, "###,##0.00")
     lblComServicio.Caption = Format(pnComServicio, "###,##0.00")
     lblCustDiferida.Caption = Format(pnCustDif, "###,##0.00")
     Me.Show 1
End Sub

Public Sub Limpiar()
     lblTotal.Caption = Format(0, "###,##0.00")
     lblCapital.Caption = Format(0, "###,##0.00")
     lblIntCompensatorio.Caption = Format(0, "###,##0.00")
     lblIntMoratorio.Caption = Format(0, "###,##0.00")
     lblComVcto.Caption = Format(0, "###,##0.00")
     lblDerRemate.Caption = Format(0, "###,##0.00")
     lblPenalidad.Caption = Format(0, "###,##0.00")
     lblComServicio.Caption = Format(0, "###,##0.00")
     lblCapital.Caption = Format(0, "###,##0.00")
     lblCustDiferida.Caption = Format(0, "###,##0.00")
End Sub

Private Sub cmdOk_Click()
    Unload Me
End Sub
