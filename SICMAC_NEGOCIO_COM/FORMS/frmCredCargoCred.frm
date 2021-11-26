VERSION 5.00
Begin VB.Form frmCredCargoCred 
   Caption         =   "Amortizar Credito con Cargo A Cuenta"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "frmCredCargoCred.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2955
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   420
      Left            =   5355
      TabIndex        =   4
      Top             =   2355
      Width           =   1260
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   420
      Left            =   4185
      TabIndex        =   3
      Top             =   2355
      Width           =   1155
   End
   Begin VB.CommandButton CmdCargar 
      Caption         =   "Cargar a Cta"
      Height          =   420
      Left            =   90
      TabIndex        =   2
      Top             =   2355
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   1650
      Left            =   45
      TabIndex        =   1
      Top             =   630
      Width           =   6570
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00800000&
         Height          =   270
         Left            =   4095
         TabIndex        =   14
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label8 
         Caption         =   "Monto Pago :"
         Height          =   270
         Left            =   2700
         TabIndex        =   13
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1335
         TabIndex        =   12
         Top             =   1110
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Prestamo :"
         Height          =   270
         Left            =   135
         TabIndex        =   11
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3660
         TabIndex        =   10
         Top             =   690
         Width           =   1410
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Vencimiento :"
         Height          =   270
         Left            =   2145
         TabIndex        =   9
         Top             =   720
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1320
         TabIndex        =   8
         Top             =   705
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Dias de Atraso :"
         Height          =   270
         Left            =   135
         TabIndex        =   7
         Top             =   735
         Width           =   1200
      End
      Begin VB.Label LblTitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   720
         TabIndex        =   6
         Top             =   285
         Width           =   5610
      End
      Begin VB.Label Label1 
         Caption         =   "Titular :"
         Height          =   270
         Left            =   105
         TabIndex        =   5
         Top             =   300
         Width           =   615
      End
   End
   Begin SICMACT.ActXCodCta CtaCred 
      Height          =   390
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   3750
      _extentx        =   6615
      _extenty        =   688
      texto           =   "Credito"
      enabledcmac     =   -1  'True
      enabledcta      =   -1  'True
      enabledprod     =   -1  'True
      enabledage      =   -1  'True
   End
End
Attribute VB_Name = "frmCredCargoCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

