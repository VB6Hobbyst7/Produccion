VERSION 5.00
Begin VB.Form frmCredBPPConfGen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Configuración General"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   Icon            =   "frmCredBPPConfGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConfigGeneral 
      Caption         =   "Configuración General"
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
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5055
      Begin VB.CommandButton cmdAgencias 
         Caption         =   "Agencias (Zonas y Cant. de Comités)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   2250
      End
      Begin VB.CommandButton cmdCoordJef 
         Caption         =   "Coordinadores, Jefes de Agencia y Territoriales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2640
         TabIndex        =   5
         Top             =   1320
         Width           =   2250
      End
      Begin VB.CommandButton cmdTpoCredNiveles 
         Caption         =   "Tipo Créditos X Niveles"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   2250
      End
      Begin VB.CommandButton cmdMoraBase 
         Caption         =   "Mora Base"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2250
      End
      Begin VB.CommandButton cmdCatAnalista 
         Caption         =   "Categorías de Analista"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   2250
      End
   End
   Begin VB.Label lblOpcionesConfig 
      AutoSize        =   -1  'True
      Caption         =   "Opciones de Configuración BPP"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2280
   End
End
Attribute VB_Name = "frmCredBPPConfGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************************
''***     Rutina         :   frmCredBPPConfGen
''***     Descripcion    :   Configuración General del BPP
''***     Creado por     :   WIOR
''***     Maquina        :   TIF-1-19
''***     Fecha-Creación :   24/05/2013 08:20:00 AM
''*****************************************************************************************
'Option Explicit
'
'Private Sub cmdAgencias_Click()
'    frmCredBPPConfigAgencias.Show 1
'End Sub
'
'Private Sub cmdCatAnalista_Click()
'    frmCredBPPCatAnalista.Show 1
'End Sub
'
'Private Sub cmdCoordJef_Click()
'    frmCredBPPConfigCoordJef.Show 1
'End Sub
'
'Private Sub cmdMoraBase_Click()
'    frmCredBPPMoraBase.Show 1
'End Sub
'
'Private Sub cmdTpoCredNiveles_Click()
'    frmCredBPPTpoCredNiveles.Inicio 2
'End Sub
