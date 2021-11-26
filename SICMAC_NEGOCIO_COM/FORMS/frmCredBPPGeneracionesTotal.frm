VERSION 5.00
Begin VB.Form frmCredBPPGeneracionesTotal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Generar Bonos del Mes"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   Icon            =   "frmCredBPPGeneracionesTotal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConfigGeneral 
      Caption         =   "Generar Bonos"
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
      Top             =   120
      Width           =   5175
      Begin VB.CommandButton cmdConsulta 
         Caption         =   "Consultar BPP Generados"
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
         TabIndex        =   5
         Top             =   2280
         Width           =   2250
      End
      Begin VB.CommandButton cmdJefesAgencia 
         Caption         =   "Bono Coordinadores y Jefes de Agencia"
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
         TabIndex        =   4
         Top             =   360
         Width           =   2250
      End
      Begin VB.CommandButton cmdAnalista 
         Caption         =   "Bono Analista"
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
         Top             =   360
         Width           =   2250
      End
      Begin VB.CommandButton cmdJefesTerritorial 
         Caption         =   "Bono Jefes Territoriales"
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
         TabIndex        =   2
         Top             =   1320
         Width           =   2250
      End
      Begin VB.CommandButton cmdCierre 
         Caption         =   "Cierre de los Bonos"
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
         TabIndex        =   1
         Top             =   1320
         Width           =   2250
      End
   End
End
Attribute VB_Name = "frmCredBPPGeneracionesTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private Sub cmdAnalista_Click()
'    frmCredBPPGenBonoMensual.Show 1
'End Sub
'
'Private Sub cmdCierre_Click()
'Dim fgFecActual As Date
'Dim oConsSist As COMDConstSistema.NCOMConstSistema
'Set oConsSist = New COMDConstSistema.NCOMConstSistema
'fgFecActual = oConsSist.LeeConstSistema(gConstSistFechaBPP)
'Set oConsSist = Nothing
'
'    Call frmCredBPPBonoCierres.Inicio(Month(fgFecActual), Year(fgFecActual), "BPP - Realizar Cierres de Bonos Generados", True)
'End Sub
'
'Private Sub cmdConsulta_Click()
'    frmCredBPPBonoConsulta.Show 1
'End Sub
'
'Private Sub cmdJefesAgencia_Click()
'    frmCredBPPBonoCoordJA.Show 1
'End Sub
'
'Private Sub cmdJefesTerritorial_Click()
'
'     frmCredBPPBonoJT.Show 1
'End Sub
