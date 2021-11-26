VERSION 5.00
Begin VB.Form frmCredBPPConfigMensual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BPP - Configuracion Mensual"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   Icon            =   "frmCredBPPConfigMensual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Configuración Mensual"
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
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin VB.CommandButton cmdConfJefesAGYT 
         Caption         =   "Configuración de Jefes de Agencia y Jefes Territoriales."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4320
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdTopesMetasCJ 
         Caption         =   "Topes y Metas de Coordinador, Jefe de Agencia y Jefe Territoriales."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6360
         TabIndex        =   10
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdPesosTopesMen 
         Caption         =   "Pesos y Topes Mensuales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   6360
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdParamCumpGere 
         Caption         =   "Parametros de Cumplimiento - Gerencia"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdRangoMoraPerm 
         Caption         =   "Rango de Mora Permitidos y Factor de Rend"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdPorcCumpMin 
         Caption         =   "Porcentaje de Cumplimiento Minimo - GM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4320
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdPenMora 
         Caption         =   "Penalidad por Mora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdParamBonoOpe 
         Caption         =   "Parametros Bono Plus y Operaciones"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2280
         TabIndex        =   3
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   2
         Top             =   2880
         Width           =   3855
      End
      Begin VB.CommandButton cmdMigrarDatosMesAnt 
         Caption         =   "Migrar Datos Mes Anterior"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   2880
         Width           =   3855
      End
   End
   Begin VB.Label lblOpcionesConfig 
      AutoSize        =   -1  'True
      Caption         =   "Opciones de Configuración BPP"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2280
   End
End
Attribute VB_Name = "frmCredBPPConfigMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmdCerrar_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdConfJefesAGYT_Click()
'    frmCredBPPConfJefesAgeYT.Show 1
'End Sub
'
'Private Sub cmdMigrarDatosMesAnt_Click()
'    frmCredBPPMigrarDatos.Show 1
'End Sub
'
'Private Sub cmdParamBonoOpe_Click()
'    frmCredBPPPlusOpe.Show 1
'End Sub
'
'Private Sub cmdParamCumpGere_Click()
'    frmCredBPPParamCumGeren.inicia 1
'End Sub
'
'Private Sub cmdPenMora_Click()
'    frmCredBPPPenalIncreMora.Show 1
'End Sub
'
'Private Sub cmdPesosTopesMen_Click()
'    frmCredBPPPesosTopesMensuales.Show 1
'End Sub
'
'Private Sub cmdPorcCumpMin_Click()
'    frmCredBPPParamCumGeren.inicia 2
'End Sub
'
'Private Sub cmdRangoMoraPerm_Click()
'    frmCredBPPRangoMoraPerm.Show 1
'End Sub
'
'Private Sub cmdTopesMetasCJ_Click()
'    frmCredBPPTopeMetaCoordJefes.Show 1
'End Sub
