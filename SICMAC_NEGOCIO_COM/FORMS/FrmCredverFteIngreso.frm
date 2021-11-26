VERSION 5.00
Begin VB.Form FrmCredverFteIngreso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Visualización de la Fuentes de Ingreso"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Fuente de Ingreso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6405
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   315
         Left            =   1740
         TabIndex        =   9
         Top             =   2220
         Width           =   1275
      End
      Begin VB.TextBox txtSaldo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         TabIndex        =   8
         Top             =   1740
         Width           =   1605
      End
      Begin VB.TextBox txtRazSoc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         TabIndex        =   6
         Top             =   1320
         Width           =   1605
      End
      Begin VB.TextBox txtFechaInicio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         TabIndex        =   4
         Top             =   900
         Width           =   1605
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1710
         TabIndex        =   2
         Top             =   450
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1065
         TabIndex        =   7
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Des Raz.Soc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   420
         TabIndex        =   5
         Top             =   1350
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   930
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   870
         TabIndex        =   1
         Top             =   450
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmCredverFteIngreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Inicia(ByVal pcNumFte As String)
    Dim oVisualizacion As COMNCredito.NCOMVisualizacion 'DVisualizacion
    Dim rs As ADODB.Recordset
    
    Set oVisualizacion = New COMNCredito.NCOMVisualizacion ' DVisualizacion
    Set rs = oVisualizacion.VerFteIngreso(pcNumFte)
    Set oVisualizacion = Nothing
    
     If Not rs.EOF And Not rs.BOF Then
        txtNombre = rs!cPersNombre
        txtFechaInicio = rs!dPersFIInicio
        txtRazSoc = rs!cRazSocDescrip
        txtSaldo = Format(rs!nSaldo, "#0.00")
     End If
     Set rs = Nothing
     Me.Show vbModal
End Sub


Private Sub CmdAceptar_Click()
    Unload Me
End Sub


