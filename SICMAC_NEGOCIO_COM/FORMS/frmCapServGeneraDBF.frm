VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCapServGeneraDBF 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "frmCapServGeneraDBF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1035
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   855
      Left            =   3600
      Picture         =   "frmCapServGeneraDBF.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1035
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Fecha Cobranza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2835
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   3315
      Begin MSComCtl2.MonthView mvwFecha 
         Height          =   2370
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   187760641
         CurrentDate     =   37113
      End
   End
End
Attribute VB_Name = "frmCapServGeneraDBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerar_Click()
Dim dFechaCobro As Date
dFechaCobro = CDate(mvwFecha.value)
If MsgBox("¿Desea Procesar la Información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsServ As NCapServicios
    Set clsServ = New NCapServicios
    On Error GoTo ErrGenera
    clsServ.GeneraDBFSedalib dFechaCobro
    Set clsServ = Nothing
End If
Exit Sub
ErrGenera:
    MsgBox Err.Description, vbExclamation, "Error"
    Set clsServ = Nothing
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Captaciones - Servicios - Generación DBF"
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
mvwFecha.value = gdFecSis
End Sub
