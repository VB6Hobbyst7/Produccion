VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAudCierraCalificacion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Calificacion"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmAudCierraCalificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3735
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   3495
         Begin MSComCtl2.Animation Animacion 
            Height          =   495
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            _Version        =   393216
            AutoPlay        =   -1  'True
            FullWidth       =   57
            FullHeight      =   33
         End
      End
      Begin VB.CommandButton cmdTranferir 
         Caption         =   "&Transferir"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label lblFechaCalif 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   195
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Cierre Calif"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label lblFechaData 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label2"
         Height          =   195
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Data :"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   375
      Left            =   120
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   120
      Picture         =   "frmAudCierraCalificacion.frx":08CA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Este proceso puede durar algunos minutos"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Aviso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   600
   End
End
Attribute VB_Name = "frmAudCierraCalificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sServer As String
Dim FechaCalif  As String
Dim sServConsol As String

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdTranferir_Click()
Dim rs As New ADODB.Recordset
Dim sql As String
Dim I As Integer
Dim opt As Integer
Dim Rcc As COMDCredito.DCOMColocEval
I = DateDiff("d", FechaCalif, lblFechaData)
If I <= 0 Then
    MsgBox "Ya se Realizo la Tranferencia...", vbInformation, "AVISO"
    Exit Sub
End If

opt = MsgBox("Esta seguro de hacer la Transferencia de Data?", vbQuestion + vbYesNo, "AVISO")
If vbNo = opt Then Exit Sub

Frame2.Visible = True
'FILECOPY.AVI
'Animacion.Open (App.path & "\Videos\Grabando.AVI")
Set Rcc = New COMDCredito.DCOMColocEval

If Rcc.VerificaDataMigradaFecha(sServConsol, gdFecData) > 0 Then
    MsgBox "La Data ya fue Transferida", vbInformation, "AVISO"
    Set Rcc = Nothing
    Exit Sub
End If

' Trannferencia de Data
Call Rcc.InsertaDataColocCalifProvTotal(sServConsol, gdFecData)

'Animacion.AutoPlay = True
Frame2.Visible = False
Rcc.ActulizaFechaCierre (gdFecData)
MsgBox "Transferencia satisfactoria", vbInformation, "AVISO"

Set rs = Nothing
Set Rcc = Nothing
End Sub

Private Sub Form_Load()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim Fecha As Date
Dim Riesgo As COMDCredito.DCOMColocEval
Dim Rcc As COMDCredito.DCOMColocEval

Set Riesgo = New COMDCredito.DCOMColocEval
Set Rcc = New COMDCredito.DCOMColocEval

lblFechaData = gdFecData
'Animacion.Open (App.path & "\Videos\Grabando.AVI")

sServConsol = Rcc.ServConsol(gConstSistServCentralRiesgos)


'FILECOPY.AVI
'AbreConexion
FechaCalif = ""
'Sql = "SELECT * FROM VARSISTEMA WHERE cNomVar = 'dCieCalifica'"
'Set Rs = Conexion(Sql, "11207")

'If Rs.EOF And Rs.BOF Then
'Else
FechaCalif = Riesgo.DiaCierreCalifMes(sServConsol)
Me.lblFechaCalif = FechaCalif
'End If
'CierraConeccion

Set rs = Nothing
Set Rcc = Nothing
Set Riesgo = Nothing
End Sub
