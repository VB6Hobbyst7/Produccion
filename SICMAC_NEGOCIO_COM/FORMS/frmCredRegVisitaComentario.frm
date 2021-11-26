VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRegVisitaComentario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comentarios"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   Icon            =   "frmCredRegVisitaComentario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   3360
      Width           =   1335
   End
   Begin TabDlg.SSTab sstComentario 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Seleccione y escriba"
      TabPicture(0)   =   "frmCredRegVisitaComentario.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblTrimestreDesc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblAnalista"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblFecha"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblNVisita"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTrimestre"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtComentario"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.TextBox txtComentario 
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1680
         Width           =   4575
      End
      Begin VB.Label lblTrimestre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblNVisita 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3240
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblAnalista 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblTrimestreDesc 
         Caption         =   "Trimestre:"
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Comentarios:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Visita:"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Analista:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCredRegVisitaComentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fsComentario As String

Public Function Inicio(ByVal pnCodigo As Integer, ByVal psAnalista As String, ByVal pdFecha As String, ByVal pnNVisita As Integer, ByVal psComentario As String) As String
lblAnalista.Caption = psAnalista
lblFecha.Caption = pdFecha
lblNVisita.Caption = pnNVisita
Dim pnTrimestre As Integer

If pnCodigo = 1 Then
    lblTrimestreDesc.Visible = False
    lblTrimestre.Visible = False
    Me.Caption = "Comentarios Analista"
ElseIf pnCodigo = 2 Then
    lblTrimestreDesc.Visible = True
    lblTrimestre.Visible = True
    pnTrimestre = Month(pdFecha)
    Me.Caption = "Comentarios Jefe de Agencia"
    lblTrimestre.Caption = IIf(pnTrimestre > 9, "IV", IIf(pnTrimestre > 6, "III", IIf(pnTrimestre > 3, "II", "I"))) & "-" & Year(pdFecha)
End If

fsComentario = psComentario
txtComentario.Text = fsComentario
Me.Show 1

Inicio = fsComentario
End Function

Private Sub cmdAceptar_Click()
Unload Me
End Sub

