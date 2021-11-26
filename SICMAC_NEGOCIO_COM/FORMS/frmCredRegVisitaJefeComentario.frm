VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRegVisitaJefeComentario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comentarios"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   Icon            =   "frmCredRegVisitaJefeComentario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
   End
   Begin TabDlg.SSTab sstComentario 
      Height          =   3135
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Seleccione y escriba"
      TabPicture(0)   =   "frmCredRegVisitaJefeComentario.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtFecVisita"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtComentario"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtNVisita"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtNVisita 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   600
         Width           =   1530
      End
      Begin VB.TextBox txtComentario 
         Height          =   1335
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1680
         Width           =   4575
      End
      Begin MSMask.MaskEdBox txtFecVisita 
         Height          =   300
         Left            =   1080
         TabIndex        =   1
         Top             =   960
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "DD/MM/YYYY"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Comentarios:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Nº Visita:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCredRegVisitaJefeComentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private fbRealizo As Boolean
Private fnCodigo As Long
Private fgFecVisita As Date

Public Function Inicio(ByVal pnCodigo As Long, ByVal pdFecha As String, ByVal pnNVisita As Integer) As Boolean
fgFecVisita = CDate(pdFecha)
fnCodigo = pnCodigo
fbRealizo = False
Me.Show 1

Inicio = fbRealizo
End Function

Private Sub cmdAceptar_Click()
If ValidaDatos Then
    If MsgBox("Estas seguro de guardar los Datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Dim oCredito As COMDCredito.DCOMCredito
        Set oCredito = New COMDCredito.DCOMCredito
        
        Call oCredito.RegistrarVisitaJefe(fnCodigo, CDate(txtFecVisita.Text), gsCodUser, CLng(txtNVisita.Text), Trim(Me.txtComentario.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        MsgBox "Se guardaron correctamente los Datos", vbInformation, "Aviso"
        fbRealizo = True
        Unload Me
    End If
End If
End Sub

Private Sub cmdCancelar_Click()
fbRealizo = False
Unload Me
End Sub

Private Sub txtNVisita_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = True

If InStr(Trim(txtFecVisita.Text), "_") > 0 Then
    MsgBox "Ingrese correctamente la Fecha de la Visita", vbInformation, "Aviso"
    ValidaDatos = False
    txtFecVisita.SetFocus
    Exit Function
End If

If Trim(txtNVisita.Text) = "" Or Trim(txtNVisita.Text) = "0" Then
    MsgBox "Ingrese el Nº de Visita", vbInformation, "Aviso"
    ValidaDatos = False
    txtNVisita.SetFocus
    Exit Function
End If

If Trim(txtComentario.Text) = "" Then
    MsgBox "Ingrese el Comentario", vbInformation, "Aviso"
    ValidaDatos = False
    txtComentario.SetFocus
    Exit Function
End If

If fgFecVisita > CDate(txtFecVisita.Text) Then
    MsgBox "La fecha de visita del Jefe de Agencia no puede ser menor a la Fecha de visita del Analista.", vbInformation, "Aviso"
    ValidaDatos = False
    txtFecVisita.SetFocus
    Exit Function
End If
End Function
