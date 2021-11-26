VERSION 5.00
Begin VB.Form frmPersComectario 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmPersComectario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   100
      TabIndex        =   1
      Top             =   960
      Width           =   5415
      Begin VB.TextBox TxtComentario 
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   100
      TabIndex        =   0
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RUC"
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   600
         Width           =   345
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DNI"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   285
      End
      Begin VB.Label LblRUC 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label LblDNI 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.Label LblNombre 
         BackColor       =   &H8000000E&
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmPersComectario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PersCod As String
Dim opt As Integer
Private Sub cmdBuscar_Click()
Dim rs As ADODB.Recordset
Dim Pers As UPersona
Set Pers = New UPersona
Set Pers = frmBuscaPersona.Inicio

Call cmdCancelar_Click
If Pers Is Nothing Then
Else
    
    PersCod = Pers.sPersCod
    lblNombre = PstaNombre(Pers.sPersNombre, False)
    LblDNI = IIf(IsNull(Pers.sPersIdnroDNI), "", Pers.sPersIdnroDNI)
    LblRUC = IIf(IsNull(Pers.sPersIdnroRUC), "", Pers.sPersIdnroRUC)
    Set rs = Pers.ObtieneComentario(PersCod)
    TxtComentario = ""
    If rs.EOF And rs.BOF Then
        opt = 1
    Else
        opt = 0
        TxtComentario = rs!cComentario
    End If
    TxtComentario.SetFocus
    
End If
Set rs = Nothing
Set Pers = Nothing
End Sub

Private Sub cmdCancelar_Click()
opt = -1
PersCod = ""
lblNombre = ""
LblRUC = ""
LblDNI = ""
TxtComentario = ""
End Sub

Private Sub cmdGrabar_Click()
Dim oFunciones As NContFunciones
Dim Pers As UPersona
Dim ban As Integer
Dim sMovNro As String
Set Pers = New UPersona
Set oFunciones = New NContFunciones
sMovNro = oFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set oFunciones = Nothing

If opt = 0 Then
    ban = MsgBox("Esta Seguro de Actualizar esta Informacion", vbInformation + vbYesNo, "AVISO")
    If vbYes = ban Then Call Pers.ActComentario(PersCod, (TxtComentario), sMovNro, 0)
End If
If opt = 1 Then
    ban = MsgBox("Esta Seguro de Guardar esta Informacion", vbQuestion + vbYesNo, "AVISO")
    If vbYes = ban Then Call Pers.ActComentario(PersCod, UCase(TxtComentario), sMovNro, 1)
End If
Call cmdCancelar_Click
Set Pers = Nothing
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
opt = -1
PersCod = ""
End Sub
