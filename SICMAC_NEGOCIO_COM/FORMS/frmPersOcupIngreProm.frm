VERSION 5.00
Begin VB.Form frmPersOcupIngreProm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Ocupacion e Ingreso Promedio"
   ClientHeight    =   2580
   ClientLeft      =   4620
   ClientTop       =   4485
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7185
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   6855
      Begin VB.TextBox txtIngresoProm 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   9
         Text            =   "0"
         Top             =   690
         Width           =   1395
      End
      Begin VB.ComboBox cboocupa 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmPersOcupIngreProm.frx":0000
         Left            =   1320
         List            =   "frmPersOcupIngreProm.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   3345
      End
      Begin VB.Label Label2 
         Caption         =   "Ingreso Promedio S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Ocupacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.CommandButton CmdPersAceptar 
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4320
      TabIndex        =   4
      ToolTipText     =   "Grabar todos los Cambios Realizados"
      Top             =   2160
      Width           =   1230
   End
   Begin VB.CommandButton CmdPersCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5745
      TabIndex        =   3
      ToolTipText     =   "Cancelar Todos los cambios Realizados"
      Top             =   2160
      Width           =   1230
   End
   Begin VB.Label lblPersNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   5055
   End
   Begin VB.Label lblPersCod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "cPersCod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmPersOcupIngreProm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Inicio(ByVal psPersCod As String, ByVal psPersNombre As String, ByVal psOcupacion As String, ByVal pnIngreProm As Double)
    Dim objPersona As COMDPersona.DCOMPersonas
    Dim rsOcupacion As Recordset
    
    Set rsOcupacion = New Recordset
    Set objPersona = New COMDPersona.DCOMPersonas
    
    Set rsOcupacion = objPersona.CargarOcupaciones()
    Call Llenar_Combo_con_Recordset(rsOcupacion, cboocupa)
    
    Me.lblPersCod.Caption = psPersCod
    Me.lblPersNombre.Caption = psPersNombre
    If psOcupacion <> "" Then
        Me.cboocupa.ListIndex = IndiceListaCombo(cboocupa, psOcupacion)
    End If
    Me.txtIngresoProm.Text = Format(pnIngreProm, "##,##0.00")
    
    Me.Show 1
   
End Sub

Private Sub cboocupa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtIngresoProm.SetFocus
    End If
End Sub

Private Sub CmdPersAceptar_Click()
    Dim objPersona As COMDPersona.DCOMPersonas
    Set objPersona = New COMDPersona.DCOMPersonas
    
    If validarDatos Then
        Exit Sub
    End If
    
    objPersona.ActualizarPersonaOcupIngreProm Me.lblPersCod, Trim(Right(Me.cboocupa.Text, 10)), CDbl(Me.txtIngresoProm.Text)
    Unload Me
    
End Sub
Private Function validarDatos() As Boolean
    validarDatos = False
    If Me.cboocupa.ListIndex = -1 Then
        MsgBox "Seleccione una ocupacion"
        validarDatos = True
        Exit Function
    End If
    If Me.txtIngresoProm.Text = "" Or Me.txtIngresoProm.Text = "0" Then
        MsgBox "Ingrese el Ingreso Promedio"
        validarDatos = True
        Exit Function
    End If
End Function

Private Sub CmdPersCancelar_Click()
    Unload Me
End Sub

Private Sub txtIngresoProm_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Me.CmdPersAceptar.SetFocus
  End If
End Sub
