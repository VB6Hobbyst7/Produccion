VERSION 5.00
Begin VB.Form frmLogValorizaComproObra 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Valorizar Contrato de Obra"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4875
   Icon            =   "frmLogValorizaComproObra.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
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
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nueva Valorización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1140
      Left            =   120
      TabIndex        =   0
      Top             =   50
      Width           =   4695
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción:"
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
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Monto:"
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
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmLogValorizaComproObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fsMatrizDatos() As String
Public Function Inicio(ByVal pnvalor As Integer) As String()
    Me.Show 1
    Inicio = fsMatrizDatos
End Function
Private Sub cmdAceptar_Click()
    If Not Valida Then Exit Sub
    fsMatrizDatos(1, 1) = Trim(txtMonto.Text)
    fsMatrizDatos(2, 1) = Trim(txtDescripcion.Text)
    Unload Me
End Sub
Private Function Valida() As Boolean
    If Len(txtMonto.Text) = 0 Then
        MsgBox "No se ha Ingresado Ningun Monto.", vbExclamation, "Aviso."
        Valida = False
        txtMonto.SetFocus
        Exit Function
    End If
    If Len(txtDescripcion.Text) = 0 Then
        MsgBox "No se ha Ingresado la Descripción", vbExclamation, "Aviso."
        Valida = False
        txtDescripcion.SetFocus
        Exit Function
    End If
    Valida = True
End Function
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    ReDim fsMatrizDatos(2, 1)
End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtMonto, KeyAscii, 10, 3)
    If KeyAscii = 13 Then
        Me.txtDescripcion.SetFocus
    End If
End Sub
