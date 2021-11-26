VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIngFechaGenDoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fecha de Generación de Documento"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4170
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   1035
   End
   Begin VB.Frame fraFecha 
      Appearance      =   0  'Flat
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1725
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
End
Attribute VB_Name = "frmIngFechaGenDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFecha As String
Dim sFechaVig As Date
Public Function Inicio(ByVal psFechaVig As Date) As String
    sFechaVig = psFechaVig
    Me.Show 1
    Inicio = sFecha
    sFecha = ""
End Function

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    End If
End Sub
Private Sub cmdAceptar_Click()
    If Not IsDate(txtFecha) Then
        MsgBox "Fecha no Valida", vbInformation, "Aviso"
        Me.txtFecha.SetFocus
        Exit Sub
    End If
    If CDate(txtFecha) > DateAdd("d", 7, sFechaVig) Then
        MsgBox "La Fecha no puede ser mayor a 7 dias más de la fecha de aprobación del crédito"
        Exit Sub
    End If
    
    If CDate(txtFecha) < sFechaVig Then
        MsgBox "Fecha no Valida"
        Exit Sub
    End If
    sFecha = Me.txtFecha
    Unload Me
End Sub
