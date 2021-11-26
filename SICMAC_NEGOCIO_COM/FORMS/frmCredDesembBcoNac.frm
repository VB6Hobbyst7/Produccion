VERSION 5.00
Begin VB.Form frmCredDesembBcoNac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Desembolso en Banco de la Nación"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nueva Clave"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
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
      Height          =   405
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1140
   End
   Begin VB.ComboBox cmbAgencia 
      Height          =   315
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   3240
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "Agencia :"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "Clave Cliente :"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   165
      Width           =   1020
   End
End
Attribute VB_Name = "frmCredDesembBcoNac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nCodigo As Integer

Private Sub cmdAceptar_Click()
    Me.Hide
End Sub

Private Sub cmdNuevo_Click()
    Randomize
    'Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)
    txtCodigo.Text = Right("000" & CStr(Int(9999 * Rnd + 1)), 4)
End Sub

Private Sub Form_Load()
    Randomize
    'Int((Límite_superior - límite_inferior + 1) * Rnd + límite_inferior)
    txtCodigo.Text = Right("000" & CStr(Int(9999 * Rnd + 1)), 4)
    CargaComboDatos cmbAgencia
End Sub

Private Sub CargaComboDatos(ByVal combo As ComboBox)
    Dim oConst As COMDConstantes.DCOMAgencias
    Dim lrs As New ADODB.Recordset
    Set oConst = New COMDConstantes.DCOMAgencias
        Set lrs = oConst.RecuperaAgenciasBancoNacion()
        combo.Clear
        If Not (lrs.EOF And lrs.BOF) Then
            Do Until lrs.EOF
                combo.AddItem lrs(1) & Space(20) & lrs(0)
                lrs.MoveNext
            Loop
        End If
    Set oConst = Nothing
    Set lrs = Nothing
End Sub
