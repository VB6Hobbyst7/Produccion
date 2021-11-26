VERSION 5.00
Begin VB.Form frmLogContTipoAdenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Adenda"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3105
   Icon            =   "frmLogContTipoAdenda.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1515
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmTipoAdenda 
      Caption         =   "Seleccione el Tipo de Adenda"
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox cboTipoAdenda 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
   End
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
      Left            =   2040
      TabIndex        =   1
      Top             =   1080
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
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frmLogContTipoAdenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fsContrato As String
Dim fnContRef As Integer
Dim fnAdenda As Integer
Public Sub Inicio(ByVal psNContrato As String, ByVal pnContRef As Integer, Optional ByVal pnAdenda As Integer = 0)
    fsContrato = psNContrato
    fnContRef = pnContRef
    fnAdenda = pnAdenda
    Me.Show 1
End Sub
Private Sub cmdAceptar_Click()
    Dim oDLogGeneral As DLogGeneral
    Set oDLogGeneral = New DLogGeneral
    If cboTipoAdenda.ListIndex = -1 Then
        MsgBox "No se ha seleccionado ningún tipo de Adednda.", vbExclamation + vbOKOnly, "Mensaje"
        cboTipoAdenda.SetFocus
        Exit Sub
    End If
    frmLogContRegAdendas.Inicio fsContrato, fnAdenda, Trim(Right(cboTipoAdenda.Text, 4)), LogTipoContrato.ContratoServicio, fnContRef, oDLogGeneral.ObtieneTipoPagoContratoServicio(fsContrato, fnContRef)
    Unload Me
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CargaDatos
End Sub
Private Sub CargaDatos()
    Dim oLog As DLogGeneral
    Dim oConst As DConstantes
    Set oConst = New DConstantes

    Set oLog = New DLogGeneral
    CargaCombo oConst.GetConstante(gsLogContTipoAdendas), Me.cboTipoAdenda
End Sub
