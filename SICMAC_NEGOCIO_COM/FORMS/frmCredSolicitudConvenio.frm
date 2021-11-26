VERSION 5.00
Begin VB.Form frmCredSolicitudConvenio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso de Datos de Solicitud por Convenio"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frmCredSolicitudConvenio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraConvenio 
      Height          =   1965
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6915
      Begin VB.Frame fraTipoPlanilla 
         Caption         =   "Tipo Planilla"
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
         Height          =   690
         Left            =   150
         TabIndex        =   11
         Top             =   1125
         Width           =   6600
         Begin VB.OptionButton optT_Plani 
            Caption         =   "CAS"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optT_Plani 
            Caption         =   "Descuento para Activos"
            Height          =   390
            Index           =   0
            Left            =   1680
            TabIndex        =   13
            Top             =   240
            Width           =   2190
         End
         Begin VB.OptionButton optT_Plani 
            Caption         =   "Descuento para Cesantes"
            Height          =   390
            Index           =   1
            Left            =   4320
            TabIndex        =   12
            Top             =   225
            Width           =   2190
         End
      End
      Begin VB.TextBox txtCARBEN 
         Height          =   315
         Left            =   6135
         MaxLength       =   4
         TabIndex        =   10
         Text            =   "0000"
         Top             =   675
         Width           =   615
      End
      Begin VB.TextBox txtCargo 
         Height          =   315
         Left            =   3960
         MaxLength       =   6
         TabIndex        =   8
         Text            =   "000000"
         Top             =   675
         Width           =   765
      End
      Begin VB.ComboBox cmbModular 
         Height          =   315
         Left            =   1170
         TabIndex        =   4
         Top             =   690
         Width           =   1815
      End
      Begin VB.ComboBox cmbInstitucion 
         Height          =   315
         Left            =   1170
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   225
         Width           =   5565
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "o DNI :"
         Height          =   195
         Left            =   75
         TabIndex        =   14
         Top             =   900
         Width           =   510
      End
      Begin VB.Label Label2 
         Caption         =   "Correlativo Sobreviviente:"
         Height          =   390
         Left            =   4935
         TabIndex        =   9
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Cargo:"
         Height          =   240
         Left            =   3285
         TabIndex        =   7
         Top             =   750
         Width           =   615
      End
      Begin VB.Label lblModular 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Modular :"
         Height          =   195
         Left            =   75
         TabIndex        =   6
         Top             =   675
         Width           =   1035
      End
      Begin VB.Label lblinstitucion 
         AutoSize        =   -1  'True
         Caption         =   "Institución :"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   285
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5565
      TabIndex        =   1
      Top             =   2025
      Width           =   1380
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   2025
      Width           =   1380
   End
End
Attribute VB_Name = "frmCredSolicitudConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fsCodInstitucion As String
Public fsCodModular As String
Public fsCargo As String
Public fsCARBEN As String
Public fsT_Plani As String

Public Function Inicio(ByVal rsInstituc As ADODB.Recordset)
    cmbInstitucion.Clear
    Do While Not rsInstituc.EOF
        cmbInstitucion.AddItem PstaNombre(rsInstituc!cPersNombre) & Space(250) & rsInstituc!cPersCod
        rsInstituc.MoveNext
    Loop
    Call CambiaTamañoCombo(cmbInstitucion, 400)

    Me.Show 1
End Function

Private Sub cmbInstitucion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmbModular.SetFocus
End If
End Sub

Private Sub cmbModular_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCargo.SetFocus
End If
End Sub

Private Sub CmdAceptar_Click()
    If cmbInstitucion.ListIndex = -1 Then
        MsgBox "Debe seleccionar una institucion", vbInformation, "Mensaje"
        Exit Sub
    End If
    If cmbModular.Text = "" Then
        MsgBox "Debe indicar un Codigo Modular", vbInformation, "Mensaje"
        Exit Sub
    End If
    If txtCargo.Text = "" Then
        MsgBox "Debe indicar un Cargo", vbInformation, "Mensaje"
        Exit Sub
    End If
    If txtCARBEN.Text = "" Then
        MsgBox "Debe indicar un Correlativo de Sobreviviente", vbInformation, "Mensaje"
        Exit Sub
    End If
    If txtCargo.Text = "" Then
        MsgBox "Debe indicar un Cargo", vbInformation, "Mensaje"
        Exit Sub
    End If
    If optT_Plani(0).Value = False And optT_Plani(1).Value = False And optT_Plani(2).Value = False Then 'WIOR 20150217 AGREGO optT_Plani(2).Value = False
        MsgBox "Debe indicar tipo de Planilla", vbInformation, "Mensaje"
        Exit Sub
    End If
    'ARCV 30-01-2007
    'If Len(cmbModular.Text) <> 10 Then
    '    MsgBox "El Codigo Modular debe tener 10 caracteres", vbInformation, "Mensaje"
    '    Exit Sub
    'End If
    If Len(txtCargo.Text) <> 6 Then
        MsgBox "El Cargo debe tener 6 caracteres", vbInformation, "Mensaje"
        Exit Sub
    End If
    If Len(txtCARBEN.Text) <> 4 Then
        MsgBox "El Correlativo de Sobreviviente debe tener 4 caracteres", vbInformation, "Mensaje"
        Exit Sub
    End If
    If txtCargo.Text = "" Then
        MsgBox "Debe indicar un Cargo", vbInformation, "Mensaje"
        Exit Sub
    End If
    
    
    fsCodInstitucion = Trim(Right(cmbInstitucion.Text, 15))
    fsCodModular = cmbModular.Text
    fsCargo = txtCargo.Text
    fsCARBEN = txtCARBEN.Text
    fsT_Plani = IIf(optT_Plani(0).Value, "A", IIf(optT_Plani(2).Value, "CA", "C")) 'WIOR 20150217
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    fsCodInstitucion = ""
    fsCodModular = ""
    fsCargo = ""
    fsCARBEN = ""
    fsT_Plani = ""
    Unload Me
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
End Sub

Private Sub optT_Plani_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub

Private Sub txtCARBEN_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    optT_Plani(0).SetFocus
End If
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCARBEN.SetFocus
End If
End Sub
