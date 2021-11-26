VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCuentasContables 
   Caption         =   "Cuentas Contables"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCuentaContable 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuscarCuentaContable 
      Caption         =   "&Buscar"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   3360
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   3360
      Width           =   1095
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2820
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   4974
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cuenta"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta Contable"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCuentasContables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
'***Nombre:         frmCuentasContables
'***Descripción:    Formulario donde se lista las cuentas contables
'                   que permiten seleccionar para su generación en
'                   el Reporte de Gastos
'***Creado por ELRO el 20111011, según Acta 278-2011/TI-D
'********************************************************

Private fsOpeCod As String
Private fsCtaCta As String

Private Sub cmdAceptar_Click()
Dim k As Integer

For k = 1 To lv.ListItems.Count
    If lv.ListItems(k).Checked Then
        fsCtaCta = fsCtaCta & ",'" & lv.ListItems(k).Text & "'"
    End If
Next
Unload Me
End Sub

Private Sub cmdBuscarCuentaContable_Click()
lv.SetFocus
    Dim k As Integer
    For k = 1 To lv.ListItems.Count
        If lv.ListItems(k).Text = txtCuentaContable Then
            DoEvents
            lv.ListItems(k).Selected = True
            lv.ListItems(k).Checked = True
            lv.SelectedItem.EnsureVisible
            Exit Sub
        End If
    Next
End Sub

Private Sub cmdCancelar_Click()
fsCtaCta = ""
Unload Me
End Sub

Private Sub Form_Load()
Dim oOperacion As DOperacion
Dim rsCuentasContables As ADODB.Recordset
Dim lvItem As ListItem
Dim k As Integer

CentraForm Me

If fsOpeCod = "760200" Then
    Set oOperacion = New DOperacion
    Set rsCuentasContables = New ADODB.Recordset
    Set rsCuentasContables = oOperacion.listarCuentasOperacion("45")
    Do While Not rsCuentasContables.EOF
        Set lvItem = lv.ListItems.Add
        lvItem.Text = rsCuentasContables!cCtaContCod
        lvItem.SubItems(1) = rsCuentasContables!cCtaContDesc
        rsCuentasContables.MoveNext
    Loop
    lv.Visible = True
Else
    fsCtaCta = ""
    
    Unload Me
End If
End Sub

Public Function inicio(ByVal psOpeCod) As String
fsOpeCod = psOpeCod
Me.Show 1
inicio = fsCtaCta
End Function

Private Sub txtCuentaContable_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call cmdBuscarCuentaContable_Click
    End If
End Sub


