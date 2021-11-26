VERSION 5.00
Begin VB.Form frmMntOperacionDoc 
   Caption         =   "Operaciones: Mantenimiento: Documentos"
   ClientHeight    =   3690
   ClientLeft      =   1290
   ClientTop       =   3075
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   7575
   Visible         =   0   'False
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   4890
      TabIndex        =   8
      Top             =   3240
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   6120
      TabIndex        =   9
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Frame Frame4 
      Caption         =   "Documento"
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
      Height          =   2265
      Left            =   90
      TabIndex        =   10
      Top             =   870
      Width           =   7365
      Begin VB.Frame Frame1 
         Height          =   1695
         Left            =   5100
         TabIndex        =   18
         Top             =   390
         Width           =   2085
         Begin VB.OptionButton OptMetodo 
            Caption         =   "Digi&tado"
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   7
            Top             =   1170
            Width           =   1395
         End
         Begin VB.OptionButton OptMetodo 
            Caption         =   "Nro. de &Movimiento"
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Top             =   300
            Value           =   -1  'True
            Width           =   1845
         End
         Begin VB.OptionButton OptMetodo 
            Caption         =   "Auto&generado"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   750
            Width           =   1395
         End
      End
      Begin VB.Frame fraEstado 
         Height          =   885
         Left            =   3450
         TabIndex        =   16
         Top             =   390
         Width           =   1605
         Begin VB.OptionButton OptEstado 
            Caption         =   "O&pcional"
            Height          =   285
            Index           =   1
            Left            =   150
            TabIndex        =   2
            Top             =   510
            Width           =   1125
         End
         Begin VB.OptionButton OptEstado 
            Caption         =   "&Obligatorio"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   1
            Top             =   210
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.ListBox lstDocTpo 
         Height          =   1620
         Left            =   180
         TabIndex        =   0
         Top             =   480
         Width           =   3225
      End
      Begin VB.Frame fraCond 
         Height          =   855
         Left            =   3450
         TabIndex        =   17
         Top             =   1230
         Width           =   1605
         Begin VB.OptionButton OptCond 
            Caption         =   "&No debe existir"
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   450
            Width           =   1395
         End
         Begin VB.OptionButton OptCond 
            Caption         =   "&Debe existir"
            Height          =   285
            Index           =   0
            Left            =   150
            TabIndex        =   3
            Top             =   180
            Value           =   -1  'True
            Width           =   1125
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Método de Numeración"
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
         Left            =   5100
         TabIndex        =   19
         Top             =   210
         Width           =   1980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Estado"
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
         Left            =   3450
         TabIndex        =   15
         Top             =   210
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Left            =   150
         TabIndex        =   14
         Top             =   210
         Width           =   390
      End
   End
   Begin VB.TextBox txtOpeCod 
      Alignment       =   2  'Center
      BackColor       =   &H00ECFFFA&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   270
      TabIndex        =   11
      Top             =   300
      Width           =   1095
   End
   Begin VB.TextBox txtOpeDesc 
      BackColor       =   &H00ECFFFA&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1410
      TabIndex        =   12
      Top             =   300
      Width           =   5865
   End
   Begin VB.Frame fraOpe 
      Caption         =   "Operación "
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
      Height          =   705
      Left            =   90
      TabIndex        =   13
      Top             =   60
      Width           =   7365
   End
End
Attribute VB_Name = "frmMntOperacionDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOpeCod As String, sOpeDesc As String
Dim sCod As String, sDoc As String
Dim nDocTpo As TpoDoc
Dim sDocEstado As String, sDocMetodo As String
Dim rs As New ADODB.Recordset

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(vOpeCod As String, vOpeDesc As String)
sOpeCod = vOpeCod
sOpeDesc = vOpeDesc
Me.Show 1
End Sub

Private Sub LstDocTpo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If OptEstado(0) Then
      OptEstado(0).SetFocus
   Else
      OptEstado(1).SetFocus
   End If
End If
End Sub

Private Sub cmdAceptar_Click()
Dim clsOpeDoc As New DOperacion
On Error GoTo AceptarErr
nDocTpo = Val(Mid(lstDocTpo.List(lstDocTpo.ListIndex), 1, 3))
sDocEstado = IIf(OptEstado(0), "1", "2") & IIf(OptCond(0), "1", "2")
sDocMetodo = IIf(OptMetodo(0), "1", IIf(OptMetodo(1), "2", "3"))
If nDocTpo <= 0 Then
   MsgBox "Faltan datos por Definir...!", vbCritical, "Error"
   Exit Sub
End If
Set rs = clsOpeDoc.CargaOpeDoc(sOpeCod, nDocTpo, , adLockReadOnly, Val(nDocTpo))
If rs.RecordCount <> 0 Then
   MsgBox "El Tipo de documento ya está asignada a la operación", vbCritical, "Error"
   Exit Sub
End If
RSClose rs

If MsgBox("¿ Esta seguro de asignar el documento a la operación ?", vbQuestion + vbOKCancel, "Confirmación") = vbOk Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   clsOpeDoc.InsertaOpeDoc sOpeCod, nDocTpo, sDocEstado, sDocMetodo, gsMovNro
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Se Asigno el |Documento : " & lstDocTpo.Text & " a la Operacion : " & txtOpeDesc.Text
            Set objPista = Nothing
            '*******
   Unload Me
End If

Set clsOpeDoc = Nothing
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub
Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
txtOpeCod.Text = sOpeCod
txtOpeDesc.Text = sOpeDesc
CargaLstDocTpo
CentraForm Me
End Sub

Private Sub CargaLstDocTpo()
Dim clsDoc As New DDocumento
Set rs = clsDoc.CargaDocumento
Do While Not rs.EOF
    lstDocTpo.AddItem Format(rs!nDocTpo, "00") & " " & rs!cDocDesc
    rs.MoveNext
Loop
If lstDocTpo.ListCount > 0 Then
   lstDocTpo.ListIndex = 0
End If
RSClose rs
Set clsDoc = Nothing
End Sub

Private Sub OptCond_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   If OptMetodo(0) Then
      OptMetodo(0).SetFocus
   ElseIf OptMetodo(1) Then
      OptMetodo(1).SetFocus
   ElseIf OptMetodo(2) Then
      OptMetodo(2).SetFocus
   End If
End If
End Sub

Private Sub OptEstado_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   If OptCond(0) Then
      OptCond(0).SetFocus
   ElseIf OptCond(1) Then
      OptCond(1).SetFocus
   End If
End If
End Sub

Private Sub OptMetodo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub
