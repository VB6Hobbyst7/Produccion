VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsolidaEstadRiesgos 
   Caption         =   "Riesgos: Consolida Estadísticas"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   Icon            =   "frmConsolidaEstadRiesgos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prg 
      Height          =   255
      Left            =   3060
      TabIndex        =   10
      Top             =   3030
      Visible         =   0   'False
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdConsolidar 
      Caption         =   "Consoli&dar"
      Height          =   405
      Left            =   4065
      TabIndex        =   2
      Top             =   2085
      Width           =   1620
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   4065
      TabIndex        =   3
      Top             =   2490
      Width           =   1620
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3720
      Begin VB.CheckBox chkTodos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   0
         Width           =   960
      End
      Begin MSComctlLib.ListView lstAgencias 
         Height          =   2250
         Left            =   120
         TabIndex        =   5
         Top             =   330
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   3969
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Agencia"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Total Registros"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "cCodAge"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox txtDesde 
      Height          =   315
      Left            =   4530
      TabIndex        =   0
      Top             =   285
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txtHasta 
      Height          =   315
      Left            =   4530
      TabIndex        =   1
      Top             =   645
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.StatusBar estado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   3015
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde :"
      Height          =   195
      Left            =   3915
      TabIndex        =   9
      Top             =   300
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta :"
      Height          =   195
      Left            =   3900
      TabIndex        =   8
      Top             =   720
      Width           =   510
   End
End
Attribute VB_Name = "frmConsolidaEstadRiesgos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCmact As DConecta
Dim oCon    As DConecta
Dim I As Integer

Private Sub chktodos_Click()
For I = 1 To Me.lstAgencias.ListItems.Count
    lstAgencias.ListItems(I).Checked = Me.chkTodos.value
Next I
End Sub

Private Sub cmdConsolidar_Click()
If Valida Then
    cmdConsolidar.Enabled = False
    Consolida
    cmdConsolidar.Enabled = True
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Set dbCmact = New DConecta
Set oCon = New DConecta
oCon.AbreConexion

LlenaAgencias
Me.txtDesde = gdFecSis
Me.txtHasta = gdFecSis
End Sub
Private Sub LlenaAgencias()
Dim sql As String
Dim rs As New ADODB.Recordset
Dim Item As ListItem

Dim oAge As New DActualizaDatosArea
Set rs = oAge.GetAgencias(, False)
If Not RSVacio(rs) Then
    lstAgencias.ListItems.Clear
    Do While Not rs.EOF
        If Trim(rs!Codigo) = gsCodAge Then
            Set Item = lstAgencias.ListItems.Add(, , "Agencia Local")
            Item.SubItems(1) = 0
            Item.SubItems(2) = Trim(rs!Codigo)
        Else
            Set Item = lstAgencias.ListItems.Add(, , Trim(rs!Descripcion))
            Item.SubItems(1) = 0
            Item.SubItems(2) = Trim(rs!Codigo)
        End If
        rs.MoveNext
    Loop
Else
    MsgBox "El Sistema no Detecta Servidores. Por Favor consulte al Area de Sistemas", vbInformation, "Aviso"
End If
rs.Close
Set rs = Nothing
End Sub
Private Function Valida() As Boolean
If CDate(Me.txtDesde) > CDate(Me.txtHasta) Then
    MsgBox "Fecha Final no puede ser menor que Inicial", vbInformation, "Aviso"
    Valida = False
    txtHasta.SetFocus
    Exit Function
Else
    Valida = True
End If
End Function
Private Sub Consolida()
Dim j As Long
Dim I As Long
Dim lbOk As Boolean
lbOk = True
For I = 1 To Me.lstAgencias.ListItems.Count
    If lstAgencias.ListItems(I).Checked = True Then
        If dbCmact.AbreConexion Then 'Remota(Right(lstAgencias.ListItems(i).SubItems(2), 2), False)
            CopiaEstadVencRiesgos gsCodCMAC & lstAgencias.ListItems(I).SubItems(2)
            dbCmact.CierraConexion
        Else
            MsgBox "Agencia " & lstAgencias.ListItems(I).SubItems(2) & " no responde", vbInformation, "¡Aviso!"
            lbOk = False
        End If
    End If
Next I
If lbOk Then
    MsgBox "Consolidacion Finalizada con éxito", vbInformation, "Aviso"
Else
    MsgBox "Algunas Agencias no se han consolidado de forma satisfactoria" & Chr(13) & "Verifique el Error con sistemas y vuelva a intentar el proceso", vbExclamation, "Aviso"
End If
End Sub

Private Sub CopiaEstadVencRiesgos(psCodAge As String)
Dim dFecha  As Date
Dim lConsol As Boolean
Dim rs   As New ADODB.Recordset
Dim sSql As String
dFecha = CDate(txtDesde)
Do While dFecha <= CDate(txtHasta)
   lConsol = True
   Me.Estado.Panels(1).Text = "Procesando Ag. " & psCodAge & " Día " & dFecha
   If ExisteEstadistica(psCodAge, dFecha) Then
      If MsgBox(" ¿ Estadística de Agencia " & psCodAge & " día " & dFecha & " ya fue Consolidada. ¿Desea volver a Consolidar?", vbQuestion + vbYesNo, "¡Aviso!") = vbNo Then
         lConsol = False
      Else
         sSql = "DELETE CredEstadRiesgos WHERE cCodAge = '" & psCodAge & "' and datediff(d,dFecha,'" & Format(dFecha, gsFormatoFecha) & "') = 0"
         oCon.Ejecutar sSql
      End If
   End If
   If lConsol Then
      sSql = "SELECT * FROM CredEstadRiesgos WHERE cCodAge = '" & psCodAge & "' and datediff(d,dFecha,'" & Format(dFecha, gsFormatoFecha) & "') = 0"
      Set rs = dbCmact.CargaRecordSet(sSql)
      If Not rs.EOF Then
         prg.Visible = True
         prg.Max = rs.RecordCount
         DoEvents
         Do While Not rs.EOF
            sSql = "INSERT CredEstadRiesgos (cCodAge, cProd, dFecha, cEstado, cMoneda, dFecVenc, nSaldo) " _
                 & "VALUES ('" & rs!cCodAge & "','" & rs!cProd & "','" & Format(rs!dFecha, gsFormatoFecha) & "','" & rs!cEstado & "','" & rs!cmoneda & "','" & Format(rs!dFecVenc, gsFormatoFecha) & "', " & rs!nSaldo & ")"
            oCon.Ejecutar sSql
            prg.value = rs.Bookmark
            rs.MoveNext
         Loop
         prg.Visible = False
      End If
      RSClose rs
   End If
   dFecha = dFecha + 1
Loop
End Sub

Private Function ExisteEstadistica(psCodAge As String, pdFecha As Date) As Boolean
Dim sSql As String
Dim prs  As New ADODB.Recordset
ExisteEstadistica = False
sSql = "SELECT TOP 1 cCodAge FROM CredEstadRiesgos WHERE cCodAge = '" & psCodAge & "' and datediff(d,dFecha,'" & Format(pdFecha, gsFormatoFecha) & "') = 0"
Set prs = oCon.CargaRecordSet(sSql)
If Not prs.EOF Then
   ExisteEstadistica = True
End If
RSClose prs
End Function

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing

End Sub

Private Sub txtDesde_GotFocus()
fEnfoque txtDesde
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtHasta.SetFocus
End If
End Sub

Private Sub txtDesde_LostFocus()
fEnfoque txtDesde
End Sub

Private Sub txtHasta_GotFocus()
fEnfoque txtHasta
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdConsolidar.SetFocus
End If
End Sub

Private Sub txtHasta_LostFocus()
fEnfoque txtHasta
End Sub

