VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsEstadisticas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consolidación de Estadisticas Diarias"
   ClientHeight    =   3885
   ClientLeft      =   2010
   ClientTop       =   2325
   ClientWidth     =   6090
   Icon            =   "frmConsEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2715
      Left            =   135
      TabIndex        =   15
      Top             =   825
      Width           =   3915
      Begin VB.CheckBox chkTodos 
         Caption         =   "&Todos"
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   0
         Width           =   960
      End
      Begin MSComctlLib.ListView lstAgencias 
         Height          =   2250
         Left            =   165
         TabIndex        =   6
         Top             =   300
         Width           =   3570
         _ExtentX        =   6297
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
      Left            =   4725
      TabIndex        =   7
      Top             =   975
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estadísticas"
      Height          =   690
      Left            =   150
      TabIndex        =   12
      Top             =   45
      Width           =   5670
      Begin VB.CheckBox chkPrendario 
         Caption         =   "&Pignoraticio"
         Height          =   285
         Left            =   4305
         TabIndex        =   4
         Top             =   240
         Width           =   1185
      End
      Begin VB.CheckBox chkCredito 
         Caption         =   "&Credito"
         Height          =   285
         Left            =   3315
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CheckBox chkCTS 
         Caption         =   "CT&S"
         Height          =   285
         Left            =   2490
         TabIndex        =   2
         Top             =   240
         Width           =   630
      End
      Begin VB.CheckBox chkPlazofijo 
         Caption         =   "Plazo &Fijo"
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   990
      End
      Begin VB.CheckBox chkAhorros 
         Caption         =   "&Ahorros"
         Height          =   285
         Left            =   270
         TabIndex        =   0
         Top             =   225
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   4365
      TabIndex        =   10
      Top             =   3090
      Width           =   1575
   End
   Begin VB.CommandButton cmdConsolidar 
      Caption         =   "Consoli&dar"
      Height          =   405
      Left            =   4365
      TabIndex        =   9
      Top             =   2685
      Width           =   1575
   End
   Begin MSComctlLib.StatusBar estado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   3630
      Width           =   6090
      _ExtentX        =   10742
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
   Begin MSMask.MaskEdBox txtHasta 
      Height          =   315
      Left            =   4725
      TabIndex        =   8
      Top             =   1335
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta :"
      Height          =   195
      Left            =   4170
      TabIndex        =   14
      Top             =   1380
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde :"
      Height          =   195
      Left            =   4155
      TabIndex        =   13
      Top             =   1035
      Width           =   555
   End
End
Attribute VB_Name = "frmConsEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkPrendario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.chkTodos.SetFocus
End If
End Sub

Private Sub chktodos_Click()
Dim I As Integer
For I = 1 To Me.lstAgencias.ListItems.Count
    lstAgencias.ListItems(I).Checked = Me.chkTodos.value
Next I
End Sub

Private Sub chkTodos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.lstAgencias.SetFocus
End If
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
Dim oConec As DConecta
Set oConec = New DConecta

oConec.AbreConexion
CentraForm Me
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
        If Trim(rs!Codigo) = Right(gsCodAge, 2) Then
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
'If (Me.chkAhorros.Value = 0 Or Me.chkPlazofijo.Value = 0 Or Me.chkCTS.Value = 0 _
 & Me.chkCredito.Value = 0 Or Me.chkPrendario.Value = 0) Then
'   MsgBox "Seleccione Alguna opción para consolidar", vbInformation, "Aviso"
'   Valida = False
'   Me.chkAhorros.SetFocus
'   Exit Function
'Else
'    Valida = True
'End If
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
Dim rs As New ADODB.Recordset
Dim sql As String
Dim Total As Long
Dim j As Long
Dim I As Long
Dim CadTemp As String
Dim lbOk As Boolean
Dim oCon As DConecta
Set oCon = New DConecta
lbOk = True
For I = 1 To Me.lstAgencias.ListItems.Count
    If Me.lstAgencias.ListItems(I).Checked = True Then
        Me.Estado.Panels(2).Text = "Conectandose a Agencia" & lstAgencias.ListItems(I).SubItems(2)
        If oCon.AbreConexion Then 'Remota(Right(lstAgencias.ListItems(I).SubItems(2), 2), False)
            If Me.chkAhorros.value = 1 Then
                EstadAhorros oCon.ConexionActiva, lstAgencias.ListItems(I).SubItems(2)
            End If
            If Me.chkPlazofijo.value = 1 Then
                'EstadPlazoFijo oCon.ConexionActiva, lstAgencias.ListItems(I).SubItems(2)
            End If
            If Me.chkCTS.value = 1 Then
                'EstadCTS oCon.ConexionActiva, lstAgencias.ListItems(I).SubItems(2)
            End If
            If Me.chkCredito.value = 1 Then
                EstadCreditos oCon.ConexionActiva, lstAgencias.ListItems(I).SubItems(2)
            End If
            If Me.chkPrendario.value = 1 Then
                EstadPrendario oCon.ConexionActiva, lstAgencias.ListItems(I).SubItems(2)
            End If
            oCon.CierraConexion
        Else
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
Private Sub EstadAhorros(dbConexion As ADODB.Connection, lsCodAge As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim I As Long
Dim m As Long
Dim lnTotal As Long
Dim poCon As New DConecta
Set poCon = New DConecta
poCon.AbreConexion

m = 0

sql = "SELECT * FROM CapEstadSaldo WHERE dEstad BETWEEN '" & Format(txtDesde, gsFormatoFecha) & "' AND '" & Format(txtHasta, gsFormatoFecha) & " 23:59:59'"
rs.Open sql, dbConexion, adOpenStatic, adLockOptimistic, adCmdText
Do While Not rs.EOF
    m = m + 1
    Estado.Panels(1).Text = "Ahorros al : " & Format(rs!dEstadAC, "dd/mm/yyyy") & " Total :" & m
    If ExisteEstadistica(rs!dEstadAC, "Ahorros", lsCodAge, , Trim(rs!cmoneda)) = True Then
        sql = "UPDATE CapEstadMovimiento Set " _
            & "nNumAperAC=" & rs!nNumAperAC & "," _
            & "nMonAperAC=" & rs!nMonAperAC & "," _
            & "nNumCancAC=" & rs!nNumCancAC & "," _
            & "nMonCancAC=" & rs!nMonCancAC & "," _
            & "nRetIntAC=" & rs!nRetIntAC & "," _
            & "nRetInacAC=" & rs!nRetInacAC & "," _
            & "nNumCancInacAC=" & rs!nNumCancInacAC & "," _
            & "nMonCancInacAC=" & rs!nMonCancInacAC & "," _
            & "nNumDepAC=" & rs!nNumDepAC & "," _
            & "nMonDepAC=" & rs!nMonDepAC & "," _
            & "nNumRetAC=" & rs!nNumRetAC & "," _
            & "nMonRetAC=" & rs!nMonRetAC & "," _
            & "nSaldoAntAC=" & rs!nSaldoAntAC & "," _
            & "nSaldoAC=" & rs!nSaldoAC & "," _
            & "nSaldCMAC=" & rs!nSaldCMAC & "," _
            & "nChqCMAC=" & rs!nChqCMAC & "," _
            & "nCtaVigAC=" & rs!nCtaVigAC & "," _
            & "nMonChqAC=" & rs!nMonChqAC & "," _
            & "nIntCapAC=" & rs!nIntCapAC & "," _
            & "cCodUsu='" & rs!cCodUsu & "'," _
            & "nMonChqVal=" & rs!nMonChqVal & " " _
            & "Where cCodAge='" & lsCodAge & "' AND Datediff(Day,dEstadAC,'" & Format(rs!dEstadAC, gsFormatoFecha) & "')=0 and cMoneda ='" & rs!cmoneda & "'"
            
        poCon.Ejecutar sql
            
    Else
        sql = "INSERT INTO CapEstadMovimiento(dEstadAC, cMoneda, nNumAperAC, nMonAperAC, nNumCancAC, nMonCancAC," _
            & "nRetIntAC, nRetInacAC, nNumCancInacAC, nMonCancInacAC, nNumDepAC, " _
            & "nMonDepAC, nNumRetAC, nMonRetAC, nSaldoAntAC, nSaldoAC, nSaldCMAC, " _
            & "nChqCMAC, nCtaVigAC, nMonChqAC, nIntCapAC, cCodUsu, nMonChqVal, cCodAge) " _
            & "Values ('" & Format(rs!dEstadAC, "mm/dd/yyyy hh:mm:ss AMPM") & "','" & rs!cmoneda & "'," & rs!nNumAperAC & "," _
            & rs!nMonAperAC & "," & rs!nNumCancAC & "," & rs!nMonCancAC & "," _
            & rs!nRetIntAC & "," & rs!nRetInacAC & "," & rs!nNumCancInacAC & "," _
            & rs!nMonCancInacAC & "," & rs!nNumDepAC & "," & rs!nMonDepAC & "," _
            & rs!nNumRetAC & "," & rs!nMonRetAC & "," & rs!nSaldoAntAC & "," _
            & rs!nSaldoAC & "," & rs!nSaldCMAC & "," & rs!nChqCMAC & "," _
            & rs!nCtaVigAC & "," & rs!nMonChqAC & "," & rs!nIntCapAC & ",'" _
            & rs!cCodUsu & "'," & rs!nMonChqVal & ",'" & lsCodAge & "')"
            
        poCon.Ejecutar sql
    End If
    rs.MoveNext
    Me.Estado.Panels(2).Text = "Agencia : " & lsCodAge
    DoEvents
Loop

m = 0
'sql = "SELECT * FROM EstadMensAho WHERE dEstadMens BETWEEN '" & Format(txtDesde, gsFormatoFecha) & "' AND '" & Format(txtHasta, gsFormatoFecha) & " 23:59:59'"
'rs.Close
'rs.Open sql, dbConexion, adOpenStatic, adLockOptimistic, adCmdText
'Do While Not rs.EOF
'   m = m + 1
'   sql = "UPDATE EstadDiaACConsol Set " _
'       & "nSaldCRAC =" & rs!nSaldCRAC & " " _
'       & "Where cCodAge='" & lsCodAge & "' AND Datediff(Day,dEstadAC,'" & Format(rs!dEstadMens, gsFormatoFecha) & "')=0 and cMoneda ='" & rs!cmoneda & "'"
'   poCon.Ejecutar sql
'
'   rs.MoveNext
'   Me.estado.Panels(2).Text = "Agencia : " & lsCodAge
'   DoEvents
'Loop
RSClose rs
poCon.CierraConexion
Set poCon = Nothing
End Sub

Private Sub EstadPlazoFijo(dbConexion As ADODB.Connection, lsCodAge As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim I As Long
Dim m As Long
Dim poCon As New DConecta
poCon.AbreConexion

sql = "SELECT * FROM EstadDiaPF WHERE datediff(d,dEstadPF,'" & Format(txtDesde, gsFormatoFecha) & "') <= 0 AND datediff(d,dEstadPF,'" & Format(txtHasta, gsFormatoFecha) & "') >= 0 "
m = 0
rs.Open sql, dbConexion, adOpenStatic, adLockOptimistic, adCmdText
Do While Not rs.EOF
    m = m + 1
    Me.Estado.Panels(1).Text = "Plazo Fijo al : " & Format(rs!dEstadPF, "dd/mm/yyyy") & " Total :" & m
    If ExisteEstadistica(rs!dEstadPF, "Plazo", lsCodAge, , Trim(rs!cmoneda)) = True Then
        
        sql = "UPDATE EstadDiaPFConsol Set " _
            & "nNumAperPF=" & rs!nNumAperPF & "," _
            & "nMonAperPF=" & rs!nMonAperPF & "," _
            & "nNumCancPF=" & rs!nNumCancPF & "," _
            & "nMonCancPF=" & rs!nMonCancPF & "," _
            & "nNumMovPF=" & rs!nNumMovPF & "," _
            & "nMonMovPF=" & rs!nMonMovPF & "," _
            & "nNumVigPF=" & rs!nNumVigPF & "," _
            & "nIntCapPF=" & rs!nSaldoAntPF & "," _
            & "nSaldoPF=" & rs!nSaldoPF & "," _
            & "cCodUsu='" & rs!cCodUsu & "'," _
            & "nMonChqVal=" & rs!nMonChqVal & ", nNumChqVal =" & rs!nNumChqVal & ", nNumCMAC=" & rs!nNumCMAC & ",nSaldCMAC=" & rs!nSaldCMAC & "  " _
            & "Where cCodAge='" & lsCodAge & "' AND Datediff(Day,dEstadPF,'" & Format(rs!dEstadPF, gsFormatoFecha) & "')=0 and cMoneda ='" & rs!cmoneda & "'"
            
        poCon.Ejecutar sql
            
    Else
        
        sql = "INSERT INTO EstadDiaPFConsol(dEstadPF, cMoneda, nNumAperPF, nMonAperPF, nNumCancPF, nMonCancPF, nNumMovPF," _
            & "nMonMovPF, nNumVigPF, nIntCapPF, nSaldoAntPF, nSaldoPF, cCodUsu, nMonChqVal, cCodAge, nNumChqVal, nNumCMAC, nSaldCMAC ) " _
            & "Values ('" & Format(rs!dEstadPF, "mm/dd/yyyy hh:mm:ss AMPM") & "','" & rs!cmoneda & "'," & rs!nNumAperPF & "," _
            & rs!nMonAperPF & "," & rs!nNumCancPF & "," & rs!nMonCancPF & "," & rs!nNumMovPF & "," _
            & rs!nMonMovPF & "," & rs!nNumVigPF & "," & rs!nIntCapPF & "," _
            & rs!nSaldoAntPF & "," & rs!nSaldoPF & ",'" _
            & rs!cCodUsu & "'," & rs!nMonChqVal & ",'" & lsCodAge & "'," & rs!nNumChqVal & "," & rs!nNumCMAC & "," & rs!nSaldCMAC & ")"
            
        poCon.Ejecutar sql
    End If
    rs.MoveNext
    Me.Estado.Panels(2).Text = "Agencia : " & lsCodAge
    DoEvents
Loop

'Saldo de Caja Rurales
sql = "SELECT * FROM EstadMensPF WHERE dEstadMens BETWEEN '" & Format(txtDesde, gsFormatoFecha) & "' AND '" & Format(txtHasta, gsFormatoFecha) & " 23:59:59'"
m = 0
rs.Close
rs.Open sql, dbConexion, adOpenStatic, adLockOptimistic, adCmdText
Do While Not rs.EOF
   m = m + 1
   sql = "UPDATE EstadDiaPFConsol Set " _
       & "nSaldCRAC =" & rs!nSaldCRAC & " " _
       & "Where cCodAge='" & lsCodAge & "' AND Datediff(Day,dEstadPF,'" & Format(rs!dEstadMens, gsFormatoFecha) & "')=0 and cMoneda ='" & rs!cmoneda & "'"
   poCon.Ejecutar sql
   rs.MoveNext
   Me.Estado.Panels(2).Text = "Agencia : " & lsCodAge
   DoEvents
Loop

rs.Close
Set rs = Nothing
poCon.CierraConexion
Set poCon = Nothing
End Sub
Private Sub EstadCTS(dbConexion As ADODB.Connection, lsCodAge As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim I As Long
Dim m As Long
Dim poCon As New DConecta
poCon.AbreConexion

sql = "SELECT * FROM EstadDiaCTS WHERE datediff(d,dEstadCTS, '" & Format(txtDesde, gsFormatoFecha) & "') <= 0  AND datediff(d,dEstadCTS,'" & Format(txtHasta, gsFormatoFecha) & "') >= 0 "
m = 0
rs.Open sql, dbConexion, adOpenStatic, adLockOptimistic, adCmdText
Do While Not rs.EOF
    m = m + 1
    Me.Estado.Panels(1).Text = "CTS al : " & Format(rs!dEstadCTS, "dd/mm/yyyy") & " Total :" & m
    If ExisteEstadistica(rs!dEstadCTS, "CTS", lsCodAge, , Trim(rs!cmoneda)) = True Then
        'dEstadCTS,cMoneda, nNumAperCTS, nMonAperCTS, nNumCancCTS, nMonCancCTS
        'nNumRetCTS, nMonRetCTS, nRetIntCTS, nNumDepCTS, nMonDepCTS, nNumVigCTS
        'nIntCapCTS, nSaldoCTS, nSaldEmp, cCodUsu, nMonChqVal, cCodAge
        
        sql = "UPDATE EstadDiaCTSConsol Set " _
            & "nNumAperCTS=" & rs!nNumAperCTS & "," _
            & "nMonAperCTS=" & rs!nMonAperCTS & "," _
            & "nNumCancCTS=" & rs!nNumCancCTS & "," _
            & "nMonCancCTS=" & rs!nMonCancCTS & "," _
            & "nNumRetCTS=" & rs!nNumRetCTS & "," _
            & "nMonRetCTS=" & rs!nMonRetCTS & "," _
            & "nRetIntCTS=" & rs!nRetIntCTS & "," _
            & "nNumDepCTS=" & rs!nNumDepCTS & "," _
            & "nMonDepCTS=" & rs!nMonDepCTS & "," _
            & "nNumVigCTS=" & rs!nNumVigCTS & "," _
            & "nIntCapCTS=" & rs!nIntCapCTS & "," _
            & "nSaldoCTS=" & rs!nSaldoCTS & "," _
            & "nSaldEmp=" & rs!nSaldEmp & "," _
            & "nMonChq=" & rs!nMonChq & "," _
            & "cCodUsu='" & rs!cCodUsu & "'," _
            & "nMonChqVal=" & rs!nMonChqVal & " " _
            & "Where cCodAge='" & lsCodAge & "' AND Datediff(Day,dEstadCTs,'" & Format(rs!dEstadCTS, gsFormatoFecha) & "')=0 and cMoneda ='" & rs!cmoneda & "'"
        
        
        'sql = "UPDATE EstadDiaCTSConsol Set " _
            & "nNumAperPF=" & rs!nNumAperPF & "," _
            & "nMonAperPF=" & rs!nMonAperPF & "," _
            & "nNumCancPF=" & rs!nNumCancPF & "," _
            & "nMonCancPF=" & rs!nMonCancPF & "," _
            & "nNumMovPF=" & rs!nNumMovPF & "," _
            & "nMonMovPF=" & rs!nMonMovPF & "," _
            & "nNumVigPF=" & rs!nNumVigPF & "," _
            & "nIntCapPF=" & rs!nSaldoAntPF & "," _
            & "nSaldoPF=" & rs!nSaldoPF & "," _
            & "cCodUsu='" & rs!cCodUsu & "'," _
            & "nMonChqVal=" & rs!nMonChqVal & " " _
            & "Where cCodAge=" & lsCodAge & " AND Datediff(Day,dEstadCTS,'" & Format(rs!dEstadCTS, gsFormatoFecha) & "')=0 and cMoneda ='" & rs!cMoneda & "'"
            
       poCon.Ejecutar sql
            
    Else
        
           sql = "INSERT INTO EstadDiaCTSConsol(dEstadCTS,cMoneda, nNumAperCTS, nMonAperCTS, nNumCancCTS, nMonCancCTS," _
            & "nNumRetCTS, nMonRetCTS, nRetIntCTS, nNumDepCTS, nMonDepCTS, nNumVigCTS," _
            & "nIntCapCTS, nSaldoCTS, nSaldEmp,nMonChq,cCodUsu, nMonChqVal , cCodAge) " _
            & "Values ('" & Format(rs!dEstadCTS, "mm/dd/yyyy hh:mm:ss AMPM") & "','" & rs!cmoneda & "'," & rs!nNumAperCTS & "," _
            & rs!nMonAperCTS & "," & rs!nNumCancCTS & "," & rs!nMonCancCTS & "," & rs!nNumRetCTS & "," & rs!nMonRetCTS & "," _
            & rs!nRetIntCTS & "," & rs!nNumDepCTS & "," & rs!nMonDepCTS & "," _
            & rs!nNumVigCTS & "," & rs!nIntCapCTS & "," & rs!nSaldoCTS & "," _
            & rs!nSaldEmp & "," & rs!nMonChq & ",'" _
            & rs!cCodUsu & "'," & rs!nMonChqVal & ",'" & lsCodAge & "')"
            
       poCon.Ejecutar sql
    End If
    rs.MoveNext
    Me.Estado.Panels(2).Text = "Agencia : " & lsCodAge
    DoEvents
Loop
RSClose rs
poCon.CierraConexion
Set poCon = Nothing
       
End Sub

Private Sub EstadCreditos(dbConexion As ADODB.Connection, lsCodAge As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim I As Long
Dim m As Long

Dim lsStringBaseContab As String
Dim oConec As New DConecta
Dim poCon As New DConecta

poCon.AbreConexion
lsStringBaseContab = poCon.ServerName & "." & poCon.DatabaseName & ".dbo."

sql = "DELETE ColocEstadDiaCred WHERE cCodAge='" & lsCodAge & "' AND dEstad BETWEEN '" & Format(txtDesde, gsFormatoFecha) & "' AND '" & Format(txtHasta, gsFormatoFecha) & " 23:59'"
poCon.Ejecutar sql

sql = "INSERT " & lsStringBaseContab & "EstadDiaCredConsol SELECT *, '" & lsCodAge & "' " _
    & " FROM EstadDiaCred WHERE dFecha BETWEEN '" & Format(txtDesde, gsFormatoFecha) & "' AND '" & Format(txtHasta, gsFormatoFecha) & " 23:59'"
dbConexion.Execute sql
Me.Estado.Panels(1).Text = "Créditos "

'sql = "SELECT * FROM EstadDiaCred WHERE dFecha BETWEEN '" & Format(txtDesde, gsFormatoFecha) & "' AND '" & Format(txtHasta, gsFormatoFecha) & " 23:59'"
'm = 0
'rs.Open sql, dbConexion, adOpenStatic, adLockOptimistic, adCmdText
'Do While Not rs.EOF
'    m = m + 1
'    Me.estado.Panels(1).Text = "Créditos al : " & Format(rs!dFecha, "dd/mm/yyyy") & " Total :" & m
'
'    If ExisteEstadistica(rs!dFecha, "Creditos", lsCodAge, Trim(rs!cCodLinCred)) = True Then
'
'        sql = "UPDATE EstadDiaCredConsol Set " _
'            & "nMontoDesembN=" & rs!nMontoDesembN & "," _
'            & "nMontoDesembR=" & rs!nMontoDesembR & "," _
'            & "nMontoRefinan=" & rs!nMontoRefinan & "," _
'            & "nMontoJudic=" & rs!nMontoJudic & "," _
'            & "nCapPag=" & rs!nCapPag & "," _
'            & "nIntPag=" & rs!nIntPag & "," _
'            & "nMoraPag=" & rs!nMoraPag & "," _
'            & "nOtrPag=" & rs!nOtrPag & "," _
'            & "nSaldoCap=" & rs!nSaldoCap & "," _
'            & "nNumDesembN=" & rs!nNumDesembN & "," _
'            & "nNumDesembR=" & rs!nNumDesembR & "," _
'            & "nNumCanc =" & rs!nNumCanc & "," _
'            & "nNumRef=" & rs!nNumRef & "," _
'            & "nNumJud=" & rs!nNumJud & "," _
'            & "nNumSaldos=" & rs!nNumSaldos & "," _
'            & "nNumPagos=" & rs!nNumPagos & " " _
'            & "Where cCodAge='" & lsCodAge & "' AND dFecha='" & Format(rs!dFecha, gsFormatoFecha) & "' and cCodLinCred ='" & rs!cCodLinCred & "'"
'
'       poCon.Ejecutar sql
'
'    Else
'
'        sql = "INSERT INTO EstadDiaCredConsol(dFecha, cCodLinCred, nMontoDesembN, nMontoDesembR, nMontoRefinan, nMontoJudic," _
'            & "nCapPag, nIntPag, nMoraPag, nOtrPag, nSaldoCap, nNumDesembN, nNumDesembR," _
'            & "nNumCanc , nNumRef, nNumJud, nNumSaldos, nNumPagos, cCodAge) " _
'            & "Values ('" & Format(rs!dFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "','" & rs!cCodLinCred & "'," & rs!nMontoDesembN & "," _
'            & rs!nMontoDesembR & "," & rs!nMontoRefinan & "," & rs!nMontoJudic & "," _
'            & rs!nCapPag & "," & rs!nIntPag & "," & rs!nMoraPag & "," _
'            & rs!nOtrPag & "," & rs!nSaldoCap & "," _
'            & rs!nNumDesembN & "," & rs!nNumDesembR & "," _
'            & rs!nNumCanc & "," & rs!nNumRef & "," _
'            & rs!nNumJud & "," & rs!nNumSaldos & "," & rs!nNumPagos & ",'" _
'            & lsCodAge & "')"
'
'       poCon.Ejecutar sql
'    End If
'    rs.MoveNext
'    Me.estado.Panels(2).Text = "Agencia : " & lsCodAge
'    DoEvents
'Loop
'RSClose rs
poCon.CierraConexion
Set poCon = Nothing
       
End Sub

Private Sub EstadPrendario(dbConexion As ADODB.Connection, lsCodAge As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim I As Long
Dim m As Long
Dim poCon As New DConecta
poCon.AbreConexion
m = 0

'sql = "SELECT * FROM EstadDiaPrenda WHERE dFecha BETWEEN '" & Format(txtDesde, gsFormatoFecha) & "' AND '" & Format(txtHasta, gsFormatoFecha) & " 23:59'"
sql = "SELECT * FROM EstadDiaPrendaBov WHERE dFecha BETWEEN '" & Format(txtDesde, gsFormatoFecha) & "' AND '" & Format(txtHasta, gsFormatoFecha) & " 23:59'"
rs.Open sql, dbConexion, adOpenStatic, adLockOptimistic, adCmdText
Do While Not rs.EOF
    m = m + 1
    Me.Estado.Panels(1).Text = "Prendario al : " & Format(rs!dFecha, "dd/mm/yyyy") & " Total :" & m
    If ExisteEstadistica(rs!dFecha, "Prendario", lsCodAge, , Trim(rs!cmoneda), Right(Trim(rs!cAgeBoveda), 2)) = True Then
        'dFecha, cMoneda, nNumCredVig, nOroVig, nCapVig, nNumCredAdj, nOroAdj, nCapAdj,
        'nNumCredDif , nOroDif, cCodAge , cAgeBoveda

        sql = "UPDATE EstadDiaPrendaConsol Set " _
            & "nNumCredVig=" & rs!nNumCredVig & "," _
            & "nOroVig=" & rs!nOroVig & "," _
            & "nCapVig=" & rs!nCapVig & "," _
            & "nNumCredAdj=" & rs!nNumCredAdj & "," _
            & "nOroAdj=" & rs!nOroAdj & "," _
            & "nCapAdj=" & rs!nCapAdj & "," _
            & "nNumCredDif =" & rs!nNumCredDif & "," _
            & "nOroDif=" & rs!nOroDif & " " _
            & "Where    cCodAge='" & lsCodAge & "' AND " _
            & "         dFecha='" & Format(rs!dFecha, gsFormatoFecha) & "' " _
            & "         and cMoneda ='" & rs!cmoneda & "' and cAgeBoveda='" & Right(rs!cAgeBoveda, 2) & "'"
            
       poCon.Ejecutar sql
    Else
        sql = "INSERT INTO EstadDiaPrendaConsol(dFecha, cMoneda, nNumCredVig, nOroVig, nCapVig, nNumCredAdj, nOroAdj, nCapAdj," _
            & "nNumCredDif , nOroDif, cCodAge, cAgeBoveda) " _
            & "Values ('" & Format(rs!dFecha, "mm/dd/yyyy hh:mm:ss AMPM") & "','" & rs!cmoneda & "'," & rs!nNumCredVig & "," _
            & rs!nOroVig & "," & rs!nCapVig & "," & rs!nNumCredAdj & "," _
            & rs!nOroAdj & "," & rs!nCapAdj & "," & rs!nNumCredDif & "," & rs!nOroDif & ",'" _
            & lsCodAge & "','" & Right(rs!cAgeBoveda, 2) & "')"
            
       poCon.Ejecutar sql
    End If
    rs.MoveNext
    Me.Estado.Panels(2).Text = "Agencia : " & lsCodAge
    DoEvents
Loop
RSClose rs
poCon.CierraConexion
Set poCon = Nothing
       
End Sub
Private Function ExisteEstadistica(lsFecha As String, lsProducto As String, lsCodAge As String, Optional lsLinea As String = "", Optional lsMoneda As String = "", Optional lsAgeBoveda As String = "") As Boolean
Dim sql As String
Dim rs As New ADODB.Recordset
Dim poCon As New DConecta
poCon.AbreConexion

Select Case lsProducto
    Case "Ahorros"
        sql = "Select dEstadAC from EstadDiaACConsol where  datediff(Day,dEstadAC,'" & Format(lsFecha, gsFormatoFecha) & "')=0 and cCodAge='" & lsCodAge & "' and cMoneda='" & lsMoneda & "'"
    Case "Plazo"
        sql = "Select dEstadPF from EstadDiaPFConsol where  Datediff(Day,dEstadPF,'" & Format(lsFecha, gsFormatoFecha) & "')=0 and cCodAge='" & lsCodAge & "' and cMoneda='" & lsMoneda & "'"
    Case "CTS"
        sql = "Select dEstadCTS from EstadDiaCTSConsol where DateDiff(Day,dEstadCTS,'" & Format(lsFecha, gsFormatoFecha) & "')=0 and cCodAge='" & lsCodAge & "' and cMoneda='" & lsMoneda & "'"
    Case "Creditos"
        sql = "Select dFecha from EstadDiaCredConsol where DateDiff(Day,dFecha,'" & Format(lsFecha, gsFormatoFecha) & "')=0 and cCodLinCred='" & Trim(lsLinea) & "' and cCodAge='" & lsCodAge & "'"
    Case "Prendario"
        sql = "Select dFecha from EstadDiaPrendaConsol where Datediff(Day,dFecha,'" & Format(lsFecha, gsFormatoFecha) & "')=0 and cCodAge='" & lsCodAge & "'  and cMoneda='" & lsMoneda & "' and  cAgeBoveda='" & lsAgeBoveda & "'"
End Select
Set rs = poCon.CargaRecordSet(sql)
If Not RSVacio(rs) Then
    ExisteEstadistica = True
Else
    ExisteEstadistica = False
End If
RSClose rs
poCon.CierraConexion
Set poCon = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub lstAgencias_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtDesde.SetFocus
End If
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
