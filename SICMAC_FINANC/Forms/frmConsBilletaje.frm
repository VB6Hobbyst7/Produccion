VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmConsBilletaje 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consolidación de Billetajes de Agencias"
   ClientHeight    =   3225
   ClientLeft      =   2475
   ClientTop       =   2610
   ClientWidth     =   5835
   Icon            =   "frmConsBilletaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   90
      TabIndex        =   9
      Top             =   75
      Width           =   3720
      Begin VB.CheckBox chkTodos 
         Caption         =   "&Todos"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   150
         TabIndex        =   0
         Top             =   0
         Width           =   960
      End
      Begin MSComctlLib.ListView lstAgencias 
         Height          =   2250
         Left            =   120
         TabIndex        =   1
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
      Left            =   4500
      TabIndex        =   2
      Top             =   240
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   405
      Left            =   4028
      TabIndex        =   5
      Top             =   2445
      Width           =   1620
   End
   Begin VB.CommandButton cmdConsolidar 
      Caption         =   "Consoli&dar"
      Height          =   405
      Left            =   4035
      TabIndex        =   4
      Top             =   2040
      Width           =   1620
   End
   Begin MSComctlLib.StatusBar estado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2970
      Width           =   5835
      _ExtentX        =   10292
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
      Left            =   4500
      TabIndex        =   3
      Top             =   600
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
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3870
      TabIndex        =   8
      Top             =   675
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Desde :"
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   3885
      TabIndex        =   7
      Top             =   255
      Width           =   555
   End
End
Attribute VB_Name = "frmConsBilletaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chktodos_Click()
Dim I As Integer
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
Dim oCon As DConecta
Set oCon = New DConecta

oCon.AbreConexion
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
Dim rs As New ADODB.Recordset
Dim sql As String
Dim Total As Long
Dim j As Long
Dim I As Long
Dim CadTemp As String
Dim lbOk As Boolean
Dim oCon As New DConecta
On Error GoTo ConsolidaErr
lbOk = True
For I = 1 To Me.lstAgencias.ListItems.Count
    If lstAgencias.ListItems(I).Checked = True Then
        If oCon.AbreConexion Then 'Remota(Right(lstAgencias.ListItems(I).SubItems(2), 2), True)
            BilletajesAgencia oCon.ConexionActiva, lstAgencias.ListItems(I).SubItems(2), "1"
            BilletajesAgencia oCon.ConexionActiva, lstAgencias.ListItems(I).SubItems(2), "2"
            oCon.CierraConexion
        Else
            lbOk = False
        End If
    End If
Next I
Set oCon = Nothing
If lbOk Then
    MsgBox "Consolidacion Finalizada con éxito", vbInformation, "Aviso"
Else
    MsgBox "Algunas Agencias no se han consolidado de forma satisfactoria" & Chr(13) & "Verifique el Error con sistemas y vuelva a intentar el proceso", vbExclamation, "Aviso"
End If
Exit Sub
ConsolidaErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub
Private Sub BilletajesAgencia(dbConexion As ADODB.Connection, lsCodAge As String, lsMoneda As String)
Dim lsMsgErr As String
Dim sql As String
Dim rs As New ADODB.Recordset
Dim rsE As ADODB.Recordset
Dim I As Long
Dim m As Long
Dim lnTotal As Long
Dim poCon As New DConecta
Dim oEfe As New DMov
Dim lsEfectivoCod As String
Dim lbConsolida   As Boolean
Dim ldFecha       As Date

On Error GoTo BilletajesAgenciaErr
sql = " SELECT  SUM(nCantidad) as nCantidad, nMoneda, cMoneda, cTipMoneda, convert(datetime,CONVERT(CHAR(10),dFecha,101)) AS FECHA " _
    & " From BILLETAJE " _
    & " WHERE dFecha BETWEEN '" & Format(txtDesde, gsFormatoFecha) & "' " _
    & " AND '" & Format(txtHasta, gsFormatoFecha) & " 23:59:59' and cMoneda = '" & lsMoneda & "' " _
    & " GROUP BY nMoneda, cMoneda, cTipMoneda, convert(datetime,CONVERT(CHAR(10),dFecha,101)) " _
    & " ORDER BY FECHA, cMoneda, nMoneda, cTipMoneda "
    
m = 0
rs.CursorLocation = adUseClient
rs.Open sql, dbConexion, adOpenStatic, adLockOptimistic, adCmdText
rs.ActiveConnection = Nothing

oEfe.BeginTrans
Do While Not rs.EOF
    lbConsolida = True
    ldFecha = rs!Fecha
    Set rsE = oEfe.CargaBilletaje(Format(rs!Fecha, gsFormatoMovFecha), Trim(rs!cmoneda), Right(lsCodAge, 2), "")
    If rsE.EOF Then
        gsMovNro = oEfe.GeneraMovNro(rs!Fecha, lsCodAge, "BOVE")
        oEfe.InsertaMov gsMovNro, gOpeHabBoveRegEfect, "Declaración de Billetaje Ag. " & lsCodAge & " Moneda: " & rs!cmoneda & " del día " & rs!Fecha, gMovEstContabNoContable, gMovFlagVigente
        gnMovNro = oEfe.GetnMovNro(gsMovNro)
    Else
        gsMovNro = rsE!cMovNro
        gnMovNro = rsE!nMovNro
        If MsgBox("Consolidación de Billetaje de Agencia " & lsCodAge & " ya fue realizada ¿Desea volver a consolidar? ", vbQuestion + vbYesNo, "¡Aviso!") = vbNo Then
            RSClose rsE
            lbConsolida = False
        End If
        oEfe.EliminaMovUserEfectivo rsE!nMovNro
    End If
    
    Do While rs!Fecha = ldFecha
        If lbConsolida Then
            m = m + 1
            Estado.Panels(1).Text = "Agencia : " & lsCodAge & " al " & Format(rs!Fecha, "dd/mm/yyyy") & " Total :" & m
        'nMovNro     cUser cEfectivoCod nMonto
            lsEfectivoCod = rs!cmoneda & rs!cTipMoneda
            Select Case rs!nMoneda
                    Case 0.05: lsEfectivoCod = lsEfectivoCod & "001"
                    Case 0.1: lsEfectivoCod = lsEfectivoCod & "002"
                    Case 0.2: lsEfectivoCod = lsEfectivoCod & "003"
                    Case 0.5: lsEfectivoCod = lsEfectivoCod & "004"
                    Case 1: lsEfectivoCod = lsEfectivoCod & "005"
                    Case 2: lsEfectivoCod = lsEfectivoCod & "006"
                    Case 5: lsEfectivoCod = lsEfectivoCod & "007"
                    Case 10: lsEfectivoCod = lsEfectivoCod & "008"
                    Case 20: lsEfectivoCod = lsEfectivoCod & "009"
                    Case 50: lsEfectivoCod = lsEfectivoCod & "010"
                    Case 100: lsEfectivoCod = lsEfectivoCod & "011"
                    Case 200: lsEfectivoCod = lsEfectivoCod & "012"
            End Select
            If rs!nCantidad <> 0 Then
'                Set rsE = oEfe.CargaBilletaje(Format(rs!Fecha, gsFormatoMovFecha), Trim(rs!cMoneda), Right(lsCodAge, 2), "", lsEfectivoCod)
'                If Not rsE.EOF = True Then
'                    oEfe.ActualizaMovUserEfectivo gnMovNro, "BOVE", lsEfectivoCod, rs!nMoneda * rs!nCantidad
'                Else
                    oEfe.InsertaMovUserEfectivo gnMovNro, "BOVE", lsEfectivoCod, rs!nMoneda * rs!nCantidad
'                End If
'                RSClose rsE
            End If
        End If
        rs.MoveNext
        Me.Estado.Panels(2).Text = "Agencia : " & lsCodAge
        DoEvents
        If rs.EOF Then
            Exit Do
        End If
    Loop
    
Loop
oEfe.CommitTrans
RSClose rs
Exit Sub
BilletajesAgenciaErr:
 lsMsgErr = Err.Description
 oEfe.RollbackTrans
 MsgBox TextErr(lsMsgErr), vbInformation, "Aviso"
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim oCon As DConecta
Set oCon = New DConecta

oCon.CierraConexion
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
