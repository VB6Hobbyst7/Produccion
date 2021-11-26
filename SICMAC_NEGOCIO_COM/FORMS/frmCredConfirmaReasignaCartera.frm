VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCredConfirmaReasignaCartera 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirmación de Reasignación de Cartera"
   ClientHeight    =   4845
   ClientLeft      =   960
   ClientTop       =   1860
   ClientWidth     =   8895
   Icon            =   "frmCredConfirmaReasignaCartera.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRefresca 
      Caption         =   "&Refrescar [Alt+R]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5160
      TabIndex        =   5
      Top             =   4395
      Width           =   1740
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
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
      Height          =   390
      Left            =   1440
      TabIndex        =   4
      Top             =   4395
      Width           =   1260
   End
   Begin VB.Frame fraDetalle 
      Height          =   4155
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   8715
      Begin MSComctlLib.ListView LstCreditos 
         Height          =   3735
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   1287
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Analista Inicial"
            Object.Width           =   4339
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Analista Final"
            Object.Width           =   4339
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nº Créditos"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nº Clientes"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Saldo Vigente"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Saldo Mora"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Saldo Cartera"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "cMovNroReg"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Seleccione un item y de doble clic para activar [Rechazar] y [Aceptar]"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
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
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   4395
      Width           =   1260
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7860
      TabIndex        =   0
      Top             =   4395
      Width           =   1020
   End
End
Attribute VB_Name = "frmCredConfirmaReasignaCartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''***************************************************************************
'''*** Nombre: frmCredConfirmaReasignaCartera
'''*** Creado por: PEAC - Pedro J. Acuña V.
'''*** Fecha: 16/05/2013 09:59 am
'''*** Utilidad: Permitir al Jefe de Negocios Territoriales Aceptar o Rechazar
'''***           las transferencias de cartera de un analista a otro.
'''***************************************************************************

Option Explicit
Dim nLima As Integer
Dim nSaldosAho As Integer
Dim fbPersNatural As Boolean
Dim sCalifiFinal As String
Dim lsPersTDoc As String
Dim sPersCod As String
Dim lsOpeCod As String
Dim nTipo As Integer 'FRHU20140220 RQ14011

Private Sub MuestraListaReasignacionCarteraCreditos(ByVal pRs As ADODB.Recordset)
Dim L As ListItem
Dim nEstado As Integer

On Error GoTo ERRORBuscaCreditos
    LstCreditos.ListItems.Clear
    
    If Not (pRs.EOF And pRs.BOF) Then
        LstCreditos.Enabled = True
    End If
    
    Do While Not pRs.EOF
        Set L = LstCreditos.ListItems.Add(, , pRs.Bookmark)
        
        L.SubItems(1) = pRs!Usuario
        L.SubItems(2) = pRs!ana_ini
        L.SubItems(3) = pRs!ana_final
        L.SubItems(4) = pRs!num_cred
        L.SubItems(5) = pRs!num_cli
        L.SubItems(6) = pRs!saldoK
        L.SubItems(7) = pRs!saldomora
        L.SubItems(8) = pRs!saldoCartera
        L.SubItems(9) = pRs!cmovnroreg
        
        pRs.MoveNext
    Loop
    
    Exit Sub
    
ERRORBuscaCreditos:
    MsgBox err.Description, vbInformation, "Aviso"

End Sub

Private Sub cmdAceptar_Click()

    If MsgBox("¿Está seguro de ACEPTAR la transferencia de cartera de" & Chr(10) & _
               Me.LstCreditos.SelectedItem.SubItems(2) & " a " & Me.LstCreditos.SelectedItem.SubItems(3) & "?.", vbYesNo + vbQuestion, "Pregunta") = vbNo Then
        Exit Sub
    End If
    
    If VerificaReasignacion(Me.LstCreditos.SelectedItem.SubItems(9)) = 1 Then
        MsgBox "Esta reasignación de cartera fue cancelada por el usuario solicitante.", vbInformation + vbOKOnly, "Atención"
        Call ObtieneReasignacionCartera
        Exit Sub
    End If
    Call CambiaEstadoReasignaCartera(Me.LstCreditos.SelectedItem.SubItems(9), 1)
    Call ObtieneReasignacionCartera
    MsgBox "Se ACEPTO la transferencia de cartera.", vbOKOnly + vbInformation, "Atención"

    Me.cmdRechazar.Enabled = False
    Me.cmdAceptar.Enabled = False
End Sub

Private Sub cmdRechazar_Click()

    If MsgBox("¿Está seguro de RECHAZAR la transferencia de cartera de" & Chr(10) & _
               Me.LstCreditos.SelectedItem.SubItems(2) & " a " & Me.LstCreditos.SelectedItem.SubItems(3) & "?.", vbYesNo + vbQuestion, "Pregunta") = vbNo Then
        Exit Sub
    End If
    
    If VerificaReasignacion(Me.LstCreditos.SelectedItem.SubItems(9)) = 1 Then
        MsgBox "Esta reasignación de cartera fue cancelada por el usuario solicitante.", vbInformation + vbOKOnly, "Atención"
        Call ObtieneReasignacionCartera
        Exit Sub
    End If
    Call CambiaEstadoReasignaCartera(Me.LstCreditos.SelectedItem.SubItems(9), 2)
    Call ObtieneReasignacionCartera
    MsgBox "Se RECHAZO la transferencia de cartera.", vbOKOnly + vbInformation, "Atención"
    Me.cmdRechazar.Enabled = False
    Me.cmdAceptar.Enabled = False
    
End Sub

Private Function VerificaReasignacion(ByVal pcMovNro As String) As Integer
        
    Dim oCredito As COMDCredito.DCOMCredito
    Dim R As ADODB.Recordset

    Set oCredito = New COMDCredito.DCOMCredito
    Set R = oCredito.ConsultaReasignacionCartera(pcMovNro)
    Set oCredito = Nothing
 
    If Not R.EOF Then
        VerificaReasignacion = R!nMovFlag
    End If
    R.Close
    Set R = Nothing
    
End Function


Private Sub CambiaEstadoReasignaCartera(ByVal pcMovNroReg As String, ByVal pnEstadoReasig As Integer)
    Dim oCredito As COMDCredito.DCOMCredito
    
    Dim oMov As COMDMov.DCOMMov
    Set oMov = New COMDMov.DCOMMov
    Dim lsMovNro As String
       
    lsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oMov = Nothing
    
    Set oCredito = New COMDCredito.DCOMCredito
    Call oCredito.MantReasignacionCartera(pcMovNroReg, 9, pnEstadoReasig, lsMovNro)
    Set oCredito = Nothing
    
End Sub

Private Sub cmdRefresca_Click()
    LstCreditos.Enabled = False
    Call ObtieneReasignacionCartera
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
'FRHU 20140220 RQ14011
Public Sub Inicio()
    Dim oCred As New COMDCredito.DCOMCredito
    Dim rs As New ADODB.Recordset
  
    If Not ValidaUsuarioPermitido(gsCodUser) Then
        MsgBox "Esta opcion solo está permitido a Jefes de Negocios Territoriales", vbInformation + vbOKOnly, "Atención"
        Exit Sub
    End If
    
    Set rs = oCred.ValidarJefeTerritorial(gsCodUser)
    If Not (rs.EOF And rs.BOF) Then
        sPersCod = rs!cPersCod
        nTipo = rs!Tipo
    Else
        MsgBox "No tiene ninguna Agencia asignada a su persona: Comunicarse con Jefatura de Creditos", vbInformation, "Atención"
        Exit Sub
    End If
    Me.Show 1
    
End Sub
Private Sub Form_Load()
    LstCreditos.Enabled = False
    Call ObtieneReasignacionCartera
End Sub
'Private Sub Form_Load()
'
'    If Not ValidaUsuarioPermitido(gsCodUser) Then
'        MsgBox "Esta opcion solo está permitido a Jefes de Negocios Territoriales", vbInformation + vbOKOnly, "Atención"
'        Exit Sub
'    End If
'
'    LstCreditos.Enabled = False
'    Call ObtieneReasignacionCartera
'
'End Sub
'FIN FRHU 20140220 RQ14011

Private Sub LstCreditos_DblClick()
    cmdRechazar.Enabled = True
    cmdAceptar.Enabled = True
    cmdRechazar.SetFocus

End Sub

Private Sub ObtieneReasignacionCartera()

    Dim oCreds As COMDCredito.DCOMCredito
    
    Dim rsCred As ADODB.Recordset

    Dim rsCalSBS As ADODB.Recordset
    Dim rsEndSBS As ADODB.Recordset
    Dim rsCalCMAC As ADODB.Recordset
    Dim bExitoBusqueda As Boolean
    Dim dFechaRep As Date
    Dim lsPersDoc As String
    Dim lsPersTDoc As String

    '**************
    
    '**************
    
    Set oCreds = New COMDCredito.DCOMCredito
    'Set rsCred = oCreds.ObtieneParaReasignacionCartera(gsCodAge)
    Set rsCred = oCreds.ObtieneParaReasignacionCartera(gsCodAge, nTipo, sPersCod) 'FRHU 20140220 RQ14011
    Set oCreds = Nothing
    
    Call MuestraListaReasignacionCarteraCreditos(rsCred)

End Sub

Public Function ValidaUsuarioPermitido(ByVal psUser As String) As Boolean
    Dim oCreds As COMDCredito.DCOMCredito
    Dim rsCred As ADODB.Recordset
    Set oCreds = New COMDCredito.DCOMCredito
    Set rsCred = oCreds.ValidaUsuarioReasigCartPermitido(psUser)
    Set oCreds = Nothing
    
    If Not (rsCred.EOF And rsCred.BOF) Then
        ValidaUsuarioPermitido = True
    Else
        ValidaUsuarioPermitido = False
    End If
    Set rsCred = Nothing
End Function
