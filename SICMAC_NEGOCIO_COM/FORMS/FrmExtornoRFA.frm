VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmExtornoRFA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno RFA"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Operaciones de Extorno"
      Height          =   3015
      Left            =   30
      TabIndex        =   3
      Top             =   1620
      Width           =   8625
      Begin SICMACT.Usuario Usuario 
         Left            =   4020
         Top             =   2400
         _ExtentX        =   820
         _ExtentY        =   820
      End
      Begin MSComctlLib.ListView LstOpExt 
         Height          =   1995
         Left            =   195
         TabIndex        =   4
         Top             =   255
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3519
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nª Cuenta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Operacion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hora"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Movimiento"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "nPrePago"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "RFA"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblTotal 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1230
         TabIndex        =   11
         Top             =   2460
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Monto Total:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   900
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
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
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Top             =   1050
         Width           =   1245
      End
      Begin VB.CommandButton cmdExtorno 
         Caption         =   "&Extorno"
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
         Left            =   7080
         TabIndex        =   8
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Height          =   465
         Left            =   1020
         TabIndex        =   7
         Top             =   690
         Width           =   5685
      End
      Begin SICMACT.TxtBuscar TxtBuscar1 
         Height          =   345
         Left            =   1020
         TabIndex        =   5
         Top             =   210
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.CommandButton cmdBusCli 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7080
         TabIndex        =   1
         Top             =   180
         Width           =   1245
      End
      Begin SICMACT.TxtBuscar TxtBuscarUser 
         Height          =   345
         Left            =   4290
         TabIndex        =   12
         Top             =   210
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   12582912
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Usuario:"
         Height          =   195
         Left            =   3300
         TabIndex        =   13
         Top             =   300
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   780
         Width           =   600
      End
      Begin VB.Label LblUsu 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   270
         TabIndex        =   2
         Top             =   330
         Width           =   525
      End
   End
End
Attribute VB_Name = "FrmExtornoRFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Extorno RFA
'Usuario:LMMD
'Ultima Actualizacion:23/11/2004

Dim cPersCod As String
Dim oGen As COMDConstSistema.DCOMGeneral

Private Sub cmdBusCli_Click()
    If cPersCod <> "" Then 'And TxtBuscarUser.Text <> "" Then
     LstOpExt.ListItems.Clear
     lblTotal.Caption = "0.00"
         CargarExtorno
    Else
        MsgBox "Debe seleccionar un cliente para la " & vbCrLf & _
               "busqueda de su movimiento de rfa", vbInformation, "AVISO"
    End If
End Sub

Private Sub cmdExtorno_Click()
Dim psDescrip As String
Dim R As ADODB.Recordset
Dim RCap As ADODB.Recordset
Dim nSaldoCtaAho As Double
Dim rs As ADODB.Recordset
Dim odRFa As COMDCredito.DCOMRFA

    If LstOpExt.ListItems.Count <= 0 Then
        MsgBox "No existen Operaciones para Extornar", vbInformation, "Aviso"
        Exit Sub
    End If
    
   
    If MsgBox("Se va a Extornar el movimiento : " & LstOpExt.SelectedItem.ListSubItems(3).Text, vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

   Set odRFa = New COMDCredito.DCOMRFA
   Call odRFa.ExtornarPago(LstOpExt.SelectedItem.SubItems(3), gdFecSis, TxtBuscarUser.Text, gsCodAge)
   Set odRFa = Nothing
   
   Call ImpreBoleta
   
    MsgBox "Extorno Finalizado", vbInformation, "Aviso"
    Call cmdBusCli_Click
    Exit Sub

End Sub

Private Sub CmdSalir_Click()
     Unload Me
End Sub

Private Sub Form_Load()
Usuario.Inicio gsCodUser
Set oGen = New COMDConstSistema.DCOMGeneral
TxtBuscarUser.psRaiz = "USUARIOS"
TxtBuscarUser.rs = oGen.GetUserAreaAgencia("026", gsCodAge, "", False)
TxtBuscarUser.Enabled = True
TxtBuscarUser.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oGen = Nothing
End Sub

Private Sub TxtBuscar1_EmiteDatos()
   Dim objDRFA As COMDCredito.DCOMRFA
   Dim rs As ADODB.Recordset
   
   Set objDRFA = New COMDCredito.DCOMRFA
   Set rs = objDRFA.BuscarPersona(TxtBuscar1.Text)
   Set objDRFA = Nothing
   If Not rs.EOF And Not rs.BOF Then
        cPersCod = rs!cPersCod
        txtNombre = Trim(rs!cPersNombre)
   End If
   
   Set rs = Nothing
End Sub

Sub CargarExtorno()
    Dim rs As ADODB.Recordset
    Dim objDRFA As New COMDCredito.DCOMRFA
    Dim nMonto As Currency
    
    Dim L As ListItem

    On Error GoTo ErrHandler
    LstOpExt.ListItems.Clear
    
    Set objDRFA = New COMDCredito.DCOMRFA
    Set rs = objDRFA.ObtenerExtornoRFA(cPersCod, TxtBuscarUser.Text, gdFecSis, gsCodAge)
    Set objDRFA = Nothing
    
    nMonto = 0
    
    Do Until rs.EOF
        Set L = LstOpExt.ListItems.Add(, , rs!cCtaCod)
        L.SubItems(1) = rs!cOpecod
        L.SubItems(2) = rs!cHora
        L.SubItems(3) = rs!nmovnro
        L.SubItems(4) = Format(rs!nMonto, "#0.00")
        L.SubItems(5) = rs!cUsuario
        L.SubItems(6) = IIf(IsNull(rs!nPrepago), "", rs!nPrepago)
        L.SubItems(7) = rs!cRFA
        
        nMonto = nMonto + rs!nMonto
        rs.MoveNext
    Loop
    Set rs = Nothing
    lblTotal = Format(nMonto, "0.00")
    Exit Sub
ErrHandler:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not objDRFA Is Nothing Then Set objDRFA = Nothing
    MsgBox "Error al momento de cargar el extorno", vbInformation, "AVISO"
End Sub

Sub ImpreBoleta()
Dim oCredDoc As COMNCredito.NCOMCredDoc
Dim oPrevio As Previo.clsPrevio
Dim sImpresion As String

Set oCredDoc = New COMNCredito.NCOMCredDoc

sImpresion = oCredDoc.ImprimeBoletaRFA(CLng(LstOpExt.ListItems(1).SubItems(3)), txtNombre.Text, gsNomAge, gdFecSis, _
                                gsCodUser, sLpt, gsInstCmac, gsCodCMAC)
Set oCredDoc = Nothing

Set oPrevio = New Previo.clsPrevio
Call oPrevio.Show(sImpresion, "Extorno de Crédito RFA")
Set oPrevio = Nothing
End Sub

'Sub ImprimeBoleta()
'    Dim odRFa As COMDCredito.DCOMRFA
'    Dim i As Integer
'    Dim sLcCtaCod As String
'    Dim sTempcCtaCod As String
'    Dim nCuotas As Integer
'    Dim oConec As COMConecta.DCOMConecta
'    Dim sSql As String
'    Dim rs As ADODB.Recordset
'    Dim snMonto As Double
'    Dim sUser As String
'    Dim oDCredito As COMDCredito.DCOMCredDoc
'
'    sTempcCtaCod = ""
'    For i = 1 To LstOpExt.ListItems.Count
'        If sTempcCtaCod <> Me.LstOpExt.ListItems(i).Text Then
'            If Len(sTempcCtaCod) = 0 Then
'                sLcCtaCod = LstOpExt.ListItems(i).Text
'            Else
'                sLcCtaCod = sLcCtaCod & "," & LstOpExt.ListItems(i).Text
'            End If
'            sTempcCtaCod = LstOpExt.ListItems(i).Text
'        End If
'    Next i
'
'    Set oDCredito = New COMDCredito.DCOMCredDoc
'    sUser = oDCredito.GetUsuario(LstOpExt.ListItems(1).SubItems(3))
'    Set oDCredito = Nothing
'
'    Set odRFa = New COMDCredito.DCOMRFA
'    nCuotas = odRFa.ObtenNumeroCuotas(LstOpExt.ListItems(1).SubItems(3))
'    Set oConec = New COMConecta.DCOMConecta
'    oConec.AbreConexion
'    sSql = "Select Sum(nMonto) as nMonto,cCtaCod"
'    sSql = sSql & " From MovColDet"
'    sSql = sSql & " Where nMOvNro=" & LstOpExt.ListItems(1).SubItems(3) & " and cOpeCod not like '107[123456789]%'"
'    sSql = sSql & " Group By cCtaCod"
'    sSql = sSql & " Order by cCtaCod"
'
'    Set rs = oConec.CargaRecordSet(sSql)
'    oConec.CierraConexion
'    Set oConec = Nothing
'    sTempcCtaCod = ""
'    Do Until rs.EOF
'        Call odRFa.ImprimeBoletaExtorno("Extorno RFA", rs!cCtaCod, txtNombre, _
'                  gsNomAge, "DOLARES", sGetCuotas(LstOpExt.ListItems(1).SubItems(3), rs!cCtaCod), gdFecSis, Time, "", IIf(IsNull(rs!nMonto), 0, rs!nMonto), 0#, gsCodUser, sLpt, gsInstCmac, gsCodCMAC, sUser)
'        rs.MoveNext
'    Loop
'    Set rs = Nothing
'    'Call odRFA.ImprimeBoletaExtorno("Extorno RFA", LstOpExt.ListItems(1).SubItems(2), txtNombre, _
'     '             gsNomAge, "DOLARES", nCuotas, gdFecSis, Time, "", nCapitalPagado, 0#, gsCodUser, sLpt, gsInstCmac, gsCodCMAC)
'    Set odRFa = Nothing
'
'End Sub
'
'Function sGetCuotas(ByVal pnMovNro, ByVal psCtaCod As String) As String
'   Dim oConec As COMConecta.DCOMConecta
'   Dim sSql As String
'   Dim rs As ADODB.Recordset
'   Dim sCuotas As String
'
'
'   Set oConec = New COMConecta.DCOMConecta
'   oConec.AbreConexion
'   sSql = "Select Distinct nNroCuota"
'   sSql = sSql & " From MovColDet"
'   sSql = sSql & " Where nMovNro=" & pnMovNro & " and cCtaCod='" & psCtaCod & "' and nNroCuota<>0"
'
'   Set rs = oConec.CargaRecordSet(sSql)
'   oConec.CierraConexion
'   Set oConec = Nothing
'
'   Do Until rs.EOF
'    If Len(sCuotas) = 0 Then
'        sCuotas = rs!nNroCuota
'    Else
'        sCuotas = sCuotas & "," & rs!nNroCuota
'    End If
'    rs.MoveNext
'   Loop
'   Set rs = Nothing
'
'   sGetCuotas = sCuotas
'End Function
