VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCapOrdPagProceso 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   Icon            =   "frmOrdPagProceso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefrescar 
      Caption         =   "&Refrescar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5115
      Width           =   900
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8985
      TabIndex        =   5
      Top             =   5130
      Width           =   855
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar y Enviar"
      Height          =   375
      Left            =   7365
      TabIndex        =   4
      Top             =   5130
      Width           =   1515
   End
   Begin VB.Frame fraOrdPag 
      Caption         =   "Orden Pago"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4980
      Left            =   90
      TabIndex        =   7
      Top             =   75
      Width           =   9765
      Begin VB.OptionButton optSeleccion 
         Caption         =   "&Ninguno"
         Height          =   255
         Index           =   1
         Left            =   1455
         TabIndex        =   1
         Top             =   285
         Width           =   1020
      End
      Begin VB.OptionButton optSeleccion 
         Caption         =   "&Todos"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   270
         Width           =   1020
      End
      Begin VB.CommandButton cmdEstado 
         Caption         =   "&Cambio Estado"
         Height          =   375
         Left            =   8130
         TabIndex        =   2
         Top             =   180
         Width           =   1485
      End
      Begin MSComctlLib.ListView lstOrdPag 
         Height          =   4275
         Left            =   105
         TabIndex        =   3
         Top             =   600
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   7541
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483639
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Inicio"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fin"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "# OP"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Titular"
            Object.Width           =   5556
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tipo"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCapOrdPagProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sEstadoOrden As String

Private Function ExistenItemsMarcados() As Boolean
Dim L As ListItem
For Each L In lstOrdPag.ListItems
    If L.Checked Then
        ExistenItemsMarcados = True
        Exit Function
    End If
Next
ExistenItemsMarcados = False
End Function

Private Function GetTitulares(ByVal sCuenta As String) As String
Dim sNombre1 As String * 60, sNombre2 As String * 60, sNombre3 As String * 60
Dim sDocId1 As String * 11, sDocId2 As String * 11, sDocId3 As String * 11
Dim sNom As String * 60, sDoc As String * 11
Dim rsPers As ADODB.Recordset
Dim I As Integer
If AbreConeccion(sCuenta, False, False) Then
    VSQL = "Select P.cNomPers, ISNULL(P.cNudoci,'') cNudoci, ISNULL(P.cNudotr,'') cNudotr, P.cTipPers From " & gcCentralPers & "Persona P INNER JOIN PersCuenta PC " _
        & "ON P.cCodPers = PC.cCodPers Where PC.cCodCta = '" & sCuenta & "' And PC.cRelaCta = 'TI'"
    
    sNombre1 = Space(60)
    sNombre2 = Space(60)
    sNombre3 = Space(60)
    sDocId1 = Space(11)
    sDocId2 = Space(11)
    sDocId3 = Space(11)
    I = 0
    Set rsPers = New ADODB.Recordset
    rsPers.CursorLocation = adUseClient
    rsPers.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
    Set rsPers.ActiveConnection = Nothing
    Do While Not rsPers.EOF
        sNom = PstaNombre(rsPers("cNomPers"), True)
        If rsPers("cTipPers") = "1" Then
            sDoc = rsPers("cNudoci")
        Else
            sDoc = rsPers("cNudotr")
        End If
        I = I + 1
        Select Case I
            Case 1
                sNombre1 = sNom
                sDocId1 = sDoc
            Case 2
                sNombre2 = sNom
                sDocId2 = sDoc
            Case 3
                sNombre3 = sNom
                sDocId3 = sDoc
        End Select
        rsPers.MoveNext
    Loop
    rsPers.Close
    Set rsPers = Nothing
    GetTitulares = sDocId1 & sDocId2 & sDocId3 & sNombre1 & sNombre2 & sNombre3
Else
    GetTitulares = ""
End If
CierraConeccion
End Function

Public Sub Inicia(ByVal sEst As String)
sEstadoOrden = sEst
If sEstadoOrden = "1" Then
    Me.Caption = "Orden Pago - Consolidación y Envío"
    cmdGenerar.Caption = "&Generar y Enviar"
ElseIf sEstadoOrden = "2" Then
    Me.Caption = "Orden Pago - Recepción"
    cmdGenerar.Caption = "&Grabar"
ElseIf sEstadoOrden = "3" Then
    Me.Caption = "Orden Pago - Entrega al Cliente"
    cmdGenerar.Caption = "&Grabar"
End If
cmdRefrescar.Enabled = False
cmdGenerar.Enabled = False
AbreConexion
ObtieneDatosOrdenPago
CierraConexion
Me.Show 1
End Sub

Private Sub ObtieneDatosOrdenPago()
Dim rsOrden As ADODB.Recordset
Dim L As ListItem

If AbreConeccion(gsAgenciaCentralOP & "2321000019", False) Then
    VSQL = "Select OP.cCodCta, OP.nInicio, OP.nFin, OP.dFecha, P.cNomPers, T.cTipo From " _
        & "OrdPagEmision OP INNER JOIN PersCuenta PC INNER JOIN " & gcCentralPers & "Persona P " _
        & "ON PC.cCodPers = P.cCodPers ON OP.cCodCta = PC.cCodCta INNER JOIN OrdPagTarifa T ON " _
        & "SUBSTRING(OP.cCodCta,6,1) = T.cMoneda And (OP.nFin - OP.nInicio + 1) = T.nNumOP " _
        & "Where OP.cEstado = '" & sEstadoOrden & "' And PC.cRelaCta = 'TI' And P.cCodPers IN " _
        & "(Select MAX(PC1.cCodPers) From PersCuenta PC1 Where PC1.cCodCta = OP.cCodCta And PC1.cRelaCta = 'TI') Order by OP.dFecha"
    Set rsOrden = New ADODB.Recordset
    rsOrden.CursorLocation = adUseClient
    rsOrden.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
    Set rsOrden.ActiveConnection = Nothing
    If Not (rsOrden.EOF And rsOrden.BOF) Then
        lstOrdPag.ListItems.Clear
        Do While Not rsOrden.EOF
            Set L = lstOrdPag.ListItems.Add(, , rsOrden("cCOdCta"))
            L.SubItems(1) = rsOrden("nInicio")
            L.SubItems(2) = rsOrden("nFin")
            L.SubItems(3) = Format$(rsOrden("dFecha"), "dd-mmm-yyyy")
            L.SubItems(4) = rsOrden("nFin") - rsOrden("nInicio") + 1
            L.SubItems(5) = PstaNombre(rsOrden("cNomPers"), False)
            L.SubItems(6) = rsOrden("cTipo")
            rsOrden.MoveNext
        Loop
        cmdRefrescar.Enabled = True
        cmdGenerar.Enabled = True
        cmdEstado.Enabled = True
        optSeleccion(0).Enabled = True
        optSeleccion(1).Enabled = True
    Else
        cmdRefrescar.Enabled = True
        cmdGenerar.Enabled = False
        cmdEstado.Enabled = False
        optSeleccion(0).Enabled = False
        optSeleccion(1).Enabled = False
        MsgBox "No Existen Ordenes de Pago para este proceso", vbInformation, "Aviso"
        lstOrdPag.ListItems.Clear
    End If
    
Else
    MsgBox "No es posible conectar con la Agencia Central. Avise al Area de Sistemas", vbExclamation, "Error"
    cmdRefrescar.Enabled = False
    cmdGenerar.Enabled = False
End If
CierraConeccion
End Sub

Private Sub cmdEstado_Click()
Dim L As ListItem
Dim sEstadoAnterior As String
Dim sCuenta As String
Dim sInicio As String, sFecha As String

If MsgBox("Desea cambiar el estado al siguiente item?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If
sFecha = FechaHora(gdFecSis)
sEstadoAnterior = Trim(CInt(sEstadoOrden) - 1)
Set L = lstOrdPag.SelectedItem
sCuenta = L.Text
sInicio = L.SubItems(1)
AbreConexion
If AbreConeccion(gsAgenciaCentralOP & "2321000019", False, False) Then
    dbCmactN.BeginTrans
    VSQL = "Update OrdPagEmision Set cEstado = '" & sEstadoAnterior & "' Where cCodCta = '" & sCuenta & "' " _
        & "And cEstado = '" & sEstadoOrden & "' And nInicio = " & sInicio
    dbCmactN.Execute VSQL
    VSQL = "Insert OrdPagEstado (cCodCta,nInicio,dFecha,cEstado,cCodUsu) " _
        & "Values ('" & sCuenta & "'," & sInicio & ",'" & sFecha & " ','" & sEstadoAnterior & "','" & gsCodUser & "')"
    dbCmactN.Execute VSQL
    dbCmactN.CommitTrans
End If
CierraConeccion
ObtieneDatosOrdenPago
CierraConexion
End Sub

Private Sub cmdRefrescar_Click()
AbreConexion
ObtieneDatosOrdenPago
CierraConexion
End Sub

Private Sub cmdGenerar_Click()
Dim L As ListItem
Dim sCuenta As String
Dim sInicio As String * 8, sFin As String * 8
Dim sNumTal As String * 2
Dim sOficCentral As String * 3
Dim sAgencia As String * 3
Dim sTitular As String
Dim sCadena As String, sFecha As String
Dim sMoneda As String * 1
Dim sCadImp As String
Dim sOPTipoTalon  As String
Dim dbAux As ADODB.Connection
If Not ExistenItemsMarcados() Then
    MsgBox "Debe seleccionar algún item para este proceso.", vbInformation, "Aviso"
    lstOrdPag.SetFocus
    Exit Sub
End If

sFecha = FechaHora(gdFecSis)
sCadImp = ""
AbreConexion
Select Case sEstadoOrden
    Case "1" 'Consolidación y Envío
        Dim sRuta As String
        sRuta = App.Path & "\SPOOLER\" & "CMT" & Right$(Year(gdFecSis), 1) & Format$(Month(gdFecSis), "00") & Format$(Day(gdFecSis), "00") & ".TXT"
        Open sRuta For Output As #1
        If AbreConeccion(gsAgenciaCentralOP & "2321000019", False, False) Then
            Set dbAux = dbCmactN
            dbAux.BeginTrans
            For Each L In lstOrdPag.ListItems
                If L.Checked Then
                    sCuenta = Trim(L.Text)
                    sInicio = L.SubItems(1)
                    sFin = L.SubItems(2)
                    sNumTal = "01"
                    sOPTipoTalon = L.SubItems(6)
                    sAgencia = "0" & Mid$(sCuenta, 1, 2)
                    sOficCentral = "0" & gsOficinaCentral
                    sMoneda = Mid$(sCuenta, 6, 1)
                    sTitular = GetTitulares(sCuenta)
                    If sTitular <> "" Then
                        sCadena = sInicio & Mid$(sCuenta, 3, 10) & sTitular & sAgencia & sOficCentral & sNumTal & gsOPTipoProd & sOPTipoTalon & sMoneda
                        sCadImp = sCadImp & sCadena
                        Print #1, sCadena
                        VSQL = "Update OrdPagEmision Set cEstado = '2' Where cCodCta = '" & sCuenta & "' " _
                            & "And cEstado = '" & sEstadoOrden & "' And nInicio = " & sInicio
                        dbAux.Execute VSQL
                        VSQL = "Insert OrdPagEstado (cCodCta,nInicio,dFecha,cEstado,cCodUsu) " _
                            & "Values ('" & sCuenta & "'," & sInicio & ",'" & sFecha & " ','2','" & gsCodUser & "')"
                        dbAux.Execute VSQL
                    Else
                        MsgBox "No se encontraron titulares para la cuenta " & sCuenta, vbInformation, "Aviso"
                    End If
                End If
            Next
            dbAux.CommitTrans
            Close #1
            MsgBox "Archivo Creado : " & sRuta, vbInformation, "Aviso"
            ObtieneDatosOrdenPago
        End If
        CierraConeccion
    Case "2" 'Recepción
        If AbreConeccion(gsAgenciaCentralOP & "2321000019", False, False) Then
            dbCmactN.BeginTrans
            For Each L In lstOrdPag.ListItems
                If L.Checked Then
                    sCuenta = Trim(L.Text)
                    sInicio = L.SubItems(1)
                    VSQL = "Update OrdPagEmision Set cEstado = '3' Where cCodCta = '" & sCuenta & "' " _
                        & "And cEstado = '2' And nInicio = " & sInicio
                    dbCmactN.Execute VSQL
                    VSQL = "Insert OrdPagEstado (cCodCta,nInicio,dFecha,cEstado,cCodUsu) " _
                        & "Values ('" & sCuenta & "'," & sInicio & ",'" & sFecha & " ','3','" & gsCodUser & "')"
                    dbCmactN.Execute VSQL
                End If
            Next
            dbCmactN.CommitTrans
            ObtieneDatosOrdenPago
        End If
        CierraConeccion
        
    Case "3" 'Entrega al Cliente
        If AbreConeccion(gsAgenciaCentralOP & "2321000019", False, False) Then
            dbCmactN.BeginTrans
            For Each L In lstOrdPag.ListItems
                If L.Checked Then
                    sCuenta = Trim(L.Text)
                    sInicio = L.SubItems(1)
                    VSQL = "Update OrdPagEmision Set cEstado = '4' Where cCodCta = '" & sCuenta & "' " _
                        & "And cEstado = '3' And nInicio = " & sInicio
                    dbCmactN.Execute VSQL
                    VSQL = "Insert OrdPagEstado (cCodCta,nInicio,dFecha,cEstado,cCodUsu) " _
                        & "Values ('" & sCuenta & "'," & sInicio & ",'" & sFecha & " ','4','" & gsCodUser & "')"
                    dbCmactN.Execute VSQL
                End If
            Next
            dbCmactN.CommitTrans
            ObtieneDatosOrdenPago
        End If
        CierraConeccion
End Select
CierraConexion
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub optSeleccion_Click(Index As Integer)
Dim L As ListItem
Select Case Index
    Case 0
        For Each L In lstOrdPag.ListItems
            L.Checked = True
        Next
    Case 1
        For Each L In lstOrdPag.ListItems
            L.Checked = False
        Next
End Select
End Sub
