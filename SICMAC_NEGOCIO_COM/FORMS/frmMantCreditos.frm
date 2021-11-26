VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMantCreditos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Estado de Credito"
   ClientHeight    =   5340
   ClientLeft      =   960
   ClientTop       =   1860
   ClientWidth     =   9015
   Icon            =   "frmMantCreditos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraDetalle 
      Height          =   3555
      Left            =   105
      TabIndex        =   11
      Top             =   1185
      Width           =   8835
      Begin MSComctlLib.ListView LstCreditos 
         Height          =   3255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5741
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Item"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Num. Credito"
            Object.Width           =   2575
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fecha Solicitud"
            Object.Width           =   2170
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Estado"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Agencia"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Tipo Crédito"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Moneda"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Monto"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Analista"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Restaurar"
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
      Left            =   6480
      TabIndex        =   8
      Top             =   4875
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
      TabIndex        =   1
      Top             =   4875
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   8820
      Begin VB.CommandButton CmdBuscar 
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
         Height          =   345
         Left            =   7530
         TabIndex        =   10
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label LblPersCod 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   870
         TabIndex        =   9
         Top             =   262
         Width           =   1755
      End
      Begin VB.Label lblDocJur 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   4440
         TabIndex        =   7
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label lblDocNat 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1395
         TabIndex        =   6
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label lblNomPers 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2760
         TabIndex        =   5
         Top             =   262
         Width           =   4620
      End
      Begin VB.Label lblDocJuridico 
         AutoSize        =   -1  'True
         Caption         =   "RUC :"
         Height          =   195
         Left            =   3045
         TabIndex        =   4
         Top             =   660
         Width           =   435
      End
      Begin VB.Label lblDocNatural 
         AutoSize        =   -1  'True
         Caption         =   "Doc. Identidad :"
         Height          =   195
         Left            =   105
         TabIndex        =   3
         Top             =   660
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   315
         Width           =   570
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Seleccione un item y de doble clic para activar Restaurar"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   4800
      Width           =   4035
   End
End
Attribute VB_Name = "frmMantCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nLima As Integer
Dim nSaldosAho As Integer
Dim fbPersNatural As Boolean
Dim sCalifiFinal As String
Dim lsPersTDoc As String
Dim sPersCod As String
Dim lsOpeCod As String

Private Sub MuestraCreditos(ByVal psPersCod As String, ByVal pRs As ADODB.Recordset)
Dim L As ListItem
Dim nEstado As Integer

On Error GoTo ERRORBuscaCreditos
    LstCreditos.ListItems.Clear
    
    If Not (pRs.EOF And pRs.BOF) Then
        LstCreditos.Enabled = True
    End If
    
    Do While Not pRs.EOF
        Set L = LstCreditos.ListItems.Add(, , pRs.Bookmark)
        
        L.SubItems(1) = pRs!cCtaCod
        L.SubItems(2) = pRs!FecSoli
        
        L.SubItems(3) = pRs!Estado
        L.SubItems(4) = pRs!Agencia
        
        L.SubItems(5) = pRs!TipoCredito
        L.SubItems(6) = pRs!Moneda
        L.SubItems(7) = pRs!Monto
        
        L.SubItems(8) = pRs!Analista
        
        pRs.MoveNext
    Loop
    
    Exit Sub
    
ERRORBuscaCreditos:
    MsgBox Err.Description, vbInformation, "Aviso"

End Sub

Private Sub CmdBuscar_Click()

Dim oPersona As COMDPersona.UCOMPersona

    'AMDO 20130727 TI-ERS086-201
    '    If Not ValidaUsuarioPermitido(gsCodUser) Then
    '        MsgBox "Esta opción solo está permitido a Jefes de Agencia.", vbInformation + vbOKOnly, "Atención"
    '        Exit Sub
    '    End If
    'END AMDO
    

    LstCreditos.Enabled = False
    
    Set oPersona = frmBuscaPersona.Inicio
    If Not oPersona Is Nothing Then
        LblPersCod.Caption = oPersona.sPersCod
        lblNomPers.Caption = oPersona.sPersNombre
        lblDocNat.Caption = Trim(oPersona.sPersIdnroDNI)
        lblDocJur.Caption = Trim(oPersona.sPersIdnroRUC)
        lsPersTDoc = "1"

        If oPersona.sPersPersoneria = "1" Then
            fbPersNatural = True

            If Trim(oPersona.sPersIdnroDNI) = "" Then
                If Not Trim(oPersona.sPersIdnroOtro) = "" Then
                    lblDocNat.Caption = Trim(oPersona.sPersIdnroOtro)
                    lsPersTDoc = Trim(oPersona.sPersTipoDoc) 'aqUI
                End If
            End If

        Else
            fbPersNatural = False
            lsPersTDoc = "3"
        End If
    Else
        Exit Sub
    End If
    sPersCod = oPersona.sPersCod
    Set oPersona = Nothing
        
    If sPersCod <> "" Then
        Call ObtieneListaCreditos(sPersCod, lsPersTDoc)
    End If
End Sub

'Private Sub BuscarPosicionCliente(ByVal psPersCod As String, Optional psPersTDoc As String = "1")
'    Dim oCreds As COMDCredito.DCOMCreditos
'    Dim rsCred As ADODB.Recordset
'    Dim rsAho As ADODB.Recordset
'    Dim rsPig As ADODB.Recordset
'    Dim rsJud As ADODB.Recordset
'    Dim rsCF As ADODB.Recordset
'    Dim rsCom As ADODB.Recordset
'
'    'Se Agrego para la Calificacion RCC
'    Dim rsCalSBS As ADODB.Recordset
'    Dim rsEndSBS As ADODB.Recordset
'    Dim rsCalCMAC As ADODB.Recordset
'    Dim bExitoBusqueda As Boolean
'    Dim dFechaRep As Date
'    Dim lsPersDoc As String '**DAOR 20080410
'    Dim lsPersTDoc As String 'ALPA 20100922
'    Set oCreds = New COMDCredito.DCOMCreditos
'
'    If fbPersNatural Then
'        lsPersDoc = IIf(lblDocNat.Caption = "", Trim(lblDocJur.Caption), Trim(lblDocNat.Caption))
'
'    Else
'        lsPersDoc = IIf(lblDocJur.Caption <= "", Trim(lblDocNat.Caption), Trim(lblDocJur.Caption))
'    End If
'    bExitoBusqueda = oCreds.BuscarPosicionCliente(psPersCod, IIf(Check1.value = 1, True, False), nLima, _
'                                    fbPersNatural, lsPersDoc, rsCred, rsAho, rsPig, rsJud, rsCF, rsCom, dFechaRep, rsCalSBS, rsEndSBS, rsCalCMAC, gdFecSis, 1, psPersTDoc)
'
'    Set oCreds = Nothing
'    Call BuscaCreditos(psPersCod, rsCred)
'    If bExitoBusqueda Then
'        Call BuscaCalificacionRCC(dFechaRep, rsCalSBS, rsEndSBS, rsCalCMAC)
'    End If
'End Sub

'Private Sub BuscaCalificacionRCC(ByVal pdFechaRep As Date, _
'                                ByVal prsCalSBS As ADODB.Recordset, _
'                                ByVal prsEndSBS As ADODB.Recordset, _
'                                ByVal prsCalCMAC As ADODB.Recordset)
'
'Dim iTem As ListItem
'Dim fil As Integer
'Dim lnCorreFinanciero As Long
'
'    lnCorreFinanciero = 0
'    lstFinanciero.ListItems.Clear
'
'    lnCorreFinanciero = lnCorreFinanciero + 1
'    Set iTem = Me.lstFinanciero.ListItems.Add(, , "Calificacion SBS-RCC ")
'    iTem.SubItems(1) = Format(pdFechaRep, "dd/mm/yyyy")
'    iTem.SubItems(2) = ""
'    iTem.SubItems(3) = ""
'    iTem.SubItems(4) = ""
'
'    lnCorreFinanciero = lnCorreFinanciero + 1
'    Set iTem = Me.lstFinanciero.ListItems.Add(, , "Normal")
'    iTem.SubItems(1) = "Potencial"
'    iTem.SubItems(2) = "Deficiente"
'    iTem.SubItems(3) = "Dudoso"
'    iTem.SubItems(4) = "Perdida"
'
'    lstFinanciero.ListItems(1).ForeColor = vbRed
'    lstFinanciero.ListItems(1).Bold = True
'    lstFinanciero.ListItems(2).ListSubItems(2).ForeColor = vbRed
'
'    lstFinanciero.ListItems(2).ForeColor = vbBlue
'    lstFinanciero.ListItems(2).Bold = True
'    For fil = 1 To 4
'        lstFinanciero.ListItems(2).ListSubItems(fil).ForeColor = vbBlue
'        lstFinanciero.ListItems(2).ListSubItems(fil).Bold = True
'    Next
'    'Calificacion SBS
'    If Len(Trim(lblDocNat.Caption)) = 0 And Len(Trim(lblDocJur.Caption)) = 0 Then
'        MsgBox "Cliente no registra documento, Favor Actualizar Datos ", vbInformation, "Aviso"
'        Exit Sub
'    End If
'
'    If prsCalSBS.BOF And prsCalSBS.EOF Then
'        lnCorreFinanciero = lnCorreFinanciero + 1
'        Set iTem = Me.lstFinanciero.ListItems.Add(, , "No Registrado")
'        iTem.SubItems(1) = "No Registrado"
'        iTem.SubItems(2) = "No Registrado"
'        iTem.SubItems(3) = "No Registrado"
'        iTem.SubItems(4) = "No Registrado"
'    Else
'        Do While Not prsCalSBS.EOF
'            lnCorreFinanciero = lnCorreFinanciero + 1
'            Set iTem = Me.lstFinanciero.ListItems.Add(, , prsCalSBS!nNormal & "%")
'            iTem.SubItems(1) = prsCalSBS!nPotencial & "%"
'            iTem.SubItems(2) = prsCalSBS!nDeficiente & "%"
'            iTem.SubItems(3) = prsCalSBS!nDudoso & "%"
'            iTem.SubItems(4) = prsCalSBS!nPerdido & "%"
'
'            lnCorreFinanciero = lnCorreFinanciero + 1
'            Set iTem = Me.lstFinanciero.ListItems.Add(, , "NRO ENTIDADES : ")
'            iTem.SubItems(1) = prsCalSBS!Can_Ents
'            iTem.SubItems(2) = ""
'            iTem.SubItems(3) = ""
'            iTem.SubItems(4) = ""
'            prsCalSBS.MoveNext
'        Loop
'
'        'Endeudamiento SBS
'        lnCorreFinanciero = lnCorreFinanciero + 1
'        Set iTem = Me.lstFinanciero.ListItems.Add(, , "Endeudamiento")
'        lstFinanciero.ListItems(lnCorreFinanciero).ForeColor = vbRed
'        lstFinanciero.ListItems(lnCorreFinanciero).Bold = True
'        iTem.SubItems(1) = ""
'        iTem.SubItems(2) = ""
'        iTem.SubItems(3) = ""
'        iTem.SubItems(4) = ""
'
'        lnCorreFinanciero = lnCorreFinanciero + 1
'        Set iTem = Me.lstFinanciero.ListItems.Add(, , "Directa Soles")
'        iTem.SubItems(1) = "Directa Dolar"
'        iTem.SubItems(2) = "Indirecta Soles"
'        iTem.SubItems(3) = "Indirecta Dolar"
'        iTem.SubItems(4) = ""
'
'        lstFinanciero.ListItems(lnCorreFinanciero).ForeColor = vbBlue
'        lstFinanciero.ListItems(lnCorreFinanciero).Bold = True
'        For fil = 1 To 3
'            lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).ForeColor = vbBlue
'            lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).Bold = True
'        Next
'
'        lnCorreFinanciero = lnCorreFinanciero + 1
'        Set iTem = Me.lstFinanciero.ListItems.Add(, , prsEndSBS!DDirSoles)
'        iTem.SubItems(1) = Format(prsEndSBS!DDirDolar, "#,#00.00")
'        iTem.SubItems(2) = Format(prsEndSBS!dIndSoles, "#,#00.00")
'        iTem.SubItems(3) = Format(prsEndSBS!dIndDolar, "#,#00.00")
'        iTem.SubItems(4) = ""
'        prsEndSBS.Close
'    End If
'    prsCalSBS.Close
'    Set prsCalSBS = Nothing
'
'
'    'Calificacion CMAC
'
'    lnCorreFinanciero = lnCorreFinanciero + 1
'    Set iTem = Me.lstFinanciero.ListItems.Add(, , "Calificacion CMAC - Riesgos")
'    iTem.SubItems(1) = ""
'    iTem.SubItems(2) = ""
'    iTem.SubItems(3) = ""
'    iTem.SubItems(4) = ""
'
'    lnCorreFinanciero = lnCorreFinanciero + 1
'    Set iTem = Me.lstFinanciero.ListItems.Add(, , "Fecha")
'
'    iTem.SubItems(1) = "Calif. Final"
'    iTem.SubItems(2) = "Calif. Riesgos"
'    iTem.SubItems(3) = "Calif. S.Financ"
'    iTem.SubItems(4) = ""
'
'    lstFinanciero.ListItems(lnCorreFinanciero - 1).ForeColor = vbRed
'    lstFinanciero.ListItems(lnCorreFinanciero - 1).Bold = True
'
'    lstFinanciero.ListItems(lnCorreFinanciero).ForeColor = vbBlue
'    lstFinanciero.ListItems(lnCorreFinanciero).Bold = True
'    For fil = 1 To 3
'        lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).ForeColor = vbBlue
'        lstFinanciero.ListItems(lnCorreFinanciero).ListSubItems(fil).Bold = True
'    Next
'
'    If Not (prsCalCMAC Is Nothing) Then
'        Do While Not prsCalCMAC.EOF
'            Set iTem = Me.lstFinanciero.ListItems.Add(, , Format(prsCalCMAC!dFecha, "dd/MM/YYYY"))
'            iTem.SubItems(1) = prsCalCMAC!nCalFinal
'            iTem.SubItems(2) = prsCalCMAC!nCalRiesgos
'            iTem.SubItems(3) = prsCalCMAC!nCalSistFinan
'            iTem.SubItems(4) = ""
'            DoEvents
'            prsCalCMAC.MoveNext
'        Loop
'        prsCalCMAC.Close
'    Else
'        lnCorreFinanciero = lnCorreFinanciero + 1
'        Set iTem = Me.lstFinanciero.ListItems.Add(, , "No Registrado")
'        iTem.SubItems(1) = "No Registrado"
'        iTem.SubItems(2) = "No Registrado"
'        iTem.SubItems(3) = "No Registrado"
'        iTem.SubItems(4) = ""
'    End If
'    Set prsCalCMAC = Nothing
'
'End Sub

'Private Sub cmdEstadoCuenta_Click()
'    If Me.LstCreditos.SelectedItem.SubItems(2) <> "" Then
'        Call ImprimeEstadoCuentaCredito(Me.LstCreditos.SelectedItem.SubItems(2))
'    End If
'End Sub

'Private Sub CmdHistorial_Click()
'    Dim oNCredDoc As COMNCredito.NCOMCredDoc
'    Dim sCadImp As String
'    Dim oPrev As PrevioCredito.clsPrevioCredito
'
'    Set oNCredDoc = New COMNCredito.NCOMCredDoc
'        sCadImp = oNCredDoc.ImpreRepor_HistorialCliente(LblPersCod.Caption, gsNomAge, gdFecSis, gsCodUser, gsNomCmac)
'    Set oNCredDoc = Nothing
'
'    If Len(sCadImp) = 0 Then
'        MsgBox "No se encontraron datos del reporte", vbInformation, "AVISO"
'    Else
'        Set oPrev = New PrevioCredito.clsPrevioCredito
'        oPrev.Show sCadImp, "LISTA DE CREDITOS DEL CLIENTE", True
'        Set oPrev = Nothing
'    End If
'End Sub

Private Sub cmdImprimir_Click()

    If MsgBox("Se cambiará el estado del crédito " & Me.LstCreditos.SelectedItem.SubItems(1) & " al último estado vigente,¿Seguro de continuar?.", vbYesNo + vbQuestion, "Pregunta") = vbNo Then
        cmdImprimir.Enabled = True
        Exit Sub
    Else
        CambiaEstadoCred (Me.LstCreditos.SelectedItem.SubItems(1))
        Call ObtieneListaCreditos(sPersCod, lsPersTDoc)
        cmdImprimir.Enabled = False
    End If
    
End Sub

Private Sub CambiaEstadoCred(ByVal pcCtaCod As String)
    Dim lcNuevoEstado As String
    Dim oCambEst As COMDCredito.DCOMCredito
    Set oCambEst = New COMDCredito.DCOMCredito
    Dim rs As ADODB.Recordset

    Dim oMov As COMDMov.DCOMMov
    Set oMov = New COMDMov.DCOMMov
    Dim lsMovNro As String
    Dim lnMovNro As Long, nMovNroRef As Long
    
    Set rs = oCambEst.CambiaEstadoCred(pcCtaCod)
    'Set oCambEst = Nothing 'Comento JOEP20171011
    
    lsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oMov.InsertaMov lsMovNro, lsOpeCod, "Camb.Estd.Cred:'" & pcCtaCod & "' de 2080-Retirado a " & IIf(IsNull(rs!cNuevoEstado), "", rs!cNuevoEstado) & ".", 13
    
    'WIOR 20140203 *************************
    If Left(IIf(IsNull(rs!cNuevoEstado), "", rs!cNuevoEstado), 4) = gColocEstAprob Then
        Call oCambEst.ActualizaCredVinculados(pcCtaCod, "", 4, 1, True)
    End If
    'WIOR FIN ******************************
    MsgBox "Se cambió el estado del credito " & pcCtaCod & " de ''2080-Retirado'' a ''" & IIf(IsNull(rs!cNuevoEstado), "", rs!cNuevoEstado) & "''."
    
    Set oCambEst = Nothing 'Agrego JOEP20171011
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oCred As COMDCredito.DCOMCreditos
Dim oAho As COMDCaptaGenerales.DCOMCaptaGenerales

    Set oCred = New COMDCredito.DCOMCreditos
    Set oAho = New COMDCaptaGenerales.DCOMCaptaGenerales

    Call oCred.CargarValoresPosicionCliente(nLima, nSaldosAho)
    
    nSaldosAho = oAho.GetVisualizaSaldoPosicion(gsCodCargo)
    Set oCred = Nothing

    CentraForm Me
    
    lsOpeCod = "200809" 'Retirado a Vigente
    
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    sCalifiFinal = ""
End Sub

'Private Sub lstAhorros_DblClick()
'    If Me.lstAhorros.SelectedItem.SubItems(4) <> "" Then
'        Dim lsNomProd As String
'        lsNomProd = IIf(Mid(Me.lstAhorros.SelectedItem.SubItems(4), 6, 3) = 232, "Ahorros", (IIf(Mid(Me.lstAhorros.SelectedItem.SubItems(4), 6, 3) = 233, "Plazo Fijo", "CTS")))
'        frmCapMantenimiento.Caption = lsNomProd
'        Call frmCapMantenimiento.MuestraPosicionCliente(Me.lstAhorros.SelectedItem.SubItems(4))
'    End If
'End Sub

'Private Sub LstCartaFianza_DblClick()
'    If LstCartaFianza.SelectedItem.SubItems(2) <> "" Then
'        Call frmCFHistorial.CargaCFHistorial(LstCartaFianza.SelectedItem.SubItems(2))
'    End If
'End Sub

Private Sub LstCreditos_DblClick()
    cmdImprimir.Enabled = True
    cmdImprimir.SetFocus

End Sub

'Private Sub lstJudicial_DblClick()
'    If lstJudicial.SelectedItem.SubItems(2) <> "" Then
'        Call frmColRecRConsulta.MuestraPosicionCliente(lstJudicial.SelectedItem.SubItems(2))
'    End If
'End Sub

'Private Sub lstPrendario_DblClick()
'    If lstPrendario.SelectedItem.SubItems(3) <> "" Then
'
'        Call frmColPMantPrestamoPig.BuscaContrato(lstPrendario.SelectedItem.SubItems(3), 1, LblPersCod)
'        frmColPMantPrestamoPig.AXCodCta.NroCuenta = lstPrendario.SelectedItem.SubItems(3)
'        frmColPMantPrestamoPig.Show 1
'
'    End If
'End Sub

'Public Sub ImprimeEstadoCuentaCredito(ByVal psCtaCod As String)
'Dim oDCred As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
'Dim oWord As Word.Application
'Dim oDoc As Word.Document
'Dim oRange As Word.Range
'Dim nTasaCompAnual As Double
'
'    Set oWord = CreateObject("Word.Application")
'        oWord.Visible = False
'
'    Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\EstadoCuentaCredito.doc")
'
'    With oWord.Selection.Find
'        .Text = "<<cFecha>>"
'        .Replacement.Text = Format(gdFecSis, "dd/mm/yyyy")
'        .Forward = True
'        .Wrap = wdFindContinue
'        .Format = False
'        .Execute Replace:=wdReplaceAll
'    End With
'
'    Set oDCred = New COMDCredito.DCOMCredito
'    Set R = oDCred.RecuperaDatosParaEstadoCuentaCredito(psCtaCod)
'    Set oDCred = Nothing
'
'    If Not (R.EOF And R.BOF) Then
'        With oWord.Selection.Find
'            .Text = "<<cNomCli>>"
'            .Replacement.Text = R!cNomCli
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<cDirCli>>"
'            .Replacement.Text = R!cDirCli
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<cCtaCod>>"
'            .Replacement.Text = R!cCtaCod
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<cNomAge>>"
'            .Replacement.Text = R!vAgencia
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<cMoneda>>"
'            .Replacement.Text = IIf(Mid(psCtaCod, 9, 1) = 1, "NUEVOS SOLES", "DOLARES")
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<nMonDes>>"
'            .Replacement.Text = Format(R!nMontoCol, "#0.00")
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        nTasaCompAnual = Format(((1 + R!nTasaInteres / 100) ^ (360 / 30) - 1) * 100, "#.00")
'        With oWord.Selection.Find
'            .Text = "<<nTEA>>"
'            .Replacement.Text = nTasaCompAnual
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'            .Text = "<<nCEA>>"
'            .Replacement.Text = R!nTasCosEfeAnu
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'            .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<nSaldo>>"
'           .Replacement.Text = Format(R!nSaldo, "#0.00")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'           .Text = "<<dFecVenPag>>"
'           .Replacement.Text = IIf(DateDiff("d", R!dVencPag, "1900-01-01") = 0, "", Format(R!dVencPag, "dd/mm/yyyy"))
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<nCapPag>>"
'           .Replacement.Text = Format(R!nCapPag, "#0.00")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<nIntPag>>"
'           .Replacement.Text = Format(R!nIntPag, "#0.00")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<nMorPag>>"
'           .Replacement.Text = Format(R!nMorPag, "#0.00")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<nSegDesPag>>"
'           .Replacement.Text = Format(R!nSegDesPag, "#0.00")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<nSegBiePag>>"
'           .Replacement.Text = Format(0, "#0.00")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<nComPorPag>>"
'           .Replacement.Text = Format(0, "#0.00")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<nItfPag>>"
'           .Replacement.Text = Format(R!nItfPag, "#0.00")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<dFecPagPag>>"
'           .Replacement.Text = IIf(DateDiff("d", R!dVencPag, "1900-01-01") = 0, "", Format(R!dPago, "dd/mm/yyyy"))
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'
'        With oWord.Selection.Find
'           .Text = "<<dFecVenAPag>>"
'           .Replacement.Text = Format(R!dVencAPag, "dd/mm/yyyy")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'        With oWord.Selection.Find
'           .Text = "<<nMontoAPag>>"
'           .Replacement.Text = Format(R!nMontoAPag, "#0.00")
'           .Forward = True
'           .Wrap = wdFindContinue
'           .Format = False
'           .Execute Replace:=wdReplaceAll
'        End With
'    End If
'
'    oDoc.SaveAs (App.path & "\FormatoCarta\EstadoCuentaCredito_" & psCtaCod & ".doc")
'    oDoc.Close
'    Set oDoc = Nothing
'
'    Set oWord = CreateObject("Word.Application")
'        oWord.Visible = True
'    Set oDoc = oWord.Documents.Open(App.path & "\FormatoCarta\EstadoCuentaCredito_" & psCtaCod & ".doc")
'    Set oDoc = Nothing
'    Set oWord = Nothing
'
'End Sub
''***Agregado por ELRO el 20121112, según OYP-RFC115-2012
'Public Sub iniciarFormulario(ByVal psPersCod As String)
'Dim oDCOMPersonas As New COMDpersona.DCOMPersonas
'Dim rsPersona As New ADODB.Recordset
'Dim oPersona As New COMDpersona.UCOMPersona
'Dim oPersona2 As New COMDpersona.UCOMPersona
'Dim sPersCod As String
'
'
'Set rsPersona = oDCOMPersonas.BuscaCliente(psPersCod, BusquedaCodigo)
'Set oDCOMPersonas = Nothing
'
'oPersona2.CargaDatos rsPersona!cPersCod, _
'                     rsPersona!cPersNombre, _
'                     Format(IIf(IsNull(rsPersona!dPersNacCreac), gdFecSis, rsPersona!dPersNacCreac), "dd/mm/yyyy"), _
'                     IIf(IsNull(rsPersona!cPersDireccDomicilio), "", rsPersona!cPersDireccDomicilio), _
'                     IIf(IsNull(rsPersona!cPersTelefono), "", rsPersona!cPersTelefono), rsPersona!nPersPersoneria, _
'                     IIf(IsNull(rsPersona!cPersIDnroDNI), "", rsPersona!cPersIDnroDNI), _
'                     IIf(IsNull(rsPersona!cPersIDnroRUC), "", rsPersona!cPersIDnroRUC), _
'                     IIf(IsNull(rsPersona!cPersIdNro), "", rsPersona!cPersIdNro), _
'                     IIf(IsNull(rsPersona!cPersnatSexo), "", rsPersona!cPersnatSexo), _
'                     IIf(IsNull(rsPersona!cActiGiro1), "", rsPersona!cActiGiro1), _
'                     IIf(IsNull(rsPersona!nTipoId), "1", rsPersona!nTipoId)
'
'    LstCreditos.Enabled = False
'    lstAhorros.Enabled = False
'    LstCartaFianza.Enabled = False
'    lstJudicial.Enabled = False
'    lstPrendario.Enabled = False
'
'    Set oPersona = oPersona2
'    If Not oPersona Is Nothing Then
'        LblPersCod.Caption = oPersona.sPersCod
'        lblNomPers.Caption = oPersona.sPersNombre
'        lblDocNat.Caption = Trim(oPersona.sPersIdnroDNI)
'        lblDocJur.Caption = Trim(oPersona.sPersIdnroRUC)
'        lsPersTDoc = "1"
'        If oPersona.sPersPersoneria = "1" Then
'            fbPersNatural = True
'            If Trim(oPersona.sPersIdnroDNI) = "" Then
'                If Not Trim(oPersona.sPersIdnroOtro) = "" Then
'                    lblDocNat.Caption = Trim(oPersona.sPersIdnroOtro)
'                    lsPersTDoc = Trim(oPersona.sPersTipoDoc)
'                End If
'            End If
'        Else
'            fbPersNatural = False
'            lsPersTDoc = "3"
'        End If
'    Else
'        Exit Sub
'    End If
'    sPersCod = oPersona.sPersCod
'    Set oPersona = Nothing
'
'    If sPersCod <> "" Then
'        Call BuscarPosicionCliente(sPersCod, lsPersTDoc)
'    End If
'
'    If sPersCod <> "" Then
'        cmdSaldosConsol.Enabled = True
'        cmdImprimir.Enabled = True
'    End If
'
'    CmdBuscar.Enabled = False
'    TabPosicion.Tab = 1
'    Show 1
'End Sub

Private Sub ObtieneListaCreditos(ByVal psPersCod As String, Optional psPersTDoc As String = "1")

    Dim oCreds As COMDCredito.DCOMCredito
    
    Dim rsCred As ADODB.Recordset

    Dim rsCalSBS As ADODB.Recordset
    Dim rsEndSBS As ADODB.Recordset
    Dim rsCalCMAC As ADODB.Recordset
    Dim bExitoBusqueda As Boolean
    Dim dFechaRep As Date
    Dim lsPersDoc As String
    Dim lsPersTDoc As String

    If fbPersNatural Then
        lsPersDoc = IIf(lblDocNat.Caption = "", Trim(lblDocJur.Caption), Trim(lblDocNat.Caption))
    Else
        lsPersDoc = IIf(lblDocJur.Caption <= "", Trim(lblDocNat.Caption), Trim(lblDocJur.Caption))
    End If
    
    Set oCreds = New COMDCredito.DCOMCredito
    Set rsCred = oCreds.RecuperaMantCreditos(psPersCod)
    Set oCreds = Nothing
    
    Call MuestraCreditos(psPersCod, rsCred)

End Sub

Public Function ValidaUsuarioPermitido(ByVal psUser As String) As Boolean
    Dim oCreds As COMDCredito.DCOMCredito
    Dim rsCred As ADODB.Recordset
    Set oCreds = New COMDCredito.DCOMCredito
    Set rsCred = oCreds.ValidaUsuarioPermitido(psUser)
    Set oCreds = Nothing
    
    If Not (rsCred.EOF And rsCred.BOF) Then
        ValidaUsuarioPermitido = True
    Else
        ValidaUsuarioPermitido = False
    End If
    Set rsCred = Nothing
End Function
