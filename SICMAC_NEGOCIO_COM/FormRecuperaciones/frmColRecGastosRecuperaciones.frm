VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColRecGastosRecuperaciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recuperacion - Asignación de Gastos"
   ClientHeight    =   6345
   ClientLeft      =   360
   ClientTop       =   1830
   ClientWidth     =   9090
   Icon            =   "frmColRecGastosRecuperaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   6120
      TabIndex        =   25
      Top             =   5670
      Width           =   945
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
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
      Left            =   2520
      TabIndex        =   24
      Top             =   5670
      Width           =   825
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eli&minar"
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
      Left            =   5220
      TabIndex        =   23
      Top             =   5670
      Width           =   825
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
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
      Left            =   3420
      TabIndex        =   22
      Top             =   5670
      Width           =   825
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Left            =   4320
      TabIndex        =   21
      Top             =   5670
      Width           =   825
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7560
      TabIndex        =   20
      Top             =   90
      Width           =   1185
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Left            =   7135
      TabIndex        =   15
      Top             =   5670
      Width           =   825
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
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   5670
      Width           =   825
   End
   Begin VB.Frame fradatos 
      Height          =   1425
      Left            =   180
      TabIndex        =   7
      Top             =   1890
      Width           =   8790
      Begin VB.ComboBox cboGastos 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3420
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   270
         Width           =   5055
      End
      Begin VB.TextBox txtMotivo 
         Height          =   675
         Left            =   3420
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   660
         Width           =   5070
      End
      Begin MSMask.MaskEdBox txtFecha 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   1245
         TabIndex        =   26
         Top             =   270
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.EditMoney AXMonto 
         Height          =   285
         Left            =   1260
         TabIndex        =   34
         Top             =   630
         Width           =   1095
         _extentx        =   2090
         _extenty        =   503
         font            =   "frmColRecGastosRecuperaciones.frx":030A
         text            =   "0"
         enabled         =   -1
      End
      Begin SICMACT.EditMoney AXMontoPagado 
         Height          =   285
         Left            =   1260
         TabIndex        =   35
         Top             =   990
         Width           =   1095
         _extentx        =   2090
         _extenty        =   503
         font            =   "frmColRecGastosRecuperaciones.frx":0336
         text            =   "0"
         enabled         =   -1
      End
      Begin VB.Label Label7 
         Caption         =   "Monto Pagado"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   36
         Top             =   1020
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Gasto"
         Height          =   195
         Index           =   2
         Left            =   2790
         TabIndex        =   32
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label7 
         Caption         =   "Motivo"
         Height          =   195
         Index           =   12
         Left            =   2820
         TabIndex        =   31
         Top             =   660
         Width           =   540
      End
      Begin VB.Label Label7 
         Caption         =   "Monto Gasto"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   28
         Top             =   630
         Width           =   915
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
         Height          =   195
         Left            =   180
         TabIndex        =   27
         Top             =   270
         Width           =   660
      End
   End
   Begin VB.Frame fratitular 
      Caption         =   "Datos del Credito"
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
      Height          =   1095
      Left            =   135
      TabIndex        =   3
      Top             =   720
      Width           =   8790
      Begin VB.Label lblTotalDeuda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1350
         TabIndex        =   14
         Top             =   630
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Total Deuda"
         Height          =   195
         Index           =   6
         Left            =   270
         TabIndex        =   13
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label lblTotalGastos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   7020
         TabIndex        =   12
         Top             =   630
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Saldo Gastos"
         Height          =   195
         Index           =   7
         Left            =   5760
         TabIndex        =   11
         Top             =   645
         Width           =   1065
      End
      Begin VB.Label lblTotalCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4140
         TabIndex        =   10
         Top             =   630
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Saldo Capital"
         Height          =   195
         Index           =   8
         Left            =   3060
         TabIndex        =   9
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lblCodPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1350
         TabIndex        =   6
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2880
         TabIndex        =   5
         Top             =   270
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   270
         TabIndex        =   4
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame fralista 
      Height          =   2085
      Left            =   180
      TabIndex        =   1
      Top             =   3300
      Width           =   8790
      Begin MSComctlLib.ListView lstGastos 
         Height          =   1845
         Left            =   165
         TabIndex        =   2
         Top             =   180
         Width           =   8490
         _ExtentX        =   14975
         _ExtentY        =   3254
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
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Motivo"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Pagado"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Estado"
            Object.Width           =   265
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CodGasto"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   195
      Left            =   7260
      TabIndex        =   8
      Top             =   330
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmColRecGastosRecuperaciones.frx":0362
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   375
      Left            =   180
      TabIndex        =   33
      Top             =   180
      Width           =   3615
      _extentx        =   6376
      _extenty        =   661
      texto           =   "Crédito"
      enabledcta      =   -1
      enabledprod     =   -1
      enabledage      =   -1
   End
   Begin VB.Label lblCredTransferido 
      Caption         =   "CREDITO TRANSFERIDO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   3960
      TabIndex        =   37
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label7 
      Caption         =   "G. Acumulados"
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   17
      Top             =   5580
      Width           =   1155
   End
   Begin VB.Label lblGastoPag 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1230
      TabIndex        =   19
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label Label7 
      Caption         =   "G. Pagados"
      Height          =   195
      Index           =   14
      Left            =   90
      TabIndex        =   18
      Top             =   5865
      Width           =   945
   End
   Begin VB.Label lblGastoAcum 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   1230
      TabIndex        =   16
      Top             =   5505
      Width           =   1080
   End
End
Attribute VB_Name = "frmColRecGastosRecuperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
' GASTOS DE RECUPERACIONES
'LAYG   :  10/08/2001.
'Resumen:  Nos permite hacer el mantenimiento de los Gastos de Recuperaciones

Option Explicit
Dim fbNuevo As Boolean
Dim fnSaldoGasto As Currency
Dim fnUltimoGasto As Currency
Dim objPista As COMManejador.Pista  '' *** PEAC 20090126
Dim fbCredTransferido As Boolean 'FRHU 20150415 ERS022-2015

Private Sub HabilitaControles(ByVal pbCmdNuevo As Boolean, ByVal pbCmdEditar As Boolean, ByVal pbCmdGrabar As Boolean, _
            ByVal pbCmdCancelar As Boolean, ByVal pbCmdEliminar As Boolean, ByVal pbCmdSalir As Boolean, _
            ByVal pbFraDatos As Boolean, ByVal pbFraLista As Boolean)
    cmdNuevo.Enabled = pbCmdNuevo
    cmdEditar.Enabled = pbCmdEditar
    cmdGrabar.Enabled = pbCmdGrabar
    cmdCancelar.Enabled = pbCmdCancelar
    cmdEliminar.Enabled = pbCmdEliminar
    fradatos.Enabled = pbFraDatos
    fralista.Enabled = pbFraLista
End Sub

Private Sub CargaGastosRecup()

Dim loGastoRec As COMDColocRec.DCOMColRecCredito
Dim lrGastos As New ADODB.Recordset

Set loGastoRec = New COMDColocRec.DCOMColRecCredito
        Set lrGastos = loGastoRec.dObtieneGastosRecup
Set loGastoRec = Nothing

If lrGastos Is Nothing Then
    Limpiar
    Set lrGastos = Nothing
    Exit Sub
End If

    Me.cboGastos.Clear
    Do While Not lrGastos.EOF
        Me.cboGastos.AddItem lrGastos!cdescripcion & Space(120) & lrGastos!nPrdConceptoCod
        lrGastos.MoveNext
    Loop
    cboGastos.ListIndex = -1

Set lrGastos = Nothing
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaCredito(AXCodCta.NroCuenta)
End Sub

'Busca el contrato ingresado
Private Sub BuscaCredito(ByVal psNroContrato As String)
Dim lbOk As Boolean
Dim lrDatCredito As New ADODB.Recordset
Dim lrListaGastos As New ADODB.Recordset
Dim loValCred As COMDColocRec.DCOMColRecCredito
'Dim loCalculos As NColRecCalculos
Dim lnDeuda As Currency
Dim lnGastoAcum As Currency, lnGastoPag As Currency
Dim iTem As ListItem

Dim lsmensaje As String

'On Error GoTo ControlError
    'FRHU 20150415 ERS022-2015
    If VerificarSiEsUnCreditoTransferido(psNroContrato) Then
        fbCredTransferido = True
        lblCredTransferido.Visible = True
    Else
        fbCredTransferido = False
        lblCredTransferido.Visible = False
    End If
    'FIN FRHU 20150415
    'Carga Datos
    Set loValCred = New COMDColocRec.DCOMColRecCredito
        'Set lrDatCredito = loValCred.dObtieneDatosAsignaGastoCredRecup(psNroContrato)
        Set lrDatCredito = loValCred.dObtieneDatosAsignaGastoCredRecup(psNroContrato, fbCredTransferido) 'FRHU 20150415 ERS022-2015
        Set lrListaGastos = loValCred.dObtieneListaGastosxCredito(psNroContrato, lsmensaje, fbCredTransferido) 'FRHU 20150415 ERS022-2015
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Set loValCred = Nothing
    
    If lrDatCredito Is Nothing Or (lrDatCredito.BOF And lrDatCredito.EOF) Then   ' Hubo un Error
        Limpiar
        Set lrDatCredito = Nothing
        Exit Sub
    End If
    
    ' Asigna Valores a las Variables
    fnSaldoGasto = lrDatCredito!nSaldoGasto
    fnUltimoGasto = lrDatCredito!nUltGasto
    
    'Muestra Datos
    Me.lblCodPers.Caption = Trim(lrDatCredito!cPersCod)
    Me.lblNomPers.Caption = PstaNombre(Trim(lrDatCredito!cPersNombre))
    Me.lblTotalDeuda.Caption = Format(lrDatCredito!nSaldo + lrDatCredito!nSaldoIntComp + lrDatCredito!nSaldoIntMor + lrDatCredito!nSaldoGasto, "####0.00")
    Me.lblTotalCapital.Caption = Format(lrDatCredito!nSaldo, "####0.00")
    Me.lblTotalGastos.Caption = Format(lrDatCredito!nSaldoGasto, "####0.00")
    
    lstGastos.ListItems.Clear
    Do While Not lrListaGastos.EOF
        Set iTem = lstGastos.ListItems.Add(, , lrListaGastos!nNroGastoCta)
            iTem.SubItems(1) = Format(lrListaGastos!dAsigna, "dd/mm/yyyy hh:mm")
            iTem.SubItems(2) = IIf(IsNull(Trim(lrListaGastos!cMotivoGasto)), "", Trim(Trim(lrListaGastos!cMotivoGasto)))    'Motivo Gasto
            iTem.SubItems(3) = Format(IIf(IsNull(Trim(lrListaGastos!nMonto)), 0, Trim(lrListaGastos!nMonto)), "####0.00")   'monto del gasto
            iTem.SubItems(4) = Format(IIf(IsNull(Trim(lrListaGastos!nMontoPagado)), 0, Trim(lrListaGastos!nMontoPagado)), "####0.00")  'monto a pagar
            iTem.SubItems(5) = lrListaGastos!nColocRecGastoEstado 'Estado del Gastos
            iTem.SubItems(6) = Trim(lrListaGastos!nPrdConceptoCod)  'codigo del gasto
            If lrListaGastos!nColocRecGastoEstado = gColRecGastoEstPendiente Then  ' Pendientes
                iTem.ForeColor = RGB(200, 20, 10)
                iTem.ListSubItems.iTem(1).ForeColor = RGB(200, 20, 10)
                iTem.ListSubItems.iTem(2).ForeColor = RGB(200, 20, 10)
            End If
        lnGastoAcum = lnGastoAcum + lrListaGastos!nMonto
        lnGastoPag = lnGastoPag + lrListaGastos!nMontoPagado
        lrListaGastos.MoveNext
    Loop
  
    Set lrDatCredito = Nothing
    Set lrListaGastos = Nothing
        
    Me.lblGastoAcum = Format(lnGastoAcum, "#,###0.00")
    Me.lblGastoPag = Format(lnGastoPag, "#,###0.00")
    AXCodCta.Enabled = False
    Call HabilitaControles(True, True, False, False, False, True, False, True)
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & err.Number & " " & err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMotivo.SetFocus
    End If
End Sub

Private Sub cboGastos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    AXMonto.Enabled = True
    AXMonto.SetFocus
End If
End Sub

Private Sub CmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona 'UPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
'lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast & "," & gColocEstTransferido 'FRHU 20150415 ERS022-2015

If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.Enabled = True
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
If fbNuevo = True Then
    cboGastos.Enabled = False
    AXMonto.Enabled = False
    cboGastos.ListIndex = -1
    txtMotivo = ""
    AXMonto = ""
    fbNuevo = False
Else
    Limpiar
    Call HabilitaControles(False, False, False, False, False, True, False, False)
    AXCodCta.Enabled = True
End If

Call HabilitaControles(True, True, False, True, True, True, False, True)
End Sub

Private Sub CmdEditar_Click()
   
    fbNuevo = False
    If lstGastos.ListItems.Count = 0 Then
       MsgBox "Seleccione Gasto a editar ", vbInformation, "Aviso"
       cmdNuevo.SetFocus
    Else
       If lstGastos.SelectedItem.SubItems(5) = gColRecGastoEstPendiente Then
            Call HabilitaControles(False, False, True, True, False, True, True, False)
            lstGastos.SelectedItem.ForeColor = vbRed
            AXMonto.Enabled = True
            AXMonto.SetFocus
       Else
            MsgBox "Gasto no esta pendiente de Cobranza", vbInformation, "Aviso"
       End If
    End If
End Sub
Private Sub cmdEliminar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnConceptoGasto As Integer
Dim lnNewSaldoGasto As Currency, lnMonto As Currency
Dim lnNroGastoCta As Integer

'Valida si puede eliminar Gasto

If lstGastos.SelectedItem.SubItems(5) <> gColRecGastoEstPendiente Then
     MsgBox "Gasto no esta pendiente de Cobranza, No se puede eliminar.", vbInformation, "Aviso"
     Exit Sub
End If

lnMonto = CCur(Format(Trim(AXMonto), "###0.00"))
lnNewSaldoGasto = fnSaldoGasto - lnMonto
lnNroGastoCta = val(lstGastos.SelectedItem.Text)


If MsgBox(" Grabar Eliminación de Gasto de Recuperaciones ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabar = New COMNColocRec.NCOMColRecCredito
           'MADM 20100505
            If Not loGrabar.GetBuscaGastoCreditoFinanciero(AXCodCta.NroCuenta, lnNroGastoCta) Then
            Call loGrabar.nGastoRecupEliminaGasto(AXCodCta.NroCuenta, lsFechaHoraGrab, _
                 lsMovNro, lnNroGastoCta, gColRecGastoEstEliminado, lnNewSaldoGasto, False)
            Else
                MsgBox "No puede eliminar el Gasto por ser Ingresado por el Financiero", vbInformation, "Aviso"
            End If
           'END MADM
            
        Set loGrabar = Nothing
        Limpiar
        BuscaCredito (AXCodCta.NroCuenta)
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        Refresco
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub CmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabar As COMNColocRec.NCOMColRecCredito

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lsFechaAsigna As String
Dim lnConceptoGasto As Long
Dim lnMonto As Currency
Dim lnMontoAntes As Currency
Dim lsMotivo As String
Dim lnNewSaldoGasto As Currency
Dim lnNroGastoCta As Integer
Dim lsNroCredito As String

    If ValidaAntesGrabar = False Then ' Valida si se ingreso los datos correctos
        Exit Sub
    End If

   lsFechaAsigna = Format$(txtFecha.Text, "mm/dd/yyyy")

   If fbNuevo = True Then
      lnMontoAntes = 0
      lnConceptoGasto = Trim(Right(Trim(cboGastos), 6))
      lnNroGastoCta = fnUltimoGasto + 1
      lnMonto = CCur(Format(Trim(AXMonto), "###0.00"))
   Else
      lnMontoAntes = Format(lstGastos.SelectedItem.SubItems(3), "###0.00")
      lnConceptoGasto = val(lstGastos.SelectedItem.SubItems(6))
      lnNroGastoCta = val(lstGastos.SelectedItem.Text)
      lnMonto = CCur(Format(Trim(AXMonto), "###0.00"))
   End If
   

   lsMotivo = Trim(Me.txtMotivo.Text)
   lnNewSaldoGasto = fnSaldoGasto + lnMonto - lnMontoAntes


If MsgBox(" Grabar Registro de Gastos de Recuperaciones ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    
    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    Set loGrabar = New COMNColocRec.NCOMColRecCredito
        If fbNuevo = True Then  ' Si es un registro nuevo
            'FRHU 20150415 ERS022-2015: Se agrego fbCredTransferido
            'Call loGrabar.nGastoRecupAsignaNuevo(AXCodCta.NroCuenta, lsFechaHoraGrab, _
            '     lsMovNro, lnNroGastoCta, lsFechaAsigna, lnConceptoGasto, lnMonto, _
            '     0, gColRecGastoEstPendiente, lsMotivo, lnNewSaldoGasto, False)
            Call loGrabar.nGastoRecupAsignaNuevo(AXCodCta.NroCuenta, lsFechaHoraGrab, _
                 lsMovNro, lnNroGastoCta, lsFechaAsigna, lnConceptoGasto, lnMonto, _
                 0, gColRecGastoEstPendiente, lsMotivo, lnNewSaldoGasto, False, fbCredTransferido)
            'FRHU 20150415
            '' *** PEAC 20090126
            objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, , AXCodCta.NroCuenta, gCodigoCuenta
                 
        Else
            'MADM 20110505
            If Not loGrabar.GetBuscaGastoCreditoFinanciero(AXCodCta.NroCuenta, lnNroGastoCta) Then
                 Call loGrabar.nGastoRecupModifica(AXCodCta.NroCuenta, lsFechaHoraGrab, _
                     lsMovNro, lnNroGastoCta, lsFechaAsigna, lnConceptoGasto, lnMonto, _
                     0, gColRecGastoEstPendiente, lsMotivo, lnNewSaldoGasto, False)
                
                '' *** PEAC 20090126
                objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, , AXCodCta.NroCuenta, gCodigoCuenta
            Else
                MsgBox "No puede actualizar el Gasto por ser Ingresado por el Financiero", vbInformation, "Aviso"
            End If
           'END MADM
        End If
    Set loGrabar = Nothing
    
    lsNroCredito = AXCodCta.NroCuenta
    Limpiar
    AXCodCta.NroCuenta = lsNroCredito
    BuscaCredito (lsNroCredito)
    Call HabilitaControles(True, False, False, True, False, True, False, True)
        
    'AXCodCta.Enabled = True
    'AXCodCta.SetFocus
        
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & err.Number & " " & err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdImprimir_Click()
Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim lsFechaHora As String
Dim loGastoRec As COMDColocRec.DCOMColRecCredito 'DColRecCredito
Dim lrGastos As New ADODB.Recordset
Dim loImprimeGasto As COMNColocRec.NCOMColRecImpre 'NColRecImpre
Dim lsCadImprimir As String
Dim loPrevio As previo.clsprevio

Dim lsmensaje As String

    If lstGastos.ListItems.Count > 0 Then
        
        Set loGastoRec = New COMDColocRec.DCOMColRecCredito
                Set lrGastos = loGastoRec.dObtieneListaGastosxCredito(Me.AXCodCta.NroCuenta, lsmensaje)
                If Trim(lsmensaje) <> "" Then
                    MsgBox lsmensaje, vbInformation, "Aviso"
                    Exit Sub
                End If
        Set loGastoRec = Nothing

        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsFechaHora = fgFechaHoraGrab(loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
        Set loContFunct = Nothing
        
        Set loImprimeGasto = New COMNColocRec.NCOMColRecImpre
            lsCadImprimir = loImprimeGasto.nPrintGastosRecuperaciones(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, Me.lblNomPers.Caption, lrGastos, lsmensaje)
            If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation, "Aviso"
                Exit Sub
            End If

        Set loImprimeGasto = Nothing
        Set loPrevio = New previo.clsprevio
            loPrevio.Show lsCadImprimir, "Gastos de Recuperación de Creditos", True
        Set loPrevio = Nothing
    Else
        MsgBox " Lista vacía ...! ", vbInformation, " Aviso "
    End If
End Sub

Private Sub cmdNuevo_Click()
    
    fbNuevo = True
    Call HabilitaControles(False, False, True, True, False, True, True, False)
    
    cboGastos.Enabled = True
    AXMonto.Enabled = True
    
    txtFecha = Format(gdFecSis, "dd/mm/yyyy")
    Me.AXMonto.Text = 0
    Me.AXMontoPagado.Text = 0
    cboGastos.ListIndex = -1
    Me.txtMotivo.Text = ""
    txtFecha.SetFocus
    
End Sub

Private Sub cmdsalir_Click()
    If cmdGrabar.Enabled = True Then
        MsgBox "Antes de Culminar Guarde o Cancele los Cambios", vbInformation, "Aviso"
        cmdGrabar.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Sub Limpiar()
    AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    lblCodPers = ""
    lblNomPers = ""
    txtFecha.Text = "__/__/____"
    AXMonto = ""
    txtMotivo = ""
    Me.lblTotalDeuda = ""
    Me.lblTotalCapital = ""
    Me.lblTotalGastos = ""
    cboGastos.ListIndex = -1
    Me.lstGastos.ListItems.Clear
    Me.lblTotalGastos = ""
    Me.AXMontoPagado.Text = ""
    Me.AXMonto.Text = ""
    lblCredTransferido.Visible = False 'FRHU 20150415 ERS022-2015
End Sub

'******************************************************************
'funcion para validar la información ingresada el momento de grabar
'******************************************************************
Private Function ValidaAntesGrabar() As Boolean
    If Len(AXCodCta.NroCuenta) = 0 Then
         ValidaAntesGrabar = False
         MsgBox "Número de Cuenta de Crédito no valido", vbInformation, "Aviso"
         AXCodCta.SetFocusAge
         Exit Function
    End If
    If fbNuevo = True Then
       If Len(cboGastos) = 0 Then
            ValidaAntesGrabar = False
            MsgBox "Gasto no valido", vbInformation, "Aviso"
            If cboGastos.Enabled = True Then
                cboGastos.SetFocus
            End If
            Exit Function
       End If
       If CCur(Format(Trim(AXMonto), "###0.00")) <= 0 Then
            ValidaAntesGrabar = False
            MsgBox "Monto de Gasto debe ser mayor a Cero", vbInformation, "Aviso"
            Exit Function
       End If
    End If
    
    ValidaAntesGrabar = True

End Function

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    Call HabilitaControles(False, False, False, False, False, True, False, False)
    CargaGastosRecup
    fbCredTransferido = False 'FRHU ERS022-2015
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gRecAsignarGastos
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub lstGastos_Click()
    Refresco
    Call HabilitaControles(True, True, False, True, True, True, False, True)
End Sub

Private Sub lstGastos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lstGastos.Sorted = False
    lstGastos.SortKey = ColumnHeader.Index - 1
  
    If lstGastos.SortOrder = lvwAscending Then
        lstGastos.SortOrder = lvwDescending
    Else
        lstGastos.SortOrder = lvwAscending
    End If
   ' Asigna a Sorted el valor True para ordenar la lista.
    lstGastos.Sorted = True
End Sub

Private Sub lstGastos_GotFocus()
    Refresco
End Sub

Private Sub Refresco()
    If lstGastos.ListItems.Count > 0 Then
        Me.txtFecha.Text = Format(CDate(Trim(lstGastos.SelectedItem.SubItems(1))), "dd/mm/yyyy")
        Me.txtMotivo.Text = Trim(lstGastos.SelectedItem.SubItems(2))
        Me.AXMonto.Text = Format(Trim(Me.lstGastos.SelectedItem.SubItems(3)), "####0.00")
        Me.AXMontoPagado.Text = Format(Trim(Me.lstGastos.SelectedItem.SubItems(4)), "####0.00")
        Call UbicaCombo(Me.cboGastos, lstGastos.SelectedItem.SubItems(6), True, 6)
        
    End If
End Sub

Private Sub lstGastos_KeyUp(KeyCode As Integer, Shift As Integer)
    Call lstGastos_Click
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If IsDate(txtFecha) Then
        Me.cboGastos.SetFocus
    End If
End If
End Sub

Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
       cmdGrabar.SetFocus
    End If
End Sub
'FRHU 20150415 ERS022-2015
Private Function VerificarSiEsUnCreditoTransferido(ByVal psCtaCod As String) As Boolean
    Dim oCredito As COMDCredito.DCOMCredito
    
    Set oCredito = New COMDCredito.DCOMCredito
    VerificarSiEsUnCreditoTransferido = oCredito.VerificaSiEsCreditoTransferido(psCtaCod)
    Set oCredito = Nothing
End Function
'FIN FRHU 20150415
