VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmChequeOpePendiente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPERACIONES PENDIENTES CON CHEQUE"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   Icon            =   "frmChequeOpePendiente.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Cheques Valorizados"
      TabPicture(0)   =   "frmChequeOpePendiente.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feCheque"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame3 
         Caption         =   "Operaciones Pendientes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3080
         Left            =   120
         TabIndex        =   1
         Top             =   3360
         Width           =   11115
         Begin VB.CommandButton cmdDetalle 
            Caption         =   "&Ver Detalle"
            Height          =   345
            Left            =   120
            TabIndex        =   6
            Top             =   2640
            Width           =   1050
         End
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "&Actualizar"
            Height          =   345
            Left            =   8830
            TabIndex        =   5
            Top             =   2640
            Width           =   1050
         End
         Begin VB.CommandButton cmdSalir 
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            Height          =   345
            Left            =   9915
            TabIndex        =   4
            Top             =   2640
            Width           =   1050
         End
         Begin SICMACT.FlexEdit feOperacion 
            Height          =   2370
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   10860
            _extentx        =   19156
            _extenty        =   4180
            cols0           =   8
            highlight       =   2
            allowuserresizing=   1
            encabezadosnombres=   "N°-Código-Tipo de Operación-Disponible-Estado-Detalle de la Operación-Glosa-Aux"
            encabezadosanchos=   "0-1200-3000-1120-1200-2800-2500-0"
            font            =   "frmChequeOpePendiente.frx":0326
            font            =   "frmChequeOpePendiente.frx":034E
            font            =   "frmChequeOpePendiente.frx":0376
            font            =   "frmChequeOpePendiente.frx":039E
            font            =   "frmChequeOpePendiente.frx":03C6
            fontfixed       =   "frmChequeOpePendiente.frx":03EE
            lbultimainstancia=   -1
            tipobusqueda    =   6
            columnasaeditar =   "X-X-X-X-X-X-X-X"
            textstylefixed  =   4
            listacontroles  =   "0-0-0-0-0-0-0-0"
            encabezadosalineacion=   "C-L-L-R-L-L-L-C"
            formatosedit    =   "0-0-0-2-0-0-0-0"
            textarray0      =   "N°"
            lbflexduplicados=   0
            lbformatocol    =   -1
            lbpuntero       =   -1
            lbbuscaduplicadotext=   -1
            rowheight0      =   300
         End
      End
      Begin SICMACT.FlexEdit feCheque 
         Height          =   2850
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   11100
         _extentx        =   19579
         _extenty        =   5027
         cols0           =   9
         highlight       =   2
         allowuserresizing=   1
         encabezadosnombres=   "N°-N° de cheque-Girador-Banco Emisor-Moneda-Monto-Fecha Reg.-EstructuraNueva-nID"
         encabezadosanchos=   "0-2500-2500-2500-1000-1500-1000-0-0"
         font            =   "frmChequeOpePendiente.frx":0414
         font            =   "frmChequeOpePendiente.frx":043C
         font            =   "frmChequeOpePendiente.frx":0464
         font            =   "frmChequeOpePendiente.frx":048C
         font            =   "frmChequeOpePendiente.frx":04B4
         fontfixed       =   "frmChequeOpePendiente.frx":04DC
         lbultimainstancia=   -1
         tipobusqueda    =   6
         columnasaeditar =   "X-X-X-X-X-X-X-X-X"
         textstylefixed  =   4
         listacontroles  =   "0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-L-L-L-C-R-C-C-C"
         formatosedit    =   "0-0-0-0-0-2-0-0-0"
         textarray0      =   "N°"
         lbformatocol    =   -1
         lbordenacol     =   -1
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
      End
   End
End
Attribute VB_Name = "frmChequeOpePendiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre : frmChequeOpePendientes
'** Descripción : Cheques valorizados aptos para realizar operaciones creado segun TI-ERS126-2013
'** Creación : EJVG, 20140305 11:00:00 AM
'************************************************************************************************
Option Explicit

Private Sub Form_Load()
    CargarChequesValorizados
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub CargarChequesValorizados()
    Dim oDR As New NCOMDocRec
    Dim oRS As New ADODB.Recordset
    Dim row As Long
    
    On Error GoTo ErrCargarCheques
    Screen.MousePointer = 11
    Set oRS = oDR.ListaChequexOperacionesPendiente(Right(gsCodAge, 2))
    FormateaFlex feCheque
    FormateaFlex feOperacion
    Do While Not oRS.EOF
        feCheque.AdicionaFila
        row = feCheque.row
        feCheque.TextMatrix(row, 1) = IIf(oRS!cNroDocComp = "", oRS!cNroDoc, oRS!cNroDocComp)
        feCheque.TextMatrix(row, 2) = oRS!cPersNombreGirador
        feCheque.TextMatrix(row, 3) = oRS!cPersNombreBancoEmisor
        feCheque.TextMatrix(row, 4) = oRS!cMoneda
        feCheque.TextMatrix(row, 5) = Format(oRS!nMonto, gsFormatoNumeroView)
        feCheque.TextMatrix(row, 6) = Format(oRS!dFecha, gsFormatoFechaView)
        feCheque.TextMatrix(row, 7) = oRS!nEstructuraNueva
        feCheque.TextMatrix(row, 8) = IIf(oRS!nEstructuraNueva = 0, CStr(oRS!nTpoDoc) & "|" & CStr(oRS!cNroDoc) & "|" & CStr(oRS!cperscod) & "|" & CStr(oRS!cIFTpo) & "|" & CStr(oRS!cIFCta), oRS!nId)
        oRS.MoveNext
    Loop
    feCheque.TopRow = 1
    Screen.MousePointer = 0
    Exit Sub
ErrCargarCheques:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub CargarOperacionesCheques(ByVal pnRow As Long)
    Dim oDR As New NCOMDocRec
    Dim oRS As New ADODB.Recordset
    Dim row As Long
    Dim lbEstructNew As Boolean
    Dim lnId As Long
    Dim lnTpoDoc As Integer
    Dim lsNroDoc As String
    Dim lsPersCod As String
    Dim lsIFTpo As String
    Dim lsIFCta As String
    Dim Matrix As Variant
        
    On Error GoTo ErrCargarOperacionesCheques
    
    If feCheque.TextMatrix(pnRow, 0) = "" Then Exit Sub
    If pnRow <= 0 Then Exit Sub
    Screen.MousePointer = 11

    lbEstructNew = IIf(feCheque.TextMatrix(pnRow, 7) = 1, True, False)
    Matrix = Split(feCheque.TextMatrix(pnRow, 8), "|")
    If lbEstructNew Then
        lnId = Matrix(0)
    Else
        lnTpoDoc = Matrix(0)
        lsNroDoc = Matrix(1)
        lsPersCod = Matrix(2)
        lsIFTpo = Matrix(3)
        lsIFCta = Matrix(4)
    End If
    
    Set oRS = oDR.ListaChequeDetxOperacionesPendiente(lbEstructNew, lnId, lnTpoDoc, lsNroDoc, lsPersCod, lsIFTpo, lsIFCta)
    Frame3.Caption = "Operaciones Pendientes con el cheque N° " & feCheque.TextMatrix(pnRow, 1)
    FormateaFlex feOperacion
    Do While Not oRS.EOF
        feOperacion.AdicionaFila
        row = feOperacion.row
        feOperacion.TextMatrix(row, 1) = oRS!cNroDoc
        feOperacion.TextMatrix(row, 2) = oRS!cTipoOperacion & space(500) & oRS!nTipoOperacion
        feOperacion.TextMatrix(row, 3) = Format(oRS!nDisponible, gsFormatoNumeroView)
        feOperacion.TextMatrix(row, 4) = oRS!cEstado
        feOperacion.TextMatrix(row, 5) = oRS!cDetalle
        feOperacion.TextMatrix(row, 6) = oRS!cGlosa
        oRS.MoveNext
    Loop
    Screen.MousePointer = 0
    Exit Sub
ErrCargarOperacionesCheques:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdActualizar_Click()
    CargarChequesValorizados
End Sub
Private Sub feCheque_DblClick()
    If feCheque.row > 0 Then
        CargarOperacionesCheques feCheque.row
    End If
End Sub
Private Sub feCheque_KeyPress(KeyAscii As Integer)
    If feCheque.row > 0 And KeyAscii = 13 Then
        CargarOperacionesCheques feCheque.row
    End If
End Sub
Private Sub feCheque_OnRowChange(pnRow As Long, pnCol As Long)
    If pnRow > 0 Then
        CargarOperacionesCheques pnRow
    End If
End Sub
Private Sub cmdDetalle_Click()
    Dim row As Integer
    Dim lnOperacion As TipoOperacionCheque
    Dim frmMntCap As frmCapMantenimientoCtas
    Dim frmLote As New frmChequeDetLote
    Dim lsDetalle As String
    
    On Error GoTo ErrcmdDetalle
    If feOperacion.TextMatrix(1, 0) = "" Then Exit Sub
    row = feOperacion.row
    If feOperacion.TextMatrix(row, 0) = "" Then Exit Sub
    
    row = feOperacion.row
    lnOperacion = CInt(Trim(Right(feOperacion.TextMatrix(row, 2), 8)))
    lsDetalle = feOperacion.TextMatrix(row, 5)

    Select Case lnOperacion '*** Constante 10034
        Case DPF_Apertura, AHO_Apertura, CTS_Apertura 'Apertura
            Dim frm As New frmChequeDetApert
            frm.Inicio lsDetalle, True
            Set frm = Nothing
        Case DPF_AumentoCapital, AHO_Deposito, CTS_Deposito 'Depósitos y Aumento de Capital
            Set frmMntCap = New frmCapMantenimientoCtas
            frmMntCap.lstCuentas.AddItem lsDetalle
            frmMntCap.Inicia
        Case CRED_Pago 'Pago de Crédito Individual
            Set frmMntCap = New frmCapMantenimientoCtas
            frmMntCap.lstCuentas.AddItem lsDetalle
            frmMntCap.Inicia
        Case CTS_DepositoLote, CRED_PagoLote, AHO_DepositoLote, AHO_AperturaLote, AHO_DepositoHaberesLote, DPF_AperturaLote 'Lote
            Set frmLote = New frmChequeDetLote
            frmLote.Inicio Val(lsDetalle), True
        Case Else
            MsgBox "Esta Operación no esta configurado para el nuevo proceso de cheques", vbCritical, "Aviso"
            Exit Sub
    End Select
        
    Set frmMntCap = Nothing
    Set frmLote = Nothing
    Exit Sub
ErrcmdDetalle:
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

