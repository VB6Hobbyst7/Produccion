VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColPRemateRegVtaListados 
   Caption         =   "Crédito Pignoraticio - Venta de Listados para Remate"
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   7125
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraMonto 
      Caption         =   "Monto"
      Height          =   735
      Left            =   4500
      TabIndex        =   6
      Top             =   1620
      Width           =   2415
      Begin SICMACT.EditMoney txtMonto 
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2460
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   2460
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Frame fraContenedor 
      Caption         =   "Cliente(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Index           =   6
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   6810
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   345
         Left            =   180
         TabIndex        =   1
         Top             =   990
         Width           =   825
      End
      Begin MSComctlLib.ListView lstCliente 
         Height          =   765
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   1349
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo del Cliente"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre / Razón Social del Cliente"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Dirección"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Teléfono"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ciudad"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Zona"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Doc.Civil"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Nro.Doc.Civil"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Doc.Tributario"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Nro.Doc.Tributario"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "frmColPRemateRegVtaListados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* VENTA DE LISTADO DE REMATE
'Archivo:  frmColPRemateRegVtaListados.frm
'LAYG   :  23/07/2001.
'Resumen:  Nos permite registrar la venta de listados para remate

Option Explicit

Dim fnVarCostoListadoParaRemate As Double
Private Sub Limpiar()
    lstCliente.ListItems.Clear
    txtMonto.Text = 0
End Sub

Private Sub CmdAgregar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String
Dim lstTmpCliente As ListItem

Dim ls As String
On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    Set lstTmpCliente = lstCliente.ListItems.Add(, , lsPersCod)
        lstTmpCliente.SubItems(1) = loPers.sPersNombre
        lstTmpCliente.SubItems(2) = loPers.sPersDireccDomicilio
        lstTmpCliente.SubItems(3) = loPers.sPersTelefono
End If

Set loPers = Nothing

Me.txtMonto.Enabled = True
Me.txtMonto.SetFocus

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    cmdAgregar.Enabled = True
    cmdAgregar.SetFocus
End Sub

Private Sub cmdGrabar_Click()
'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarVta As COMNColoCPig.NCOMColPContrato

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim vNumDupl As Integer

If MsgBox(" Grabar Venta de Listados Para Remate de Credito Pignoraticio ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarVta = New COMNColoCPig.NCOMColPContrato
            'Grabar Venta de Listado para Remate
            'Call loGrabarVta.nDuplicadoContratoCredPignoraticio(AXCodCta.NroCuenta, vNumDupl, lsFechaHoraGrab, _
            '      lsMovNro, CCur(Me.lblCostoDuplicado.Caption), False)
        Set loGrabarVta = Nothing

        ' *** Impresion
        If MsgBox("Desea realizar impresión del Recibo ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            'ImprimirContrato 1
            Do While True
                If MsgBox("Desea reimprimir ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    'ImprimirContrato 1
                Else
                    Exit Do
                End If
            Loop
        End If
        Limpiar
        
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CargaParametros
    Limpiar
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    fnVarCostoListadoParaRemate = loParam.dObtieneColocParametro(gConsColPCostoDuplicadoContrato)
Set loParam = Nothing
End Sub

Private Sub txtMonto_Change()
If txtMonto.value = 0 Then
    cmdGrabar.Enabled = False
Else
    cmdGrabar.Enabled = True
End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.Enabled = True
    cmdGrabar.SetFocus
End If
End Sub
