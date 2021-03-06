VERSION 5.00
Begin VB.Form frmCredSugAprob 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12810
   ForeColor       =   &H00000000&
   Icon            =   "frmCredSugAprob.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   320
      Left            =   11400
      TabIndex        =   6
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalExp 
      Caption         =   "Salida Exp."
      Height          =   320
      Left            =   4800
      TabIndex        =   5
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton cmdRIngObs 
      Caption         =   "Re-Ingreso Obs."
      Height          =   320
      Left            =   3240
      TabIndex        =   4
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdSalObs 
      Caption         =   "Salida x Obs."
      Height          =   320
      Left            =   1680
      TabIndex        =   3
      Top             =   6960
      Width           =   1575
   End
   Begin VB.CommandButton cmdIngExp 
      Caption         =   "Ingreso Exp."
      Height          =   320
      Left            =   120
      TabIndex        =   2
      Top             =   6960
      Width           =   1575
   End
   Begin VB.Frame fraCreditosSugAprob 
      Caption         =   " Lista de Cr?ditos "
      ForeColor       =   &H00FF0000&
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   12615
      Begin SICMACT.FlexEdit FECredSugAprob 
         Height          =   5055
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   8916
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Agencia-Cr?dito-Titular-Producto-Moneda-Monto-Ingreso Exp-Salida Observ-Re Ingreso Obs-cSalExpediente-nEstRevExpediente"
         EncabezadosAnchos=   "400-1800-1700-2600-1800-1000-1000-1800-1800-1800-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-L-R-L-L-C-R-C-C-C-R-R"
         FormatosEdit    =   "0-0-3-0-0-0-2-5-5-5-5-3"
         AvanceCeldas    =   1
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Opciones de Filtrado "
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   320
         Left            =   5280
         TabIndex        =   15
         Top             =   910
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar..."
         Height          =   320
         Left            =   5280
         TabIndex        =   14
         Top             =   610
         Width           =   1455
      End
      Begin VB.ComboBox cboAgencia 
         Height          =   315
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   920
         Width           =   3735
      End
      Begin VB.TextBox txtTitular 
         Height          =   285
         Left            =   1410
         TabIndex        =   11
         Top             =   620
         Width           =   3735
      End
      Begin VB.OptionButton optAge 
         Caption         =   "Agencia"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton optTit 
         Caption         =   "Titular"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   650
         Width           =   975
      End
      Begin VB.OptionButton optCredito 
         Caption         =   "Cr?dito"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   1290
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCredSugAprob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCredSugAprob
'*** Descripci?n : Formulario para registrar validacion expediente credito.
'*** Creaci?n : MIOL el 20120802, seg?n OYP-RFC065-2012
'********************************************************************
Option Explicit
Dim oNCOMColocEval As NCOMColocEval
Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
Dim rsCredSugAprob As Recordset
Dim rsSaldosVinculados As Recordset
Dim lnTipEstado As Integer
Dim lnDescTitulo As String
Dim lnTipoReg As TipoRegControl
Dim lnFilaSelec As Integer

Public Sub Inicio(ByVal pnTipEstado As Integer, ByVal pnDescTitulo As String, Optional ByVal pnTipoReg As TipoRegControl = gTpoRegCtrlInformeLegal)
    lnTipEstado = pnTipEstado
    lnDescTitulo = pnDescTitulo
    lnTipoReg = pnTipoReg
    Me.Show 1
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub cboAgencia_KeyPress(KeyAscii As Integer)
    cmdBuscar.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    Dim Msj As String
    Msj = ValidaFiltros
    If Msj = "" Then
        Screen.MousePointer = 11
        Call CargarDatos
        Call HabilitaOpciones(0, True)
        Screen.MousePointer = 0
    Else
        MsgBox Msj, vbInformation, "Alerta"
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdIngExp_Click()
    Set oNCOMColocEval = New NCOMColocEval
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Set rsCredSugAprob = New ADODB.Recordset
    Dim lcMovNro As String
    
    lnFilaSelec = FECredSugAprob.row
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call oNCOMColocEval.insEstadosExpediente(FECredSugAprob.TextMatrix(FECredSugAprob.row, 2), "En Asesoria Legal", lcMovNro, "", "", "", 1, 2001, lnTipoReg)
    MsgBox "Expediente Ingreso a Asesoria Legal", vbInformation, "Aviso"
    Me.cmdSalObs.Enabled = True
    Me.cmdIngExp.Enabled = False
    Me.cmdRIngObs.Enabled = False
    Me.cmdSalExp.Enabled = True
    Set oNCOMColocEval = Nothing
    Call CargarDatos
    Me.cmdRIngObs.Enabled = False
    Me.cmdIngExp.Enabled = False
    Me.cmdSalObs.Enabled = False
    Me.cmdSalExp.Enabled = False
    FECredSugAprob.row = lnFilaSelec
    FECredSugAprob.TopRow = lnFilaSelec
    Call FECredSugAprob_Click
End Sub

Private Sub cmdLimpiar_Click()
    Call LimpiarFormulario
End Sub

Private Sub cmdRIngObs_Click()
    Set oNCOMColocEval = New NCOMColocEval
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Set rsCredSugAprob = New ADODB.Recordset
    Dim lcMovNro As String

    lnFilaSelec = FECredSugAprob.row
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call oNCOMColocEval.updateEstadoExpediente(FECredSugAprob.TextMatrix(FECredSugAprob.row, 2), lnTipoReg) 'BY ARLO MODIFY 20171027
    Call oNCOMColocEval.insEstadosExpediente(FECredSugAprob.TextMatrix(FECredSugAprob.row, 2), "En Asesoria Legal", "", "", lcMovNro, "", 1, 2001, lnTipoReg)
    MsgBox "Re Ingreso de Expediente a Asesoria Legal", vbInformation, "Aviso"
    Me.cmdRIngObs.Enabled = False
    Me.cmdIngExp.Enabled = False
    Me.cmdSalObs.Enabled = False
    Me.cmdSalExp.Enabled = True
    Set oNCOMColocEval = Nothing
    Call CargarDatos
    Me.cmdRIngObs.Enabled = False
    Me.cmdIngExp.Enabled = False
    Me.cmdSalObs.Enabled = False
    Me.cmdSalExp.Enabled = False
    FECredSugAprob.row = lnFilaSelec
    FECredSugAprob.TopRow = lnFilaSelec
    Call FECredSugAprob_Click
End Sub

Private Sub cmdSalExp_Click()
    Set oNCOMColocEval = New NCOMColocEval
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Set rsCredSugAprob = New ADODB.Recordset
    Dim lcMovNro As String

    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call oNCOMColocEval.updateEstadoExpediente(FECredSugAprob.TextMatrix(FECredSugAprob.row, 2), lnTipoReg) 'BY ARLO MODIFY 20171027
    Call oNCOMColocEval.insEstadosExpediente(FECredSugAprob.TextMatrix(FECredSugAprob.row, 2), "Analista de Credito", "", "", "", lcMovNro, 2, 2001, lnTipoReg)
    MsgBox "Expediente Salio de Asesoria Legal", vbInformation, "Aviso"
    FECredSugAprob.EliminaFila FECredSugAprob.row
    Me.cmdRIngObs.Enabled = False
    Me.cmdIngExp.Enabled = False
    Me.cmdSalObs.Enabled = False
    Me.cmdSalExp.Enabled = False
    Set oNCOMColocEval = Nothing
    Call CargarDatos
    Me.cmdRIngObs.Enabled = False
    Me.cmdIngExp.Enabled = False
    Me.cmdSalObs.Enabled = False
    Me.cmdSalExp.Enabled = False
    Call cmdLimpiar_Click
End Sub

Private Sub cmdSalObs_Click()
    Set oNCOMColocEval = New NCOMColocEval
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Set rsCredSugAprob = New ADODB.Recordset
    Dim lcMovNro As String

    lnFilaSelec = FECredSugAprob.row
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call oNCOMColocEval.updateEstadoExpediente(FECredSugAprob.TextMatrix(FECredSugAprob.row, 2), lnTipoReg) 'BY ARLO MODIFY 20171027
    Call oNCOMColocEval.insEstadosExpediente(FECredSugAprob.TextMatrix(FECredSugAprob.row, 2), "Analista de Creditos", "", lcMovNro, "", "", 1, 2001, lnTipoReg)
    MsgBox "Expediente Salio por Observaci?n de Asesoria Legal", vbInformation, "Aviso"
    Me.cmdRIngObs.Enabled = True
    Me.cmdIngExp.Enabled = False
    Me.cmdSalObs.Enabled = False
    Me.cmdSalExp.Enabled = False
    Set oNCOMColocEval = Nothing
    Call CargarDatos
    Me.cmdRIngObs.Enabled = False
    Me.cmdIngExp.Enabled = False
    Me.cmdSalObs.Enabled = False
    Me.cmdSalExp.Enabled = False
    FECredSugAprob.row = lnFilaSelec
    FECredSugAprob.TopRow = lnFilaSelec
    Call FECredSugAprob_Click
End Sub

Private Sub FECredSugAprob_Click()
    Call ActivaBotones
End Sub

Private Sub FECredSugAprob_RowColChange()
    Call ActivaBotones
End Sub

Private Sub Form_Activate()
    FECredSugAprob.SetFocus
    Me.cmdRIngObs.Enabled = False
    Me.cmdIngExp.Enabled = False
    Me.cmdSalObs.Enabled = False
    Me.cmdSalExp.Enabled = False
    Call HabilitaOpciones(1)
End Sub

Private Sub Form_Load()
    Me.Caption = lnDescTitulo
    CargarAgencias
End Sub

Private Sub cargarCreditosSugAprob(ByVal pnTpoRegCtrl As TipoRegControl, ByVal psCtaCod As String, ByVal psTitular As String, ByVal psAgeCod As String)
    Set oNCOMColocEval = New NCOMColocEval
    Set rsCredSugAprob = New ADODB.Recordset
    Dim i As Integer
    
    Call LimpiaFlex(FECredSugAprob)
    
    
    Set rsCredSugAprob = oNCOMColocEval.obtenerCreditosSugeridosHipotecarios(pnTpoRegCtrl, psCtaCod, psTitular, psAgeCod)
    If Not rsCredSugAprob.BOF And Not rsCredSugAprob.EOF Then
        i = 1
        FECredSugAprob.lbEditarFlex = True
        Do While Not rsCredSugAprob.EOF
            FECredSugAprob.AdicionaFila
            FECredSugAprob.TextMatrix(i, 1) = rsCredSugAprob!cAgeDescripcion
            FECredSugAprob.TextMatrix(i, 2) = rsCredSugAprob!cCtaCod
            FECredSugAprob.TextMatrix(i, 3) = rsCredSugAprob!cPersNombre
            FECredSugAprob.TextMatrix(i, 4) = rsCredSugAprob!cConsDescripcion
            FECredSugAprob.TextMatrix(i, 5) = rsCredSugAprob!cMoneda
            FECredSugAprob.TextMatrix(i, 6) = Format(rsCredSugAprob!nMonto, gsFormatoNumeroView)
            FECredSugAprob.TextMatrix(i, 7) = rsCredSugAprob!cIngExpediente
            FECredSugAprob.TextMatrix(i, 8) = rsCredSugAprob!cSalObsExpediente
            FECredSugAprob.TextMatrix(i, 9) = rsCredSugAprob!cReIngObsExpediente
            FECredSugAprob.TextMatrix(i, 10) = rsCredSugAprob!cSalExpediente
            FECredSugAprob.TextMatrix(i, 11) = rsCredSugAprob!nEstRevExpediente
            i = i + 1
            rsCredSugAprob.MoveNext
        Loop
    Else
        MsgBox "No se encontraron datos", vbInformation, "Alerta"
    End If
    Set rsCredSugAprob = Nothing
    Set oNCOMColocEval = Nothing
End Sub

Private Sub CargarAgencias()
    Dim oAge As New COMDConstantes.DCOMAgencias
    Dim RS As New ADODB.Recordset
    
    Set RS = oAge.ObtieneAgencias
    Call CargaCombo(cboAgencia, RS)
    cboAgencia.ListIndex = 0
End Sub

Private Function ValidaFiltros() As String
    ValidaFiltros = ""
    If optCredito.value = True Then
        If ActXCodCta.NroCuenta = "" Then
            ValidaFiltros = "Ingrese el n?mero de cr?dito"
        ElseIf Len(ActXCodCta.NroCuenta) < 18 Then
            ValidaFiltros = "N?mero de cr?dito incorrecto"
        End If
    ElseIf optTit.value = True And txtTitular.Text = "" Then
        ValidaFiltros = "Ingrese el nombre del Titular"
    End If
End Function

Private Function HabilitaOpciones(ByVal pnOpcion As Integer, Optional ByVal pbBuscar As Boolean)
    If pbBuscar = True Then
        pnOpcion = 0
    End If
    optCredito.Enabled = Not pbBuscar
    optTit.Enabled = Not pbBuscar
    optAge.Enabled = Not pbBuscar
    cmdBuscar.Enabled = Not pbBuscar
    
    ActXCodCta.Enabled = IIf(pnOpcion = 1, True, False)
    txtTitular.Enabled = IIf(pnOpcion = 2, True, False)
    cboAgencia.Enabled = IIf(pnOpcion = 3, True, False)
    If pnOpcion = 1 Then
        ActXCodCta.SetFocus
        txtTitular.Text = ""
        cboAgencia.ListIndex = 0
    ElseIf pnOpcion = 2 Then
        txtTitular.SetFocus
        cboAgencia.ListIndex = 0
        ActXCodCta.NroCuenta = ""
    ElseIf pnOpcion = 3 Then
        cboAgencia.SetFocus
        ActXCodCta.NroCuenta = ""
        txtTitular.Text = ""
    End If
End Function

Private Sub optAge_Click()
    Call HabilitaOpciones(3)
End Sub

Private Sub optAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboAgencia.SetFocus
    End If
End Sub

Private Sub optCredito_Click()
    Call HabilitaOpciones(1)
End Sub

Private Sub optCredito_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ActXCodCta.SetFocus
    End If
End Sub

Private Sub optTit_Click()
    Call HabilitaOpciones(2)
End Sub

Private Sub optTit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTitular.SetFocus
    End If
End Sub

Private Sub txtTitular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub LimpiarFormulario()
    ActXCodCta.NroCuenta = ""
    txtTitular.Text = ""
    cboAgencia.ListIndex = 0
    FECredSugAprob.Clear
    FormateaFlex FECredSugAprob
    optCredito.value = True
    Call HabilitaOpciones(1, False)
End Sub

Private Sub CargarDatos()
    If optCredito.value = True Then
        Call cargarCreditosSugAprob(lnTipoReg, ActXCodCta.NroCuenta, "", "")
    ElseIf optTit.value = True Then
        Call cargarCreditosSugAprob(lnTipoReg, "", txtTitular.Text, "")
    Else
        Call cargarCreditosSugAprob(lnTipoReg, "", "", Mid(cboAgencia.Text, Len(cboAgencia.Text) - 1, 2))
    End If
End Sub

Private Sub ActivaBotones()
    cmdIngExp.Enabled = True
    If FECredSugAprob.TextMatrix(FECredSugAprob.row, 7) <> "" Then
        If FECredSugAprob.TextMatrix(FECredSugAprob.row, 8) <> "" Then
            cmdIngExp.Enabled = False
            cmdSalObs.Enabled = False
            cmdRIngObs.Enabled = True
            cmdSalExp.Enabled = False
        ElseIf FECredSugAprob.TextMatrix(FECredSugAprob.row, 9) <> "" Then
            cmdIngExp.Enabled = False
            cmdSalObs.Enabled = True
            cmdRIngObs.Enabled = False
            cmdSalExp.Enabled = True
        Else
            cmdIngExp.Enabled = False
            cmdSalObs.Enabled = True
            cmdRIngObs.Enabled = False
            cmdSalExp.Enabled = True
        End If
    Else
        cmdIngExp.Enabled = True    'BY ARLO MODIFY 20171027
        cmdSalObs.Enabled = False
        cmdRIngObs.Enabled = False
        cmdSalExp.Enabled = False
    End If
End Sub

Private Sub txtTitular_LostFocus()
    txtTitular.Text = UCase(txtTitular.Text)
End Sub
