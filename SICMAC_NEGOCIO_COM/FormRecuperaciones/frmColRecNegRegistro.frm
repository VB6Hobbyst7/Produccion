VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColRecNegRegistro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recuperaciones - Registro de Negociación "
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmColRecNegRegistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   27
      Top             =   5160
      Width           =   6735
      Begin VB.CommandButton cmdAnular 
         Caption         =   "&Anular"
         Height          =   390
         Left            =   480
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   390
         Left            =   5640
         TabIndex        =   33
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   390
         Left            =   480
         TabIndex        =   32
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   390
         Left            =   1512
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   390
         Left            =   2544
         TabIndex        =   30
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   3576
         TabIndex        =   29
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   390
         Left            =   4608
         TabIndex        =   28
         Top             =   240
         Width           =   870
      End
   End
   Begin VB.Frame fraComenta 
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   4440
      Width           =   6735
      Begin VB.TextBox txtNegComenta 
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   6375
      End
   End
   Begin VB.Frame fraNegociacion 
      Caption         =   "Negociación"
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
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6750
      Begin VB.CommandButton cmdCalendario 
         Caption         =   "Calen&dario"
         Height          =   345
         Left            =   4740
         TabIndex        =   24
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtNegNro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtNegCuotas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   3240
         TabIndex        =   15
         Top             =   600
         Width           =   1260
      End
      Begin VB.TextBox txtNegMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNegEstado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   5280
         TabIndex        =   13
         Top             =   240
         Width           =   1290
      End
      Begin MSMask.MaskEdBox TxtNegVigencia 
         Height          =   300
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lvwCalendario 
         Height          =   1440
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   2540
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto Pagado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Estado"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Nro Negoc"
         Height          =   225
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label11 
         Caption         =   "Cuotas"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label10 
         Caption         =   "Estado"
         Height          =   225
         Left            =   4680
         TabIndex        =   11
         Top             =   240
         Width           =   675
      End
      Begin VB.Label Label9 
         Caption         =   "Vigencia"
         Height          =   225
         Left            =   2520
         TabIndex        =   10
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "Monto Neg"
         Height          =   255
         Left            =   255
         TabIndex        =   8
         Top             =   600
         Width           =   825
      End
   End
   Begin VB.Frame fraCredito 
      Caption         =   "Credito"
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
      Height          =   1845
      Left            =   120
      TabIndex        =   0
      Top             =   30
      Width           =   6705
      Begin VB.ListBox lstCreditos 
         Height          =   645
         ItemData        =   "frmColRecNegRegistro.frx":030A
         Left            =   4800
         List            =   "frmColRecNegRegistro.frx":030C
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar ..."
         Height          =   315
         Left            =   3840
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   465
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   3615
         _extentx        =   6376
         _extenty        =   820
         texto           =   "Crédito"
         enabledcta      =   -1  'True
         enabledprod     =   -1  'True
         enabledage      =   -1  'True
      End
      Begin VB.Label lblSaldoCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5160
         TabIndex        =   18
         Top             =   1140
         Width           =   1380
      End
      Begin VB.Label lblEstudioJuridico 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1080
         TabIndex        =   17
         Top             =   1140
         Width           =   3015
      End
      Begin VB.Label lblCondicionCredito 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1080
         TabIndex        =   16
         Top             =   1440
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Condicion"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Estado"
         Height          =   270
         Index           =   1
         Left            =   4320
         TabIndex        =   7
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Est. Jurid."
         Height          =   240
         Index           =   4
         Left            =   240
         TabIndex        =   6
         Top             =   1185
         Width           =   735
      End
      Begin VB.Label lblNombreCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1080
         TabIndex        =   5
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo Cap"
         Height          =   240
         Index           =   2
         Left            =   4320
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblEstadoCredito 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5160
         TabIndex        =   3
         Top             =   1440
         Width           =   1380
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   660
      End
   End
   Begin RichTextLib.RichTextBox rtfImp 
      Height          =   315
      Left            =   60
      TabIndex        =   35
      Top             =   5040
      Visible         =   0   'False
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   556
      _Version        =   393217
      TextRTF         =   $"frmColRecNegRegistro.frx":030E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmColRecNegRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'***************************************************************************
'*  APLICACION : Recuperaciones - Negociaciones
'*  CREACION: 15/06/2004      AUTOR :LAYG
'*  MODIFICACION
'*  RESUMEN: PERMITE REGISTRAR LA NEGOCIACION DIRECTA PARA RECUPARAR UN CREDITO
'*************************************************************************
Option Explicit

Dim fsCodCta As String
Dim lbConexion As Boolean
Dim fbNuevaNegoc As Boolean
Dim fsNegAnterior As String
Dim fntipo As Integer


Public Sub Inicia(Optional ByVal pbRegistra As Boolean = True)
    fntipo = 2
    If pbRegistra = True Then
        cmdAnular.Visible = False
    Else
        cmdNuevo.Visible = False
        cmdModificar.Visible = False
        cmdCalendario.Visible = False
        cmdAnular.Visible = True
    End If
    Me.Show 1
End Sub


Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
If Len(Trim(AXCodCta.NroCuenta)) = 18 Then
    fsCodCta = Trim(AXCodCta.NroCuenta)
    'Mostrar Datos del Credito
    Call MuestraDatos(fsCodCta, fbNuevaNegoc)
 End If
End Sub

Private Sub cmdAnular_Click()

Dim lsSQL As String
Dim lsFechaHora As String
Dim lsNroNegocia As String
Dim lcRec As COMDColocRec.DCOMColRecNegociacion

If Len(fsCodCta) > 0 Then
    fsCodCta = AXCodCta.NroCuenta
    lsNroNegocia = Me.txtNegNro.Text
    
    If MsgBox("Desea Grabar Anulacion de Negociacion ? ", vbYesNo + vbQuestion, "Aviso") = vbYes Then
       
            'Anula negociacion Anterior
            If Len(Trim(lsNroNegocia)) > 0 Then
                Set lcRec = New COMDColocRec.DCOMColRecNegociacion
                    '27-12 Se Modifico
                    lcRec.AnularNegociacion lsNroNegocia, gdFecSis, gsCodUser
                Set lcRec = Nothing
            Else
                MsgBox "No se encontro Nro de NEGOCIACION"
                Exit Sub
            End If
        'Set loConec = Nothing ' Destruye la conexion
        cmdAnular.Enabled = False
        Call HabilitaControles(True, False, False, False, False, True, False, True, False)
        LimpiaDatos
        fbNuevaNegoc = False
        AXCodCta.NroCuenta = Mid(AXCodCta.NroCuenta, 1, 5)
        AXCodCta.SetFocusAge
    End If
End If
End Sub

Private Sub cmdBuscar_Click()
On Error GoTo ControlError

Dim loPers As comdpersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito 'DColRecCredito
Dim lrCreditos As ADODB.Recordset
Dim loCuentas As comdpersona.UCOMProdPersona 'UProdPersona

On Error GoTo ControlError

Set loPers = New comdpersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast

If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If

Set loCuentas = New comdpersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.Enabled = True
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    End If
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdCalendario_Click()
    lvwCalendario.ListItems.Clear
    Call frmColRecNegCalculaCalendario.Inicio(2, AXCodCta.NroCuenta)
End Sub

Private Sub cmdCancelar_Click()
    fbNuevaNegoc = False
    Call HabilitaControles(True, False, False, True, False, True, False, True)
    LimpiaDatos
End Sub

Private Sub cmdGrabar_Click()
Dim lsSQL As String
Dim lsFechaHora As String
Dim lsNroNegocia As String
Dim lnNumTranCta As Integer
Dim i As Integer, NumCuotas As Integer
Dim lsNegocOperac As String
Dim rs As New ADODB.Recordset
Dim lcRec As COMDColocRec.DCOMColRecNegociacion
Dim oCons As COMDConstSistema.DCOMGeneral
lsNegocOperac = "134001"

'********* VERIFICAR VISTO AVMM - 13-12-2006 **********************
'Dim loVisto As COMDColocRec.DCOMColRecCredito
'Set loVisto = New COMDColocRec.DCOMColRecCredito
'    '2=Negociacion
'    If loVisto.bVerificarVisto(AXCodCta.NroCuenta, 2) = False Then
'        MsgBox "No existe Visto para realizar Negociación", vbInformation, "Aviso"
'        Exit Sub
'    End If
'Set loVisto = Nothing
'********************************************************************

If ValidaDatos Then
    fsCodCta = AXCodCta.NroCuenta
    If MsgBox("Desea Grabar Registro de Negociacion ? ", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        'crear RS que contega  datos del calendario
   
    
With rs
    'Crear RecordSet
    .Fields.Append "dFecha", adDate
    .Fields.Append "nMonto", adCurrency
    .Open
    'Llenar Recordset
    For i = 1 To Me.lvwCalendario.ListItems.Count
        .AddNew
        .Fields("dFecha") = Me.lvwCalendario.ListItems.Item(i).SubItems(1)
        .Fields("nMonto") = Me.lvwCalendario.ListItems.Item(i).SubItems(2)
    Next i
    
End With
        Set oCons = New COMDConstSistema.DCOMGeneral
            lsFechaHora = oCons.FechaHora(gdFecSis)
        Set oCons = Nothing
        lnNumTranCta = 1
        lsNroNegocia = ObtieneSgteNroNegocia
        Set lcRec = New COMDColocRec.DCOMColRecNegociacion
            lcRec.InsertarNegociacion TxtNegVigencia.Text, fsNegAnterior, fsCodCta, lsNroNegocia, txtNegMonto.Text, _
                                      txtNegCuotas.Text, txtNegComenta.Text, gsCodUser, lsFechaHora, lnNumTranCta, _
                                      lsNegocOperac, gsCodAge, rs
        Set lcRec = Nothing
        cmdGrabar.Enabled = False
        Call HabilitaControles(True, False, False, False, False, True, False, True)
        LimpiaDatos
        fbNuevaNegoc = False

        AXCodCta.NroCuenta = Mid(AXCodCta.NroCuenta, 1, 5)
        AXCodCta.SetFocusAge

    End If
End If
End Sub

Private Sub cmdImprimir_Click()
Dim lsCadena As String
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim lcRecImp As COMNColocRec.NCOMColRecImpre

' llenar recordset con datos del calendario
With rs
    'Crear RecordSet
    .Fields.Append "nNro", adInteger
    .Fields.Append "dFecha", adDate
    .Fields.Append "nMonto", adCurrency
    .Open
    'Llenar Recordset
    For i = 1 To Me.lvwCalendario.ListItems.Count
        .AddNew
        .Fields("nNro") = Me.lvwCalendario.ListItems.Item(i).Text
        .Fields("dFecha") = Me.lvwCalendario.ListItems.Item(i).SubItems(1)
        .Fields("nMonto") = Me.lvwCalendario.ListItems.Item(i).SubItems(2)
    Next i
    
End With

Set lcRecImp = New COMNColocRec.NCOMColRecImpre
    lsCadena = lcRecImp.ImprimeNegociacion(gsNomCmac, gdFecSis, gsNomAge, gsCodUser, Me.AXCodCta.NroCuenta, lblNombreCliente, lblEstudioJuridico, txtNegMonto.Text, rs, fntipo, gImpresora)
Set lcRecImp = Nothing

rtfImp.Text = lsCadena

Dim loPrevio As previo.clsPrevio
    If Len(Trim(lsCadena)) > 0 Then
        Set loPrevio = New previo.clsPrevio
        loPrevio.Show lsCadena, "Recuperaciones - Negociaciones ", True
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte", vbInformation, "Aviso"
    End If

End Sub

Private Sub CmdModificar_Click()
    If Me.TxtNegVigencia = gdFecSis Then
        fbNuevaNegoc = False
        Call HabilitaControles(False, False, True, True, False, True, True, False)
    Else
        MsgBox "NO se puede modificar Convenio de fecha anterior", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdNuevo_Click()
    'fraNegociacion.Enabled = True
    fbNuevaNegoc = True
    LimpiaDatos
    Call HabilitaControles(False, False, True, True, False, True, True, True)
    cmdBuscar.SetFocus
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 LimpiaDatos
 Call HabilitaControles(True, False, False, False, False, True, False, True)
End Sub

Private Sub lstCreditos_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Len(lstCreditos.Text) > 0 Then
     AXCodCta.NroCuenta = Mid(Trim(lstCreditos.Text), 1, 18)
     lstCreditos.Visible = False
     AXCodCta.Enabled = True
     AXCodCta.SetFocusAge
  End If
End Sub
Private Sub MuestraDatos(ByVal psCodCta As String, ByVal pbNuevaNeg As Boolean)
'**** Datos del Credito
Dim lcRec As COMDColocRec.DCOMColRecNegociacion
Call MuestraDatosCredito(psCodCta)
If pbNuevaNeg = False Then  ' Muestra datos de Negociacion Actual
    Call MuestraDatosNegocia(psCodCta)
Else ' Negociacion Nueva
    'Verifica si tiene Negociacion Activa
    Dim lrNegAct As New ADODB.Recordset
    Dim lsSQL As String
   
    Set lcRec = New COMDColocRec.DCOMColRecNegociacion
        Set lrNegAct = lcRec.ObtenerNroNeg(psCodCta)
    Set lcRec = Nothing
    
    If Not (lrNegAct.BOF And lrNegAct.EOF) Then
        fsNegAnterior = lrNegAct!cNroNeg
        If MsgBox("El cliente Tiene una negociacion activa, Desea Anularla ? ", vbInformation + vbYesNo, "Aviso") = vbNo Then
            Call MuestraDatosNegocia(psCodCta)
            Call HabilitaControles(True, True, False, True, True, True, False, False, True)
            Exit Sub
        End If
    Else
        fsNegAnterior = ""
    End If
        Set lrNegAct = Nothing
    cmdCalendario.SetFocus
End If

End Sub

Private Sub MuestraDatosCredito(ByVal psCtaCod As String)
On Error GoTo ControlError
Dim lcRec As COMDColocRec.DCOMColRecNegociacion
Dim reg As New ADODB.Recordset
Dim lsSQL As String

' Busca el Credito
    Set lcRec = New COMDColocRec.DCOMColRecNegociacion
        Set reg = lcRec.ObtenerDatosCredito(psCtaCod)
    Set lcRec = Nothing
    If reg.BOF And reg.EOF Then
        reg.Close
        Set reg = Nothing
        MsgBox " No se encuentra el Credito " & fsCodCta, vbInformation, " Aviso "
        LimpiaDatos
        AXCodCta.Enabled = True
        Exit Sub
    Else
        ' Mostrar los datos del Credito
        Me.lblNombreCliente = PstaNombre(reg!cNomClie, False)
        Me.lblEstudioJuridico = PstaNombre(reg!cNomAbog, False)
        Me.lblEstadoCredito = fgEstadoColRecupDesc(reg!nPrdEstado)
        Me.lblCondicionCredito = fgCondicionColRecupDesc(reg!nPrdEstado)
        Me.lblSaldoCapital = Format(reg!nSaldo, "#0.00")
        AXCodCta.Enabled = False
     End If
    Set reg = Nothing
Exit Sub
ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub MuestraDatosNegocia(ByVal psCodCta As String)
On Error GoTo ControlError
Dim lcRec As COMDColocRec.DCOMColRecNegociacion
Dim reg As New ADODB.Recordset
Dim lsSQL As String
Dim L As ListItem
    ' Busca la Negociacion
    Set lcRec = New COMDColocRec.DCOMColRecNegociacion
        Set reg = lcRec.ObtenerDatosCredNegociacion(psCodCta)
    Set lcRec = Nothing
    
    If reg.BOF And reg.EOF Then
        MsgBox " No se Tiene Negociaciones Vigentes para Credito  " & psCodCta, vbInformation, " Aviso "
        'LimpiaDatos
        AXCodCta.Enabled = True
        Exit Sub
    Else
        ' Mostrar los datos de Negociacion
        Me.txtNegNro.Text = reg!cNroNeg
        Me.TxtNegVigencia.Text = Format(reg!dFecVig, "dd/mm/yyyy")
        Me.txtNegEstado.Text = IIf(reg!cEstado = "V", "Vigente", "Cancelado")
        Me.txtNegMonto.Text = Format(reg!nMontoNeg, "#0.00")
        Me.txtNegCuotas.Text = Format(reg!nCuotasNeg, "#0.00")
        Me.txtNegComenta.Text = IIf(IsNull(reg!cComenta), "", reg!cComenta)
        reg.Close
        Set reg = Nothing
    End If
        
   ' Busca Plan de Pagos de Negociacion
    Set lcRec = New COMDColocRec.DCOMColRecNegociacion
        Set reg = lcRec.ObtenerPlanPagosNegocia(psCodCta, Me.txtNegNro.Text)
    Set lcRec = Nothing

    If reg.BOF And reg.EOF Then
        MsgBox " Negociacion No Posee Plan Pagos " & psCodCta, vbInformation, " Aviso "
    Else
        ' Mostrar Plan de Pagos
        reg.MoveFirst
        Do While Not reg.EOF
            Set L = lvwCalendario.ListItems.Add(, , Trim(Str(reg!nNroCuota)))
            L.SubItems(1) = Format(reg!dFecVenc, "dd/mm/yyyy")
            L.SubItems(2) = Format(reg!nMonto, "#0.00")
            L.SubItems(3) = Format(reg!nMontoPag, "#0.00")
            L.SubItems(4) = reg!cEstado
            reg.MoveNext
        Loop
    End If
 
    Call HabilitaControles(True, True, False, True, True, True, False, True, True)

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "


End Sub

'*  VALIDACION DE DATOS DEL FORMULARIO ANTES DE GRABAR
Function ValidaDatos() As Boolean

Dim reg As New ADODB.Recordset
Dim lsSQL As String
Dim MonGarant As Currency
 
ValidaDatos = True
'valida fecha Vigencia
    If ValidaFecha(Me.TxtNegVigencia) <> "" Then
        MsgBox "No se registro Fecha de Vigencia", vbInformation, "Aviso"
        ValidaDatos = False
        TxtNegVigencia.Enabled = True
        TxtNegVigencia.SetFocus
        Exit Function
    End If
'Valida que se ha llenado el calendario
    If lvwCalendario.ListItems.Count < 1 Then
        MsgBox "No se ha generado el Calendario", vbInformation, "Aviso"
        ValidaDatos = False
        cmdCalendario.SetFocus
        Exit Function
    End If


End Function

'****************************************************************
'*  LIMPIA LOS DATOS DE LA PANTALLA PARA UNA NUEVA APROBACION
'****************************************************************
Sub LimpiaDatos()
    lblNombreCliente.Caption = ""
    lblEstudioJuridico.Caption = ""
    lblCondicionCredito.Caption = ""
    lblEstadoCredito.Caption = ""
    lblSaldoCapital.Caption = ""
    txtNegNro.Text = ""
    TxtNegVigencia.Text = "__/__/____"
    txtNegMonto.Text = ""
    txtNegCuotas.Text = ""
    txtNegEstado.Text = ""
    txtNegComenta.Text = ""
    lvwCalendario.ListItems.Clear
    AXCodCta.Texto = ""
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
    fsCodCta = ""
    fsNegAnterior = ""
    AXCodCta.NroCuenta = fgIniciaAxCuentaRecuperaciones
    'fraNegociacion.Enabled = False
End Sub

Private Sub HabilitaControles(ByVal pbNuevo As Boolean, ByVal pbModifica As Boolean, _
        ByVal pbGrabar As Boolean, ByVal pbCancelar As Boolean, ByVal pbImprimir As Boolean, _
        ByVal pbSalir As Boolean, ByVal pbEditaCtrl As Boolean, ByVal pbBuscar As Boolean, _
        Optional ByVal pbAnular As Boolean = False)

cmdNuevo.Enabled = pbNuevo
cmdModificar.Enabled = pbModifica
cmdGrabar.Enabled = pbGrabar
cmdCancelar.Enabled = pbCancelar
cmdImprimir.Enabled = pbImprimir
cmdSalir.Enabled = pbSalir
cmdAnular.Enabled = pbAnular

'axCodCta.Enabled = pbEditaCtrl
cmdCalendario.Enabled = pbEditaCtrl
fraComenta.Enabled = pbEditaCtrl

cmdBuscar.Enabled = pbBuscar
End Sub

Private Function ObtieneSgteNroNegocia() As String
Dim lsSQL As String
Dim lr As ADODB.Recordset
Dim lsNroNeg As String
Dim lcRec As COMDColocRec.DCOMColRecNegociacion

    Set lcRec = New COMDColocRec.DCOMColRecNegociacion
        Set lr = lcRec.ObtenerNroNegMax
    Set lcRec = Nothing
    If IsNull(lr!UltNeg) Then
        lsNroNeg = FillNum(1, 6, "0")
    Else
        lsNroNeg = FillNum(lr!UltNeg + 1, 6, "0")
    End If
    ObtieneSgteNroNegocia = lsNroNeg

End Function


Private Sub txtNegComenta_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
     If KeyAscii = 13 Then
        cmdGrabar.SetFocus
     End If
End Sub

Private Sub txtNegMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.txtNegCuotas.SetFocus
End If
End Sub
