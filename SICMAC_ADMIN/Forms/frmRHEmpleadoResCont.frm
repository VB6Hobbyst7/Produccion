VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRHEmpleadoResCont 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frmRHEmpleadoResCont.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2445
      TabIndex        =   11
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlanillaLiquidacion 
      Caption         =   "&Liquidación"
      Height          =   375
      Left            =   1260
      TabIndex        =   9
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Frame fraComentario 
      Caption         =   "Motivo de Rescisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1110
      Left            =   60
      TabIndex        =   5
      Top             =   3075
      Width           =   7920
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   300
         Left            =   6540
         TabIndex        =   10
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar TxtBuscar 
         Height          =   300
         Left            =   105
         TabIndex        =   8
         Top             =   255
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         Appearance      =   0
         BackColor       =   -2147483624
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   105
         MaxLength       =   300
         TabIndex        =   6
         Top             =   600
         Width           =   7725
      End
      Begin VB.Label lblMotivo 
         Height          =   180
         Left            =   1815
         TabIndex        =   7
         Top             =   330
         Width           =   4665
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6900
      TabIndex        =   2
      Top             =   4260
      Width           =   1095
   End
   Begin VB.CommandButton cmdRescindir 
      Caption         =   "&Rescindir"
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Top             =   4260
      Width           =   1095
   End
   Begin Sicmact.ctrRRHHGen ctrRRHHGen 
      Height          =   1200
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   2117
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame fraResCon 
      Caption         =   "Contratos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1860
      Left            =   45
      TabIndex        =   3
      Top             =   1200
      Width           =   7935
      Begin Sicmact.FlexEdit FlexContrato 
         Height          =   1545
         Left            =   90
         TabIndex        =   4
         ToolTipText     =   "Haga doble Click Sobre el contrato que dese visualizar"
         Top             =   210
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   2725
         Cols0           =   13
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-NumContrato-Tipo-Fecha Fin-Contrato Ini-Contrato Fin-Area Asig-Age Asig-Area Actual-Age Actual-Cargo-Sueldo-Comentario"
         EncabezadosAnchos=   "300-1200-2000-1200-1200-1200-2500-2500-2500-2500-2500-1500-4000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-L-R-R-R-L-L-L-L-L-R-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmRHEmpleadoResCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ALPA 20090122***********************
Dim objPista As COMManejador.Pista
'************************************
Private Sub cmdCancelar_Click()
    ClearScreen
    Activa False
End Sub

Private Sub cmdPlanillaLiquidacion_Click()
    Me.Enabled = False
        frmRHPlanillas.IniResincion gTipoProcesoRRHHCalculo, "RECURSOS HUMANOS:PROCESOS:CALCULO DE PLANILLAS", Me, gsRHPlanillaLiquidacion, Me.ctrRRHHGen.psCodigoPersona, Me.ctrRRHHGen.psNombreEmpledo, Me.ctrRRHHGen.psCodigoEmpleado
    Me.Enabled = True
End Sub

Private Sub cmdRescindir_Click()
    Dim oRh As NActualizaDatosRRHH
    Set oRh = New NActualizaDatosRRHH
    If Not Valida Then Exit Sub
    If MsgBox("Desea Rescindir el Último Contrato ?", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
     'ALPA 20090122*************************************
    glsMovNro = GetMovNro(gsCodUser, gsCodAge)
    '**************************************************
    oRh.RescindeRRHH Me.ctrRRHHGen.psCodigoPersona, Me.FlexContrato.TextMatrix(1, 1), Format(CDate(Me.mskFecha.Text), gsFormatoFecha), Me.txtbuscar.Text, Me.txtComentario.Text, glsMovNro
    
    CargaData Me.ctrRRHHGen.psCodigoPersona, False
   
    'ALPA 20090122*************************************
    gsOpeCod = LogPistaRescindirContrato
    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "2", , Me.FlexContrato.TextMatrix(1, 1), gNumeroContrato
     '**************************************************
    Me.cmdCancelar.Enabled = False
    Me.cmdRescindir.Enabled = False
    Me.cmdPlanillaLiquidacion.SetFocus
    Me.fraComentario.Enabled = False
    Me.cmdSalir.Enabled = True
End Sub
 
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub ctrRRHHGen_EmiteDatos()
    Dim oPersona As UPersona
    Dim oRRHH As DActualizaDatosRRHH
    ClearScreen
    Set oRRHH = New DActualizaDatosRRHH
    Set oPersona = New UPersona
    Set oPersona = frmBuscaPersona.Inicio(True)
    If Not oPersona Is Nothing Then
        Me.ctrRRHHGen.psCodigoPersona = oPersona.sPersCod
        Me.ctrRRHHGen.psNombreEmpledo = oPersona.sPersNombre
        Me.ctrRRHHGen.psCodigoEmpleado = oRRHH.GetCodigoEmpleado(Me.ctrRRHHGen.psCodigoPersona)
        CargaData Me.ctrRRHHGen.psCodigoPersona
    End If
End Sub

Private Sub CargaData(psPersCod As String, Optional pbVerfFechas As Boolean = True)
    Dim oRh As DActualizaDatosRRHH
    Dim rsRH As ADODB.Recordset
    Set rsRH = New ADODB.Recordset
    Set oRh = New DActualizaDatosRRHH

    Set rsRH = oRh.GetRRHHContratos(psPersCod)
    SetFlexEdit Me.FlexContrato, rsRH
    
    rsRH.Close
    Set rsRH = Nothing
    Activa True
    
    If Me.FlexContrato.TextMatrix(1, 3) <> "" And pbVerfFechas Then
        MsgBox "El ultimo contrato ya ha rescindido. Solo puede Liquidarlo.", vbInformation, "Aviso"
        Me.cmdRescindir.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    
    Set rsC = oCon.GetPeriodosNoLabTpo(RHEstadosTpoPorLiquidar)
    Me.txtbuscar.rs = rsC
    Set oCon = Nothing
    Activa False
    'ALPA 20090122 ***************************************************************************
    Set objPista = New COMManejador.Pista
    '*****************************************************************************************
End Sub

Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 10
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComentario.SetFocus
    End If
End Sub

Private Sub txtBuscar_EmiteDatos()
    Me.lblMotivo.Caption = Me.txtbuscar.psDescripcion
    Me.mskFecha.SetFocus
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFecha.SetFocus
    End If
End Sub

Private Sub txtComentario_GotFocus()
    txtComentario.SelStart = 0
    txtComentario.SelLength = 320
End Sub

Private Sub Activa(pbvalor As Boolean)
    Me.ctrRRHHGen.Enabled = Not pbvalor
    Me.cmdSalir.Enabled = Not pbvalor
    Me.cmdPlanillaLiquidacion.Enabled = pbvalor
    Me.fraComentario.Enabled = pbvalor
    Me.cmdRescindir.Enabled = pbvalor
    Me.cmdCancelar.Enabled = pbvalor
End Sub

Private Sub ClearScreen()
    Me.txtbuscar.Text = ""
    Me.txtComentario.Text = ""
    Me.mskFecha.Text = "__/__/____"
    Me.ctrRRHHGen.ClearScreen
    FlexContrato.Clear
    FlexContrato.Rows = 2
    Me.FlexContrato.FormaCabecera
    
End Sub

Private Function Valida() As Boolean
    If Me.txtbuscar.Text = "" Then
        MsgBox "Debe Ingresar un Tipo de Rescisión del Contrato.", vbInformation, "Aviso"
        Me.txtbuscar.SetFocus
        Valida = False
    ElseIf Not IsDate(Me.mskFecha.Text) Then
        MsgBox "Debe Ingresar una fecha valida para la Rescisión del Contrato.", vbInformation, "Aviso"
        Me.mskFecha.SetFocus
        Valida = False
    ElseIf Me.txtComentario.Text = "" Then
        MsgBox "Debe Ingresar un comentario de la Rescisión del Contrato.", vbInformation, "Aviso"
        Me.txtComentario.SetFocus
        Valida = False
    ElseIf Not (CDate(Me.mskFecha.Text) >= CDate(Me.FlexContrato.TextMatrix(1, 4)) And CDate(Me.mskFecha.Text) <= CDate(Me.FlexContrato.TextMatrix(1, 5))) Then
        MsgBox "Debe Ingresar una fecha que sea mayor a la fecha de inicio de contrato y menor o igual a la fecha de fin el contrato.", vbInformation, "Aviso"
        Me.txtComentario.SetFocus
        Valida = False
    Else
        Valida = True
    End If
End Function

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdRescindir.SetFocus
    End If
End Sub

Public Sub Ini(psCaption As String, pMdi As Form)
    Caption = psCaption
    Me.Show , pMdi
End Sub

Private Sub ctrRRHHGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oRRHH As DActualizaDatosRRHH
        Dim rsR As ADODB.Recordset
        Set oRRHH = New DActualizaDatosRRHH
        ctrRRHHGen.psCodigoEmpleado = Left(ctrRRHHGen.psCodigoEmpleado, 1) & Format(Trim(Mid(ctrRRHHGen.psCodigoEmpleado, 2)), "00000")
        Dim oCon As DActualizaDatosContrato
        Set oCon = New DActualizaDatosContrato
        
        Set rsR = oRRHH.GetRRHH(ctrRRHHGen.psCodigoEmpleado, gPersIdDNI)
           
        If Not (rsR.EOF And rsR.BOF) Then
            ctrRRHHGen.SpinnerValor = CInt(Right(ctrRRHHGen.psCodigoEmpleado, 5))
            ctrRRHHGen.psCodigoPersona = rsR.Fields("Codigo")
            ctrRRHHGen.psNombreEmpledo = rsR.Fields("Nombre")
            rsR.Close
            Set rsR = oRRHH.GetRRHHGeneralidades(ctrRRHHGen.psCodigoEmpleado)
            CargaData Me.ctrRRHHGen.psCodigoPersona
        Else
            MsgBox "Codigo no Reconocido.", vbInformation, "Aviso"
            ClearScreen
            ctrRRHHGen.SetFocus
        End If
        
        rsR.Close
        Set rsR = Nothing
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

