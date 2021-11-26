VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColBienesAdjudicacion 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   Icon            =   "frmColRecListaAdjudBienes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SICMACT.EditMoney txtValorEmb 
      Height          =   285
      Left            =   1920
      TabIndex        =   41
      Top             =   3240
      Width           =   1215
      _extentx        =   2355
      _extenty        =   503
      font            =   "frmColRecListaAdjudBienes.frx":030A
      appearance      =   0
      text            =   "0"
      enabled         =   -1
   End
   Begin VB.CommandButton cmdProDepre 
      Caption         =   "Prorrogar Depreciación"
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame fraVenta 
      Caption         =   "Datos de la Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   34
      Top             =   4320
      Width           =   8655
      Begin MSMask.MaskEdBox txtFecVenta 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   35
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin SICMACT.EditMoney lblValVenta 
         Height          =   285
         Left            =   1680
         TabIndex        =   36
         Top             =   240
         Width           =   1215
         _extentx        =   2566
         _extenty        =   661
         font            =   "frmColRecListaAdjudBienes.frx":0336
         appearance      =   0
         text            =   "0"
         enabled         =   -1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor de Venta:"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha de Venta:"
         Height          =   255
         Left            =   3000
         TabIndex        =   37
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4200
      TabIndex        =   14
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7320
      TabIndex        =   16
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame fraAdjudicacion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar ..."
         CausesValidation=   0   'False
         Height          =   375
         Left            =   7410
         TabIndex        =   42
         Top             =   240
         Width           =   1005
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos del Bien"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1335
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   8415
         Begin VB.TextBox txtDescripcion 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   4
            Top             =   600
            Width           =   6495
         End
         Begin VB.TextBox txtPElectronica 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   2
            Top             =   240
            Width           =   1575
         End
         Begin SICMACT.TxtBuscar txtAF 
            Height          =   315
            Left            =   4440
            TabIndex        =   3
            Top             =   240
            Width           =   3810
            _extentx        =   6720
            _extenty        =   556
            appearance      =   0
            appearance      =   0
            font            =   "frmColRecListaAdjudBienes.frx":0362
            appearance      =   0
            enabledtext     =   0
         End
         Begin VB.Label Label16 
            Caption         =   "Tipo Bien:"
            Height          =   255
            Left            =   3600
            TabIndex        =   28
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "Descripción del Bien:"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label5 
            Caption         =   "N° Partida Electrónica:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos de Adjudicación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1455
         Left            =   120
         TabIndex        =   18
         Top             =   2520
         Width           =   8415
         Begin VB.ComboBox cboTpoAdj 
            Height          =   315
            ItemData        =   "frmColRecListaAdjudBienes.frx":038E
            Left            =   1680
            List            =   "frmColRecListaAdjudBienes.frx":039B
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   240
            Width           =   1935
         End
         Begin VB.ComboBox cboMoneda 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmColRecListaAdjudBienes.frx":03CF
            Left            =   6960
            List            =   "frmColRecListaAdjudBienes.frx":03D9
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   960
            Width           =   1215
         End
         Begin MSMask.MaskEdBox txtFecAdjudic 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin SICMACT.EditMoney txtValComercial 
            Height          =   285
            Left            =   4440
            TabIndex        =   9
            Top             =   600
            Width           =   1215
            _extentx        =   2566
            _extenty        =   661
            font            =   "frmColRecListaAdjudBienes.frx":03ED
            appearance      =   0
            text            =   "0"
            enabled         =   -1
         End
         Begin SICMACT.EditMoney txtValTasacion 
            Height          =   285
            Left            =   6960
            TabIndex        =   10
            Top             =   600
            Width           =   1215
            _extentx        =   2566
            _extenty        =   661
            font            =   "frmColRecListaAdjudBienes.frx":0419
            appearance      =   0
            text            =   "0"
            enabled         =   -1
         End
         Begin MSMask.MaskEdBox txtFecTasacion 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   3
            EndProperty
            Height          =   255
            Left            =   4440
            TabIndex        =   12
            Top             =   960
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin SICMACT.EditMoney txtCapital 
            Height          =   285
            Left            =   4440
            TabIndex        =   6
            Top             =   240
            Width           =   1215
            _extentx        =   2355
            _extenty        =   503
            font            =   "frmColRecListaAdjudBienes.frx":0445
            appearance      =   0
            text            =   "0"
            enabled         =   -1
         End
         Begin SICMACT.EditMoney txtInteres 
            Height          =   285
            Left            =   6960
            TabIndex        =   7
            Top             =   240
            Width           =   1215
            _extentx        =   2143
            _extenty        =   503
            font            =   "frmColRecListaAdjudBienes.frx":0471
            appearance      =   0
            text            =   "0"
            enabled         =   -1
         End
         Begin VB.Label txtValAdjudicacion 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   1680
            TabIndex        =   8
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Int. y Gastos:"
            Height          =   255
            Left            =   5760
            TabIndex        =   33
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "Capital:"
            Height          =   255
            Left            =   3600
            TabIndex        =   32
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo:"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   5640
            TabIndex        =   27
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha Tasación:"
            Height          =   255
            Left            =   3120
            TabIndex        =   25
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Fecha Adjudic:"
            Height          =   255
            Left            =   360
            TabIndex        =   22
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Tasación:"
            Height          =   255
            Left            =   5640
            TabIndex        =   21
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Comercial:"
            Height          =   255
            Left            =   3120
            TabIndex        =   20
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Valor Adjudicación:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1455
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3615
         _extentx        =   6376
         _extenty        =   661
         texto           =   "Crédito"
         enabledcta      =   -1
         enabledprod     =   -1
         enabledage      =   -1
         cmac            =   "109"
      End
      Begin VB.Label lblPersCod 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   720
         TabIndex        =   31
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lbltitular 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2640
         TabIndex        =   29
         Top             =   720
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Titular:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modificar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   39
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "frmColBienesAdjudicacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTipoProceso As Integer
Dim nNumAdjudicacion As Integer
Dim bModificar As Boolean
Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = AXCodCta.NroCuenta
        ObtieneDatosCuenta sCta
    End If
End Sub
Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsPers As ADODB.Recordset
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPers = New ADODB.Recordset
    Set rsPers = clsMant.GetPersonaCuenta(sCuenta)

    If Not (rsPers.EOF And rsPers.BOF) Then
        lblPersCod.Caption = rsPers("cPersCod")
        lbltitular.Caption = UCase(PstaNombre(rsPers("Nombre")))
        Me.txtPElectronica.SetFocus
        cmdAceptar.Enabled = True
    Else
        MsgBox "No se ha encontrado información de la cuenta ingresada"
        AXCodCta.SetFocus
    End If
End Sub
'JIPR20190520 INICIO
Private Sub cmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersCredito  As COMDColocRec.DCOMColRecCredito 'DColRecCredito
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
lsEstados = gColocEstRecVigJud & "," & gColocEstRecVigCast & "," & gColocEstVigNorm & "," & gColocEstVigVenc & "," & gColocEstVigMor & "," & gColocEstRefNorm & "," & gColocEstRefVenc & "," & gColocEstRefMor & "," & gColocEstRecRefiJud

AXCodCta.Enabled = True
    
If Trim(lsPersCod) <> "" Then
    Set loPersCredito = New COMDColocRec.DCOMColRecCredito
        Set lrCreditos = loPersCredito.dObtieneCreditosDePersona(lsPersCod, lsEstados)
    Set loPersCredito = Nothing
End If


Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        lblPersCod.Caption = lsPersCod
        lbltitular.Caption = UCase(lsPersNombre)
        AXCodCta.SetFocusCuenta
    End If
     cmdAceptar.Enabled = True
Set loCuentas = Nothing

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub
'JIPR20190520 FIN

Private Sub CmdModificar_Click()
Dim oCnt2 As COMNContabilidad.NCOMContFunciones
Set oCnt2 = New COMNContabilidad.NCOMContFunciones
Dim pRs2 As ADODB.Recordset
Set pRs2 = oCnt2.ObtenerDatosBienAdjudicado(nNumAdjudicacion)

    bModificar = True
    cmdModificar.Visible = False
    cmdAceptar.Visible = True
    cmdCancelar.Visible = True
    cmdAceptar.Enabled = True
    fraAdjudicacion.Enabled = True
    If CDec(lblValVenta.Text) > 0 Then
        fraVenta.Enabled = True
    End If
    
    If pRs2!nTipoAdjud = 3 Then
    Me.txtValorEmb.Visible = True
    Me.txtValAdjudicacion.Visible = False
    End If
    
    txtPElectronica.SetFocus
End Sub

Private Sub cmdProDepre_Click()
Dim oCnt As COMNContabilidad.NCOMContFunciones
Set oCnt = New COMNContabilidad.NCOMContFunciones
Dim sMovNro As String
    If MsgBox("¿Está seguro de dar la prórroga de depreciación?", vbQuestion + vbYesNo, "Advertencia") = vbNo Then Exit Sub
    sMovNro = oCnt.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    If oCnt.RegistraProrrogaDepreAdjudicado(nNumAdjudicacion, sMovNro) Then MsgBox "La prórroga se ha realizado correctamente.", vbInformation, "¡Aviso!": cmdProDepre.Enabled = False
End Sub

Private Sub Form_Load()
'JIPR20190520 INICIO
    'Dim oConst As DCOMConstantes
    'Dim rs As ADODB.Recordset
    'Set oConst = New DCOMConstantes
    'Me.txtAF.rs = oConst.ObtieneTipoBienGarantia
    'Set oConst = Nothing
'JIPR20190520 FIN

    Me.cboTpoAdj.ListIndex = 0
    Me.cboMoneda.ListIndex = 0
    cmdAceptar.Enabled = False
End Sub
'Public Sub Inicio(nTpoProc As Integer, Optional nNumAdj As Integer = 0)
Public Sub Inicio(nTpoProc As Integer, Optional nNumAdj As Integer = 0, Optional pbBienAdj As Boolean = False) 'PASI20161103 ERS0572016
'PASI20161103 ERS0572016************
Dim oPers As COMDPersona.UCOMAcceso
Dim oCnt As COMNContabilidad.NCOMContFunciones
Dim gsGrupo As String
Set oCnt = New COMNContabilidad.NCOMContFunciones
Dim oConst As DCOMConstantes 'JIPR20190520
Set oConst = New DCOMConstantes 'JIPR20190520
Dim rs As ADODB.Recordset 'JIPR20190520

'PASI END ***************************
    nTipoProceso = nTpoProc
    Call cmdCancelar_Click
    If nTipoProceso = 1 Then ' Adjudicacion
        Me.fraVenta.Enabled = False
        Me.fraAdjudicacion.Enabled = True
        Me.Caption = "REGISTRO DE BIENES ADJUDICADOS"
        Me.cmdModificar.Visible = False
        
         'JIPR20190520 INICIO
        cboTpoAdj.Clear
        cboTpoAdj.AddItem "1. Adjudicación", 0
        cboTpoAdj.AddItem "2. Dación de Pago", 1
        Me.txtAF.rs = oConst.ObtieneTipoBienGarantia
        Me.txtValorEmb.Visible = False
        Me.txtValAdjudicacion.Visible = True
        Set oConst = Nothing
        Me.cmdBuscar.Visible = True
         'JIPR20190520 FIN
        
    ElseIf nTipoProceso = 2 Then ' Venta
        nNumAdjudicacion = nNumAdj
        Me.fraVenta.Enabled = True
        Me.fraAdjudicacion.Enabled = False
        Call ObtenerDatosBienAdjud
        Me.cmdAceptar.Enabled = True
        Me.cmdModificar.Visible = False
        
        'JIPR20190520 INICIO
        'Me.Caption = "REGISTRO DE VENTAS DE BIENES ADJUDICADOS"
        Me.Caption = "REGISTRO DE VENTAS DE BIENES ADJUDICADOS/EMBARGADOS"
        Me.cmdAceptar.Visible = True
        Me.cmdCancelar.Visible = True
        Me.txtValorEmb.Visible = False
        Me.txtValAdjudicacion.Visible = True
        Me.cmdBuscar.Visible = False
        
    'Else
    ElseIf nTipoProceso = 3 Then ' Detalle
     'JIPR20190520 FIN
     
        'PASI20161103 ERS08572016
        Set oPers = New COMDPersona.UCOMAcceso
            gsGrupo = oPers.CargaUsuarioGrupo(gsCodUser, gsDominio)
        Set oPers = Nothing
        'PASI END ***************
        nNumAdjudicacion = nNumAdj
        Me.fraVenta.Enabled = False
        Me.fraAdjudicacion.Enabled = False
        Me.cmdAceptar.Visible = False
        Me.cmdCancelar.Visible = False
        If oCnt.EsGrupoHabilxProDepreAdj(gsGrupo) And pbBienAdj Then cmdProDepre.Visible = IIf(oCnt.ExisteProrrogaDepreAdjudicado(nNumAdjudicacion), False, True) 'PASI20161103 ERS0572016
        If gsCodArea = "042" Then Me.cmdModificar.Enabled = True
        Call ObtenerDatosBienAdjud
        
        'JIPR20190520 INICIO
       'Me.Caption = "DETALLE DE BIENES ADJUDICADOS/VENDIDOS"
        Me.txtValorEmb.Visible = False
        Me.txtValAdjudicacion.Visible = True
        Me.Caption = "DETALLE DE BIENES ADJUDICADOS/EMBARGADOS/VENDIDOS"
        Me.cmdBuscar.Visible = False
       
    Else
        cboTpoAdj.Clear
        cboTpoAdj.AddItem "3. Embargo", 0
        
        Me.fraVenta.Enabled = False
        Me.fraAdjudicacion.Enabled = True
        Me.Caption = "REGISTRO DE BIENES EMBARGADOS"
        Me.cmdModificar.Visible = False
        Me.Frame3.Caption = "Datos de Embargo"
        Me.Label6.Caption = "Valor Embargo:"
        Me.Label9.Caption = "Fecha Embargo:"
        Me.Label5.Caption = "Nro. Placa u Otros:"
        Me.txtAF.rs = oConst.ObtieneTipoBienEmbargo
        Me.txtValAdjudicacion.Visible = False
        Me.txtValorEmb.Visible = True
        Me.cmdBuscar.Visible = True
        Set oConst = Nothing
       'JIPR20190520 FIN
        End If
    Me.Show 1
End Sub
Private Sub ObtenerDatosBienAdjud()
    Dim oCnt As COMNContabilidad.NCOMContFunciones
    Set oCnt = New COMNContabilidad.NCOMContFunciones
    Dim pRs As ADODB.Recordset
    Set pRs = oCnt.ObtenerDatosBienAdjudicado(nNumAdjudicacion)

    If Not pRs.EOF Then
        AXCodCta.CMAC = Left(pRs!cCtaCod, 3)
        AXCodCta.Age = Mid(pRs!cCtaCod, 4, 2)
        AXCodCta.Prod = Mid(pRs!cCtaCod, 6, 3)
        AXCodCta.Cuenta = Right(pRs!cCtaCod, 10)
        lblPersCod.Caption = pRs!cperscod
        lbltitular.Caption = pRs!cPersNombre
        txtPElectronica.Text = pRs!cNumPartElectronica
        txtAF.Text = pRs!cTipoBien
        txtDescripcion.Text = pRs!cdescripcion
        cboTpoAdj.ListIndex = pRs!nTipoAdjud - 1
        'txtValAdjudicacion.Caption = pRs!nValAdjudicacion JIPR20190520
        txtValAdjudicacion.Caption = Format(pRs!nValAdjudicacion, "#,###,##0.00")
        txtValComercial.Text = Format(pRs!nValComercial, "#,###,##0.00")
        txtValTasacion.Text = Format(pRs!nValTasacion, "#,###,##0.00")
        txtFecAdjudic.Text = Format(pRs!dFecAdjudicacion, "dd/MM/yyyy")
        txtFecTasacion.Text = Format(pRs!dFecTasacion, "dd/MM/yyyy")
        cboMoneda.ListIndex = pRs!nmoneda - 1
        txtCapital.Text = Format(pRs!nCapital, "#,###,##0.00")
        txtInteres.Text = Format(pRs!nIntereses, "#,###,##0.00")
        If nTipoProceso = 3 Then
            lblValVenta.Text = Format(pRs!nValVenta, "#,###,##0.00")
            txtFecVenta.Text = Format(pRs!nFecVenta, "dd/MM/yyyy")
        End If
        
        'JIPR20190520 INICIO
        Me.cmdModificar.Visible = False
        If pRs!nTipoAdjud = 3 Then
        Me.Frame3.Caption = "Datos de Embargo"
        Me.Label6.Caption = "Valor Embargo:"
        Me.Label9.Caption = "Fecha Embargo:"
        Me.txtValorEmb.Text = Format(pRs!nValAdjudicacion, "#,###,##0.00")
        Me.cmdModificar.Visible = False
        Me.Label5.Caption = "Nro. Placa u Otros:"
        End If
        'JIPR20190520 FIN
        
    End If
End Sub
Private Sub cmdAceptar_Click()
    
    Dim oCnt As COMNContabilidad.NCOMContFunciones
    Set oCnt = New COMNContabilidad.NCOMContFunciones
    Dim clsMovN As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String

    If MsgBox("¿Está seguro de haber ingresado correctamente los datos?", vbQuestion + vbYesNo, "Advertencia") = vbYes Then
        If ValidarDatos Then
            If bModificar = False Then
                If nTipoProceso = 1 Then
                    Set clsMovN = New COMNContabilidad.NCOMContFunciones
                    sMovNro = clsMovN.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
                    Call oCnt.InsertarBienAdjudicado(AXCodCta.NroCuenta, Me.txtPElectronica.Text, txtAF.Text, Trim(txtDescripcion.Text), Left(cboTpoAdj.Text, 1), CDec(txtValAdjudicacion.Caption), CDec(txtValComercial.Text), _
                        CDec(txtValTasacion.Text), txtFecAdjudic.Text, txtFecTasacion.Text, IIf(cboMoneda.Text = "SOLES", 1, 2), CDec(txtCapital.Text), CDec(txtInteres.Text), 0, 1, sMovNro)
                 'JIPR20190520 INICIO
                 ElseIf nTipoProceso = 4 Then
                    Set clsMovN = New COMNContabilidad.NCOMContFunciones
                    sMovNro = clsMovN.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
                    Call oCnt.InsertarBienAdjudicado(AXCodCta.NroCuenta, Me.txtPElectronica.Text, txtAF.Text, Trim(txtDescripcion.Text), Left(cboTpoAdj.Text, 1), CDec(txtValorEmb.Text), CDec(txtValComercial.Text), _
                        CDec(txtValTasacion.Text), txtFecAdjudic.Text, txtFecTasacion.Text, IIf(cboMoneda.Text = "SOLES", 1, 2), CDec(txtCapital.Text), CDec(txtInteres.Text), 0, 3, sMovNro)
                'JIPR20190520 FIN
                Else
                    Call oCnt.InsertarVentaBienAdjudicado(nNumAdjudicacion, Me.txtFecVenta.Text, Me.lblValVenta.Text)
                End If
                MsgBox "El registro se realizó correctamente"
            Else
                Set clsMovN = New COMNContabilidad.NCOMContFunciones
                sMovNro = clsMovN.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
                Call oCnt.ModificarBienAdjudicado(nNumAdjudicacion, AXCodCta.NroCuenta, Me.txtPElectronica.Text, txtAF.Text, Trim(txtDescripcion.Text), Left(cboTpoAdj.Text, 1), CDec(txtValAdjudicacion.Caption), CDec(txtValComercial.Text), _
                CDec(txtValTasacion.Text), txtFecAdjudic.Text, txtFecTasacion.Text, IIf(cboMoneda.Text = "SOLES", 1, 2), CDec(txtCapital.Text), CDec(txtInteres.Text), 0, 1, sMovNro)
                MsgBox "El registro se modificó correctamente"
            End If
            Unload Me
        End If
    End If
    Set clsMovN = Nothing
    Set oCnt = Nothing
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub txtAF_EmiteDatos()
    txtAF.Text = txtAF.psDescripcion
End Sub
Private Sub txtCapital_Change()
'JIPR20180520 INICIO
'txtValAdjudicacion.Caption = Format(CDbl(txtCapital.Text) + CDbl(txtInteres.Text), "#,###,##0.00")
    If nTipoProceso = 4 Then
    Else
    txtValAdjudicacion.Caption = Format(CDbl(txtCapital.Text) + CDbl(txtInteres.Text), "#,###,##0.00")
    End If
'JIPR20180520 FIN
End Sub
Private Sub txtGastos_Change()
    txtValAdjudicacion.Caption = Format(CDbl(txtCapital.Text) + CDbl(txtInteres.Text), "#,###,##0.00")
End Sub
Private Sub txtInteres_Change()
'JIPR20180520 INICIO
    'txtValAdjudicacion.Caption = Format(CDbl(txtCapital.Text) + CDbl(txtInteres.Text), "#,###,##0.00")
    If nTipoProceso = 4 Then
    Else
    txtValAdjudicacion.Caption = Format(CDbl(txtCapital.Text) + CDbl(txtInteres.Text), "#,###,##0.00")
    End If
'JIPR20180520 FIN
End Sub
Private Sub txtValAdjudicacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtValComercial.SetFocus
    End If
End Sub
Private Sub txtValComercial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtValTasacion.SetFocus
    End If
End Sub
Private Sub txtValTasacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFecAdjudic.SetFocus
    End If
End Sub
Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        Me.cboTpoAdj.SetFocus
    End If
End Sub
Private Sub txtPElectronica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtAF.SetFocus
    End If
End Sub
Private Sub txtFecVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    End If
End Sub
Private Sub txtFecAdjudic_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFecTasacion.SetFocus
    End If
End Sub
Private Sub txtFecTasacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
End Sub
Private Sub txtCapital_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtInteres.SetFocus
    End If
End Sub
Private Sub txtInteres_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    'JIPR20180520 INICIO
        'txtValComercial.SetFocus
        If nTipoProceso = 4 Then
        txtValorEmb.SetFocus
        Else
        Me.txtValComercial.SetFocus
        End If
     'JIPR20180520 FIN
    End If
End Sub

'JIPR20180520 INICIO
Private Sub txtValorEmb_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If nTipoProceso = 4 Then
        Me.txtValComercial.SetFocus
        End If
    End If
End Sub
'JIPR20180520 FIN


Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    End If
End Sub
Private Sub lblValVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecVenta.SetFocus
    End If
End Sub
Public Function ValidarDatos() As Boolean
    If nTipoProceso = 1 Or bModificar = True Then
        If txtPElectronica.Text = "" Then
            MsgBox "Debe ingresar el N° de Partida Electrónica", vbInformation, "SICMACM"
            ValidarDatos = False
            txtPElectronica.SetFocus
            Exit Function
        End If
        If txtAF.Text = "" Then
            MsgBox "Debe seleccionar el tipo de bien", vbInformation, "SICMACM"
            ValidarDatos = False
            txtAF.SetFocus
            Exit Function
        End If
        If txtDescripcion.Text = "" Then
            MsgBox "Debe ingresar la descripción del bien", vbInformation, "SICMACM"
            ValidarDatos = False
            txtDescripcion.SetFocus
            Exit Function
        End If
        If cboTpoAdj.Text = "" Then
            MsgBox "Debe seleccionar el tipo de adjudicación", vbInformation, "SICMACM"
            ValidarDatos = False
            cboTpoAdj.SetFocus
            Exit Function
        End If
        If val(txtValAdjudicacion.Caption) <= 0 Then
            MsgBox "El Valor total adjudicado debe ser mayor a cero(0). Verifique el capital y los intereses", vbInformation, "SICMACM"
            ValidarDatos = False
            txtCapital.SetFocus
            Exit Function
        End If
        If val(txtValComercial.Text) <= 0 Then
            MsgBox "El valor comercial debe ser mayor a cero(0)", vbInformation, "SICMACM"
            ValidarDatos = False
            txtValComercial.SetFocus
            Exit Function
        End If
        If val(txtValTasacion.Text) <= 0 Then
            MsgBox "El valor de tasación debe ser mayor a cero(0)", vbInformation, "SICMACM"
            ValidarDatos = False
            txtValTasacion.SetFocus
            Exit Function
        End If
        If Not IsDate(txtFecAdjudic.Text) Then
            MsgBox "La fecha de adjudicación  no tiene el formato correcto", vbInformation, "SICMACM"
            ValidarDatos = False
            txtFecAdjudic.SetFocus
            Exit Function
        End If
        If Not IsDate(txtFecTasacion.Text) Then
            MsgBox "La fecha de tasación  no tiene el formato correcto", vbInformation, "SICMACM"
            ValidarDatos = False
            txtFecTasacion.SetFocus
            Exit Function
        End If
        If cboMoneda.Text = "" Then
            MsgBox "Debe seleccionar la moneda", vbInformation, "SICMACM"
            ValidarDatos = False
            cboMoneda.SetFocus
            Exit Function
        End If
    'Else JIPR20190520
    ElseIf nTipoProceso = 2 Or nTipoProceso = 3 Then
        If lblValVenta.Text = "" Then
            MsgBox "Debe ingresar el valor de venta", vbInformation, "SICMACM"
            ValidarDatos = False
            lblValVenta.SetFocus
            Exit Function
        End If
        If val(lblValVenta.Text) <= 0 Then
            MsgBox "El valor de venta debe ser mayor a cero(0).", vbInformation, "SICMACM"
            ValidarDatos = False
            txtCapital.SetFocus
            Exit Function
        End If
        If Not IsDate(txtFecVenta.Text) Then
            MsgBox "La fecha de venta no tiene el formato correcto", vbInformation, "SICMACM"
            ValidarDatos = False
            txtFecVenta.SetFocus
            Exit Function
        End If
    'Else
    
    End If
    ValidarDatos = True
End Function
Private Sub cmdCancelar_Click()
    If bModificar = False Then
    
       'JIPR20180520 INICIO
        If Me.Caption = "REGISTRO DE VENTAS DE BIENES ADJUDICADOS/EMBARGADOS" Then
        Me.lblValVenta.Text = "__/__/____"
        Me.txtFecVenta.Text = "__/__/____"
        Else
         'JIPR20180520 FIN
    
        AXCodCta.CMAC = "109"
        AXCodCta.Age = ""
        AXCodCta.Prod = ""
        AXCodCta.Cuenta = ""
        lblPersCod.Caption = ""
        lbltitular.Caption = ""
        txtPElectronica.Text = ""
        txtAF.Text = ""
        txtDescripcion.Text = ""
        'cboTpoAdj.ListIndex = 0 JIPR20180520
        txtValorEmb.Text = "" 'JIPR20180520
        txtValAdjudicacion.Caption = ""
        txtValComercial.Text = ""
        txtValTasacion.Text = ""
        cboMoneda.ListIndex = 0
        txtCapital.Text = ""
        txtInteres.Text = ""
        lblValVenta.Text = ""
        Me.txtFecAdjudic.Text = "__/__/____"
        Me.txtFecTasacion.Text = "__/__/____"

        'JIPR20180520 INICIO
        If Me.Caption = "REGISTRO DE BIENES EMBARGADOS" Then
        'cboTpoAdj.Clear
        cboTpoAdj.AddItem "3. Embargo", 0
        Me.cboTpoAdj.Enabled = False
        ElseIf Me.Caption = "REGISTRO DE BIENES ADJUDICADOS" Then
        cboTpoAdj.Clear
        cboTpoAdj.AddItem "1. Adjudicación", 0
        cboTpoAdj.AddItem "2. Dación de Pago", 1
        Me.cboTpoAdj.Enabled = True
        End If
        End If
        'JIPR20180520 FIN
    Else
        bModificar = False
        cmdModificar.Visible = True
        cmdAceptar.Visible = False
        cmdCancelar.Visible = False
        fraAdjudicacion.Enabled = False
        fraVenta.Enabled = False
    End If
End Sub
