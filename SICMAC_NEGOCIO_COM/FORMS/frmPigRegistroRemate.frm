VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPigRegistroRemate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preparación de Remate"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   ForeColor       =   &H00000000&
   Icon            =   "frmPigRegistroRemate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAvisoRemate 
      Caption         =   "Aviso de Remate"
      Height          =   375
      Left            =   765
      TabIndex        =   24
      Top             =   4005
      Width           =   4470
   End
   Begin VB.Frame Frame4 
      Caption         =   "Registro de Remate"
      ForeColor       =   &H8000000D&
      Height          =   3735
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6495
      Begin MSDataListLib.DataCombo cboUbicacion 
         Height          =   315
         Left            =   4440
         TabIndex        =   25
         Top             =   795
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   5445
         TabIndex        =   23
         Top             =   3270
         Width           =   930
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   4455
         TabIndex        =   22
         Top             =   3285
         Width           =   900
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   345
         Left            =   5445
         TabIndex        =   21
         Top             =   3270
         Width           =   915
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   345
         Left            =   4440
         TabIndex        =   20
         Top             =   3270
         Width           =   915
      End
      Begin VB.Frame Frame2 
         Height          =   1170
         Left            =   75
         TabIndex        =   15
         Top             =   2040
         Width           =   6315
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "..."
            Height          =   315
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Buscar ..."
            Top             =   225
            Width           =   345
         End
         Begin VB.TextBox txtcodigo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   855
            TabIndex        =   17
            Top             =   255
            Width           =   1470
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   150
            TabIndex        =   16
            Top             =   705
            Width           =   5910
         End
         Begin VB.Label Label11 
            Caption         =   "Martillero"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   315
            Width           =   645
         End
      End
      Begin VB.ComboBox cboProceso 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   930
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   750
         Width           =   1995
      End
      Begin VB.TextBox txtRemate 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   5505
         TabIndex        =   2
         Top             =   300
         Width           =   840
      End
      Begin MSMask.MaskEdBox txtinicio 
         Height          =   315
         Left            =   1755
         TabIndex        =   7
         Top             =   1185
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfinal 
         Height          =   315
         Left            =   5175
         TabIndex        =   9
         Top             =   1200
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecRef 
         Height          =   315
         Left            =   1740
         TabIndex        =   11
         Top             =   1620
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtFecAviso 
         Height          =   315
         Left            =   5175
         TabIndex        =   13
         Top             =   1650
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Aviso"
         Height          =   195
         Left            =   3345
         TabIndex        =   14
         Top             =   1725
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Referencia"
         Height          =   180
         Left            =   195
         TabIndex        =   12
         Top             =   1710
         Width           =   1290
      End
      Begin VB.Label Label5 
         Caption         =   "Fecha Final"
         Height          =   180
         Left            =   3345
         TabIndex        =   10
         Top             =   1290
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha Inicio"
         Height          =   165
         Left            =   180
         TabIndex        =   8
         Top             =   1275
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "Ubicacion"
         Height          =   225
         Left            =   3360
         TabIndex        =   6
         Top             =   855
         Width           =   750
      End
      Begin VB.Label Label3 
         Caption         =   "Proceso"
         Height          =   195
         Left            =   195
         TabIndex        =   5
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Número"
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
         Height          =   180
         Left            =   4635
         TabIndex        =   3
         Top             =   405
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   5460
      TabIndex        =   0
      Top             =   4020
      Width           =   1065
   End
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   120
      Top             =   3915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar prgAvisos 
      Height          =   330
      Left            =   135
      TabIndex        =   26
      Top             =   4470
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmPigRegistroRemate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim lnMant As Integer
'Option Explicit
'
'Private Sub cboProceso_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If cboProceso.Text <> "" Then
'        cboUbicacion.SetFocus
'    End If
'End If
'End Sub
'
'Private Sub cboUbicacion_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If cboProceso.Text <> "" Then
'        txtinicio.SetFocus
'    End If
'End If
'End Sub
'
'Private Sub cmdAvisoRemate_Click()
'Dim oDatos As DPigFunciones
'Dim oRemate As DPigRemate
'Dim oImpAviso As NPigRemate
'Dim oPrevio As previo.clsPrevio
'Dim rs As Recordset
'Dim lnRemate As Integer, ldFecIni As Date, ldFecFin As Date, ldFecRef As Date
'Dim lnComSer As Currency, lnDerRem As Currency, lnComVen As Currency
'Dim lnPagoMin As Currency, lnDiasVenc As Integer
'Dim lsCadImp As String
'
'    Set oDatos = New DPigFunciones
'
'    lnDiasVenc = oDatos.GetParamValor(gPigParamDiasAtrasoCartaVenc)
'    lnPagoMin = oDatos.GetParamValor(gPigParamAmortizaMin)
'
'    Set rs = oDatos.GetConceptoValor(gColPigConceptoCodComiServ)
'    lnComSer = rs!nValor
'    Set rs = Nothing
'
'    Set rs = oDatos.GetConceptoValor(gColPigConceptoCodComiVencida)
'    lnComVen = rs!nValor
'    Set rs = Nothing
'
'    Set rs = oDatos.GetConceptoValor(gColPigConceptoCodPreparaRemate)
'    lnDerRem = rs!nValor
'    Set rs = Nothing
'    Set oDatos = Nothing
'
'    Set oRemate = New DPigRemate
'    Set rs = oRemate.GetNumRemate
'
'    If Not (rs.EOF And rs.BOF) Then
'        lnRemate = rs!NumRemate
'        ldFecIni = rs!dInicio
'        ldFecFin = rs!dFin
'        ldFecRef = rs!dReferencia
'        txtFecRef = Format(ldFecRef, "dd/mm/yyyy")
'        txtRemate = lnRemate
'        txtinicio = Format(ldFecIni, "dd/mm/yyyy")
'        txtfinal = Format(ldFecFin, "dd/mm/yyyy")
'    End If
'
'    Set rs = Nothing
'
'        lsCadImp = ImpreAvisoRemate(lnRemate, ldFecIni, ldFecFin, lnComSer, lnComVen, '                lnDerRem, lnPagoMin, lnDiasVenc, ldFecRef)
'
'        dlgGrabar.CancelError = True
'        dlgGrabar.InitDir = App.path & "\spooler"
'        dlgGrabar.Filter = "Archivos de Texto (*.TXT)|*.TXT"
'        On Error GoTo dError
'        dlgGrabar.ShowSave
'
'        If dlgGrabar.FileName <> "" Then
'           Open dlgGrabar.FileName For Output As #1
'            Print #1, lsCadImp
'            Close #1
'        End If
'
'    Set oRemate = Nothing
'    Exit Sub
'
'dError:
'    MsgBox "Grabación cancelada", vbInformation, "Aviso"
'End Sub
'Private Function ImpreAvisoRemate(ByVal pnNumRem As Long, pdFecIni As Date, pdFecFin As Date, ByVal pnComSer As Currency, '                    ByVal pnComVen As Currency, pnDerRem As Currency, ByVal pnPagoMin As Currency, '                    ByVal pnDiasVenAviso As Integer, ByVal pdFecRef As Date) As String
'
'
'Dim oCalc As NPigCalculos
'Dim oRemate As DPigRemate
'Dim rs As Recordset
'Dim rsCont As Recordset
'Dim lnTIComp As Double, lnTIMor As Double, ldFecVen As Date, lnDiasVenc As Integer
'Dim lnSaldo As Currency, ldFecUltPago As Date, lnAmortMin As Currency, lnPagoMin As Currency
'Dim lnComVen As Currency, lnComSer As Currency, lnDerRem As Currency
'Dim lnIntComp As Currency, lnIntMor As Currency, lnDiasTrans As Integer
'Dim lsCadImp As String, lnLineas As Integer
'Dim ldFecVenc As Date
'Dim i As Integer
'Dim lnPag As Long
'Dim lsCadImpTemp As String
'
'    lsCadImp = ""
'    lsCadImpTemp = ""
'    lnPag = 0
'
'    Set oRemate = New DPigRemate
'    Set rs = oRemate.GetClientesAvisoRemate(pnDiasVenAviso, pdFecIni)
'
'    If Not (rs.EOF And rs.BOF) Then
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(108) & Chr$(0)  'Tipo letra : 0,1,2 - Roman,SansS,Courier
'        lsCadImp = lsCadImp & Chr$(27) & Chr$(77)             'Tamaño  : 80, 77, 103
'
'        Me.Caption = "Avisos:" & rs.RecordCount & " - En Proc:   000000"
'        prgAvisos.Max = rs.RecordCount
'
'        Do While Not rs.EOF
'
'            lnLineas = 0
'            lsCadImp = lsCadImp & "." & Space(83) & Space(33) & pnNumRem & Chr(10) & Chr(10)
'            lsCadImp = lsCadImp & Space(84) & Space(33) & Format(pdFecRef, "dd/mm/yyyy") & Chr(10)
'            lsCadImp = lsCadImp & Space(84) & "Sr. (a)(ita)" & Chr(10)
'            lsCadImp = lsCadImp & Space(84) & PstaNombre(rs!cPersNombre) & Chr(10)
'            lsCadImp = lsCadImp & Space(84) & rs!cPersDireccDomicilio & "  -  " & rs!CodPostal & Chr(10)
'            lsCadImp = lsCadImp & Space(84) & rs!cUbiGeoDescripcion & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
'
'            lnLineas = lnLineas + 13
'
'            Set rsCont = oRemate.GetContratosAvisoRemate(rs!cPersCod, pdFecFin)
'            If Not (rsCont.EOF And rsCont.BOF) Then
'                Do While Not rsCont.EOF
'
'                    ldFecVenc = rsCont!dvenc
'                    lnIntComp = rsCont!Interes
'                    lnIntMor = rsCont!Mora
'
'                    lnAmortMin = pnPagoMin + lnIntComp + lnIntMor + pnComVen + pnComSer + pnDerRem
'                    Set oCalc = Nothing
'
'                    lsCadImp = lsCadImp & Space(84) & rsCont!cCtaCod & Space(8) & ImpreFormat(lnAmortMin, 11, 2, True)
'                    lsCadImp = lsCadImp & Space(10) & Format(ldFecVenc, "dd/mm/yyyy") & Chr(10)
'
'                    lnLineas = lnLineas + 1
'
'                    rsCont.MoveNext
'                Loop
'            End If
'
'            For i = lnLineas To 30
'                lsCadImp = lsCadImp & Chr(10)
'            Next i
'
'            Me.Caption = "Avisos: " & rs.RecordCount & " - En Proc: " & lnPag
'            prgAvisos.value = lnPag
'
'            lnPag = lnPag + 1
'            If lnPag Mod 2 = 0 Then
'                lsCadImp = lsCadImp & Chr(12)
'            End If
'
'            If lnPag Mod 30 = 30 Then
'                lsCadImpTemp = lsCadImpTemp & lsCadImp
'                lsCadImp = ""
'            End If
'            rs.MoveNext
'        Loop
'        lsCadImpTemp = lsCadImpTemp & lsCadImp
'        lsCadImp = ""
'
'        Me.Caption = "Avisos: " & rs.RecordCount & " - En Proc: " & lnPag
'        Me.Caption = "Registro de Remate"
'
'    End If
'
'    Set oRemate = Nothing
'
'    ImpreAvisoRemate = lsCadImpTemp
'
'End Function
'
'Private Sub cmdBuscar_Click()
'Dim loPers As UPersona
'Dim lsPersCod As String, lsPersNombre As String
'Dim lsEstados As String
'Dim loPersContrato As DColPContrato
'Dim loPersCredito As DPigContrato
'Dim lrContratos As ADODB.Recordset
'Dim loCuentas As UProdPersona
'Dim i As Integer
'Dim liEvalCli As Integer
'On Error GoTo ControlError
'
'Set loPers = New UPersona
'    Set loPers = frmBuscaPersona.Inicio
'    lsPersCod = loPers.sPersCod
'    lsPersNombre = loPers.sPersNombre
'Set loPers = Nothing
'
'txtcodigo.Text = lsPersCod
'txtnombre.Text = lsPersNombre
'
'
'Exit Sub
'ControlError:
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'
'End Sub
'
'Private Sub cmdCancelar_Click()
'
'    Limpia
'    cboProceso.Enabled = False
'
'End Sub
'
'Private Sub CmdEditar_Click()
'
'    HabilitaControles False, False, True, True
'    Frame2.Enabled = True
'    txtRemate.SetFocus
'    lnMant = 2
'
'End Sub
'
'Private Function Limpia()
'    txtRemate.Text = ""
'    cboProceso.ListIndex = -1
'    cboUbicacion.BoundText = 0
'    txtinicio.Text = "__/__/____"
'    txtfinal.Text = "__/__/____"
'    TxtFecAviso.Text = "__/__/____"
'    txtFecRef.Text = "__/__/____"
'    txtcodigo.Text = ""
'    txtnombre.Text = ""
'
'    HabilitaControles True, True, False, False
'    Frame2.Enabled = False
'
'End Function
'
'Private Sub cmdGrabar_Click()
'Dim oGraba As DPigActualizaBD
'
'    If lnMant = 1 Then          'NUEVO
'        If ValidaGrabar Then
'            Set oGraba = New DPigActualizaBD
'            oGraba.dInsertRemate txtRemate.Text, 0, cboUbicacion.BoundText, CDate(txtinicio.Text), CDate(txtfinal.Text), '                                CDate(txtFecRef.Text), txtcodigo.Text, CDate(TxtFecAviso.Text)
'            Set oGraba = Nothing
'        End If
'    ElseIf lnMant = 2 Then      'EDITAR
'        If ValidaGrabar Then
'            Set oGraba = New DPigActualizaBD
'            oGraba.dUpdateColocPigRemate txtRemate.Text, cboUbicacion.BoundText, txtinicio.Text, '                                txtfinal.Text, txtFecRef.Text, TxtFecAviso.Text, txtcodigo.Text
'            Set oGraba = Nothing
'        End If
'    End If
'
'    Limpia
'
'End Sub
'
'Private Sub CmdNuevo_Click()
'Dim loremate As DPigContrato
'Dim lrMaxRemate As Long
'
'    HabilitaControles False, False, True, True
'    Frame2.Enabled = True
'    cboUbicacion.SetFocus
'    lnMant = 1
'    Set loremate = New DPigContrato
'    lrMaxRemate = loremate.dObtieneMaxRemate()
'    Set loremate = Nothing
'    txtRemate.Text = lrMaxRemate
'    cboProceso.ListIndex = 0
'    cboProceso.Enabled = True
'End Sub
'
'Private Sub cmdsalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'Dim loPersCredito As DPigContrato
'Dim rs As Recordset
'Dim oCons As DConstante
'Dim oRemate As DPigRemate
'
'    HabilitaControles True, True, False, False
'    Frame2.Enabled = False
'
'    Set oCons = New DConstante
'    Set rs = oCons.RecuperaConstantes(gColocPigTipoProceso)
'    Do While Not rs.EOF
'        cboProceso.AddItem (rs!cConsDescripcion & Space(45) & rs!nConsValor)
'        rs.MoveNext
'    Loop
'    Set rs = Nothing
'
'    CargaCombo cboUbicacion, gColocPigUbicacion
'    Set rs = Nothing
'    Set oCons = Nothing
'
'    Set oRemate = New DPigRemate
'    Set rs = oRemate.GetNumRemate
'
'    If Not (rs.EOF And rs.BOF) Then
'        txtRemate = rs!NumRemate
'        TxtFecAviso.Text = Format(rs!dAviso, "dd/mm/yyyy")
'        txtFecRef.Text = Format(rs!dReferencia, "dd/mm/yyyy")
'        cboProceso.ListIndex = rs!nTipoProceso
'        cboUbicacion.BoundText = rs!cUbicacion
'        txtinicio.Text = Format(rs!dInicio, "dd/mm/yyyy")
'        txtfinal.Text = Format(rs!dFin, "dd/mm/yyyy")
'        txtcodigo.Text = rs!cPersCod
'        txtnombre.Text = PstaNombre(IIf(IsNull(rs!cPersNombre), "", rs!cPersNombre))
'    End If
'
'    Set rs = Nothing
'    Set oRemate = Nothing
'
'    Me.Icon = LoadPicture(App.path & "\bmps\cm.ico")
'
'End Sub
'
'Private Sub txtfecaviso_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If IsDate(TxtFecAviso.Text) Then
'        cmdBuscar.SetFocus
'    Else
'        MsgBox "Fecha no válida", vbInformation, "Aviso"
'        TxtFecAviso.SetFocus
'    End If
'End If
'End Sub
'
'Private Sub txtFecRef_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If IsDate(txtinicio.Text) Then
'        TxtFecAviso.SetFocus
'    Else
'        MsgBox "Fecha no válida", vbInformation, "Aviso"
'        txtFecRef.SetFocus
'    End If
'End If
'End Sub
'
'Private Sub txtfinal_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If txtinicio <> "" Then
'        If IsDate(txtinicio.Text) Then
'            If IsDate(txtfinal) Then
'                If (CDate(txtinicio.Text) > CDate(txtfinal)) Then
'                    MsgBox "Fecha Fin de Remate no puede ser menor que Fecha Inicio", vbInformation, "Aviso"
'                Else
'                    Me.txtFecRef.SetFocus
'                End If
'            End If
'        Else
'            MsgBox "Fecha no válida", vbInformation, "Aviso"
'            txtfinal.SetFocus
'        End If
'    End If
'End If
'End Sub
'
'Private Sub txtInicio_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If IsDate(txtinicio.Text) Then
'        txtfinal.SetFocus
'    Else
'        MsgBox "Fecha no válida", vbInformation, "Aviso"
'        txtinicio.SetFocus
'    End If
'End If
'End Sub
'
'Private Sub txtRemate_KeyPress(KeyAscii As Integer)
'  Dim lrRemate As DPigContrato
'  Dim lrDatosRemate As ADODB.Recordset
'  Dim liproceso As Integer
'  Dim liRemate As Integer
'  If KeyAscii = 13 Then 'se presione enter y el boton nuevo en falso
'     If lnMant = 2 Then
'           Set lrRemate = New DPigContrato
'           Set lrDatosRemate = lrRemate.dObtieneDatosRemate(txtRemate.Text)
'           Set lrRemate = Nothing
'            If Not (lrDatosRemate.EOF) Then
'                cboProceso.ListIndex = lrDatosRemate!nTipoProceso
'                cboUbicacion.BoundText = lrDatosRemate!cUbicacion
'                txtinicio.Text = lrDatosRemate!dInicio
'                txtfinal.Text = lrDatosRemate!dFin
'                txtFecRef.Text = lrDatosRemate!dReferencia
'                TxtFecAviso = lrDatosRemate!dAviso
'                txtcodigo.Text = lrDatosRemate!cPersCod
'                txtnombre.Text = PstaNombre(IIf(IsNull(lrDatosRemate!cPersNombre), "", lrDatosRemate!cPersNombre))
'                If cmdBuscar.Enabled Then cmdBuscar.SetFocus
'            Else
'                MsgBox "Remate no Registrado", vbInformation, "Aviso"
'            End If
'        Else
'            If cboProceso.Enabled Then cboProceso.SetFocus
'       End If
'   End If
'End Sub
'
'Private Sub HabilitaControles(ByVal pbNuevo As Boolean, ByVal pbEdita As Boolean, ByVal pbGraba As Boolean, ByVal pbCancela As Boolean)
'
'    cmdNuevo.Visible = pbNuevo
'    cmdEditar.Visible = pbEdita
'    cmdGrabar.Visible = pbGraba
'    cmdCancelar.Visible = pbCancela
'    cmdNuevo.Enabled = pbNuevo
'    cmdEditar.Enabled = pbEdita
'    cmdGrabar.Enabled = pbGraba
'    cmdCancelar.Enabled = pbCancela
'
'End Sub
'
'Private Function ValidaGrabar() As Boolean
'
'ValidaGrabar = True
'
'If cboUbicacion = "" Then
'    MsgBox "Seleccione la ubicacion donde se efectuara el Remate ", vbInformation, "Aviso"
'    ValidaGrabar = False
'    cboUbicacion.SetFocus
'    Exit Function
'End If
'
'If txtinicio = "" Then
'    MsgBox "Fecha no válida", vbInformation, "Aviso"
'    ValidaGrabar = False
'    txtinicio.SetFocus
'    Exit Function
'End If
'
'If txtfinal = "" Then
'    MsgBox "Fecha no válida", vbInformation, "Aviso"
'    ValidaGrabar = False
'    txtfinal.SetFocus
'    Exit Function
'End If
'
'If txtFecRef = "" Then
'    MsgBox "Fecha no válida", vbInformation, "Aviso"
'    ValidaGrabar = False
'    txtFecRef.SetFocus
'    Exit Function
'End If
'
'If TxtFecAviso = "" Then
'    MsgBox "Fecha no válida", vbInformation, "Aviso"
'    ValidaGrabar = False
'    TxtFecAviso.SetFocus
'    Exit Function
'End If
'
'If IsDate(txtfinal) And IsDate(txtinicio) Then
'    If CDate(txtinicio) > CDate(txtfinal) Then
'        MsgBox "Fecha Fin de Remate no puede ser menor que Fecha de Inicio", vbInformation, "Aviso"
'        ValidaGrabar = False
'        Exit Function
'    End If
'Else
'    MsgBox "Fecha no válida", vbInformation, "Aviso"
'    ValidaGrabar = False
'    Exit Function
'End If
'If CDate(TxtFecAviso) > CDate(txtinicio) Then
'    MsgBox "Fecha de Aviso de Remate no puede ser mayor que Fecha de Inicio", vbInformation, "Aviso"
'    ValidaGrabar = False
'    TxtFecAviso.SetFocus
'    Exit Function
'End If
'
'End Function
'
'Private Sub CargaCombo(Combo As DataCombo, ByVal psConsCod As String)
'Dim oPigFunc As DPigFunciones
'Dim rs As Recordset
'
'Set oPigFunc = New DPigFunciones
'
'    Set rs = oPigFunc.GetConstante(psConsCod)
'    Set Combo.RowSource = rs
'    Combo.ListField = "cConsDescripcion"
'    Combo.BoundColumn = "nConsValor"
'
'    Set rs = Nothing
'
'    Set oPigFunc = Nothing
'
'End Sub
