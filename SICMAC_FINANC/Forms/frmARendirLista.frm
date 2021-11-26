VERSION 5.00
Begin VB.Form frmARendirLista 
   Caption         =   "A rendir Cuenta: Pendientes de regularizar"
   ClientHeight    =   5835
   ClientLeft      =   885
   ClientTop       =   2145
   ClientWidth     =   11850
   Icon            =   "frmARendirLista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   11850
   Begin VB.CommandButton cmdRendicion 
      Caption         =   "&Rendición"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   15
      Top             =   4170
      Width           =   1380
   End
   Begin VB.CommandButton cmdDsctoPlanilla 
      Caption         =   "Rendición X &Dscto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   14
      ToolTipText     =   "Regularizar con Documentos sustentatorios"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton cmdBuscar 
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
      Left            =   8280
      TabIndex        =   13
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdProrroga 
      Cancel          =   -1  'True
      Caption         =   "&Prórroga"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   12
      Top             =   4200
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.CheckBox chkTodos 
      Caption         =   "Incluir Arendir Cuenta Sustentados"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   9000
      TabIndex        =   11
      Top             =   4560
      Width           =   2805
   End
   Begin VB.CheckBox chkSelec 
      Caption         =   "&Todos"
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
      Height          =   210
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   900
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10290
      TabIndex        =   7
      Top             =   4170
      Width           =   1380
   End
   Begin VB.CommandButton cmdRegulariza 
      Caption         =   "S&ustentación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8925
      TabIndex        =   6
      ToolTipText     =   "Regularizar con Documentos sustentatorios"
      Top             =   4170
      Width           =   1380
   End
   Begin VB.Frame FraSeleccion 
      Enabled         =   0   'False
      Height          =   945
      Left            =   90
      TabIndex        =   9
      Top             =   120
      Width           =   8085
      Begin Sicmact.TxtBuscar txtBuscarAgenciaArea 
         Height          =   330
         Left            =   1425
         TabIndex        =   1
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         lbUltimaInstancia=   0   'False
      End
      Begin VB.Label lblAgeDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1410
         TabIndex        =   3
         Top             =   525
         Width           =   6420
      End
      Begin VB.Label lblAgenciaArea 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2520
         TabIndex        =   2
         Top             =   195
         Width           =   5310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Area/Agencia :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   135
         TabIndex        =   10
         Top             =   225
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Glosa"
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
      Height          =   1155
      Left            =   90
      TabIndex        =   8
      Top             =   4575
      Width           =   11640
      Begin VB.TextBox txtMovDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   135
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   11370
      End
   End
   Begin Sicmact.Usuario usu 
      Left            =   885
      Top             =   5760
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.FlexEdit fgAtenciones 
      Height          =   2985
      Left            =   90
      TabIndex        =   4
      Top             =   1125
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   5265
      Cols0           =   25
      HighLight       =   2
      AllowUserResizing=   3
      EncabezadosNombres=   $"frmARendirLista.frx":030A
      EncabezadosAnchos=   "350-450-1000-900-1100-900-2500-1000-0-0-0-0-1000-2000-0-0-0-2000-0-0-0-0-0-0-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C-L-C-L-R-C-L-L-L-R-L-L-C-C-C-C-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-1-2-0-0-0-0-2-0-0-0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "N°"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbPuntero       =   -1  'True
      lbOrdenaCol     =   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
   End
End
Attribute VB_Name = "frmARendirLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cOpeCod As String, cOpeDesc As String
Dim lTransActiva As Boolean
Dim lRindeCajaG As Boolean, lEsChica As Boolean
Dim lRindeCajaCh As Boolean, lRindeViaticos As Boolean
Dim sObjRendir As String
Dim sDocTpoRecibo As String
Dim sCtaPendiente As String
'************************************************************************************
'************************************************************************************
Dim oContFunc As NContFunciones
Dim oAreas As DActualizaDatosArea
Dim oNArendir As NARendir
Dim oOperacion As DOperacion

Dim lnTipoArendir As ArendirTipo
Dim lbEsChica  As Boolean
Dim lsCtaArendir As String
Dim lsCtaPendiente As String
Dim lsDocTpoRecibo As String

Dim lsTpoDocVoucher  As String
Dim lSalir As Boolean
Dim lsMovNroSolicitud As String
Dim lnArendirFase As ARendirFases

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(ByVal pnTipoArendir As ArendirTipo, ByVal pnArendirFase As ARendirFases, Optional pbEsCajaChica As Boolean = False)
lnArendirFase = pnArendirFase
lnTipoArendir = pnTipoArendir
lbEsChica = pbEsCajaChica
Me.Show 1
End Sub

Private Function GetReciboEgreso() As Boolean
Dim lnFila As Long
Dim rs As ADODB.Recordset
GetReciboEgreso = False
lSalir = False
Set rs = New ADODB.Recordset
fgAtenciones.Clear
fgAtenciones.FormaCabecera
fgAtenciones.Rows = 2
'***Comentado por ELRO el 20120425, según OYP-RFC005-2012
'If TxtBuscarArendir = "" Then
'    MsgBox "Seleccione el Area/Agencia a quien solicitó el Arendir", vbInformation, "Aviso"
'    TxtBuscarArendir.SetFocus
'    Exit Function
'End If
'***Fin Comentado por ELRO*******************************
If chkSelec.value = 0 Then
    If txtBuscarAgenciaArea = "" Then
        MsgBox "Ingrese el Area a la cual Pertenece el Arendir", vbInformation, "Aviso"
        txtBuscarAgenciaArea.SetFocus
        Exit Function
    End If
End If
Me.MousePointer = 11
'***Modificado por ELRO el 20120425, según OYP-RFC005-2012
'Set rs = oNArendir.GetAtencionPendArendir(chkSelec.value, Mid(txtBuscarAgenciaArea.Text, 4, 2), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir, Mid(gsOpeCod, 3, 1), Mid(TxtBuscarArendir, 1, 3), Mid(TxtBuscarArendir, 4, 2), chkTodos.value = vbChecked)
If gsOpeCod = gCGArendirViatRend2MN Or gsOpeCod = gCGArendirViatRend2ME Then
    Set rs = oNArendir.obtenerARendirViaticosParaRendir(Mid(gsOpeCod, 3, 1), Mid(txtBuscarAgenciaArea.Text, 1, 3), chkSelec.value)

ElseIf gsOpeCod = gCGArendirCtaRend2MN Or gsOpeCod = gCGArendirCtaRend2ME Then
    Set rs = oNArendir.obtenerARendirCuentasParaRendir(Mid(gsOpeCod, 3, 1), Mid(txtBuscarAgenciaArea.Text, 1, 3), chkSelec.value)
Else
    Set rs = oNArendir.GetAtencionPendArendir(chkSelec.value, Mid(txtBuscarAgenciaArea.Text, 4, 2), Mid(txtBuscarAgenciaArea.Text, 1, 3), lnTipoArendir, lsCtaArendir, Mid(gsOpeCod, 3, 1), "025", "", chkTodos.value = vbChecked)
End If
'***Fin Modificado por ELRO*******************************
If Not rs.EOF And Not rs.BOF Then
   Set fgAtenciones.Recordset = rs
   fgAtenciones.FormatoPersNom 6
Else
   If lnTipoArendir = gArendirTipoCajaChica Then
      MsgBox "Caja Chica sin egresos pendientes de A rendir", vbInformation, "Aviso"
   Else
      MsgBox "Area funcional sin A rendir Cuenta Pendientes", vbInformation, "Aviso"
   End If
End If
rs.Close: Set rs = Nothing
GetReciboEgreso = True
Me.MousePointer = 0
End Function

Private Sub chkSelec_Click()
If chkSelec.value = 0 Then
    FraSeleccion.Enabled = True
    txtBuscarAgenciaArea.SetFocus
    LimpiaFlex fgAtenciones
    txtMovDesc = ""
Else
    FraSeleccion.Enabled = False
    txtBuscarAgenciaArea.Text = ""
    lblAgenciaArea = ""
    lblAgeDesc = ""
    LimpiaFlex fgAtenciones
    txtMovDesc = ""
  End If
End Sub

Private Sub chkSelec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdBuscar.SetFocus
End If
End Sub

Private Sub cmdBuscar_Click()
If GetReciboEgreso Then
   fgAtenciones.SetFocus
Else
   lblAgenciaArea = ""
   lblAgenciaArea = ""
End If
End Sub

Private Sub cmdDsctoPlanilla_Click()

If fgAtenciones.TextMatrix(fgAtenciones.row, 9) = gsCodPersUser Then
    MsgBox "Usted no puede realizar su propia Rendición por Dscto. Planilla", vbInformation, "Aviso"
    cmdDsctoPlanilla.SetFocus
    Exit Sub
End If

If MsgBox("¿Esta seguro que deseas registrar la Rendición por Dscto por Planilla " & fgAtenciones.TextMatrix(fgAtenciones.row, 1) & " - " & fgAtenciones.TextMatrix(fgAtenciones.row, 4) & "?", vbYesNo, "Aviso") = vbYes Then
    Set oNArendir = New NARendir
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim oDOperacion As DOperacion
    Set oDOperacion = New DOperacion
    Dim lsMovNro, lsCtaContArendir, lsCtaContPendiente As String
    Dim lnMovNroConfirmacion As Long
    Dim lsOpeCod As String
    
    
    If gsOpeCod = CStr(gCGArendirViatRend2MN) Or gsOpeCod = CStr(gCGArendirViatRend2ME) Then
        lsOpeCod = IIf(Mid(gsOpeCod, 3, 1) = "1", gCGArendirViatRendDsctPlanMN, gCGArendirViatRendDsctPlanME)
        lsCtaContArendir = oDOperacion.EmiteOpeCta(lsOpeCod, "H")
        If lsCtaContArendir = "" Then
            MsgBox "Faltan asignar Cuentas Contables a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Contabilidad", vbInformation, "Aviso"
            lSalir = True
            Exit Sub
        End If
        
        lsCtaContPendiente = oDOperacion.EmiteOpeCta(lsOpeCod, "D")
        If lsCtaContArendir = "" Then
            MsgBox "Faltan asignar Cuentas Contables a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Contabilidad", vbInformation, "Aviso"
            lSalir = True
            Exit Sub
        End If
        
    ElseIf gsOpeCod = CStr(gCGArendirCtaRend2MN) Or gsOpeCod = CStr(gCGArendirCtaRend2ME) Then
        lsOpeCod = IIf(Mid(gsOpeCod, 3, 1) = "1", gCGArendirCtaRendDsctPlanMN, gCGArendirCtaRendDsctPlanME)
        lsCtaContArendir = oDOperacion.EmiteOpeCta(lsOpeCod, "H")
        If lsCtaContArendir = "" Then
            MsgBox "Faltan asignar Cuentas Contables a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Contabilidad", vbInformation, "Aviso"
            lSalir = True
            Exit Sub
        End If
        
        lsCtaContPendiente = oDOperacion.EmiteOpeCta(lsOpeCod, "D")
        If lsCtaContArendir = "" Then
            MsgBox "Faltan asignar Cuentas Contables a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Contabilidad", vbInformation, "Aviso"
            lSalir = True
            Exit Sub
        End If
    End If
    
    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    If gsOpeCod = CStr(gCGArendirViatRend2MN) Or gsOpeCod = CStr(gCGArendirViatRend2ME) Then
        Call oNArendir.registrarRendicionDsctoPlanillaViaticos(lsMovNro, _
                                                               lsOpeCod, _
                                                               txtMovDesc, _
                                                               CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), _
                                                               CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 7)), _
                                                               lsCtaContArendir, _
                                                               lsCtaContPendiente, _
                                                               lnTipoArendir, _
                                                               lnMovNroConfirmacion)
     ElseIf gsOpeCod = CStr(gCGArendirCtaRend2MN) Or gsOpeCod = CStr(gCGArendirCtaRend2ME) Then
        Call oNArendir.registrarRendicionDsctoPlanillaARendirCuenta(lsMovNro, _
                                                                    lsOpeCod, _
                                                                    txtMovDesc, _
                                                                    CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), _
                                                                    CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 10)), _
                                                                    CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 7)), _
                                                                    lsCtaContArendir, _
                                                                    lsCtaContPendiente, _
                                                                    lnTipoArendir, _
                                                                    lnMovNroConfirmacion)
     End If
     
    If lnMovNroConfirmacion > 0 Then
        MsgBox "Se registro correcta la Rendición por Descuento por Planilla"
        ImprimeAsientoContable lsMovNro, , , , , , txtMovDesc, fgAtenciones.TextMatrix(fgAtenciones.row, 9), fgAtenciones.TextMatrix(fgAtenciones.row, 12), 2, fgAtenciones.TextMatrix(fgAtenciones.row, 1) & " - " & fgAtenciones.TextMatrix(fgAtenciones.row, 4), , , "RENDICIÓN POR DESCUENTO POR PLANILLA", "17"
        fgAtenciones.EliminaFila fgAtenciones.row
    End If
    lsOpeCod = ""
    lnMovNroConfirmacion = 0
    lsMovNro = ""
    lsCtaContArendir = ""
    lsCtaContPendiente = ""
    Set oDOperacion = Nothing
    Set oNContFunciones = Nothing
End If
End Sub

Private Sub cmdProrroga_Click()
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Set oNArendir = New NARendir
    Dim nProrroga As Integer
    Dim lsMovNro As String
    
    'TORE - 18062018-Dias prorrogas
    Dim vProrroga As Integer
    Dim fechaLLegada As Date
    Dim fechaSis As Date
    Dim diasAtraso As Integer
    Dim rsDias As ADODB.Recordset
    Dim respuesta As Integer
    Dim cantidadDias As Integer
    Dim nDiasParametro As Integer
    Dim dLimiteSustentar As Date
    Dim cCargoAut As String
    Dim cCargoCod As String
    Dim cCargoDes As String
    Dim nCantidadProrroga As Integer
    Dim Glosa As String
    Dim cGlosa As String
    Dim i As Integer
    Dim nProrrogasAut As Integer
    'END TORE

    If fgAtenciones.TextMatrix(1, 0) = "" Then
        MsgBox "No existen Atenciones Pendientes", vbInformation, "Aviso"
        Exit Sub
    End If

    If fgAtenciones.TextMatrix(fgAtenciones.row, 9) = gsCodPersUser Then
        MsgBox "Usted no puede registrar su propia prorroga.", vbInformation, "Aviso"
        cmdProrroga.SetFocus
        Exit Sub
    End If

    lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    fechaLLegada = oNArendir.obtenerFechaLlegadaColaborador(CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), CInt(fgAtenciones.TextMatrix(fgAtenciones.row, 24)))
    fechaSis = gdFecSis

    Set rsDias = New ADODB.Recordset
    Set rsDias = oNArendir.obtenerDiasTrancurridos(Format(fechaLLegada, "yyyy/MM/dd"), fgAtenciones.TextMatrix(fgAtenciones.row, 18), CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), Format(fechaSis, "yyyy/MM/dd"), gsCodUser, lnTipoArendir)
    
    If Not rsDias.BOF And Not rsDias.EOF Then
        For i = 1 To rsDias.RecordCount
            dLimiteSustentar = rsDias!dPlazoRendir
            diasAtraso = rsDias!nDiasAtraso
            nDiasParametro = rsDias!nDiasParam
            cCargoAut = rsDias!cCargoCodAut
            cCargoCod = rsDias!cCargoCod
            cCargoDes = rsDias!cCargoDes
            nCantidadProrroga = rsDias!nCantidadProrroga
            nProrrogasAut = rsDias!ProrrogasAut
            rsDias.MoveNext
        Next
    Else
        MsgBox "Inconvenientes en el cálculo de los días a sustentar.", vbCritical, "Comprobar"
        Exit Sub
    End If
    
    If diasAtraso <= 0 Then
        If dLimiteSustentar = gdFecSis Then
            MsgBox fgAtenciones.TextMatrix(fgAtenciones.row, 6) & " tiene plazo hasta hoy para sustentar.", vbInformation, "Aviso"
            Exit Sub
        End If
            MsgBox fgAtenciones.TextMatrix(fgAtenciones.row, 6) & " tiene plazo hasta " & dLimiteSustentar & " para sustentar.", vbInformation, "Aviso"
        Exit Sub
    End If

    'TORE - Automatización de Prórrogas
    If (diasAtraso >= nDiasParametro) Or (nProrrogasAut > 0) Then
        'condicional de para el cargo
        If cCargoAut <> cCargoCod Then
            MsgBox "Sólo la jefatura del Área de Contabilidad puede otorgar prórroga con autorización del Gerente de Administración.", vbInformation, "Aviso"
            Exit Sub
        End If
        
        vProrroga = verificarProrrogaSustentar(fgAtenciones.TextMatrix(fgAtenciones.row, 23), CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), CInt(fgAtenciones.TextMatrix(fgAtenciones.row, 24)))
        If vProrroga = 2 Then
            If nProrrogasAut = 0 Then
                respuesta = MsgBox("La fecha de sustentación de " & fgAtenciones.TextMatrix(fgAtenciones.row, 6) & " venció el día de hoy. Al asignar la prórroga, ud. esta asumiendo que visualizó la autorización del Gerente de Administración." & _
                vbNewLine & "¿Está seguro de otorgar la prórroga?", vbInformation + vbYesNo, "Confirmar Autorización")
            Else
                respuesta = MsgBox("El plazo de sustentación de " & fgAtenciones.TextMatrix(fgAtenciones.row, 6) & IIf(diasAtraso > 1, " superó los " & diasAtraso & " días", " superó en " & diasAtraso & " día") & ". Al asignar la prórroga, ud. esta asumiendo que visualizó la autorización del Gerente de Administración." & _
                vbNewLine & "¿Está seguro de otorgar la prórroga?", vbInformation + vbYesNo, "Confirmar Autorización")
            End If
            Select Case respuesta
                Case 6 'SI
                    cGlosa = InputBox("Motivo de la Ampliación", "Motivo")
                    If Not ValidaCadena(cGlosa) Then
                        MsgBox "No se puede permitir caracteres especiales en la glosa.", vbInformation, "Aviso"
                        Exit Sub
                    End If
                    If cGlosa <> "" Then
                        'Se otorgo la prórroga con autorización del Gerente de Administración(4)
                        Call oNArendir.ActualizarProrrogaSustentar(DateDiff("d", fechaLLegada, fechaSis) + 2, CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, 3)
                        Call oNArendir.RegProrrogaSustentarAut(CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, cGlosa)
                        MsgBox "Se registro correctamente la prórroga", vbInformation, "Aviso"
                        Set objPista = New COMManejador.Pista
                        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se otorgó la prórroga con autorización del Gerente de Administración al  N°: " & fgAtenciones.TextMatrix(fgAtenciones.row, 1) & ", debido a que se superó el límite de fecha para la sustentación"
                        Set objPista = Nothing
                        Exit Sub
                    Else
                        MsgBox "No se registro la ampliación de la prórroga. Es necesario que ingrese el motivo de la ampliación para el registro de la ampliación.", vbInformation, "Aviso"
                        Exit Sub
                    End If
                Case 7 'NO
                    Exit Sub
    
            End Select
        ElseIf vProrroga = 3 Then 'No se otorgo ninguna prorroga y se paso la fecha limite de sustentacion (4 Dias)
            If verificarLimiteSustentar(fgAtenciones.TextMatrix(fgAtenciones.row, 23), CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), CInt(fgAtenciones.TextMatrix(fgAtenciones.row, 24))) = False Then
                If diasAtraso = 0 Then
                    respuesta = MsgBox("La fecha de sustentación de " & fgAtenciones.TextMatrix(fgAtenciones.row, 6) & " venció el día de hoy. Al asignar la prórroga, ud. esta asumiendo que visualizó la autorización del Gerente de Administración." & _
                    vbNewLine & "¿Está seguro de otorgar la prórroga?", vbInformation + vbYesNo, "Confirmar Autorización")
                Else
                     respuesta = MsgBox("El plazo de sustentación de " & fgAtenciones.TextMatrix(fgAtenciones.row, 6) & IIf(diasAtraso > 1, " superó los " & diasAtraso & " días", " superó en " & diasAtraso & " día") & ". Al asignar la prórroga, ud. esta asumiendo que visualizó  la autorización del Gerente de Administración." & _
                     vbNewLine & "¿Está seguro de otorgar la prórroga?", vbInformation + vbYesNo, "Confirmar Autorización")
                End If
               
                Select Case respuesta
                     Case 6 'SI
                        cGlosa = InputBox("Motivo de la Ampliación", "Motivo")
                        If Not ValidaCadena(cGlosa) Then
                            MsgBox "No se puede permitir caracteres especiales en la glosa.", vbInformation, "Aviso"
                            Exit Sub
                        End If
                        If cGlosa <> "" Then
                            Call oNArendir.registrarProrrogaSustentar(DateDiff("d", fechaLLegada, fechaSis) + 2, CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, 3)
                            Call oNArendir.RegProrrogaSustentarAut(CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, cGlosa)
                            MsgBox "Se registro correctamente la prórroga", vbInformation, "Aviso"
                            Set objPista = New COMManejador.Pista
                            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se dio Prórroga Por Primera VEZ al  N°: " & fgAtenciones.TextMatrix(fgAtenciones.row, 1) & ".Con autorización del Gerente de Administración debido a que se superó la fecha se sustentación"
                            Set objPista = Nothing
                            Set oNArendir = Nothing
                            Exit Sub
                        Else
                            MsgBox "No se registro la ampliación de la prórroga. Es necesario que ingrese el motivo de la ampliación para el registro de la ampliación.", vbInformation, "Aviso"
                            Exit Sub
                        End If
                     Case 7 'NO
                        Exit Sub
                End Select
            End If
        End If
    End If
    'END TORE
    
    nProrroga = -1
    nProrroga = oNArendir.obtenerExistenciaProrrogaSustentar(CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)))
    
    If nProrroga = 0 Then
        Call oNArendir.registrarProrrogaSustentar(2, CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, 1) '1 Agregado por PASI20140102 TI-ERS107-2013
        MsgBox "Se realizó la primera prórroga correctamente", vbInformation, "Aviso"
        'ARLO20170208
        Set objPista = New COMManejador.Pista
        'gsOpeCod = LogPistaCierreDiarioCont
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se dio Prórroga Por Primera VEZ al  N°: " & fgAtenciones.TextMatrix(fgAtenciones.row, 1)
        Set objPista = Nothing
        '*******
        Set oNArendir = Nothing
    ElseIf nProrroga > 0 Then
         '***Modificado por PASI20140102 TI-ERS107-2013
        'Set oNArendir = Nothing
        'MsgBox "El " & fgAtenciones.TextMatrix(fgAtenciones.Row, 4) & " ya fue registrado su prorroga", vbInformation, "Aviso"
        'Exit Sub
        If (nProrroga = 1) Then
           If MsgBox("La primera prorroga ya fue registrada; Esta seguro de dar una segunda prorroga al " & fgAtenciones.TextMatrix(fgAtenciones.row, 1) & " - " & fgAtenciones.TextMatrix(fgAtenciones.row, 4), vbYesNo, "Aviso") = vbYes Then
                Call oNArendir.ActualizarProrrogaSustentar(2, CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, 2)
                MsgBox "Se registro correctamente", vbInformation, "Aviso"
                'ARLO20170208
                Set objPista = New COMManejador.Pista
                'gsOpeCod = LogPistaCierreDiarioCont
                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se dio Prórroga Por Segunda VEZ al  N°: " & fgAtenciones.TextMatrix(fgAtenciones.row, 1)
                Set objPista = Nothing
                '*******
           Else
                Exit Sub
           End If
           
        End If
            If (nProrroga = 2 Or nProrroga = 3) Then
                 '*** Comentado por TORE - Automatizacion de prorrogas ***
'                Set oNArendir = Nothing
'                MsgBox "La segunda prórroga para " & fgAtenciones.TextMatrix(fgAtenciones.Row, 6) & ", ya se encuentra registrado.", vbInformation, "Aviso"
'                Exit Sub
                 ' ***** END *****
                
                If cCargoAut <> cCargoCod Then
                    MsgBox "Sólo la Jefatura del Área de Contabilidad puede otorgar prórroga con autorización del Gerente de Administración.", vbInformation, "Aviso"
                    Exit Sub
                End If
        
                'Call ValidarCantProrrogas(nCantidadProrroga)
                
                vProrroga = verificarProrrogaSustentar(fgAtenciones.TextMatrix(fgAtenciones.row, 23), CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), CInt(fgAtenciones.TextMatrix(fgAtenciones.row, 24)))
                If vProrroga = 2 Then
                    If diasAtraso = 0 Then
                        respuesta = MsgBox("La fecha de sustentación de " & fgAtenciones.TextMatrix(fgAtenciones.row, 6) & " venció el día de hoy. Al asignar la prórroga, ud. esta asumiendo que visualizó la autorización del Gerente de Administración." & _
                        "¿Está seguro de otorgar la prórroga?", vbInformation + vbYesNo, "Confirmar Autorización")
                    Else
                        respuesta = MsgBox("La fecha para asignar prórroga superó los " & diasAtraso & " día(s). Al asignar la prórroga, ud. esta asumiendo que visualizó la autorización del Gerente de Administración." & _
                        "¿Está seguro de otorgar la prórroga?", vbInformation + vbYesNo, "Confirmar Autorización")
                    End If
        '            respuesta = MsgBox("La fecha para asignar prórroga superó los " & diasAtraso & " día(s). Al asignar la prórroga, ud. esta asumiendo que visualizó la autorización del Gerente de Administración." & _
        '            "¿Está seguro de otorgar la prórroga?", vbInformation + vbYesNo, "Confirmar Autorización")
                    Select Case respuesta
                        Case 6 'SI
                            cGlosa = InputBox("Motivo de la Ampliación", "Motivo")
                            If Not ValidaCadena(cGlosa) Then
                                MsgBox "No se puede permitir caracteres especiales en la glosa.", vbInformation, "Aviso"
                                Exit Sub
                            End If
                            If cGlosa <> "" Then
                                'Se otorgo la prórroga con autorización del Gerente de Administración(4)
                                Call oNArendir.ActualizarProrrogaSustentar(DateDiff("d", fechaLLegada, fechaSis) + 2, CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, 3)
                                Call oNArendir.RegProrrogaSustentarAut(CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, cGlosa)
                                MsgBox "Se registro correctamente la prórroga", vbInformation, "Aviso"
                                Set objPista = New COMManejador.Pista
                                objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se otorgó la prórroga con autorización del Gerente de Administración al  N°: " & fgAtenciones.TextMatrix(fgAtenciones.row, 1) & ", debido a que se superó el límite de fecha para la sustentación"
                                Set objPista = Nothing
                                Exit Sub
                            Else
                                MsgBox "No se registro la ampliación de la prórroga. Es necesario que ingrese el motivo de la ampliación para el registro de la ampliación.", vbInformation, "Aviso"
                            Exit Sub
                        End If
                        Case 7 'NO
                            'MsgBox "Ampliación de prorroga cancelada", vbInformation, "Aviso"
                            Exit Sub
            
                    End Select
                ElseIf vProrroga = 3 Then
                    If verificarLimiteSustentar(fgAtenciones.TextMatrix(fgAtenciones.row, 23), CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), CInt(fgAtenciones.TextMatrix(fgAtenciones.row, 24))) = False Then
                        If diasAtraso = 0 Then
                            respuesta = MsgBox("La fecha de sustentación de " & fgAtenciones.TextMatrix(fgAtenciones.row, 6) & " venció el día de hoy. Al asignar la prórroga, ud. esta asumiendo que visualizó la autorización del Gerente de Administración." & _
                            "¿Está seguro de otorgar la prórroga?", vbInformation + vbYesNo, "Confirmar Autorización")
                        Else
                             respuesta = MsgBox("La fecha para asignar la prórroga superó los " & diasAtraso & " día(s). Al asignar la prorroga ud. esta asumiendo que visualizó la autorización del Gerente de Administración." & _
                             "¿Está seguro de otorgar la prórroga?", vbInformation + vbYesNo, "Confirmar Autorización")
                        End If
                       
                        Select Case respuesta
                             Case 6 'SI
                                cGlosa = InputBox("Motivo de la Ampliación", "Motivo")
                                If Not ValidaCadena(cGlosa) Then
                                    MsgBox "No se puede permitir caracteres especiales en la glosa.", vbInformation, "Aviso"
                                    Exit Sub
                                End If
                                If cGlosa <> "" Then
                                    'Se otorgo la prórroga con autorización del Gerente de Administración(3)
                                    Call oNArendir.registrarProrrogaSustentar(DateDiff("d", fechaLLegada, fechaSis) + 2, CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, 3)
                                    Call oNArendir.RegProrrogaSustentarAut(CLng(fgAtenciones.TextMatrix(fgAtenciones.row, 16)), lsMovNro, cGlosa)
                                    MsgBox "Se registro correctamente la prórroga", vbInformation, "Aviso"
                                    Set objPista = New COMManejador.Pista
                                    objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se dio Prórroga Por Primera VEZ al  N°: " & fgAtenciones.TextMatrix(fgAtenciones.row, 1) & ".Con autorización del Gerente de Administración debido a que se superó la fecha se sustentación"
                                    Set objPista = Nothing
                                    Set oNArendir = Nothing
                                    Exit Sub
                                Else
                                    MsgBox "No se registro la ampliación de la prórroga. Es necesario que ingrese el motivo de la ampliación para el registro de la ampliación.", vbInformation, "Aviso"
                                    Exit Sub
                                End If
                             Case 7 'NO
                                'MsgBox "Ampliación de prorroga cancelada", vbInformation, "Aviso"
                                Exit Sub
                        End Select
                    End If
                End If
            End If

        '***FIN PASI
    End If
    nProrroga = -1

    'If MsgBox("¿Esta seguro que deseas dar una única prorroga al " & fgAtenciones.TextMatrix(fgAtenciones.Row, 1) & " - " & fgAtenciones.TextMatrix(fgAtenciones.Row, 4), vbYesNo, "Aviso") = vbYes Then
    '
    'End If
End Sub

Private Sub ValidarCantProrrogas(ByVal nCantProrroga As Integer)
    If nCantProrroga = 3 Then
        MsgBox "Ud. otorgó la primera prórroga con autorización de la Gerencia de Administración, al otorgar esta prórroga ya no se permitirá más prórrogas automáticas", vbInformation, "Aviso"
    End If
    If nCantProrroga = 4 Then
        MsgBox "Las prórrogas automáticas esta bloquedo, debido a que se superó el límite de asignación de prórroga. Consultar con el área de Tegnología de Información", vbInformation, "Aviso"
        Exit Sub
    End If
End Sub

Private Function ValidaCadena(ByVal Cadena As String) As Boolean
Dim tamanioCadena As String
Dim cadenaResultado As String
Dim caracteresValidos As String
Dim caracteresActual As String
Dim i As Integer

tamanioCadena = Len(Cadena)
If tamanioCadena > 0 Then
    caracteresValidos = " 0123456789abcdefghijklmnñopqrstuwxyzABCDEFGHIJKLMNÑOPQRSTUWXYZ-óíÓÍÑ,:."
    
    For i = 1 To tamanioCadena
        caracteresActual = Mid(Cadena, i, 1)
        
        If InStr(caracteresValidos, caracteresActual) Then
            ValidaCadena = True
        Else
            ValidaCadena = False
        End If
    Next
End If
End Function

Private Function verificarProrrogaSustentar(ByVal psRHCargoCategoria As String, ByVal psnMovNro As Long, ByVal pnTipoArendir As Integer) As Integer
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim rsSustentar As ADODB.Recordset
Set rsSustentar = New ADODB.Recordset
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim ldFechaProrroga As Date
Dim ldFechaLimite As Date
Dim X, Y As Integer
Dim i As Integer

Set rsSustentar = oNArendir.devolverProrrogaSustentar(psnMovNro)

If Not rsSustentar.BOF And Not rsSustentar.EOF Then
    i = rsSustentar!nNroProrroga * rsSustentar!nDias
    ldFechaLimite = FechaLimiteSustentar(psRHCargoCategoria, psnMovNro, pnTipoArendir)
    Do While X < i
       
        If oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", Y + 1, ldFechaLimite), gsCodAge) = True Then
            X = X + 1
        End If
        Y = Y + 1
    Loop
    ldFechaProrroga = DateAdd("D", Y, CDate(ldFechaLimite))

    If ldFechaProrroga >= gdFecSis Then
        verificarProrrogaSustentar = 1
    ElseIf ldFechaProrroga < gdFecSis Then
        verificarProrrogaSustentar = 2
    End If
Else
    verificarProrrogaSustentar = 3
End If

ldFechaProrroga = "01/01/1900"
i = 0
Set oNArendir = Nothing
Set rsSustentar = Nothing
Set oNContFunciones = Nothing

End Function
Private Function FechaLimiteSustentar(ByVal psRHCargoCategoria As String, ByVal pnViaticoMovNro As Long, ByVal pnTipoArendir As Integer) As Date
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim rsFecha As ADODB.Recordset
Set rsFecha = New ADODB.Recordset
Dim lnDiasLimite, i, j As Integer
Dim X, Y As Integer
Dim ldFechaLlegada, ldFechaEnInstitucion, ldFechaLimite As Date

ldFechaLlegada = oNArendir.obtenerFechaLlegadaColaborador(pnViaticoMovNro, pnTipoArendir)
lnDiasLimite = oNArendir.obtenerLimiteSustentarCategoria(psRHCargoCategoria, pnTipoArendir)

If ldFechaLlegada <> "01/01/1900" Then
    ldFechaEnInstitucion = ldFechaLlegada
Else
  Exit Function
End If

If Trim(psRHCargoCategoria) = "" Then
    If pnTipoArendir = gArendirTipoViaticos Then
        j = 5
    Else
        j = 3
    End If
Else
    j = lnDiasLimite
End If

X = 0
Y = 0

    Do While X < j
        If oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", Y + 1, ldFechaEnInstitucion), gsCodAge) = True Then
            X = X + 1
        End If
        Y = Y + 1
    Loop
    ldFechaLimite = DateAdd("D", Y, ldFechaEnInstitucion)
    FechaLimiteSustentar = ldFechaLimite
End Function

Private Function verificarLimiteSustentar(ByVal psRHCargoCategoria As String, ByVal pnViaticoMovNro As Long, ByVal pnTipoArendir As Integer) As Boolean
Dim oNArendir As NARendir
Set oNArendir = New NARendir
Dim oNContFunciones As NContFunciones
Set oNContFunciones = New NContFunciones
Dim rsFecha As ADODB.Recordset
Set rsFecha = New ADODB.Recordset
Dim lnDiasLimite, i, j As Integer
Dim X, Y As Integer '************ Agregado por PASI20131119 segun TI-ERS107-2013
Dim ldFechaLlegada, ldFechaEnInstitucion, ldFechaLimite As Date


verificarLimiteSustentar = False

ldFechaLlegada = oNArendir.obtenerFechaLlegadaColaborador(pnViaticoMovNro, pnTipoArendir)
lnDiasLimite = oNArendir.obtenerLimiteSustentarCategoria(psRHCargoCategoria, pnTipoArendir)

If ldFechaLlegada <> "01/01/1900" Then
    ldFechaEnInstitucion = ldFechaLlegada
   
Else
  Exit Function
End If

If Trim(psRHCargoCategoria) = "" Then
    If pnTipoArendir = gArendirTipoViaticos Then
        j = 5
    Else
        j = 3
    End If
Else
    j = lnDiasLimite
End If

X = 0
Y = 0

    Do While X < j
        If oNContFunciones.obtenerSiFechaEsLaborable(DateAdd("D", Y + 1, ldFechaEnInstitucion), gsCodAge) = True Then
            X = X + 1
        End If
        Y = Y + 1
    Loop
ldFechaLimite = DateAdd("D", Y, ldFechaEnInstitucion)

If ldFechaEnInstitucion >= gdFecSis Or ldFechaLimite >= gdFecSis Then
    verificarLimiteSustentar = True
End If

End Function

Private Sub cmdRegulariza_Click()
Dim sRecEstado As String
Dim lsNroArendir As String
Dim lsNroDoc As String
Dim lsFechaDoc As String
Dim lsPersCod As String
Dim lsPersNomb As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsDescDoc As String
Dim lnImporte As Currency
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String
Dim lsMovNroSolicitud As String
Dim lsAgeCod As String
Dim lsAgeDesc As String

If fgAtenciones.TextMatrix(1, 0) = "" Then
   Exit Sub
End If
lsNroArendir = fgAtenciones.TextMatrix(fgAtenciones.row, 4)
lsNroDoc = fgAtenciones.TextMatrix(fgAtenciones.row, 2)
lsFechaDoc = fgAtenciones.TextMatrix(fgAtenciones.row, 5)
lsPersCod = fgAtenciones.TextMatrix(fgAtenciones.row, 9)
lsPersNomb = fgAtenciones.TextMatrix(fgAtenciones.row, 6)
lsAreaCod = fgAtenciones.TextMatrix(fgAtenciones.row, 14)
lsAreaDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 13)

lsDescDoc = fgAtenciones.TextMatrix(fgAtenciones.row, 15)
lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 7))
lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 12))
lsMovNroAtenc = fgAtenciones.TextMatrix(fgAtenciones.row, 10)
lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.row, 16)
If lnTipoArendir = gArendirTipoViaticos Then
    lsAgeDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 17)
    lsAgeCod = fgAtenciones.TextMatrix(fgAtenciones.row, 18)
Else
    lsAgeDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 17)
    lsAgeCod = fgAtenciones.TextMatrix(fgAtenciones.row, 17)
End If
frmOpeRegDocs.Inicio lnArendirFase, lnTipoArendir, False, lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod, _
                     lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, lsDescDoc, lsMovNroAtenc, lnImporte, lsCtaArendir, _
                     lsCtaPendiente, lnSaldo, lsMovNroSolicitud, , , , chkTodos.value = vbChecked

fgAtenciones.TextMatrix(fgAtenciones.row, 12) = Format(frmOpeRegDocs.lnSaldo, gsFormatoNumeroView)
'If Val(fgAtenciones.TextMatrix(fgAtenciones.Row, 12)) = 0 Then  'Doc. Regularizado
'    fgAtenciones.EliminaFila fgAtenciones.Row
'End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Sustento Cuentas Pendientes de Regularizar N°: " & lsNroDoc
            Set objPista = Nothing
            '*******
fgAtenciones.SetFocus
End Sub
Private Sub cmdRendicion_Click()
Dim sRecEstado As String
Dim lsNroArendir As String
Dim lsNroDoc As String
Dim lsFechaDoc As String
Dim lsPersCod As String
Dim lsPersNomb As String
Dim lsAreaCod As String
Dim lsAreaDesc As String
Dim lsDescDoc As String
Dim lnImporte As Currency
Dim lnSaldo As Currency
Dim lsMovNroAtenc As String
Dim lsMovNroSolicitud As String
Dim lsAgeCod As String
Dim lsAgeDesc As String
Dim lsAbrevDoc As String

If fgAtenciones.TextMatrix(1, 0) = "" Then
    MsgBox "No existen Atenciones Pendientes", vbInformation, "Aviso"
    Exit Sub
End If

'***Agregado por ELRO el 20120525, según OYP-RFC005-2012 y OYP-RFC016-2012
If fgAtenciones.TextMatrix(fgAtenciones.row, 9) = gsCodPersUser Then
    MsgBox "Usted no puede realizar su propia Rendición", vbInformation, "Aviso"
    cmdRendicion.SetFocus
    Exit Sub
End If
'***Fin Agregado por ELRO*************************************************

lsNroArendir = fgAtenciones.TextMatrix(fgAtenciones.row, 4)
lsNroDoc = fgAtenciones.TextMatrix(fgAtenciones.row, 2)
lsFechaDoc = fgAtenciones.TextMatrix(fgAtenciones.row, 5)
lsPersCod = fgAtenciones.TextMatrix(fgAtenciones.row, 9)
lsPersNomb = fgAtenciones.TextMatrix(fgAtenciones.row, 6)
lsAreaCod = fgAtenciones.TextMatrix(fgAtenciones.row, 14)
lsAreaDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 13)
lsAbrevDoc = fgAtenciones.TextMatrix(fgAtenciones.row, 1)

lsDescDoc = fgAtenciones.TextMatrix(fgAtenciones.row, 15)
lnImporte = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 7))
lnSaldo = CCur(fgAtenciones.TextMatrix(fgAtenciones.row, 12))
lsMovNroAtenc = fgAtenciones.TextMatrix(fgAtenciones.row, 10)
lsMovNroSolicitud = fgAtenciones.TextMatrix(fgAtenciones.row, 16)
lsAgeDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 17)
lsAgeCod = fgAtenciones.TextMatrix(fgAtenciones.row, 17)

frmArendirRendicion.Inicio lnArendirFase, lnTipoArendir, lsNroArendir, lsNroDoc, lsFechaDoc, lsPersCod, _
                     lsPersNomb, lsAreaCod, lsAreaDesc, lsAgeCod, lsAgeDesc, lsDescDoc, lsMovNroAtenc, lsAbrevDoc, lnImporte, lsCtaArendir, _
                     lsCtaPendiente, lnSaldo, lsMovNroSolicitud, txtMovDesc

If frmArendirRendicion.vbOk Then
    fgAtenciones.EliminaFila fgAtenciones.row
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgAtenciones_Click()
If lnArendirFase <> ArendirExtornoAtencion And lnArendirFase <> ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 8)
    Else
        txtMovDesc = ""
    End If
End Sub

Private Sub fgAtenciones_GotFocus()
If lnArendirFase <> ArendirExtornoAtencion And lnArendirFase <> ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 8)
    Else
        txtMovDesc = ""
    End If

End Sub

Private Sub fgAtenciones_OnRowChange(pnRow As Long, pnCol As Long)
If fgAtenciones.TextMatrix(1, 0) <> "" Then
    If lnArendirFase <> ArendirExtornoAtencion And lnArendirFase <> ArendirExtornoRendicion Then
        txtMovDesc = fgAtenciones.TextMatrix(fgAtenciones.row, 8)
    Else
        txtMovDesc = ""
    End If
End If

End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
Dim lvItem As ListItem
Dim rsPer As New ADODB.Recordset
Dim sOpeCod As String

Set oContFunc = New NContFunciones
Set oNArendir = New NARendir
Set oAreas = New DActualizaDatosArea
Set oOperacion = New DOperacion
lSalir = False
Me.Caption = gsOpeDesc

CentraForm Me

chkTodos.Visible = False

If Mid(gsOpeCod, 3, 1) = gMonedaNacional Then
   gsSimbolo = gcMN
Else
   gsSimbolo = gcME
End If
lsTpoDocVoucher = oOperacion.EmiteDocOpe(gsOpeCod, OpeDocEstOpcionalDebeExistir, OpeDocMetAutogenerado)
lsCtaArendir = oOperacion.EmiteOpeCta(gsOpeCod, "H", "0")
If lsCtaArendir = "" Then
   MsgBox "Faltan asignar Cuentas Contables a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Sistemas", vbInformation, "Aviso"
   lSalir = True
   Exit Sub
End If
Select Case gsOpeCod
    '***Modificado por ELRO el 20120425, según OYP-RFC005-2012
    'Case gCGArendirCtaRendMN, gCGArendirCtaRendME, gCGArendirViatRendMN, gCGArendirViatRendME
    Case gCGArendirCtaRendMN, gCGArendirCtaRendME, gCGArendirViatRendMN, gCGArendirViatRendME, gCGArendirViatRend2MN, gCGArendirViatRend2ME, gCGArendirCtaRend2MN, gCGArendirCtaRend2ME
        lsCtaPendiente = oOperacion.EmiteOpeCta(gsOpeCod, "D", "1")
    Case Else
        lsCtaPendiente = oOperacion.EmiteOpeCta(gsOpeCod, "H", "1")
End Select

If lsCtaPendiente = "" Then
   MsgBox "Falta asignar Cuenta de Pendiente a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Sistemas", vbInformation, "Aviso"
   lSalir = True
   Exit Sub
End If

'***Comentado por ELRO el 20120425, según OYP-RFC005-2012
'TxtBuscarArendir.psRaiz = "A Rendir de..."
'***Fin Comentado por ELRO*******************************
If lnTipoArendir = gArendirTipoAgencias Then
    Set rsPer = oOperacion.CargaOpeObj(gCGArendirCtaSolMNAge, 1)
Else
    Set rsPer = oOperacion.CargaOpeObj(gCGArendirCtaSolMN, 1)
End If
If Not rsPer.EOF Then
'***Comentado por ELRO el 20120425, según OYP-RFC005-2012
'    TxtBuscarArendir.rs = oAreas.GetAgenciasAreas(rsPer!cOpeObjFiltro, 1)
'***Fin Comentado por ELRO*******************************
End If
txtBuscarAgenciaArea.rs = oAreas.GetAgenciasAreas

Select Case lnArendirFase
    Case ArendirSustentacion
        cmdRendicion.Visible = False
        
    Case ArendirRendicion
        cmdRendicion.Visible = True
    Case ArendirExtornoAtencion, ArendirExtornoRendicion
        cmdRendicion.Visible = False
        cmdRegulariza.Visible = False
        txtMovDesc.Locked = False
End Select
Select Case lnTipoArendir
    Case gArendirTipoCajaChica
         Me.Height = 5550
         cmdRendicion.Top = 5030 - cmdRendicion.Height
         cmdRegulariza.Top = 5030 - cmdRegulariza.Height
         cmdSalir.Top = 5030 - cmdSalir.Height
         cmdRendicion.Visible = True
         lsDocTpoRecibo = oOperacion.EmiteDocOpe(gsOpeCod, OpeDocEstObligatorioDebeExistir, OpeDocMetDigitado)
         If lsDocTpoRecibo = "" Then
            MsgBox "No se asignó Tipo de Documento Recibo de A rendir a Operación", vbCritical, "Error"
            lSalir = True
            Exit Sub
         End If
         fgAtenciones.ColWidth(2) = fgAtenciones.ColWidth(2) - 300
         fgAtenciones.ColWidth(3) = fgAtenciones.ColWidth(3) + 300
    Case gArendirTipoViaticos
         fgAtenciones.EncabezadosAnchos = "350-600-0-0-1200-900-3000-1200-0-0-0-0-1200-2000-0-0-0-2000-0"
         'fgAtenciones.FormaCabecera
End Select
'***Agregado por ELRO el 20120425, según OYP-RFC005-2012
If gsOpeCod = gCGArendirViatRend2MN Or gsOpeCod = gCGArendirViatRend2ME Or gsOpeCod = gCGArendirCtaRend2MN Or gsOpeCod = gCGArendirCtaRend2ME Then
    cmdProrroga.Visible = True
    cmdDsctoPlanilla.Visible = True
End If
'***Fin Agregado por ELRO*******************************
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set oContFunc = Nothing
Set oNArendir = Nothing
Set oAreas = Nothing
End Sub
Public Property Get sPendiente() As String
sPendiente = sCtaPendiente
End Property
Public Property Let sPendiente(ByVal vNewValue As String)
sCtaPendiente = sPendiente
End Property
Private Sub txtBuscarAgenciaArea_EmiteDatos()
lblAgenciaArea = oAreas.GetNombreAreas(Mid(txtBuscarAgenciaArea, 1, 3))
'***Modificado por ELRO el 20120505, según OYP-RFC005-2012
'lblAgeDesc = oAreas.GetNombreAgencia(Mid(txtBuscarAgenciaArea, 4, 2))
lblAgeDesc = oAreas.GetNombreAgencia(IIf(Mid(txtBuscarAgenciaArea, 4, 2) = "", "01", Mid(txtBuscarAgenciaArea, 4, 2)))
'***Fin Modificado por ELRO*******************************
If txtBuscarAgenciaArea <> "" Then
   cmdBuscar.SetFocus
Else
   txtBuscarAgenciaArea.SetFocus
End If
End Sub

'***Comentado por ELRO el 20120425, según OYP-RFC005-2012
'Private Sub TxtBuscarArendir_EmiteDatos()
'lblDescArendir = Trim(TxtBuscarArendir.psDescripcion)
'If TxtBuscarArendir.Enabled Then
'    chkSelec.SetFocus
'End If
'End Sub
'***Fin Comentado por ELRO*******************************
