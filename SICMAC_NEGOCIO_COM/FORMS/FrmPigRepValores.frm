VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmPigRepValores 
   Caption         =   "Reportes de Boveda"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   Icon            =   "FrmPigRepValores.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmVentaMaterial 
      Caption         =   "Venta por Material"
      Height          =   1065
      Left            =   6045
      TabIndex        =   21
      Top             =   1410
      Visible         =   0   'False
      Width           =   2670
      Begin VB.TextBox txtCodTienda 
         Height          =   285
         Left            =   1680
         TabIndex        =   23
         Top             =   225
         Width           =   375
      End
      Begin VB.TextBox txtMesAno 
         Height          =   285
         Left            =   1680
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Cod Tienda:"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha(mm/yyyy):"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   615
         Width           =   1215
      End
   End
   Begin VB.Frame FrmCliente 
      Caption         =   "Cliente"
      Height          =   705
      Left            =   6135
      TabIndex        =   17
      Top             =   3105
      Visible         =   0   'False
      Width           =   2685
      Begin VB.TextBox txtcodper 
         Height          =   300
         Left            =   720
         TabIndex        =   19
         Top             =   285
         Width           =   1125
      End
      Begin VB.CommandButton cmdBuscar 
         Height          =   390
         Left            =   2040
         Picture         =   "FrmPigRepValores.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Buscar ..."
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label4 
         Caption         =   "Cliente"
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   323
         Width           =   510
      End
   End
   Begin MSComctlLib.ProgressBar PgrAvance 
      Height          =   435
      Left            =   5385
      TabIndex        =   13
      Top             =   5895
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame frmPeriodo 
      Caption         =   "Periodo"
      Height          =   1065
      Left            =   6495
      TabIndex        =   8
      Top             =   135
      Visible         =   0   'False
      Width           =   1830
      Begin MSMask.MaskEdBox txtFecini 
         Height          =   300
         Left            =   570
         TabIndex        =   14
         Top             =   255
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtfecFin 
         Height          =   300
         Left            =   570
         TabIndex        =   15
         Top             =   660
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecINi 
         Caption         =   "Del:"
         Height          =   255
         Left            =   165
         TabIndex        =   10
         Top             =   285
         Width           =   390
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Al:"
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   690
         Width           =   375
      End
   End
   Begin VB.Frame frmRemate 
      Caption         =   "Remate"
      Height          =   705
      Left            =   6495
      TabIndex        =   5
      Top             =   2340
      Visible         =   0   'False
      Width           =   1830
      Begin VB.TextBox txtnRemate 
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblnRemate 
         Caption         =   "Nro:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame frmFecha 
      Caption         =   "Fecha"
      Height          =   705
      Left            =   6495
      TabIndex        =   3
      Top             =   1380
      Visible         =   0   'False
      Width           =   1815
      Begin MSMask.MaskEdBox txtFech 
         Height          =   300
         Left            =   450
         TabIndex        =   16
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Al:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   255
         Width           =   255
      End
   End
   Begin VB.Timer Timer1 
      Left            =   7935
      Top             =   5100
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
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
      Left            =   7095
      TabIndex        =   2
      Top             =   6435
      Width           =   1335
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Left            =   5385
      TabIndex        =   1
      Top             =   6420
      Width           =   1455
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   912
      Top             =   5256
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPigRepValores.frx":040C
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPigRepValores.frx":075E
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPigRepValores.frx":0AB0
            Key             =   "Hijito"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPigRepValores.frx":0E02
            Key             =   "Bebe"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TVRep 
      Height          =   6552
      Left            =   168
      TabIndex        =   0
      Top             =   96
      Width           =   4716
      _ExtentX        =   8334
      _ExtentY        =   11562
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "imglstFiguras"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   5850
      SizeMode        =   1  'Stretch
      TabIndex        =   12
      Top             =   1260
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblmensaje 
      Caption         =   "Espere un momento ................"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5535
      TabIndex        =   11
      Top             =   5580
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "FrmPigRepValores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Dim fso As Object
'Dim sCad As NCredReporte
'Dim sCad1 As String
'
'Dim sOpePadre As String
'Dim sOpeHijo As String
'Dim sOpeHijito As String
'
'Dim xlAplicacion As Excel.Application
'Dim xlLibro As Excel.Workbook
'Dim xlHoja1 As Excel.Worksheet
'Dim xlHoja2 As Excel.Worksheet
'
'Dim Ruta As String
'
'Private Sub cmdImprimir_Click()
'Dim oPrevio As Previo.clsPrevio
'Dim sCad As NPigReporte
'Dim strs As String
'Dim oRep As NPigReporte
'Dim sNomRepo As String
'
'    Set oRep = New NPigReporte
'
'    Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
'    Case gColCredRepIngxPagoCred
'   Case 158105
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepStockLotesTransito
'        sNomRepo = "Reportes Para Boveda"
'         Case 158201
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'              If IsDate(txtFecini.Text) = False Then
'            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
'            txtFecini.SetFocus
'            Exit Sub
'      End If
'      If txtfecFin.Visible = True Then
'        If IsDate(txtfecFin.Text) = False Then
'            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
'            txtfecFin.SetFocus
'            Exit Sub
'        End If
'      End If
'        If Len(txtnRemate) = 0 Then
'            MsgBox "Ingrese el número de Remate", vbExclamation, "Aviso"
'            txtnRemate.SetFocus
'            Exit Sub
'        Else
'        strs = oRep.RepPolizasEmitidas(txtFecini, txtfecFin, txtnRemate)
'        sNomRepo = "Reportes de Tienda"
'        End If
'    Case 158106
'
'      If IsDate(txtFech.Text) = False Then
'            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
'            txtFecini.SetFocus
'            Exit Sub
'      End If
'         oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'       strs = oRep.RepInventarioJoyas(txtFech)
'
'    Case 158101
'        strs = oRep.Report_158101(Me.txtFech)
'        sNomRepo = "Reportes - Garantias en Custodia"
'    Case 158108
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepResumenSaldoAgencia
'        sNomRepo = "Reportes Para Boveda"
'    Case 158104
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepBoveda(gsCodAge)
'        sNomRepo = "Reportes Para Boveda"
'        Case 158913
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepMovimientosMensualesClientes(txtcodper)
'        sNomRepo = "Reportes Estadisticos"
'
'Case 158905
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepDistribuciónColocaciones()
'        sNomRepo = "Reportes Para Boveda"
'    Case 158202
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        If IsDate(txtFecini.Text) = False Then
'            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
'            txtFecini.SetFocus
'            Exit Sub
'        End If
'        If txtfecFin.Visible = True Then
'            If IsDate(txtfecFin.Text) = False Then
'               MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
'               txtfecFin.SetFocus
'               Exit Sub
'            End If
'        End If
'         If Len(txtnRemate) = 0 Then
'            MsgBox "Ingrese el número de Remate", vbExclamation, "Aviso"
'            txtnRemate.SetFocus
'            Exit Sub
'        Else
'        strs = oRep.RepPolizasEmitidasResumen(txtFecini, txtfecFin, txtnRemate)
'        sNomRepo = "Reportes Para Tienda"
'        End If
'    Case 158203
'    oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'    If Len(txtCodTienda) = 0 Then
'         MsgBox "Ingrese el código de la Tienda", vbExclamation, "Aviso"
'         txtCodTienda.SetFocus
'         Exit Sub
'    End If
'    If Len(txtMesAno) = 0 Then
'         MsgBox "Ingrese la fecha", vbExclamation, "Aviso"
'         txtMesAno.SetFocus
'         Exit Sub
'    End If
'     strs = oRep.RepVentaMaterial(txtCodTienda, txtMesAno)
'        sNomRepo = "Reportes de Tienda"
'    Case 158301
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepPiezasAdjTasador("580")
'        sNomRepo = "Reportes Para Boveda"
'    Case 158302
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'         If Len(txtnRemate) = 0 Then
'            MsgBox "Ingrese el número de Remate", vbExclamation, "Aviso"
'            txtnRemate.SetFocus
'            Exit Sub
'        Else
'        strs = oRep.RepPiezasRematadas(CInt(txtnRemate.Text))
'        sNomRepo = "Reportes de Remate y Adjudicados"
'        End If
'    Case 158306
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepPiezasSelFundicion
'        sNomRepo = "Reportes Para Boveda"
'    Case 158307
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepPiezasSelVenta
'        sNomRepo = "Reportes Para Boveda"
'    Case 158308
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepPiezasSelFundicionPorRemate
'        sNomRepo = "Reportes Para Boveda"
'    Case 158309
'        oRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
'        strs = oRep.RepPiezasSelVentaPorRemate
'        sNomRepo = "Reportes Para Boveda"
'    Case "158401", "158402", "158403", "158404"
'        PgrAvance.Visible = True
'        PgrAvance.value = 0
'        lblmensaje.Visible = True
'        Call ReporteExcell(CInt(Mid(txtFech.Text, 4, 2)), Mid(txtFech.Text, 7, 4))
'        lblmensaje.Visible = False
'    End Select
'
'    If Mid(TVRep.SelectedItem.Text, 1, 4) <> "1584" Then
'        Set oRep = Nothing
'        Set oPrevio = New Previo.clsPrevio
'        oPrevio.Show strs, sNomRepo, True, 66, gEPSON
'        Set oPrevio = Nothing
'    End If
'
'    'Call HabilitaControles(True, False, False, False, False, False)
'
'
'
'End Sub
'
'Private Sub Command1_Click()
'
'End Sub
'
'Private Sub cmdsalir_Click()
'Unload Me
'End Sub
'
'Private Sub Form_Load()
'    Dim i As Integer
'    Dim lsCadena As String
'    Dim lsCad As String
'    Dim oPersona As UPersona
'    'Dim oRRHH As DActualizaDatosRRHH
'    Dim oPrevio As Previo.clsPrevio
'    Set oPrevio = New Previo.clsPrevio
'    Dim oRep As NPigReporte
'    Set oRep = New NPigReporte
'    Dim oRepr As NPigRemate
''    Set oRepr = New NPigReporte
'
'Dim lnNumRep As Integer
'Dim lsCadenaBuscar As String
'Dim lsRep() As String
''Llenar el arbol
'LlenaArbol
'
'   ' TxtFecIni = Format(gdFecSis, gsFormatoFechaView)
'   ' TxtFecFin = Format(gdFecSis, gsFormatoFechaView)
'
'End Sub
'
'Private Sub LlenaArbol()
'Dim clsGen As DGeneral
'Dim rsUsu As Recordset
'Dim sOperacion As String, sOpeCod As String
'Dim nodOpe As Node
'Dim lsTipREP As String
'
''Para filtrar el tipo de reporte de la tabla OpeTipo
'    lsTipREP = "158"
'
'    Set clsGen = New DGeneral
'
'    Set rsUsu = clsGen.GetOperacionesUsuario(gsCodUser, lsTipREP, MatOperac, NroRegOpe)
'    Set clsGen = Nothing
'
'    Do While Not rsUsu.EOF
'        sOpeCod = rsUsu("cOpeCod")
'        sOperacion = sOpeCod & " - " & UCase(rsUsu("cOpeDesc"))
'        Select Case rsUsu("nOpeNiv")
'            Case "1"
'                sOpePadre = "P" & sOpeCod
'                Set nodOpe = TVRep.Nodes.Add(, , sOpePadre, sOperacion, "Padre")
'                nodOpe.Tag = sOpeCod
'            Case "2"
'                sOpeHijo = "H" & sOpeCod
'                Set nodOpe = TVRep.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, sOperacion, "Hijo")
'                nodOpe.Tag = sOpeCod
'            Case "3"
'                sOpeHijito = "J" & sOpeCod
'                Set nodOpe = TVRep.Nodes.Add(sOpeHijo, tvwChild, sOpeHijito, sOperacion, "Hijito")
'                nodOpe.Tag = sOpeCod
'            Case "4"
'                Set nodOpe = TVRep.Nodes.Add(sOpeHijito, tvwChild, "B" & sOpeCod, sOperacion, "Bebe")
'                nodOpe.Tag = sOpeCod
'        End Select
'        rsUsu.MoveNext
'    Loop
'    rsUsu.Close
'    Set rsUsu = Nothing
'End Sub
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
'    If Not loPers Is Nothing Then
'        lsPersCod = loPers.sPersCod
'        lsPersNombre = loPers.sPersNombre
'    End If
'Set loPers = Nothing
'
'txtcodper.Text = lsPersCod
''txtNombre.Text = lsPersNombre
'
'
'Exit Sub
'
'ControlError:
'    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & '        " Avise al Area de Sistemas ", vbInformation, " Aviso "
'End Sub
'
'
'Private Sub TVRep_Click()
'' Se inserto una constante gColPR-epBovedaLotes en el proyecto
' Call HabilitaControles(True, False, False, False, False, False)
'   Select Case Mid(TVRep.SelectedItem.Text, 1, 6)
''       Case "158101"
''               Call HabilitaControles(True, True, False, False, False)
'       Case "158101", "158106", "158401", "158402", "158403", "158404", "158104", "158105", "158306", "158307", "158308", "158309", "158108", "158905"
'               Call HabilitaControles(True, False, True, False, False, False)
'       Case "158201", "158202"
'               Call HabilitaControles(True, True, False, True, False, False)
'        Case "158302", "158301"
'               Call HabilitaControles(True, False, False, True, False, False)
'        Case "158203"
'                Call HabilitaControles(True, False, False, False, True, False)
'        Case "158913"
'                Call HabilitaControles(True, False, False, False, False, True)
'     End Select
'             cmdImprimir.Enabled = True
'End Sub
'
'
'Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If txtFecini.Visible = True Then
'        If IsDate(txtfecFin.Text) = False Then
'            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
'            txtfecFin.SetFocus
'            Exit Sub
'        End If
'    cmdImprimir.SetFocus
'Else
'            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
'            txtfecFin.SetFocus
'            Exit Sub
'
'    End If
'End If
'End Sub
'
'Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    If txtFecini.Visible = True Then
'        If IsDate(txtFecini.Text) = False Then
'            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
'            txtFecini.SetFocus
'            Exit Sub
'        End If
'Else
'            MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
'            txtFecini.SetFocus
'            Exit Sub
'
'    End If
'End If
'
'End Sub
'
'Public Sub ReporteExcell(nMes As Integer, nano As String)
'Dim mesano As String
'On Error Resume Next
'
'mesano = fgDameNombreMes(nMes) & "_" & nano
'
'Set fso = CreateObject("Scripting.FileSystemObject")
'fso.CreateFolder (App.path & "\Spooler\" & "riesgos")
'fso.CreateFolder (App.path & "\Spooler\riesgos\" & mesano)
'Ruta = App.path & "\Spooler\riesgos\" & mesano & "\"
'
'Set fso = Nothing
'GenerarExcell (Ruta)
'End Sub
'
''***********************************************************
'' Inicia Trabajo con EXCEL, crea variable Aplicacion y Libro
''***********************************************************
'Private Function ExcelBegin(psArchivo As String, '        xlAplicacion As Excel.Application, '        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
'
'Dim fs As New Scripting.FileSystemObject
'On Error GoTo ErrBegin
'Set fs = New Scripting.FileSystemObject
'Set xlAplicacion = New Excel.Application
'
'If fs.FileExists(psArchivo) Then
'   If pbBorraExiste Then
'      fs.DeleteFile psArchivo, True
'      Set xlLibro = xlAplicacion.Workbooks.Add
'   Else
'      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
'   End If
'Else
'   Set xlLibro = xlAplicacion.Workbooks.Add
'End If
'ExcelBegin = True
'Exit Function
'ErrBegin:
'  MsgBox Err.Description, vbInformation, "Aviso"
'  ExcelBegin = False
'End Function
'
''***********************************************************
'Private Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
'On Error GoTo ErrEnd
'   If plSave Then
'        xlHoja1.SaveAs psArchivo
'   End If
'   xlLibro.Close
'   xlAplicacion.Quit
'   Set xlAplicacion = Nothing
'   Set xlLibro = Nothing
'   Set xlHoja1 = Nothing
'Exit Sub
'ErrEnd:
'   MsgBox Err.Description, vbInformation, "Aviso"
'End Sub
'
'Private Sub GenerarExcell(sRuta As String)
'    Dim lsArchivoN  As String
'    Dim lbLibroOpen As Boolean
'  '  Mid(TVRep.SelectedItem.Text, 1, 6)
'    lsArchivoN = sRuta & TVRep.SelectedItem.Text & ".xls"
'    OleExcel.Class = "ExcelWorkSheet"
'    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
'       If lbLibroOpen Then
'          Set xlHoja1 = xlLibro.Worksheets(1)
'            ExcelAddHoja "hola", xlLibro, xlHoja1
'          End If
'          Call GeneraReporteHoja1(Mid(TVRep.SelectedItem.Text, 1, 6))
'          OleExcel.Class = "ExcelWorkSheet"
'          ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
''
'          OleExcel.SourceDoc = lsArchivoN
'          OleExcel.Verb = 1
'          OleExcel.Action = 1
'          'OleExcel.DoVerb  -1
''       End If
'End Sub
'
'Private Sub GeneraReporteHoja1(ncodigo As String)
'    Dim i As Integer
'    Dim z As Integer
'    Dim j As Integer
'    Dim nFila As Integer
'    Dim nIni  As Integer
'    Dim lNegativo As Boolean
'    Dim sConec As String
'    Dim lnAcum As Currency
'    Dim lnPosI As Integer
'    Dim lnPosJ As Integer
'
'    Dim sTipoGara As String
'    Dim sTipoCred As String
'
'    Dim lbUsu As Boolean
'    Dim LbNom As Boolean
'    Dim lbTot As Boolean
'
'    Dim lsCadAux As String
'    Dim lsCadAuxMes As String
'
'    Dim oRep As DpigReportes
'    Dim rs As ADODB.Recordset
'
'    Dim CantidadAntes As Double
'    Dim CantidadDespues As Double
'
'    Set oRep = New DpigReportes
'
'
'    Select Case ncodigo
'
'    Case "158401"
'    Set rs = oRep.RepClientesCalifManual
'    PgrAvance.Max = rs.RecordCount
'
'    If Not (rs.EOF And rs.BOF) Then
'       rs.MoveFirst
'        i = 4
'        xlHoja1.Columns("B:B").NumberFormat = "@"
'        xlHoja1.Cells(4, 2) = "Codigo"
'        xlHoja1.Cells(4, 3) = "Nombre"
'        xlHoja1.Cells(4, 4) = "Eval Actual"
'        xlHoja1.Cells(4, 5) = "Eval Anterior"
'        xlHoja1.Cells(4, 6) = "Fecha Eval"
'        xlHoja1.Cells(4, 7) = "Usuario"
'        xlHoja1.Range("B4:G4").Font.Bold = True
'        xlHoja1.Range("B4:G4").Interior.ColorIndex = 15
'        xlHoja1.Range("B4:G4").Interior.Pattern = xlSolid
'
'       Do
'            PgrAvance.value = i - 3
'            i = i + 1
'            xlHoja1.Cells(i, 2) = CStr(rs!cPersCod)
'            xlHoja1.Cells(i, 3) = rs!cPersNombre
'            xlHoja1.Cells(i, 4) = CStr(rs!cEvalPigno)
'            xlHoja1.Cells(i, 5) = CStr(rs!cEvalPignoAnterior)
'            xlHoja1.Cells(i, 6) = rs!fEvalPigno
'            xlHoja1.Cells(i, 7) = rs!cUsuarioEvalPigno
'         rs.MoveNext
'       Loop Until rs.EOF
'
'    End If
'    xlHoja1.Columns("B:B").EntireColumn.AutoFit
'    xlHoja1.Range("B4:B" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("C:C").EntireColumn.AutoFit
'    xlHoja1.Range("C4:C" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("D:D").EntireColumn.AutoFit
'    xlHoja1.Range("D4:D" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("E:E").EntireColumn.AutoFit
'    xlHoja1.Range("E4:E" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("F:F").NumberFormat = "m/d/yyyy"
'    xlHoja1.Columns("F:F").EntireColumn.AutoFit
'    xlHoja1.Range("F4:F" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("G:G").EntireColumn.AutoFit
'    xlHoja1.Range("G4:G" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Range("B4:G4").BorderAround 1, xlMedium
'    xlHoja1.Range("B1:G" & Trim(Str(i))).Font.Name = "Arial"
'    xlHoja1.Range("B1:G" & Trim(Str(i))).Font.Size = 8
'
'    Case "158402"
'    Set rs = oRep.RepClientesAntDes
'    PgrAvance.Max = rs.RecordCount
'
'    If Not (rs.EOF And rs.BOF) Then
'       rs.MoveFirst
'       i = 4
'       z = 4
'       CantidadAntes = 0
'       CantidadDespues = 0
'        xlHoja1.Cells(4, 2) = "Tipo de Cliente"
'        xlHoja1.Cells(4, 3) = "Antes"
'        xlHoja1.Cells(4, 4) = "Después "
'        xlHoja1.Range("B4:D4").Font.Bold = True
'         xlHoja1.Range("B4:D4").Interior.ColorIndex = 15
'        xlHoja1.Range("B4:D4").Interior.Pattern = xlSolid
'        xlHoja1.Columns("C:D").NumberFormat = "#,##0"
'
'       Do
'            PgrAvance.value = i - 3
'
'            If rs!Situacion = "Ant" Then
'                i = i + 1
'                xlHoja1.Cells(i, 2) = rs!cConsDescripcion
'                xlHoja1.Cells(i, 3) = rs!Cantidad
'                CantidadAntes = CantidadAntes + CDbl(rs!Cantidad)
'            Else
'                z = z + 1
'                xlHoja1.Cells(z, 4) = rs!Cantidad
'                CantidadDespues = CantidadDespues + CDbl(rs!Cantidad)
'
'            End If
'         rs.MoveNext
'
'       Loop Until rs.EOF
'    End If
'    xlHoja1.Cells(i + 1, 2) = "Total :"
'    xlHoja1.Cells(i + 1, 3) = CantidadAntes
'    xlHoja1.Cells(i + 1, 4) = CantidadDespues
'
'    xlHoja1.Columns("B:B").EntireColumn.AutoFit
'    xlHoja1.Range("B4:B" & Trim(Str(i + 1))).BorderAround 1, xlMedium
'    xlHoja1.Columns("C:C").EntireColumn.AutoFit
'    xlHoja1.Range("C4:C" & Trim(Str(i + 1))).BorderAround 1, xlMedium
'    xlHoja1.Columns("D:D").EntireColumn.AutoFit
'    xlHoja1.Range("D4:D" & Trim(Str(i + 1))).BorderAround 1, xlMedium
'    xlHoja1.Range("B4:D4").BorderAround 1, xlMedium
'    xlHoja1.Range("B" & Trim(Str(i + 1)) & ":D" & Trim(Str(i + 1))).BorderAround 1, xlMedium
'    xlHoja1.Range("B" & Trim(Str(i + 1)) & ":D" & Trim(Str(i + 1))).Font.Bold = True
'    Case "158403"
'    Set rs = oRep.RepVariacionCalifTipCliente
'    PgrAvance.Max = rs.RecordCount
'
'    If Not (rs.EOF And rs.BOF) Then
'       rs.MoveFirst
'       i = 4
'       xlHoja1.Cells(4, 2) = "Eval Anterior"
'        xlHoja1.Cells(4, 3) = "Eval Actual"
'        xlHoja1.Cells(4, 4) = "Cantidad "
'        xlHoja1.Range("B4:D4").Font.Bold = True
'         xlHoja1.Range("B4:D4").Interior.ColorIndex = 15
'        xlHoja1.Range("B4:D4").Interior.Pattern = xlSolid
'        xlHoja1.Columns("C:D").NumberFormat = "#,##0"
'       Do
'            PgrAvance.value = i - 3
'            i = i + 1
'            xlHoja1.Cells(i, 2) = rs!des2
'            xlHoja1.Cells(i, 3) = rs!des1
'            xlHoja1.Cells(i, 4) = rs!CANT
'         rs.MoveNext
'       Loop Until rs.EOF
'    End If
'    xlHoja1.Columns("B:B").EntireColumn.AutoFit
'    xlHoja1.Range("B4:B" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("C:C").EntireColumn.AutoFit
'    xlHoja1.Range("C4:C" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("D:D").EntireColumn.AutoFit
'    xlHoja1.Range("D4:D" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Range("B4:D4").BorderAround 1, xlMedium
'
'    Case "158404"
'    Set rs = oRep.RepClientesVarCalifi
'    PgrAvance.Max = rs.RecordCount
'
'   If Not (rs.EOF And rs.BOF) Then
'      rs.MoveFirst
'       i = 4
'       xlHoja1.Cells(4, 2) = "Codigo"
'        xlHoja1.Cells(4, 3) = "Nombre"
'        xlHoja1.Cells(4, 4) = "Distrito "
'        xlHoja1.Cells(4, 5) = "Fecha Ing"
'        xlHoja1.Cells(4, 6) = "Sexo"
'        xlHoja1.Cells(4, 7) = "Eval Anterior"
'        xlHoja1.Cells(4, 8) = "Eval Actual"
'
'
'        xlHoja1.Range("B4:H4").Font.Bold = True
'         xlHoja1.Range("B4:H4").Interior.ColorIndex = 15
'        xlHoja1.Range("B4:H4").Interior.Pattern = xlSolid
'       xlHoja1.Columns("B:B").NumberFormat = "@"
'       Do
'            PgrAvance.value = i - 3
'            i = i + 1
'            xlHoja1.Cells(i, 2) = CStr(rs!cPersCod)
'            xlHoja1.Cells(i, 3) = rs!cPersNombre
'            xlHoja1.Cells(i, 4) = rs!cUbiGeoDescripcion
'            xlHoja1.Cells(i, 5) = rs!dPersIng
'            xlHoja1.Cells(i, 6) = rs!cPersNatSexo
'            xlHoja1.Cells(i, 7) = CStr(rs!cCalifiClienteAnt)
'            xlHoja1.Cells(i, 8) = CStr(rs!cCalifiCliente)
'
'         rs.MoveNext
'       Loop Until rs.EOF
'    End If
'    xlHoja1.Columns("B:B").EntireColumn.AutoFit
'    xlHoja1.Range("B4:B" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("C:C").EntireColumn.AutoFit
'    xlHoja1.Range("C4:C" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("D:D").EntireColumn.AutoFit
'    xlHoja1.Range("D4:D" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("E:E").NumberFormat = "m/d/yyyy"
'    xlHoja1.Columns("E:E").EntireColumn.AutoFit
'    xlHoja1.Range("E4:E" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("E:E").NumberFormat = "m/d/yyyy"
'    xlHoja1.Columns("F:F").EntireColumn.AutoFit
'    xlHoja1.Range("F4:F" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("G:G").EntireColumn.AutoFit
'    xlHoja1.Range("G4:G" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Columns("H:H").EntireColumn.AutoFit
'    xlHoja1.Range("H4:H" & Trim(Str(i))).BorderAround 1, xlMedium
'    xlHoja1.Range("B4:H4").BorderAround 1, xlMedium
'   End Select
'
'   xlHoja1.Range("B1:B2").Font.Bold = True
'   xlHoja1.Cells(1, 2) = "CAJA METROPOLITANA - SISTEMA DE CREDITOS PIGNORATICIOS"
'   xlHoja1.Cells(2, 2) = Mid(TVRep.SelectedItem.Text, 10) & " AL " & txtFech
'   xlHoja1.Range("B1:G" & Trim(Str(i + 1))).Font.Name = "Arial"
'   xlHoja1.Range("B1:G" & Trim(Str(i + 1))).Font.Size = 8
'   lblmensaje.Visible = False
'   PgrAvance.Visible = False
'   MsgBox "El reporte se genero correctamente", vbInformation, "Aviso"
'
'
'End Sub
'
'
'Private Sub HabilitaControles(ByVal pbcmdImprimir As Boolean, '        ByVal pbfrmPeriodo As Boolean, ByVal pbfrmFecha As Boolean, ByVal pbfrmRemate As Boolean, pbfrmVentaMaterial As Boolean, pbfrmCliente As Boolean)
'
'    Me.cmdImprimir.Visible = pbcmdImprimir
'    Me.frmPeriodo.Visible = pbfrmPeriodo
'    Me.frmFecha.Visible = pbfrmFecha
'    Me.frmRemate.Visible = pbfrmRemate
'    Me.frmVentaMaterial.Visible = pbfrmVentaMaterial
'    Me.FrmCliente.Visible = pbfrmCliente
'
'End Sub
'
'
'
