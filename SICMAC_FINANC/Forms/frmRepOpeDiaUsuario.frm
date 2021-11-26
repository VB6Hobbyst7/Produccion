VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRepOpeDiaUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones: Reporte de Operaciones Por Usuario"
   ClientHeight    =   5985
   ClientLeft      =   645
   ClientTop       =   1740
   ClientWidth     =   11370
   Icon            =   "frmRepOpeDiaUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkEliminados 
      Caption         =   "Incluir Movimientos Eliminados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   7875
      TabIndex        =   6
      Top             =   225
      Width           =   2025
   End
   Begin TabDlg.SSTab tTipo 
      Height          =   900
      Left            =   195
      TabIndex        =   0
      Top             =   225
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   1588
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Usuario"
      TabPicture(0)   =   "frmRepOpeDiaUsuario.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cboUsuario"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Operación"
      TabPicture(1)   =   "frmRepOpeDiaUsuario.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtOpeDes"
      Tab(1).Control(1)=   "txtOpecod"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtOpeDes 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73650
         TabIndex        =   3
         Top             =   450
         Width           =   4035
      End
      Begin VB.ComboBox cboUsuario 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   435
         Width           =   5265
      End
      Begin Sicmact.TxtBuscar txtOpecod 
         Height          =   330
         Left            =   -74850
         TabIndex        =   2
         Top             =   450
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
         sTitulo         =   ""
      End
   End
   Begin VB.Frame fraFechas 
      Caption         =   "Rango de Fechas"
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
      Height          =   975
      Left            =   5835
      TabIndex        =   18
      Top             =   165
      Width           =   1980
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   300
         Left            =   540
         TabIndex        =   5
         Top             =   585
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   300
         Left            =   570
         TabIndex        =   4
         Top             =   240
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Del "
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
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label lblFecFin 
         Caption         =   "Al"
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
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   615
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdAsientos 
      Caption         =   "&Asientos"
      Height          =   345
      Left            =   6810
      TabIndex        =   12
      ToolTipText     =   "Imprime Asientos Contables"
      Top             =   5145
      Width           =   1395
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9765
      TabIndex        =   7
      ToolTipText     =   "Procesa según Rango de Fechas y Opciones"
      Top             =   840
      Width           =   1365
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Planilla"
      Height          =   345
      Left            =   8280
      TabIndex        =   13
      ToolTipText     =   "Imprime Planilla de Provisión"
      Top             =   5145
      Width           =   1395
   End
   Begin MSComctlLib.ListView lvProvis 
      Height          =   3045
      Left            =   180
      TabIndex        =   10
      Top             =   1245
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   5371
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgRec"
      SmallIcons      =   "imgRec"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nro. Mov."
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Documento"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Persona"
         Object.Width           =   5116
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "cMovDesc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Importe"
         Object.Width           =   2222
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "cCodPers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "nMovNro"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "cDocTpo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "cDocNro"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Estado"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "OpeCod"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Fecha Reg."
         Object.Width           =   2011
      EndProperty
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   555
      Left            =   195
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   4365
      Width           =   10980
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   9765
      TabIndex        =   14
      Top             =   5145
      Width           =   1395
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   195
      Left            =   1785
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   16
      Top             =   5700
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
            MinWidth        =   2999
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   15875
            MinWidth        =   15875
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Caption         =   "Imprimir"
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
      Height          =   540
      Left            =   210
      TabIndex        =   17
      Top             =   5010
      Width           =   1950
      Begin VB.OptionButton OptS 
         Caption         =   "Todo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   210
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton OptS 
         Caption         =   "Selección"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   9
         Top             =   210
         Width           =   1065
      End
   End
   Begin MSComctlLib.ImageList imgRec 
      Left            =   75
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRepOpeDiaUsuario.frx":0342
            Key             =   "recibo"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   60
      TabIndex        =   21
      Top             =   -15
      Width           =   11250
      Begin VB.Line Line1 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   11085
         X2              =   7815
         Y1              =   735
         Y2              =   735
      End
   End
End
Attribute VB_Name = "frmRepOpeDiaUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim sSql  As String
Dim lSalir As Boolean
Dim lMN As Boolean
Dim oCon As DConecta
Dim lsOpeFiltro As String

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Public Sub Inicio(psOpeFiltro As String)
lsOpeFiltro = psOpeFiltro
Me.Show 1
End Sub

Private Function CabeceraRepo(ByRef nLin As Integer) As String
Dim sTit    As String
If Left(Me.cboUsuario.Text, 4) = "XXXX" Then
   sTit = " P L A N I L L A   G E N E R A L   D E   O P E R A C I O N E S   D I A R I A S "
Else
   sTit = " P L A N I L L A   D E   U S U A R I O   D E   O P E R A C I O N E S   D I A R I A S "
End If
If nLin > gnLinPage - 4 Then
   xlHoja1.Cells(1, 1) = gsInstCmac
   xlHoja1.Cells(1, 6) = gdFecSis & " - " & Format(Time, "hh:mm:ss")
   xlHoja1.Cells(2, 2) = sTit
   xlHoja1.Cells(3, 2) = " M O N E D A   " & IIf(Not lMN, "E X T R A N J E R A ", "N A C I O N A L ")
   xlHoja1.Cells(4, 2) = "( DEL " & txtFechaDel & " AL " & txtFechaAl & ")"
   xlHoja1.Range("B2:E2").MergeCells = True
   xlHoja1.Range("B3:E3").MergeCells = True
   xlHoja1.Range("B4:E4").MergeCells = True
   xlHoja1.Range("B2:E4").HorizontalAlignment = xlCenter
   xlHoja1.Range("A1:F4").Font.Bold = True
   xlHoja1.Cells(6, 1) = "Item"
   xlHoja1.Cells(6, 2) = "Fecha de"
   xlHoja1.Cells(6, 3) = "Comprobante"
   xlHoja1.Cells(6, 4) = "Persona"
   xlHoja1.Cells(6, 5) = "Detalle"
   xlHoja1.Cells(6, 6) = "Importe " & IIf(lMN, "M.N.", "M.E.")
   xlHoja1.Cells(6, 7) = "Estado"
   xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 1)).ColumnWidth = 8
   xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 2)).ColumnWidth = 26
   xlHoja1.Range(xlHoja1.Cells(6, 3), xlHoja1.Cells(6, 3)).ColumnWidth = 15
   xlHoja1.Range(xlHoja1.Cells(6, 4), xlHoja1.Cells(6, 4)).ColumnWidth = 32
   xlHoja1.Range(xlHoja1.Cells(6, 5), xlHoja1.Cells(6, 5)).ColumnWidth = 40
   xlHoja1.Range(xlHoja1.Cells(6, 6), xlHoja1.Cells(6, 6)).ColumnWidth = 12
   xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(6, 7)).Font.Bold = True
   ExcelCuadro xlHoja1, 1, 6, 7, 6, True, True
   nLin = 7
End If
End Function

Private Sub GeneraListado(Optional psCodUser As String = "", Optional psOpeCod As String = "")
Dim sCond As String
Dim nItem As Integer
Dim lvItem As ListItem
Dim nImporte As Currency, nTipCambio As Currency
Dim sTpoDoc  As String
Dim rsDoc As ADODB.Recordset
Dim lsFiltro As String
Dim lsOperaciones As String
Dim lssqlAdd As String

lsFiltro = "RIGHT(a.cMovNro,4) IN ('" & psCodUser & "') "
If Not psOpeCod = "" Then
   sSql = "SELECT cOpeCod, nOpeNiv FROM OpeTpo WHERE cOpeCod = '" & psOpeCod & "' "
   Set rs = oCon.CargaRecordSet(sSql)
   sSql = "SELECT cOpeCod, nOpeNiv FROM OpeTpo WHERE cOpeCod > '" & psOpeCod & "' and nOpeNiv = " & rs!nOpeNiv
   Set rs = oCon.CargaRecordSet(sSql)
   If Not rs.EOF Then
      sSql = "SELECT cOpeCod FROM OpeTpo WHERE cOpeCod > '" & psOpeCod & "' and cOpeCod < '" & rs!cOpeCod & "'"
      Set rs = oCon.CargaRecordSet(sSql)
      If Not rs.EOF Then
         lsOperaciones = RSMuestraLista(rs)
      Else
         lsOperaciones = "'" & psOpeCod & "'"
      End If
   Else
      lsOperaciones = "'" & psOpeCod & "'"
   End If
   lsFiltro = lsFiltro & " and a.cOpeCod IN ( " & lsOperaciones & " ) "
End If

'No se quiere que aparezcan
'lssqlAdd = "'" & gCGArendirCtaSolMN & "','" & gCGArendirCtaSolME & "','" & gCGArendirCtaSustMN & "','" & gCGArendirCtaSustME & "','" & gCGArendirViatSolMN _
         & "','" & gCGArendirViatSolME & "','" & gCGArendirViatSustMN & "','" & gCGArendirViatSustME & "','" & gCGArendirViatAmpMN & "','" & gCGArendirViatAmpME _
         & "','" & gCHArendirCtaSolMN & "','" & gCHArendirCtaSolME & "','" & gCHArendirCtaSustMN & "','" & gCHArendirCtaSustME & "','" & gCHEgreDirectoSolMN & "','" & gCHEgreDirectoSolME & "'"
         
'Modificado PASI_VAPI_20150604************************************************************
'sSql = " SELECT DISTINCT Left(a.cMovNro,8) as cFecha, a.nMovNro, ISNULL(e.cPersNombre,ISNULL(p.cPersNombre,'')) cNomPers, a.cMovDesc, ISNULL(d.cPersCod,ISNULL(p.cPersCod,'')) cPersCod, "
'sSql = sSql & "        RIGHT(a.cMovNro,4), a.cMovNro, a.cOpeCod, a.nMovFlag, a.nMovEstado, SUM(" & IIf(lMN, "c.nMovImporte ", "isnull(ME.nMovMEImporte,0)") & ") as nDocImporte "
'sSql = sSql & " FROM   Mov a JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(lMN, "", " JOIN MovME ME ON (ME.nMovNro=c.nMovNro and ME.nMovItem=c.nMovItem)")
'sSql = sSql & "        LEFT  JOIN MovGasto d ON d.nMovNro = a.nMovNro LEFT JOIN Persona e ON e.cPersCod = d.cPersCod "
'sSql = sSql & "        LEFT  JOIN MovObjIF mif ON mif.nMovNro = c.nMovNro and mif.nMovItem = c.nMovItem LEFT JOIN Persona p ON p.cPersCod = mif.cPersCod "
'sSql = sSql & " WHERE  " & lsFiltro & " and (SubString(a.cOpeCod,3,1) = '" & Mid(gsOpeCod, 3, 1) & "' or  (a.cOpeCod in ('402581','402582') ))  and " & IIf(Me.chkEliminados.value = Unchecked, " not nMovFlag = " & gMovFlagEliminado & " and ", "")
'sSql = sSql & "        Left(a.cMovNro,8) >='" & Format(txtFechaDel, gsFormatoMovFecha) & "' and Left(a.cMovNro,8) <='" & Format(txtFechaAl, gsFormatoMovFecha) & "' and c.nMovImporte < 0 and a.nmovestado = '10'"
'If Len(Trim(lssqlAdd)) > 0 Then
'    sSql = sSql & " And a.cOpeCod Not In (" & lssqlAdd & ") "
'End If
'sSql = sSql & " GROUP BY Left(a.cMovNro,8), a.nMovNro, ISNULL(e.cPersNombre,ISNULL(p.cPersNombre,'')), a.cMovDesc, ISNULL(d.cPersCod,ISNULL(p.cPersCod,'')), RIGHT(a.cMovNro,4), a.cMovNro, a.cOpeCod, a.nMovFlag, a.nMovEstado "
'sSql = sSql & "  ORDER BY Left(a.cMovNro,8), RIGHT(a.cMovNro,4), a.cOpeCod, a.cMovNro "

sSql = "exec stp_sel_ListaOperacionesxUsuario '" & psCodUser & "','" & Format(CDate(txtFechaDel.Text), "yyyyMMdd") & "','" & Format(CDate(txtFechaAl.Text), "yyyyMMdd") & "','" & Mid(gsOpeCod, 3, 1) & "','" & Replace(lsOperaciones, "'", "") & "'," & CInt(Me.chkEliminados.value)
'END PASI_VAPI*************************************************************************
Set rs = oCon.CargaRecordSet(sSql)
lSalir = False
If RSVacio(rs) Then
   MsgBox "No existen Operaciones de Usuario " & psCodUser, vbInformation, "¡Aviso!"
   lSalir = True
   Exit Sub
End If
lvProvis.ListItems.Clear

prg.Visible = True
prg.Min = 0
prg.Max = rs.RecordCount
Dim lsEst As String

Do While Not rs.EOF
   prg.value = rs.Bookmark
   Status.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
   nItem = nItem + 1
   Set lvItem = lvProvis.ListItems.Add(, , Format(nItem, "000"))
   lvItem.SmallIcon = 1
   
   sSql = "SELECT md.nDocTpo, cDocNro, dDocFecha, cDocAbrev FROM MovDoc md JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
        & "WHERE md.nMovNRo = " & rs!nMovNro
   Set rsDoc = oCon.CargaRecordSet(sSql)
   If Not rsDoc.EOF Then
      lvItem.SubItems(2) = Mid(rsDoc!cDocAbrev & Space(3), 1, 3) & " " & rsDoc!cDocNro
      lvItem.SubItems(8) = rsDoc!nDocTpo
      lvItem.SubItems(9) = rsDoc!cDocNro
   End If
   RSClose rsDoc
   lvItem.SubItems(1) = rs!cMovNro
   lvItem.SubItems(3) = rs!cNomPers
   lvItem.SubItems(4) = rs!cMovDesc
   lvItem.SubItems(5) = Format(Abs(rs!nDocImporte), gsFormatoNumeroView)
   lvItem.SubItems(6) = rs!cPersCod
   lvItem.SubItems(7) = rs!nMovNro
   lvItem.SubItems(11) = rs!cOpeCod
   lvItem.SubItems(12) = rs!cFecha
   lvItem.ListSubItems(5).Bold = True
   If Not lMN Then
      lvItem.ListSubItems(5).ForeColor = gsColorME
   Else
      lvItem.ListSubItems(5).ForeColor = gsColorMN
   End If
   lsEst = ""
   Select Case rs!nMovFlag
      Case 0: lsEst = "VIG"
      Case 1: lsEst = "ELI"
      Case 2: lsEst = "EXT"
      Case 3: lsEst = "DEX"
      Case 5: lsEst = "MOD"
   End Select
   Select Case rs!nMovEstado
      Case 10: lsEst = lsEst & "-CNT"
      Case 11: lsEst = lsEst & "-PND"
      Case 13: lsEst = lsEst & "-NOC"
   End Select
   lvItem.SubItems(10) = lsEst
   rs.MoveNext
Loop
rs.Close
lvProvis.ListItems(1).Selected = True
txtMovDesc = lvProvis.ListItems(1).SubItems(4)
prg.Visible = False
End Sub

Private Sub cboUsuario_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtFechaDel.SetFocus
End If
End Sub

Private Sub cmdAsientos_Click()
Dim N As Integer
Dim lOk As Boolean
Dim lsImpre As String
Dim sMovs As String
Dim oContImp As NContImprimir

If lvProvis.ListItems.Count = 0 Then
    MsgBox "No Hay elementos en la lista", vbInformation, "Aviso"
    Exit Sub
End If

Set oContImp = New NContImprimir
prg.Min = 0
prg.Max = lvProvis.ListItems.Count
prg.Visible = True
lsImpre = ""
For N = 1 To lvProvis.ListItems.Count
   lOk = True
   If OptS(1).value Then
      lOk = IIf(lvProvis.ListItems(N).Selected, True, False)
   End If
   prg.value = N
   Status.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
   If lOk Then
      sMovs = sMovs & ",'" & lvProvis.ListItems(N).SubItems(1) & "'"
   End If
Next
If sMovs <> "" Then
    sMovs = Mid(sMovs, 2, Len(sMovs))
    lsImpre = ImprimeAsientosContables(sMovs, prg, Status, "( DEL " & txtFechaDel & " AL " & txtFechaAl & ")")
    EnviaPrevio lsImpre, "ASIENTOS DE OPERACIONES DIARIAS", gnLinPage, False
End If
prg.Visible = False
lvProvis.SetFocus
End Sub

Private Sub cmdImprimir_Click()
Dim N As Integer
Dim nLin As Integer, P As Integer
Dim nLinIni As Integer
Dim lsSumaOpe As String
Dim lsSumaUsu As String
Dim nTot As Currency
Dim lOk As Boolean
Dim lsImpre As String
Dim lsUsuario As String
Dim lsFecha   As String
Dim nU As Integer

If lvProvis.ListItems.Count = 0 Then
   MsgBox "No existen elementos que Imprimir...!", vbInformation, "Error"
   Exit Sub
End If
nLin = gnLinPage
nTot = 0
prg.Min = 0
prg.Max = lvProvis.ListItems.Count
prg.Visible = True
Me.Enabled = False
lsUsuario = ""

Dim lsArchivo As String
Dim lbLibroOpen As Boolean
Dim lsOpeCod    As String
Dim oOpe        As New DOperacion
Dim nOpeCnt     As Integer
Dim nUsuCnt     As Integer
Dim nOpeCntI    As Integer

On Error GoTo ErrImprime
lsFecha = lvProvis.ListItems(1).SubItems(12)

lsArchivo = App.path & "\Spooler\RP_OPEUSU_" & Format(Now, "yyyymmddhhmmss") & ".xls"
lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
If lbLibroOpen Then
   Set xlHoja1 = xlLibro.Worksheets(1)
   ExcelAddHoja lsFecha, xlLibro, xlHoja1
End If
xlHoja1.PageSetup.Zoom = 68
CabeceraRepo nLin
nLinIni = nLin
nOpeCnt = 0
nOpeCntI = 0
For N = 1 To lvProvis.ListItems.Count
   If lsFecha <> lvProvis.ListItems(N).SubItems(12) Then
      'Ultimo Grupo Operacion
      xlHoja1.Cells(nLin, 2) = "Total " & gsOpeDesc
      xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=SUM(F" & nLinIni & ":F" & nLin - 1 & ")"
      lsSumaOpe = lsSumaOpe & "+" & "F" & nLin
      nLin = nLin + 1
      nOpeCntI = 0
      
      'Ultimo usuario
      If lsSumaOpe <> "" Then
         xlHoja1.Cells(nLin, 1) = "TOTAL " & lsUsuario
         xlHoja1.Cells(nLin, 3) = nUsuCnt & " operaciones"
         nUsuCnt = 0
         xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=" & lsSumaOpe
         nLin = nLin + 1
      End If
      
      xlHoja1.Cells(nLin, 1) = "TOTAL OPERACIONES "
      xlHoja1.Cells(nLin, 3) = nOpeCnt & " operaciones"
      nOpeCnt = 0
      nOpeCntI = 0
      
      If lsSumaUsu <> "" Then
         xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=" & lsSumaUsu
      ElseIf lsSumaOpe <> "" Then
         xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=" & lsSumaOpe
      Else
         xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=SUM(F7:F" & nLin - 1 & ")"
      End If
      xlHoja1.Range(xlHoja1.Cells(nLin, 1), xlHoja1.Cells(nLin, 7)).Font.Bold = True
      xlHoja1.Range(xlHoja1.Cells(7, 3), xlHoja1.Cells(nLin, 7)).NumberFormat = "##,##0.00"
      
      ExcelCuadro xlHoja1, 1, 7, 7, nLin - 1, True, False
      ExcelCuadro xlHoja1, 1, nLin, 7, nLin, True, False
      
      lsFecha = lvProvis.ListItems(N).SubItems(12)
      ExcelAddHoja lsFecha, xlLibro, xlHoja1
      nLin = gnLinPage
      CabeceraRepo nLin
      nLinIni = nLin
      lsOpeCod = ""
      lsUsuario = ""
      nOpeCnt = 0
      nUsuCnt = 0
   End If
   lOk = True
   If OptS(1).value Then
      lOk = IIf(lvProvis.ListItems(N).Selected, True, False)
   End If
   prg.value = N
   Status.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
   If lOk Then
      With lvProvis.ListItems(N)
         If lsUsuario <> Right(.SubItems(1), 4) Then
            If lsOpeCod <> "" Then
               xlHoja1.Cells(nLin, 2) = "Total " & gsOpeDesc
               xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=SUM(F" & nLinIni & ":F" & nLin - 1 & ")"
               lsSumaOpe = lsSumaOpe & "+" & "F" & nLin
               nOpeCntI = 0
               nLin = nLin + 1
            End If
            If lsUsuario <> "" Then
               xlHoja1.Cells(nLin, 1) = "TOTAL " & lsUsuario
               xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=" & lsSumaOpe
               lsSumaOpe = ""
               lsSumaUsu = lsSumaUsu & "+" & "F" & nLin
               xlHoja1.Cells(nLin, 3) = nUsuCnt & " operaciones"
               nUsuCnt = 0
               nLin = nLin + 1
            End If
            lsUsuario = Right(.SubItems(1), 4)
            For nU = 0 To cboUsuario.ListCount - 2
               If Left(cboUsuario.List(nU), 4) = Right(.SubItems(1), 4) Then
                  nLin = nLin + 1
                  xlHoja1.Cells(nLin, 1) = "USUARIO : " & cboUsuario.List(nU)
                  xlHoja1.Range(xlHoja1.Cells(nLin, 1), xlHoja1.Cells(nLin, 1)).Font.Bold = True
                  nLin = nLin + 1
                  lsOpeCod = ""
               End If
            Next
         End If
         If lsOpeCod <> .SubItems(11) Then
            If lsOpeCod <> "" Then
               xlHoja1.Cells(nLin, 2) = "Total " & gsOpeDesc
               xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=SUM(F" & nLinIni & ":F" & nLin - 1 & ")"
               lsSumaOpe = lsSumaOpe & "+" & "F" & nLin
               nOpeCntI = 0
               nLin = nLin + 1
            End If
            nLin = nLin + 1
            gsOpeDesc = oOpe.GetOperacionDesc(.SubItems(11))
            xlHoja1.Cells(nLin, 2) = "OPERACION : " & .SubItems(11) & " - " & gsOpeDesc
            xlHoja1.Range(xlHoja1.Cells(nLin, 2), xlHoja1.Cells(nLin, 2)).Font.Bold = True
            nLinIni = nLin
            lsOpeCod = .SubItems(11)
            nLin = nLin + 1
         End If
         nOpeCnt = nOpeCnt + 1
         nOpeCntI = nOpeCntI + 1
         nUsuCnt = nUsuCnt + 1
         xlHoja1.Cells(nLin, 1) = Format(nOpeCntI, "000")
         xlHoja1.Cells(nLin, 2) = .SubItems(1)
         xlHoja1.Cells(nLin, 3) = .SubItems(2)
         xlHoja1.Cells(nLin, 4) = .SubItems(3)
         xlHoja1.Cells(nLin, 5) = Replace(Replace(.SubItems(4), Chr(13), ""), Chr(10), "")
         xlHoja1.Cells(nLin, 6) = .SubItems(5)
         xlHoja1.Cells(nLin, 7) = .SubItems(10)
         
      End With
      nLin = nLin + 1
   End If
Next
'Ultimo Grupo Operacion
xlHoja1.Cells(nLin, 2) = "Total " & gsOpeDesc
xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=SUM(F" & nLinIni & ":F" & nLin - 1 & ")"
lsSumaOpe = lsSumaOpe & "+" & "F" & nLin
nLin = nLin + 1

'Ultimo usuario
If lsSumaOpe <> "" Then
   xlHoja1.Cells(nLin, 1) = "TOTAL " & lsUsuario
   xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=" & lsSumaOpe
   xlHoja1.Cells(nLin, 3) = nUsuCnt & " operaciones"
   nUsuCnt = 0
   nLin = nLin + 1
End If
xlHoja1.Cells(nLin, 1) = "TOTAL OPERACIONES "
xlHoja1.Cells(nLin, 3) = nOpeCnt & " operaciones"

If lsSumaUsu <> "" Then
   xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=" & lsSumaUsu
ElseIf lsSumaOpe <> "" Then
   xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=" & lsSumaOpe
Else
   xlHoja1.Range(xlHoja1.Cells(nLin, 6), xlHoja1.Cells(nLin, 6)).Formula = "=SUM(F7:F" & nLin - 1 & ")"
End If
xlHoja1.Range(xlHoja1.Cells(nLin, 1), xlHoja1.Cells(nLin, 7)).Font.Bold = True
xlHoja1.Range(xlHoja1.Cells(7, 3), xlHoja1.Cells(nLin, 7)).NumberFormat = "##,##0.00"

ExcelCuadro xlHoja1, 1, 7, 7, nLin - 1, True, False
ExcelCuadro xlHoja1, 1, nLin, 7, nLin, True, False
ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
CargaArchivo lsArchivo, App.path & "\spooler"
prg.Visible = False
Me.Enabled = True
lvProvis.SetFocus
Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
   ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
   Me.Enabled = True
End Sub

Private Sub cmdProcesar_Click()
Dim lsUsuario As String
Dim nU As Integer
Me.lvProvis.ListItems.Clear
If cboUsuario.ListIndex >= 0 Then
   If Left(cboUsuario, 4) = "XXXX" Then
      lsUsuario = Left(Me.cboUsuario.List(0), 4)
      For nU = 1 To cboUsuario.ListCount - 1
         lsUsuario = lsUsuario & "','" & Left(Me.cboUsuario.List(nU), 4)
      Next
   Else
      lsUsuario = Left(Me.cboUsuario, 4)
   End If
End If
If Me.tTipo.Tab = 0 Then
   If cboUsuario.ListIndex >= 0 Then
      GeneraListado lsUsuario
   Else
      MsgBox "Seleccione un Usuario", vbInformation, "¡Aviso!"
      cboUsuario.SetFocus
   End If
Else
   If txtOpecod = "" Then
      MsgBox "Seleccione una Operación", vbInformation, "¡Aviso!"
      txtOpecod.SetFocus
   Else
      GeneraListado lsUsuario, txtOpecod
   End If
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If lSalir Then
   Unload Me
End If
End Sub

Private Sub Form_Load()
lSalir = False
CentraForm Me
txtFechaDel = gdFecSis
txtFechaAl = gdFecSis
lMN = True
If Mid(gsOpeCod, 3, 1) = gMonedaExtranjera Then
   lMN = False
End If
Set oCon = New DConecta
oCon.AbreConexion
'sSql = "Select Distinct rh.cUser, P.cPersNombre "
'sSql = sSql & " From rrhh rh JOIN Persona P ON P.cPersCod = RH.cPersCod, "
'sSql = sSql & "      OpeObj O "
'sSql = sSql & " Where "
'sSql = sSql & "      (O.cOpeCod = '" & gsOpeCod & "' and cAreaCodActual LIKE o.cOpeObjFiltro )"
'sSql = sSql & "       OR (rh.cUser IN ('RPLM','YOAQ','JYCA','TATA','MMAP','MSPG','PAGT','SAHM','NAMG'))"
'sSql = sSql & " Order By cUser "

'Modificacion

'sSql = "Select Distinct rh.cUser, P.cPersNombre "
'sSql = sSql & " From rrhh rh JOIN Persona P ON P.cPersCod = RH.cPersCod "
'sSql = sSql & " Where rh.cuser<>'XXXX' and rh.nRHEstado='201'"
'sSql = sSql & " Order By cUser "

sSql = "Select  rh.cUser, P.cPersNombre"
sSql = sSql & " From rrhh rh JOIN Persona P ON P.cPersCod = RH.cPersCod"
sSql = sSql & " Where (rh.cuser not in ('XXXX','')  and rh.nRHEstado='201') or"
sSql = sSql & " (rh.nRHEstado in ('803','802','801')"
sSql = sSql & " and dcese between getdate()-7 and getdate()+7)"
sSql = sSql & " Order By cUser"

Set rs = oCon.CargaRecordSet(sSql)
RSLlenaCombo rs, cboUsuario, , , False
cboUsuario.AddItem "XXXX  TODOS LOS USUARIOS"
RSClose rs
Dim clsOpe As New DOperacion
lsOpeFiltro = "4[012345]" & Mid(gsOpeCod, 3, 1)
txtOpecod.psRaiz = "Operaciones"
txtOpecod.rs = clsOpe.CargaOpeTpo(lsOpeFiltro, True, , , 1)

End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub lvProvis_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvProvis.SortKey = ColumnHeader.Index - 1
lvProvis.Sorted = True
End Sub

Private Sub lvProvis_KeyUp(KeyCode As Integer, Shift As Integer)
Dim nPos As Variant
If lvProvis.ListItems.Count > 0 Then
   nPos = lvProvis.SelectedItem.Index
   txtMovDesc = lvProvis.ListItems(nPos).SubItems(4)
End If
End Sub

Private Sub lvProvis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nPos As Variant
If lvProvis.ListItems.Count > 0 Then
   nPos = lvProvis.SelectedItem.Index
   txtMovDesc = lvProvis.ListItems(nPos).SubItems(4)
End If
End Sub
 

Private Sub txtFechaAl_GotFocus()
fEnfoque txtFechaAl
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaAl.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      txtFechaAl.SetFocus
   End If
   cmdProcesar.SetFocus
End If
End Sub

Private Sub txtFechaAl_Validate(Cancel As Boolean)
If ValidaFecha(txtFechaAl.Text) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Error"
   Cancel = True
End If
End Sub

Private Sub txtFechaDel_GotFocus()
fEnfoque txtFechaDel
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If ValidaFecha(txtFechaDel.Text) <> "" Then
      MsgBox "Fecha no válida...!", vbInformation, "Error"
      txtFechaDel.SetFocus
   End If
   txtFechaAl.SetFocus
End If
End Sub

Private Sub txtFechaDel_Validate(Cancel As Boolean)
If ValidaFecha(txtFechaDel.Text) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Error"
   Cancel = True
End If
End Sub

Private Sub txtOpeCod_EmiteDatos()
txtOpeDes = txtOpecod.psDescripcion
If txtOpeDes <> "" Then
   txtFechaDel.SetFocus
End If
End Sub




