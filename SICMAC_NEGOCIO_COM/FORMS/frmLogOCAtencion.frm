VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLogOCAtencion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logística: "
   ClientHeight    =   5880
   ClientLeft      =   705
   ClientTop       =   2205
   ClientWidth     =   11970
   Icon            =   "frmLogOCAtencion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraRefrescar 
      Appearance      =   0  'Flat
      Caption         =   "Actualizar "
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
      Height          =   735
      Left            =   120
      TabIndex        =   20
      Top             =   30
      Width           =   7935
      Begin VB.ComboBox cmbTipoOC 
         Height          =   315
         ItemData        =   "frmLogOCAtencion.frx":030A
         Left            =   4560
         List            =   "frmLogOCAtencion.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   285
         Width           =   1335
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   345
         Left            =   6600
         TabIndex        =   25
         Top             =   270
         Width           =   1230
      End
      Begin MSMask.MaskEdBox mskIni 
         Height          =   300
         Left            =   885
         TabIndex        =   21
         Top             =   285
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   300
         Left            =   3060
         TabIndex        =   22
         Top             =   285
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFinal 
         Caption         =   "Final"
         Height          =   195
         Left            =   2355
         TabIndex        =   24
         Top             =   330
         Width           =   525
      End
      Begin VB.Label lblInicial 
         Caption         =   "Inicial"
         Height          =   195
         Left            =   210
         TabIndex        =   23
         Top             =   330
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imp Excel"
      Height          =   345
      Left            =   8175
      TabIndex        =   18
      Top             =   5460
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame fraPeriodo 
      Appearance      =   0  'Flat
      Caption         =   "Periodo del Saldo Anual"
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
      Height          =   795
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Visible         =   0   'False
      Width           =   4875
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Anual"
         Height          =   375
         Index           =   12
         Left            =   3630
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Semestral"
         Height          =   375
         Index           =   6
         Left            =   2460
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Trimestral"
         Height          =   375
         Index           =   3
         Left            =   1290
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   300
         Width           =   1155
      End
      Begin VB.OptionButton optPeriodo 
         Caption         =   "Mensual"
         Height          =   375
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg 
      Height          =   4215
      Left            =   105
      TabIndex        =   0
      Top             =   825
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   3
      Cols            =   12
      FixedRows       =   2
      BackColorBkg    =   -2147483643
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   2
      GridLinesFixed  =   1
      GridLinesUnpopulated=   3
      Appearance      =   0
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
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "&Eliminar"
      Height          =   345
      Left            =   8175
      TabIndex        =   1
      Top             =   5460
      Width           =   1185
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   9405
      TabIndex        =   2
      Top             =   5460
      Width           =   1185
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   10635
      TabIndex        =   3
      Top             =   5460
      Width           =   1185
   End
   Begin VB.TextBox txtMovNro 
      Height          =   315
      Left            =   6480
      TabIndex        =   4
      Top             =   5400
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.TextBox txtTot 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   10250
         SubFormatType   =   1
      EndProperty
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
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   8835
      TabIndex        =   5
      Top             =   5055
      Width           =   1290
   End
   Begin VB.Frame fraFechas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Fechas"
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Visible         =   0   'False
      Width           =   6225
      Begin VB.ComboBox cmbTipoOCR 
         Height          =   315
         ItemData        =   "frmLogOCAtencion.frx":03BF
         Left            =   3240
         List            =   "frmLogOCAtencion.frx":03CC
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   315
         Left            =   4920
         TabIndex        =   12
         Top             =   420
         Width           =   1185
      End
      Begin MSMask.MaskEdBox mskFecIni 
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   435
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFecFin 
         Height          =   285
         Left            =   1725
         TabIndex        =   9
         Top             =   435
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   1350
         X2              =   1590
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Label lblFin 
         Caption         =   "Al"
         Height          =   225
         Left            =   1725
         TabIndex        =   11
         Top             =   225
         Width           =   825
      End
      Begin VB.Label lblFecIni 
         Caption         =   "Del"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   225
         Width           =   825
      End
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   240
      Left            =   6360
      SizeMode        =   1  'Stretch
      TabIndex        =   19
      Top             =   5160
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7650
      TabIndex        =   6
      Top             =   5115
      Width           =   885
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   315
      Left            =   7380
      Top             =   5055
      Width           =   2745
   End
End
Attribute VB_Name = "frmLogOCAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql  As String
Dim rs As ADODB.Recordset
Dim lSalir As Boolean
Dim sCtaProvis As String
Dim sDocTpoOC As String
Dim sDocDesc As String
Dim txtImporte As Currency
Dim lbBienes As Boolean
Dim lbPresu As Boolean
Dim lbModifica As Boolean
Dim lbImprime As Boolean
Dim lsOpeCod As String

Dim lbReporteFechas As Boolean

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet


Public Sub InicioRep(pbBienes As Boolean, psOpeCod As String, Optional pbPresu As Boolean = False, Optional pbModifica As Boolean = False, Optional pbImprime As Boolean = False)
    
    lbBienes = pbBienes
    lbPresu = pbPresu
    lbModifica = pbModifica
    lbImprime = pbImprime
    lbReporteFechas = True
    lsOpeCod = psOpeCod
    Me.Show 1
End Sub

Public Sub Inicio(pbBienes As Boolean, psOpeCod As String, Optional pbPresu As Boolean = False, Optional pbModifica As Boolean = False, Optional pbImprime As Boolean = False, Optional pbRangoFec As Boolean = False)
    
    lbBienes = pbBienes
    lbPresu = pbPresu
    lbModifica = pbModifica
    lbImprime = pbImprime
    lbReporteFechas = pbRangoFec
    lsOpeCod = psOpeCod
    Me.Show 1
    
End Sub

Private Sub FormatoOCompra(Optional STIPOoc As String = "")
fg.Cols = 13
fg.TextMatrix(0, 0) = " "
fg.TextMatrix(1, 0) = " "
fg.TextMatrix(0, 1) = "Documento"
fg.TextMatrix(0, 2) = "Documento"
fg.TextMatrix(0, 3) = "Documento"
fg.TextMatrix(0, 4) = "Documento"

fg.TextMatrix(1, 1) = "Tipo"
fg.TextMatrix(1, 2) = "Número"

fg.TextMatrix(1, 3) = "NºOC Di/Pr"
fg.TextMatrix(1, 4) = "Fecha"

fg.TextMatrix(0, 5) = "Proveedor"
fg.TextMatrix(1, 5) = "Proveedor"

fg.TextMatrix(0, 6) = "Importe"
fg.TextMatrix(1, 6) = "Importe"
fg.TextMatrix(0, 7) = "Observaciones"
fg.TextMatrix(1, 7) = "Observaciones"
fg.TextMatrix(1, 8) = "cMovNro"
fg.TextMatrix(1, 9) = "nImporte"

fg.TextMatrix(1, 10) = "Saldo"
fg.TextMatrix(1, 11) = "Estado"
fg.TextMatrix(1, 12) = "Monto ($)"

fg.RowHeight(-1) = 285
fg.ColWidth(0) = 400
fg.ColWidth(1) = 500
fg.ColWidth(2) = 1200
fg.ColWidth(3) = 1200
fg.ColWidth(4) = 1100

fg.ColWidth(5) = 3200
fg.ColWidth(6) = 1200
fg.ColWidth(7) = 3770
fg.ColWidth(8) = 0
fg.ColWidth(9) = 0
If lbPresu Then
   fg.ColWidth(10) = 0
   fg.ColWidth(11) = 0
Else
   fg.ColWidth(10) = 1200
   fg.ColWidth(11) = 1700
End If

If lbReporteFechas Then
    fg.ColWidth(12) = 1700
Else
    fg.ColWidth(12) = 0
End If

fg.MergeCells = flexMergeRestrictColumns
fg.MergeCol(0) = True
fg.MergeCol(1) = True
fg.MergeCol(2) = True
fg.MergeCol(3) = True
fg.MergeCol(4) = True
fg.MergeCol(5) = True
fg.MergeCol(6) = True
fg.MergeCol(7) = True

fg.MergeRow(0) = True
fg.MergeRow(1) = True
fg.RowHeight(0) = 200
fg.RowHeight(1) = 200
fg.ColAlignmentFixed(-1) = flexAlignCenterCenter
fg.ColAlignment(1) = flexAlignCenterCenter
fg.ColAlignment(3) = flexAlignCenterCenter
fg.ColAlignment(6) = flexAlignLeftCenter
End Sub


Private Sub GetOCPendientes(Optional STIPOoc As String = "")
Dim lsCadSerAdm As String
Dim nItem As Integer
Dim nTot  As Currency
Dim oCon As DConecta

Dim sTipDoc As String
Dim STIPOC As String

Set oCon = New DConecta
oCon.AbreConexion

sSql = "Select '[' +  Rtrim(cNomSer) + '].' + RTrim(cDataBase) + '.dbo.' from servidor Where cCodAge = '11207' And cNroSer = '02'"
sSql = "Select RTrim(cDataBase) + '.dbo.' from servidor Where cCodAge = '11207' And cNroSer = '02'"
lsCadSerAdm = "" 'oCon.CargaRecordSet(sSQL).Fields(0)

 If lbBienes = True Then
    sTipDoc = "('" & LogTipoOC.gLogOCompraDirecta & "','" & LogTipoOC.gLogOCompraProceso & "'  )"
    Else
    sTipDoc = "('" & LogTipoOC.gLogOServicioDirecta & "','" & LogTipoOC.gLogOServicioProceso & "'  )"
 End If

If lbPresu Or lbModifica Or lbImprime Then

        If STIPOoc = "T" Then
       sSql = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN  " & sTipDoc & " ),''),   dd.cPersCod, " _
        & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
        & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
        & "     FROM Persona PE" _
        & "     WHERE PE.cPersCod = dd.cPersCod)" _
        & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, " & IIf(Mid(lsOpeCod, 3, 1) = 2 And gsCodCMAC = "102", "me.nMovMEImporte", "c.nMovImporte") & " * -1 as nDocImporte" _
        & " FROM Mov a" _
        & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
        & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
        & " JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(Mid(lsOpeCod, 3, 1) = 2 And gsCodCMAC = "102", " LEFT JOIN MovMe me ON me.nMovNro = c.nMovNro and me.nMovItem = c.nMovItem ", "") _
        & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
        & " WHERE  ((a.nMovEstado IN (" & IIf(lbPresu, gMovEstPresupPendiente, IIf(lbImprime, gMovEstPresupRechazado & "," & gMovEstPresupAceptado, gMovEstPresupPendiente & "," & gMovEstPresupAceptado)) & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
        & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
        & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
        & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro )" _
        & " ORDER BY b.cDocNro"
        Else

        sSql = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''),dd.cPersCod, " _
        & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = ( " _
        & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
        & "     FROM Persona PE" _
        & "     WHERE PE.cPersCod = dd.cPersCod)" _
        & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, " & IIf(Mid(lsOpeCod, 3, 1) = 2 And gsCodCMAC = "102", "me.nMovMEImporte", "c.nMovImporte") & " * -1 as nDocImporte" _
        & " FROM Mov a" _
        & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
        & " join movcotizac mc on a.nMovNro  = mc.nMovNro " _
        & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
        & " JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(Mid(lsOpeCod, 3, 1) = 2 And gsCodCMAC = "102", " LEFT JOIN MovMe me ON me.nMovNro = c.nMovNro and me.nMovItem = c.nMovItem ", "") _
        & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
        & " WHERE  ((a.nMovEstado IN (" & IIf(lbPresu, gMovEstPresupPendiente, IIf(lbImprime, gMovEstPresupRechazado & "," & gMovEstPresupAceptado, gMovEstPresupPendiente & "," & gMovEstPresupAceptado)) & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
        & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
        & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
        & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro And cAgeCodRef = '') and mc.ctipoOC ='" & STIPOoc & "'" _
        & " ORDER BY b.cDocNro"


        End If

'   sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
        & " cNomPers = (" _
        & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
        & "     FROM Persona PE" _
        & "     WHERE PE.cPersCod = dd.cPersCod)" _
        & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
        & " FROM Mov a" _
        & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
        & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
        & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
        & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
        & " WHERE  ((a.nMovEstado IN (" & IIf(lbPresu, gMovEstPresupPendiente, IIf(lbImprime, gMovEstPresupRechazado & "," & gMovEstPresupAceptado, gMovEstPresupPendiente & "," & gMovEstPresupAceptado)) & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
        & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
        & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
        & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE (h.cagecodref is null or len(h.cagecodref)=0) and  m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro)" _
        & " ORDER BY b.cDocNro"

        'Agregado: (h.cagecodref is null or len(h.cagecodref)=0) and

    If lbImprime Then
        If gbBitCentral Then
            'Centralizado
'            If STIPOoc = "T" Then
'            sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''), cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "'" _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
'                 & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
'                 & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro)" _
'                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro) ORDER BY b.cDocNro"
           
            
            If STIPOoc = "T" Then

            sSql = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), b.cDocNro, dd.cPersCod, " _
                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a" _
                 & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) " _
                 & " And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "'  And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE Left(m.cOpeCod,2) Not in ('56','59') And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro)  " _
                 & " ORDER BY b.cDocNro"
            Else
'            sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro join movcotizac mc on a.nMovNro  = mc.nMovNro " _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
'                 & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
'                 & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro) and mc.ctipoOC ='" & STIPOoc & "' " _
'                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro join movcotizac mc on a.nMovNro  = mc.nMovNro " _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro) and mc.ctipoOC ='" & STIPOoc & "' ORDER BY b.cDocNro"

            
            sSql = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a" _
                 & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " join movcotizac mc on a.nMovNro  = mc.nMovNro " _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
                 & " And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "'  And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE Left(m.cOpeCod,2) Not in ('56','59') And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro)  " _
                 & " and mc.ctipoOC ='" & STIPOoc & "' " _
                 & " ORDER BY b.cDocNro"

            End If
        Else

            'Centralizado
            sSql = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
                 & " cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a" _
                 & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE a.cMovNro Between '" & Format(CDate(Me.mskIni.Text), gsFormatoMovFecha) & "' And '" & Format(CDate(Me.mskFin.Text) + 1, gsFormatoMovFecha) & "' And  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "'  And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE Left(m.cOpeCod,2) Not in ('56','59') And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro  And cAgeCodRef = '')  " _
                 & " ORDER BY b.cDocNro"

            'MsgBox "s"
            'Hibrido
            'sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
                 & " cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a" _
                 & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "'  And NOT EXISTS (SELECT h.cMovNro FROM  " & lsCadSerAdm & "MovRef h JOIN " & lsCadSerAdm & "Mov M ON m.cMovNro = h.cMovNro WHERE Left(m.cOpeCod,2) Not in ('56','59') And m.cMovFlag NOT IN ('X','Y','N') And h.cMovNroRef = a.cMovNro)  " _
                 & " ORDER BY b.cDocNro"
        End If
    ElseIf lbModifica Then
        If gbBitCentral Then
            'Centralizado

            If STIPOoc = "T" Then
            sSql = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''), cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "'" _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
                 & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
                 & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro)" _
                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro) ORDER BY b.cDocNro"
            Else
            sSql = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro join movcotizac mc on a.nMovNro  = mc.nMovNro " _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
                 & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
                 & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro) and mc.ctipoOC ='" & STIPOoc & "' " _
                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro join movcotizac mc on a.nMovNro  = mc.nMovNro " _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro) and mc.ctipoOC ='" & STIPOoc & "' ORDER BY b.cDocNro"

            End If



        Else

            'Centralizado
            sSql = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
                 & " cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM  Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
                 & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
                 & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE a.cMovNro Between '" & Format(CDate(Me.mskIni.Text), gsFormatoMovFecha) & "' And '" & Format(CDate(Me.mskFin.Text) + 1, gsFormatoMovFecha) & "' And m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro And IsNull(h.cAgeCodRef,'') = '')   " _
                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
                 & " cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  a.cMovNro Between '" & Format(CDate(Me.mskIni.Text), gsFormatoMovFecha) & "' And '" & Format(CDate(Me.mskFin.Text) + 1, gsFormatoMovFecha) & "' And ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro And IsNull(h.cAgeCodRef,'') = '') ORDER BY b.cDocNro"

            'Hibrido
            'sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
                 & " cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, ISNULL(me.nMovMEImporte, c.nMovImporte) * -1 as nDocImporte" _
                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro LEFT JOIN MovMe me ON me.nMovNro = c.nMovNro and me.nMovItem = c.nMovItem " _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
                 & " NOT EXISTS (SELECT h.cMovNro FROM  " & lsCadSerAdm & "MovRef h" _
                 & " JOIN " & lsCadSerAdm & "Mov M ON m.cMovNro = h.cMovNro WHERE m.cMovFlag NOT IN ('X','E','N') and h.cMovNroRef = a.cMovNro)" _
                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
                 & " cNomPers = (" _
                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
                 & "     FROM Persona PE" _
                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro) ORDER BY b.cDocNro"
        End If
    End If

   'SELECT b.dDocFecha, g.cDocAbrev, b.cDocTpo, b.cDocNro, cNomPers = (SELECT cNomPers FROM " & gcCentralPers & "persona WHERE cCodPers = substring(d.cObjetoCod,3,10)), " _
     & "       a.cMovDesc, d.cObjetoCod, " _
     & "       a.cMovNro, a.cMovEstado, a.cMovFlag, c.cCtaContCod, " & IIf(GSSIMBOLO = gcME, "ME.nMovMeImporte ", "c.nMovImporte") & " * -1 as nDocImporte " _
     & "FROM   Mov a JOIN MovDoc b ON b.cMovNro = a.cMovNro " _
     & "             JOIN dbComunes.dbo.Documento g ON g.cDocTpo = b.cDocTpo " _
     & "             JOIN MovCta c ON c.cMovNro = a.cMovNro " & IIf(GSSIMBOLO = gcME, " JOIN MovMe me ON me.cMovNro = c.cMovNro and me.cMovItem = c.cMovItem", "") _
     & "             JOIN MovObj d ON d.cMovNro = c.cMovNro and d.cMovItem = c.cMovItem " _
     & "WHERE  ( (a.cMovEstado IN ('" & IIf(lbPresu, "8", IIf(lbImprime, "7','9", "8','9")) & "') and a.cMovFlag NOT IN (" & IIf(lbImprime, "'M'", "'X','E','N','M'") & ")) " _
     & IIf(lbImprime, " or (a.cMovEstado IN ('7','8','9') and a.cMovFlag = 'X') ) ", ")") _
     & "       and c.cCtaContCod = '" & sCtaProvis & "' and b.cDocTpo = '" & sDocTpoOC & "' and " _
     & "       NOT EXISTS (SELECT h.cMovNro FROM  MovRef h JOIN Mov M ON m.cMovNro = h.cMovNro  " _
     & "                   WHERE m.cMovFlag NOT IN ('X','E','N') and h.cMovNroRef = a.cMovNro) " _
     & "ORDER BY b.cDocNro"
Else
   sSql = " SELECT Distinct b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro," _
        & " cNomPers = (" _
        & "             SELECT cPersNombre + space(50) + PE.cPersCod" _
        & "             FROM Persona PE" _
        & "             WHERE PE.cPersCod = dd.cPersCod)," _
        & "     a.cMovDesc, d.cBSCod, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod," _
        & "     c.nMovImporte * -1 as nDocImporte, ISNULL(ref.nMontoA,0) * -1 nMontoA, IsNull(ce.nMovMEImporte,0) * -1 as nDocMEImporte" _
        & " FROM   Mov a" _
        & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
        & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
        & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
        & " Left Join MovME ce ON c.nMovNro = ce.nMovNro And c.nMovItem = ce.nMovItem" _
        & " JOIN MovGasto dd ON dd.nMovNro = c.nMovNro" _
        & " Left JOIN MovBS d ON d.nMovNro = c.nMovNro" _
        & " LEFT JOIN (" _
        & "            SELECT h.nMovNroRef, SUM(nMovImporte) nMontoA" _
        & "            FROM  MovRef h" _
        & "            JOIN Mov m ON m.nMovNro = h.nMovNro" _
        & "             JOIN MovCta mc ON mc.nMovNro = h.nMovNro" _
        & "             WHERE m.nMovEstado = 10 and m.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & 5 & ") And mc.cCtaContCod = '" & sCtaProvis & "'" _
        & "             GROUP BY h.nMovNroRef) ref ON ref.nMovNroRef = a.nMovNro" _
        & " WHERE  a.nMovEstado IN ('16') and a.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagExtornado & ") " _
        & " And c.cCtaContCod = '" & sCtaProvis & "' And b.nDocTpo = '" & sDocTpoOC & "'" _
        & " ORDER BY b.cDocNro"

   'sSql = "SELECT b.dDocFecha, g.cDocAbrev, b.cDocTpo, b.cDocNro, cNomPers = (SELECT cNomPers FROM " & gcCentralPers & "persona WHERE cCodPers = substring(d.cObjetoCod,3,10)), " _
     & "       a.cMovDesc, d.cObjetoCod, " _
     & "       a.cMovNro, a.cMovEstado, a.cMovFlag, c.cCtaContCod, " & IIf(GSSIMBOLO = gcME, "ME.nMovMeImporte ", "c.nMovImporte") & " * -1 as nDocImporte, ISNULL(ref.nMontoA,0) * -1 nMontoA " _
     & "FROM   Mov a JOIN MovDoc b ON b.cMovNro = a.cMovNro " _
     & "             JOIN dbComunes.dbo.Documento g ON g.cDocTpo = b.cDocTpo " _
     & "             JOIN MovCta c ON c.cMovNro = a.cMovNro " & IIf(GSSIMBOLO = gcME, " JOIN MovMe me ON me.cMovNro = c.cMovNro and me.cMovItem = c.cMovItem", "") _
     & "             JOIN MovObj d ON d.cMovNro = c.cMovNro and d.cMovItem = c.cMovItem " _
     & "        LEFT JOIN (SELECT h.cMovNroRef, SUM(nMov" & IIf(GSSIMBOLO = gcME, "ME", "") & "Importe) nMontoA " _
     & "                   FROM  MovRef h JOIN Mov m ON m.cMovNro = h.cMovNro " _
     & "                                  JOIN MovCta mc ON mc.cMovNro = h.cMovNro " & IIf(GSSIMBOLO = gcME, " JOIN MovMe me ON me.cMovNro = mc.cMovNro and me.cMovItem = mc.cMovItem ", "") _
     & "                   WHERE m.cMovEstado = '0' and m.cMovFlag NOT IN ('X','E','N') " _
     & "                         and mc.cCtaContCod = '" & sCtaProvis & "' " _
     & "                   GROUP BY h.cMovNroRef " _
     & "                  ) ref ON ref.cMovNroRef = a.cMovNro " _
     & "WHERE  a.cMovEstado IN ('" & IIf(lbPresu, "8", "9") & "') and a.cMovFlag NOT IN ('X','E','N','M') and c.cCtaContCod = '" & sCtaProvis & "' " _
     & "       and b.cDocTpo = '" & sDocTpoOC & "' " _
     & "ORDER BY b.cDocNro"
End If

If lbReporteFechas Then
        If STIPOoc = "T" Then
             sSql = " SELECT  Distinct b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''),dd.cPersCod, " _
             & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
             & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
             & "     FROM Persona PE" _
             & "     WHERE PE.cPersCod = dd.cPersCod)" _
             & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte, IsNull(ce.nMovMEImporte,0) * -1 as nDocMEImporte" _
             & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
             & " Inner Join Documento g ON g.nDocTpo = b.nDocTpo" _
             & " Inner Join MovCta c ON c.nMovNro = a.nMovNro" _
             & " Left  Join MovME ce ON c.nMovNro = ce.nMovNro And c.nMovItem = ce.nMovItem" _
             & " Inner Join MovGasto dd ON a.nMovNro = dd.nMovNro" _
             & " WHERE  a.nMovFlag NOT IN (" & gMovFlagModificado & ")" _
             & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And b.dDocFecha  Between '" & Format(Me.mskFecIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFecFin.Text, gsFormatoFecha) & "'" _
             & " ORDER BY b.cDocNro"
           Else
           sSql = " SELECT  Distinct b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & "  ),''),dd.cPersCod, " _
             & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
             & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
             & "     FROM Persona PE" _
             & "     WHERE PE.cPersCod = dd.cPersCod)" _
             & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte, IsNull(ce.nMovMEImporte,0) * -1 as nDocMEImporte" _
             & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
             & " join movcotizac mc on a.nMovNro  = mc.nMovNro " _
             & " Inner Join Documento g ON g.nDocTpo = b.nDocTpo" _
             & " Inner Join MovCta c ON c.nMovNro = a.nMovNro" _
             & " Left  Join MovME ce ON c.nMovNro = ce.nMovNro And c.nMovItem = ce.nMovItem" _
             & " Inner Join MovGasto dd ON a.nMovNro = dd.nMovNro" _
             & " WHERE  a.nMovFlag NOT IN (" & gMovFlagModificado & ")" _
             & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And b.dDocFecha  Between '" & Format(Me.mskFecIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFecFin.Text, gsFormatoFecha) & "'" _
             & " and mc.ctipoOC ='" & STIPOoc & "' " _
             & " ORDER BY b.cDocNro"
         End If
End If

Set rs = oCon.CargaRecordSet(sSql)
fg.Rows = 3
nItem = 1
nTot = 0
Do While Not rs.EOF
   If nItem <> 1 Then
      AdicionaRow fg
   End If
   nItem = fg.Row
   fg.TextMatrix(nItem, 0) = nItem - 1
   fg.TextMatrix(nItem, 1) = rs!cDocAbrev
   fg.TextMatrix(nItem, 2) = rs!cDocNro

   fg.TextMatrix(nItem, 3) = IIf(IsNull(rs!tipoOc), "", rs!tipoOc) + "-" + rs!cdocnroOCD





   fg.TextMatrix(nItem, 4) = rs!dDocFecha
   If Not IsNull(rs!cNomPers) Then
      fg.TextMatrix(nItem, 5) = PstaNombre(Trim(Mid(rs!cNomPers, 1, Len(rs!cNomPers) - 50)), True)
   End If
   fg.TextMatrix(nItem, 6) = Format(rs!nDocImporte, gcFormView)
   fg.TextMatrix(nItem, 7) = rs!cMovDesc
   fg.TextMatrix(nItem, 8) = rs!nmovnro
   fg.TextMatrix(nItem, 9) = Right(rs!cNomPers, 13) & "" ' rs!cBSCod 'CODIGO PERSONA
   If Not (lbPresu Or lbModifica Or lbImprime) Then
      fg.TextMatrix(nItem, 10) = Format(rs!nDocImporte - rs!nMontoA, gcFormView)
   Else
      fg.TextMatrix(nItem, 10) = Format(rs!nDocImporte, gcFormView)
   End If
   If rs!nMovEstado = 16 And Not rs!nMovFlag = gMovFlagExtornado Then
      fg.TextMatrix(nItem, 11) = "Aprobado"
   ElseIf rs!nMovEstado = 15 And Not rs!nMovFlag = gMovFlagExtornado Then
      fg.TextMatrix(nItem, 11) = "Pendiente"
      fg.Col = 11
      fg.CellBackColor = "&H00C0C0FF"
   ElseIf rs!nMovEstado = 14 Then
      fg.TextMatrix(nItem, 11) = "RECHAZADO"
      fg.Col = 11
      fg.CellBackColor = "&H0080FF80"
   Else
      If rs!nMovFlag = gMovFlagExtornado Or rs!nMovFlag = gMovFlagEliminado Then
         fg.TextMatrix(nItem, 11) = "ELIMINADO"
         fg.Col = 11
         fg.CellBackColor = "&H0080FF80"
      End If
   End If
   If lbReporteFechas Then fg.TextMatrix(nItem, 12) = Format(rs!nDocMEImporte, gcFormView)

   nTot = nTot + rs!nDocImporte
   rs.MoveNext
Loop
RSClose rs
txtTot = Format(nTot, gcFormView)
fg.Row = 2
fg.Col = 1
End Sub






'Private Sub GetOCPendientes(Optional STIPOoc As String = "")
'Dim lsCadSerAdm As String
'Dim nItem As Integer
'Dim nTot  As Currency
'Dim oCon As DConecta
'Dim sTipDoc As String
'Dim STIPOC As String
'
'Set oCon = New DConecta
'
'oCon.AbreConexion
'sSQL = "Select '[' +  Rtrim(cNomSer) + '].' + RTrim(cDataBase) + '.dbo.' from servidor Where cCodAge = '11207' And cNroSer = '02'"
'
'sSQL = "Select RTrim(cDataBase) + '.dbo.' from servidor Where cCodAge = '11207' And cNroSer = '02'"
'
'lsCadSerAdm = "" 'oCon.CargaRecordSet(sSQL).Fields(0)
' If lbBienes = True Then
'    sTipDoc = "('" & LogTipoOC.gLogOCompraDirecta & "','" & LogTipoOC.gLogOCompraProceso & "'  )"
'    Else
'    sTipDoc = "('" & LogTipoOC.gLogOServicioDirecta & "','" & LogTipoOC.gLogOServicioProceso & "'  )"
' End If
'If lbPresu Or lbModifica Or lbImprime Then
'        If STIPOoc = "T" Then
'       sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN  " & sTipDoc & " ),''),   dd.cPersCod, " _
'        & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'        & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'        & "     FROM Persona PE" _
'        & "     WHERE PE.cPersCod = dd.cPersCod)" _
'        & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, " & IIf(Mid(lsOpeCod, 3, 1) = 2 And gsCodCMAC = "102", "me.nMovMEImporte", "c.nMovImporte") & " * -1 as nDocImporte" _
'        & " FROM Mov a" _
'        & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'        & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'        & " JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(Mid(lsOpeCod, 3, 1) = 2 And gsCodCMAC = "102", " LEFT JOIN MovMe me ON me.nMovNro = c.nMovNro and me.nMovItem = c.nMovItem ", "") _
'        & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'        & " WHERE  ((a.nMovEstado IN (" & IIf(lbPresu, gMovEstPresupPendiente, IIf(lbImprime, gMovEstPresupRechazado & "," & gMovEstPresupAceptado, gMovEstPresupPendiente & "," & gMovEstPresupAceptado)) & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'        & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
'        & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
'        & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro And cAgeCodRef = '')" _
'        & " ORDER BY b.cDocNro"
'
'        Else
'        sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''),dd.cPersCod, " _
'        & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = ( " _
'        & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'        & "     FROM Persona PE" _
'        & "     WHERE PE.cPersCod = dd.cPersCod)" _
'        & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, " & IIf(Mid(lsOpeCod, 3, 1) = 2 And gsCodCMAC = "102", "me.nMovMEImporte", "c.nMovImporte") & " * -1 as nDocImporte" _
'        & " FROM Mov a" _
'        & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'        & " join movcotizac mc on a.nMovNro  = mc.nMovNro " _
'        & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'        & " JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(Mid(lsOpeCod, 3, 1) = 2 And gsCodCMAC = "102", " LEFT JOIN MovMe me ON me.nMovNro = c.nMovNro and me.nMovItem = c.nMovItem ", "") _
'        & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'        & " WHERE  ((a.nMovEstado IN (" & IIf(lbPresu, gMovEstPresupPendiente, IIf(lbImprime, gMovEstPresupRechazado & "," & gMovEstPresupAceptado, gMovEstPresupPendiente & "," & gMovEstPresupAceptado)) & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'        & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
'        & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
'        & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro And cAgeCodRef = '') and mc.ctipoOC ='" & STIPOoc & "'" _
'        & " ORDER BY b.cDocNro"
'        End If
'
'
'
''   sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
'        & " cNomPers = (" _
'        & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'        & "     FROM Persona PE" _
'        & "     WHERE PE.cPersCod = dd.cPersCod)" _
'        & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'        & " FROM Mov a" _
'        & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'        & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'        & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'        & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'        & " WHERE  ((a.nMovEstado IN (" & IIf(lbPresu, gMovEstPresupPendiente, IIf(lbImprime, gMovEstPresupRechazado & "," & gMovEstPresupAceptado, gMovEstPresupPendiente & "," & gMovEstPresupAceptado)) & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'        & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
'        & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
'        & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE (h.cagecodref is null or len(h.cagecodref)=0) and  m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro)" _
'        & " ORDER BY b.cDocNro"
'
'        'Agregado: (h.cagecodref is null or len(h.cagecodref)=0) and
'
'    If lbImprime Then
'
'        If gbBitCentral Then
'            'Centralizado
'            If STIPOoc = "T" Then
'            sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), b.cDocNro, dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a" _
'                 & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) " _
'                 & " And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "'  And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE Left(m.cOpeCod,2) Not in ('56','59') And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro)  " _
'                 & " ORDER BY b.cDocNro"
'            Else
'
'            sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a" _
'                 & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " join movcotizac mc on a.nMovNro  = mc.nMovNro " _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'                 & " And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "'  And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE Left(m.cOpeCod,2) Not in ('56','59') And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro)  " _
'                 & " and mc.ctipoOC ='" & STIPOoc & "' " _
'                 & " ORDER BY b.cDocNro"
'
'            End If
'        Else
'            'Centralizado
'
'            sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
'                 & " cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a" _
'                 & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE a.cMovNro Between '" & Format(CDate(Me.mskIni.Text), gsFormatoMovFecha) & "' And '" & Format(CDate(Me.mskFin.Text) + 1, gsFormatoMovFecha) & "' And  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "'  And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE Left(m.cOpeCod,2) Not in ('56','59') And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro  And cAgeCodRef = '')  " _
'                 & " ORDER BY b.cDocNro"
'            'MsgBox "s"
'
'            'Hibrido
'
'            'sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
'                 & " cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a" _
'                 & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "'  And NOT EXISTS (SELECT h.cMovNro FROM  " & lsCadSerAdm & "MovRef h JOIN " & lsCadSerAdm & "Mov M ON m.cMovNro = h.cMovNro WHERE Left(m.cOpeCod,2) Not in ('56','59') And m.cMovFlag NOT IN ('X','Y','N') And h.cMovNroRef = a.cMovNro)  " _
'                & " ORDER BY b.cDocNro"
'        End If
'
'    ElseIf lbModifica Then
'        If gbBitCentral Then
'           'Centralizado
'
'
'            If STIPOoc = "T" Then
'
'            sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''), cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "'" _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
'                 & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
'                 & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro)" _
'                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro) ORDER BY b.cDocNro"
'            Else
'
'            sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro join movcotizac mc on a.nMovNro  = mc.nMovNro " _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
'                 & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
'                 & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro) and mc.ctipoOC ='" & STIPOoc & "' " _
'                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''), dd.cPersCod, " _
'                 & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro join movcotizac mc on a.nMovNro  = mc.nMovNro " _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7))) And b.dDocFecha  Between '" & Format(Me.mskIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFin.Text, gsFormatoFecha) & "' " _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro) and mc.ctipoOC ='" & STIPOoc & "' ORDER BY b.cDocNro"
'            End If
'       Else
'            'Centralizado
'
'            sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
'                 & " cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM  Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
'                 & " NOT EXISTS (SELECT h.nMovNro FROM  MovRef h" _
'                 & " JOIN Mov M ON m.nMovNro = h.nMovNro WHERE a.cMovNro Between '" & Format(CDate(Me.mskIni.Text), gsFormatoMovFecha) & "' And '" & Format(CDate(Me.mskFin.Text) + 1, gsFormatoMovFecha) & "' And m.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ") and h.nMovNroRef = a.nMovNro And IsNull(h.cAgeCodRef,'') = '')   " _
'                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
'                 & " cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  a.cMovNro Between '" & Format(CDate(Me.mskIni.Text), gsFormatoMovFecha) & "' And '" & Format(CDate(Me.mskFin.Text) + 1, gsFormatoMovFecha) & "' And ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro And IsNull(h.cAgeCodRef,'') = '') ORDER BY b.cDocNro"
'
'
'
'            'Hibrido
'
'            'sSQL = " SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
'                 & " cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, ISNULL(me.nMovMEImporte, c.nMovImporte) * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'                 & " JOIN MovCta c ON c.nMovNro = a.nMovNro LEFT JOIN MovMe me ON me.nMovNro = c.nMovNro and me.nMovItem = c.nMovItem " _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro" _
'                 & " WHERE  ((a.nMovEstado IN (" & gMovEstPresupPendiente & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' and" _
'                 & " NOT EXISTS (SELECT h.cMovNro FROM  " & lsCadSerAdm & "MovRef h" _
'                 & " JOIN " & lsCadSerAdm & "Mov M ON m.cMovNro = h.cMovNro WHERE m.cMovFlag NOT IN ('X','E','N') and h.cMovNroRef = a.cMovNro)" _
'                 & " Union SELECT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro, dd.cPersCod, " _
'                 & " cNomPers = (" _
'                 & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'                 & "     FROM Persona PE" _
'                 & "     WHERE PE.cPersCod = dd.cPersCod)" _
'                 & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte" _
'                 & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'                 & " JOIN Documento g ON g.nDocTpo = b.nDocTpo JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'                 & " JOIN MovGasto dd ON a.nMovNro = dd.nMovNro WHERE  ((a.nMovEstado IN (" & gMovEstPresupAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagDeExtorno & "," & gMovFlagEliminado & "," & gMovFlagExtornado & ",5,7)))" _
'                 & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And NOT EXISTS (SELECT h.nMovNro FROM  MovRef h JOIN Mov M ON m.nMovNro = h.nMovNro WHERE m.cOpeCod Not Like '56%' And m.nMovFlag NOT IN (3,1,2,5,7) And h.nMovNroRef = a.nMovNro) ORDER BY b.cDocNro"
'
'        End If
'
'    End If
'
'
'
'   'SELECT b.dDocFecha, g.cDocAbrev, b.cDocTpo, b.cDocNro, cNomPers = (SELECT cNomPers FROM " & gcCentralPers & "persona WHERE cCodPers = substring(d.cObjetoCod,3,10)), " _
'     & "       a.cMovDesc, d.cObjetoCod, " _
'     & "       a.cMovNro, a.cMovEstado, a.cMovFlag, c.cCtaContCod, " & IIf(GSSIMBOLO = gcME, "ME.nMovMeImporte ", "c.nMovImporte") & " * -1 as nDocImporte " _
'     & "FROM   Mov a JOIN MovDoc b ON b.cMovNro = a.cMovNro " _
'     & "             JOIN dbComunes.dbo.Documento g ON g.cDocTpo = b.cDocTpo " _
'     & "             JOIN MovCta c ON c.cMovNro = a.cMovNro " & IIf(GSSIMBOLO = gcME, " JOIN MovMe me ON me.cMovNro = c.cMovNro and me.cMovItem = c.cMovItem", "") _
'     & "             JOIN MovObj d ON d.cMovNro = c.cMovNro and d.cMovItem = c.cMovItem " _
'     & "WHERE  ( (a.cMovEstado IN ('" & IIf(lbPresu, "8", IIf(lbImprime, "7','9", "8','9")) & "') and a.cMovFlag NOT IN (" & IIf(lbImprime, "'M'", "'X','E','N','M'") & ")) " _
'     & IIf(lbImprime, " or (a.cMovEstado IN ('7','8','9') and a.cMovFlag = 'X') ) ", ")") _
'     & "       and c.cCtaContCod = '" & sCtaProvis & "' and b.cDocTpo = '" & sDocTpoOC & "' and " _
'     & "       NOT EXISTS (SELECT h.cMovNro FROM  MovRef h JOIN Mov M ON m.cMovNro = h.cMovNro  " _
'     & "                   WHERE m.cMovFlag NOT IN ('X','E','N') and h.cMovNroRef = a.cMovNro) " _
'     & "ORDER BY b.cDocNro"
'
'Else
'
'   sSQL = " SELECT Distinct b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro," _
'        & " cNomPers = (" _
'        & "             SELECT cPersNombre + space(50) + PE.cPersCod" _
'        & "             FROM Persona PE" _
'        & "             WHERE PE.cPersCod = dd.cPersCod)," _
'        & "     a.cMovDesc, d.cBSCod, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod," _
'        & "     c.nMovImporte * -1 as nDocImporte, ISNULL(ref.nMontoA,0) * -1 nMontoA, IsNull(ce.nMovMEImporte,0) * -1 as nDocMEImporte" _
'        & " FROM   Mov a" _
'        & " JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'        & " JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
'        & " JOIN MovCta c ON c.nMovNro = a.nMovNro" _
'        & " Left Join MovME ce ON c.nMovNro = ce.nMovNro And c.nMovItem = ce.nMovItem" _
'        & " JOIN MovGasto dd ON dd.nMovNro = c.nMovNro" _
'        & " Left JOIN MovBS d ON d.nMovNro = c.nMovNro" _
'        & " LEFT JOIN (" _
'        & "            SELECT h.nMovNroRef, SUM(nMovImporte) nMontoA" _
'        & "            FROM  MovRef h" _
'        & "            JOIN Mov m ON m.nMovNro = h.nMovNro" _
'        & "             JOIN MovCta mc ON mc.nMovNro = h.nMovNro" _
'        & "             WHERE m.nMovEstado = 10 and m.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & 5 & ") And mc.cCtaContCod = '" & sCtaProvis & "'" _
'        & "             GROUP BY h.nMovNroRef) ref ON ref.nMovNroRef = a.nMovNro" _
'        & " WHERE  a.nMovEstado IN ('16') and a.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagExtornado & ") " _
'        & " And c.cCtaContCod = '" & sCtaProvis & "' And b.nDocTpo = '" & sDocTpoOC & "'" _
'        & " ORDER BY b.cDocNro"
'
'
'
'   'sSql = "SELECT b.dDocFecha, g.cDocAbrev, b.cDocTpo, b.cDocNro, cNomPers = (SELECT cNomPers FROM " & gcCentralPers & "persona WHERE cCodPers = substring(d.cObjetoCod,3,10)), " _
'     & "       a.cMovDesc, d.cObjetoCod, " _
'     & "       a.cMovNro, a.cMovEstado, a.cMovFlag, c.cCtaContCod, " & IIf(GSSIMBOLO = gcME, "ME.nMovMeImporte ", "c.nMovImporte") & " * -1 as nDocImporte, ISNULL(ref.nMontoA,0) * -1 nMontoA " _
'     & "FROM   Mov a JOIN MovDoc b ON b.cMovNro = a.cMovNro " _
'     & "             JOIN dbComunes.dbo.Documento g ON g.cDocTpo = b.cDocTpo " _
'     & "             JOIN MovCta c ON c.cMovNro = a.cMovNro " & IIf(GSSIMBOLO = gcME, " JOIN MovMe me ON me.cMovNro = c.cMovNro and me.cMovItem = c.cMovItem", "") _
'     & "             JOIN MovObj d ON d.cMovNro = c.cMovNro and d.cMovItem = c.cMovItem " _
'     & "        LEFT JOIN (SELECT h.cMovNroRef, SUM(nMov" & IIf(GSSIMBOLO = gcME, "ME", "") & "Importe) nMontoA " _
'     & "                   FROM  MovRef h JOIN Mov m ON m.cMovNro = h.cMovNro " _
'     & "                                  JOIN MovCta mc ON mc.cMovNro = h.cMovNro " & IIf(GSSIMBOLO = gcME, " JOIN MovMe me ON me.cMovNro = mc.cMovNro and me.cMovItem = mc.cMovItem ", "") _
'     & "                   WHERE m.cMovEstado = '0' and m.cMovFlag NOT IN ('X','E','N') " _
'     & "                         and mc.cCtaContCod = '" & sCtaProvis & "' " _
'     & "                   GROUP BY h.cMovNroRef " _
'     & "                  ) ref ON ref.cMovNroRef = a.cMovNro " _
'     & "WHERE  a.cMovEstado IN ('" & IIf(lbPresu, "8", "9") & "') and a.cMovFlag NOT IN ('X','E','N','M') and c.cCtaContCod = '" & sCtaProvis & "' " _
'     & "       and b.cDocTpo = '" & sDocTpoOC & "' " _
'     & "ORDER BY b.cDocNro"
'
'End If
'
'
'
'If lbReporteFechas Then
'
'        If STIPOoc = "T" Then
'
'             sSQL = " SELECT  Distinct b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & " ),''),dd.cPersCod, " _
'             & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'             & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'             & "     FROM Persona PE" _
'             & "     WHERE PE.cPersCod = dd.cPersCod)" _
'             & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte, IsNull(ce.nMovMEImporte,0) * -1 as nDocMEImporte" _
'             & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'             & " Inner Join Documento g ON g.nDocTpo = b.nDocTpo" _
'             & " Inner Join MovCta c ON c.nMovNro = a.nMovNro" _
'             & " Left  Join MovME ce ON c.nMovNro = ce.nMovNro And c.nMovItem = ce.nMovItem" _
'             & " Inner Join MovGasto dd ON a.nMovNro = dd.nMovNro" _
'             & " WHERE  a.nMovFlag NOT IN (" & gMovFlagModificado & ")" _
'             & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And b.dDocFecha  Between '" & Format(Me.mskFecIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFecFin.Text, gsFormatoFecha) & "'" _
'             & " ORDER BY b.cDocNro"
'
'           Else
'
'           sSQL = " SELECT  Distinct b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro,cdocnroOCD  = isnull((select cDocNro  from movdoc md where md.nMovNro = a.nMovNro and nDocTpo IN " & sTipDoc & "  ),''),dd.cPersCod, " _
'             & " TipoOC = isnull((select ctipoOC  from movcotizac mx where mx.nMovNro = a.nMovNro ),''),cNomPers = (" _
'             & "     SELECT cPersNombre + space(50) + PE.cPersCod " _
'             & "     FROM Persona PE" _
'             & "     WHERE PE.cPersCod = dd.cPersCod)" _
'             & " , a.cMovDesc, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte, IsNull(ce.nMovMEImporte,0) * -1 as nDocMEImporte" _
'             & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
'             & " join movcotizac mc on a.nMovNro  = mc.nMovNro " _
'             & " Inner Join Documento g ON g.nDocTpo = b.nDocTpo" _
'             & " Inner Join MovCta c ON c.nMovNro = a.nMovNro" _
'             & " Left  Join MovME ce ON c.nMovNro = ce.nMovNro And c.nMovItem = ce.nMovItem" _
'             & " Inner Join MovGasto dd ON a.nMovNro = dd.nMovNro" _
'             & " WHERE  a.nMovFlag NOT IN (" & gMovFlagModificado & ")" _
'             & " and c.cCtaContCod = '" & sCtaProvis & "' and b.nDocTpo = '" & sDocTpoOC & "' And b.dDocFecha  Between '" & Format(Me.mskFecIni.Text, gsFormatoFecha) & "' And '" & Format(Me.mskFecFin.Text, gsFormatoFecha) & "'" _
'             & " and mc.ctipoOC ='" & STIPOoc & "' " _
'             & " ORDER BY b.cDocNro"
'         End If
'End If
'Set rs = oCon.CargaRecordSet(sSQL)
'fg.Rows = 3
'nItem = 1
'nTot = 0
'Do While Not rs.EOF
'   If nItem <> 1 Then
'      AdicionaRow fg
'   End If
'   nItem = fg.Row
'   fg.TextMatrix(nItem, 0) = nItem - 1
'   fg.TextMatrix(nItem, 1) = rs!cDocAbrev
'   fg.TextMatrix(nItem, 2) = rs!cDocNro
'   fg.TextMatrix(nItem, 3) = IIf(IsNull(rs!tipoOc), "", rs!tipoOc) + "-" + rs!cdocnroOCD
'   fg.TextMatrix(nItem, 4) = rs!dDocFecha
'   If Not IsNull(rs!cNomPers) Then
'      fg.TextMatrix(nItem, 5) = PstaNombre(Trim(Mid(rs!cNomPers, 1, Len(rs!cNomPers) - 50)), True)
'   End If
'   fg.TextMatrix(nItem, 6) = Format(rs!nDocImporte, gcFormView)
'   fg.TextMatrix(nItem, 7) = rs!cMovDesc
'   fg.TextMatrix(nItem, 8) = rs!nMovNro
'   fg.TextMatrix(nItem, 9) = Right(rs!cNomPers, 13) & "" ' rs!cBSCod 'CODIGO PERSONA
'   If Not (lbPresu Or lbModifica Or lbImprime) Then
'      fg.TextMatrix(nItem, 10) = Format(rs!nDocImporte - rs!nMontoA, gcFormView)
'   Else
'      fg.TextMatrix(nItem, 10) = Format(rs!nDocImporte, gcFormView)
'   End If
'   If rs!nMovEstado = 16 And Not rs!nMovFlag = gMovFlagExtornado Then
'      fg.TextMatrix(nItem, 11) = "Aprobado"
'   ElseIf rs!nMovEstado = 15 And Not rs!nMovFlag = gMovFlagExtornado Then
'      fg.TextMatrix(nItem, 11) = "Pendiente"
'      fg.Col = 11
'      fg.CellBackColor = "&H00C0C0FF"
'   ElseIf rs!nMovEstado = 14 Then
'      fg.TextMatrix(nItem, 11) = "RECHAZADO"
'      fg.Col = 11
'      fg.CellBackColor = "&H0080FF80"
'   Else
'      If rs!nMovFlag = gMovFlagExtornado Or rs!nMovFlag = gMovFlagEliminado Then
'         fg.TextMatrix(nItem, 11) = "ELIMINADO"
'         fg.Col = 11
'         fg.CellBackColor = "&H0080FF80"
'      End If
'   End If
'   If lbReporteFechas Then fg.TextMatrix(nItem, 12) = Format(rs!nDocMEImporte, gcFormView)
'   nTot = nTot + rs!nDocImporte
'   rs.MoveNext
'Loop
'
'RSClose rs
'txtTot = Format(nTot, gcFormView)
'fg.Row = 2
'fg.Col = 1
'End Sub
 



Private Sub CmdAceptar_Click()
Dim N As Integer
Dim sMovAnt As String
Dim sMovNro As String
Dim nCont   As Integer
Dim sCta    As String
Dim nSaldo  As Currency
Dim oConect As DConecta
Set oConect = New DConecta

On Error GoTo ErrSub

If fg.TextMatrix(2, 1) = "" Then
   Exit Sub
End If
If Not lbPresu And Not lbModifica And Not lbImprime Then
   If fg.TextMatrix(fg.Row, 10) = "Pendiente" Then
      MsgBox "O/C aún no aprobado por Presupuesto", vbInformation, "¡Aviso!"
      Exit Sub
   End If
End If
If Not lbImprime Then
   If MsgBox(" ¿ Seguro de Actualizar datos de " & sDocDesc & " Nro. " & fg.TextMatrix(fg.Row, 2) & " ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
      Exit Sub
   End If
End If

gcMovNro = fg.TextMatrix(fg.Row, 8)
gsGlosa = fg.TextMatrix(fg.Row, 7)
gcPersona = fg.TextMatrix(fg.Row, 5)
gcDocTpo = sDocTpoOC
gcDocNro = fg.TextMatrix(fg.Row, 2)
gdFecha = CDate(fg.TextMatrix(fg.Row, 4))


If lbPresu = False Then
   If lbModifica Or lbImprime Then
      Set frmLogOCompra = Nothing
      
      frmLogOCompra.Inicio True, lsOpeCod, fg.TextMatrix(fg.Row, 9), lbBienes, True, lbImprime, fg.TextMatrix(fg.Row, 11)
      
      Set frmLogOCompra = Nothing
      oConect.AbreConexion
      If frmLogOCompra.lOk Then
         GetOCPendientes
      End If
   Else
      gnSaldo = nVal(fg.TextMatrix(fg.Row, 9))
      nSaldo = gnSaldo
      'frmLogOCIngBien.Inicio True, lbBienes, lsOpeCod, fg.TextMatrix(fg.Row, 8)

'      If frmLogOCIngBien.lOk Then
'         If gnSaldo <= 0 Then
'            EliminaRow2 fg, fg.Row, 2
'            ActualizaTot nSaldo * -1
'         Else
'            fg.TextMatrix(fg.Row, 9) = Format(gnSaldo, gcFormView)
'            ActualizaTot (nSaldo - gnSaldo) * -1
'         End If
'      End If
   End If
Else
    'Opción para ventana en PRESUPUESTO
   gcOpeCod = lsOpeCod
   'frmPlaOCIngBien.Inicio True, lbBienes, fg.TextMatrix(fg.Row, 9), IIf(optPeriodo(1).value, 1, IIf(optPeriodo(3).value, 3, IIf(optPeriodo(6).value, 6, 12)))
   oConect.AbreConexion
'   If frmPlaOCIngBien.lOk Then
'      EliminaRow2 fg, fg.Row, 2
'      ActualizaTot nVal(txtTot.Text) * -1
'   End If
End If
Exit Sub
ErrSub:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdActualizar_Click()
    GetOCPendientes Right(cmbTipoOC.Text, 1)
End Sub

Private Sub CmdExtornar_Click()
Dim sMov As String
Dim N As Integer
Dim oMov As DMov
Set oMov = New DMov
If MsgBox(" ¿ Seguro de Extornar " & sDocDesc & " Nro. " & fg.TextMatrix(fg.Row, 2) & " ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
   Exit Sub
End If

oMov.ActualizaMov fg.TextMatrix(fg.Row, 8), , , gMovFlagEliminado
EliminaRow fg, fg.Row
Set oMov = Nothing

'dbCmact.BeginTrans
'   sSql = "UPDATE Mov SET cMovFlag = 'X' WHERE cMovNro = '" & fg.TextMatrix(fg.Row, 7) & "'"
'   dbCmact.Execute sSql
'   EliminaRow fg, fg.Row
'dbCmact.CommitTrans
'txtTot = ""
End Sub

Private Sub cmdImprimir_Click()
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    
    If Me.fg.TextMatrix(1, 1) = "" Then
        MsgBox "No existen datos.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.mskFecFin.Text), "yyyymmdd") & Format(Time, "hhmmss") & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       GeneraReporte fg, xlHoja1, 10
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0
End Sub

Public Function ExcelBegin(psArchivo As String, _
        xlAplicacion As Excel.Application, _
        xlLibro As Excel.Workbook, Optional pbBorraExiste As Boolean = True) As Boolean
        
Dim fs As New Scripting.FileSystemObject
On Error GoTo ErrBegin
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application

If fs.FileExists(psArchivo) Then
   If pbBorraExiste Then
      fs.DeleteFile psArchivo, True
      Set xlLibro = xlAplicacion.Workbooks.Add
   Else
      Set xlLibro = xlAplicacion.Workbooks.Open(psArchivo)
   End If
Else
   Set xlLibro = xlAplicacion.Workbooks.Add
End If
ExcelBegin = True
Exit Function
ErrBegin:
  MsgBox TextErr(Err.Description), vbInformation, "Aviso"
  ExcelBegin = False
End Function

Public Sub ExcelEnd(psArchivo As String, xlAplicacion As Excel.Application, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional plSave As Boolean = True)
On Error GoTo ErrEnd
   If plSave Then
        xlHoja1.SaveAs psArchivo
   End If
   xlLibro.Close
   xlAplicacion.Quit
   Set xlAplicacion = Nothing
   Set xlLibro = Nothing
   Set xlHoja1 = Nothing
Exit Sub
ErrEnd:
   MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdProcesar_Click()
    If ValidaFecha(Me.mskFecIni.Text) <> "" Then
       MsgBox "¡ Fecha no Válida !", vbInformation, "¡Aviso!"
       mskFecIni.SetFocus
       Exit Sub
    End If
    If ValidaFecha(Me.mskFecFin.Text) <> "" Then
       MsgBox "¡ Fecha no Válida !", vbInformation, "¡Aviso!"
       mskFecFin.SetFocus
       Exit Sub
    End If
    GetOCPendientes Right(cmbTipoOCR.Text, 1)
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub


Private Sub Form_Activate()
If lSalir Then
   RSClose rs
   Unload Me
End If
If lbReporteFechas Then
   ' Me.mskFecIni.SetFocus
End If
End Sub

Private Sub Form_Load()
    Dim lnColorBien As Double
    Dim lnColorServ As Double
    Dim oDoc As COMNAuditoria.DOperacion
    Set oDoc = New COMNAuditoria.DOperacion
    
    cmbTipoOC.ListIndex = 0
    If lbReporteFechas Then
        Me.fraFechas.Visible = True
        cmdImprimir.Visible = True
    Else
        Me.fraFechas.Visible = False
        cmdImprimir.Visible = False
    End If
    
    cmbTipoOCR.ListIndex = 0
    
    Me.mskIni.Text = "01/01/" & Format(gdFecSis, "yyyy")
    Me.mskFin.Text = Format(gdFecSis, gsFormatoFechaView)
    
    lnColorBien = "&H00F0FFFF"
    lnColorServ = "&H00FFFFC0"
    If Mid(lsOpeCod, 3, 1) = "2" Then
        gsSimbolo = gcME
    Else
        gsSimbolo = gcMN
    End If
    If lbBienes Then
       sDocDesc = "Orden de Compra"
       fg.BackColor = lnColorBien
    Else
       sDocDesc = "Orden de Servicio"
       fg.BackColor = lnColorServ
    End If
    Me.Caption = Me.Caption & sDocDesc
    lSalir = False
    Set rs = oDoc.CargaOpeCta(lsOpeCod, "H")
    If rs.EOF And rs.BOF Then
       MsgBox "Cuenta Contable de Provisión no fue asignada a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Sistemas", vbInformation, "¡Aviso!"
       lSalir = True
       Exit Sub
    End If
    sCtaProvis = rs!cCtaContCod
    
    Set rs = oDoc.CargaOpeDoc(lsOpeCod, "1")
    If rs.EOF Then
       MsgBox "No se asignó Tipo de Documento " & sDocDesc & " a Operación", vbInformation, "¡Aviso!"
       lSalir = True
       Exit Sub
    End If
    sDocTpoOC = rs!nDocTpo
    RSClose rs
    FormatoOCompra
    If Not lbReporteFechas Then
        Me.fraRefrescar.Visible = True
    Else
        'GetOCPendientes
        Me.fraRefrescar.Visible = False
    End If
    If lbImprime Then
       cmdExtornar.Visible = False
       cmdAceptar.Caption = "&Detalle..."
    End If
    If lbPresu Then
       cmdExtornar.Visible = False
       fraPeriodo.Visible = True
    End If
    
    If lsOpeCod = 501221 Or lsOpeCod = 502221 Then
       Me.Caption = "Aprobación Orden Compra - Presupuesto" & Space(2) & "-  " & IIf(Mid(lsOpeCod, 3, 1) = 1, "SOLES", "DOLARES")
    End If
    If lsOpeCod = 501222 Or lsOpeCod = 502222 Then
       Me.Caption = "Aprobación Orden Servicio - Presupuesto" & Space(2) & "-  " & IIf(Mid(lsOpeCod, 3, 1) = 1, "SOLES", "DOLARES")
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Obj As COMConecta.DCOMConecta
    Set Obj = New COMConecta.DCOMConecta
    Obj.CierraConexion
End Sub

Private Sub ActualizaTot(pnMonto As Currency)
    txtTot = Format(Val(Format(txtTot, gcFormDato)) + pnMonto, gcFormView)
End Sub

Private Sub mskFecFin_GotFocus()
    mskFecFin.SelStart = 0
    mskFecFin.SelLength = 50
End Sub

Private Sub mskFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmdProcesar.SetFocus
End Sub

Private Sub mskFecIni_GotFocus()
    mskFecIni.SelStart = 0
    mskFecIni.SelLength = 50
End Sub

Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.mskFecFin.SetFocus
End Sub

Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdActualizar.SetFocus
    End If
End Sub

Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub

Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 50
End Sub

Private Sub mskIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFin.SetFocus
    End If
End Sub

