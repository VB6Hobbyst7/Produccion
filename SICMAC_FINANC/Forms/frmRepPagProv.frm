VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRepPagProv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones: Reporte de Provisión de Proveedores"
   ClientHeight    =   5565
   ClientLeft      =   645
   ClientTop       =   1740
   ClientWidth     =   11775
   Icon            =   "frmRepPagProv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Cmdsaldo 
      Caption         =   "Saldo Diario"
      Height          =   335
      Left            =   6960
      TabIndex        =   25
      Top             =   4740
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox chk1 
      Caption         =   "Cancelados"
      Height          =   345
      Left            =   5370
      TabIndex        =   24
      Top             =   4770
      Width           =   1905
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   30
      Left            =   7170
      TabIndex        =   23
      Top             =   2760
      Width           =   30
   End
   Begin VB.CommandButton cmdRUC 
      Caption         =   "&Buscar Nro Ruc"
      Height          =   315
      Left            =   3720
      TabIndex        =   22
      Top             =   4770
      Width           =   1455
   End
   Begin VB.TextBox txtNroRuc 
      Height          =   285
      Left            =   1770
      TabIndex        =   21
      Top             =   4770
      Width           =   1815
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
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   8430
      TabIndex        =   19
      Top             =   60
      Width           =   1950
      Begin VB.OptionButton OptS 
         Caption         =   "Todo"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   705
      End
      Begin VB.OptionButton OptS 
         Caption         =   "Selección"
         Height          =   285
         Index           =   1
         Left            =   855
         TabIndex        =   6
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdAsientos 
      Caption         =   "&Asientos"
      Height          =   345
      Left            =   8385
      TabIndex        =   18
      ToolTipText     =   "Imprime Asientos Contables"
      Top             =   4710
      Width           =   1095
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   345
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Procesa según Rango de Fechas y Opciones"
      Top             =   4710
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opciones"
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
      Height          =   645
      Left            =   3825
      TabIndex        =   15
      Top             =   60
      Width           =   4560
      Begin VB.OptionButton Opt 
         Caption         =   "&Entregados"
         Height          =   285
         Index           =   3
         Left            =   3315
         TabIndex        =   20
         Top             =   270
         Width           =   1110
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Pen&dientes"
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   3
         Top             =   255
         Width           =   1095
      End
      Begin VB.OptionButton Opt 
         Caption         =   "&Emitidos"
         Height          =   285
         Index           =   2
         Left            =   2190
         TabIndex        =   4
         Top             =   255
         Width           =   975
      End
      Begin VB.OptionButton Opt 
         Caption         =   "To&do"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H8000000D&
      Height          =   645
      Left            =   105
      TabIndex        =   12
      Top             =   60
      Width           =   3675
      Begin MSComCtl2.DTPicker txtFechaDel 
         Height          =   315
         Left            =   540
         TabIndex        =   0
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70254593
         CurrentDate     =   36509
      End
      Begin MSComCtl2.DTPicker txtFechaAl 
         Height          =   315
         Left            =   2220
         TabIndex        =   1
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   70254593
         CurrentDate     =   36509
      End
      Begin VB.Label Label2 
         Caption         =   "Al"
         Height          =   225
         Left            =   1950
         TabIndex        =   14
         Top             =   300
         Width           =   225
      End
      Begin VB.Label Label1 
         Caption         =   "Del"
         Height          =   225
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Planilla"
      Height          =   345
      Left            =   9480
      TabIndex        =   10
      ToolTipText     =   "Imprime Planilla de Provisión"
      Top             =   4710
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvProvis 
      Height          =   3315
      Left            =   60
      TabIndex        =   7
      Top             =   720
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   5847
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
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Comprobante"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Emisión"
         Object.Width           =   2010
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Proveedor"
         Object.Width           =   7145
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ruc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "cMovDesc"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total a Pagar"
         Object.Width           =   2222
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "cCodPers"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Fecha Provision"
         Object.Width           =   2029
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "cDocTpo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "cDocNro"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Voucher"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Fecha Pago"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Cta.Cont."
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtMovDesc 
      Height          =   555
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4050
      Width           =   11580
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   10575
      TabIndex        =   11
      Top             =   4710
      Width           =   1095
   End
   Begin MSComctlLib.ImageList imgRec 
      Left            =   120
      Top             =   4290
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
            Picture         =   "frmRepPagProv.frx":030A
            Key             =   "recibo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   285
      Left            =   1755
      TabIndex        =   16
      Top             =   5250
      Visible         =   0   'False
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   5190
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2999
            MinWidth        =   2999
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRepPagProv"
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

Private Function CabeceraRepo(ByRef nLin As Integer, ByRef P As Integer) As String
Dim sTit    As String
Dim lsImpre As String

If Opt(0).value Then
    If chk1.value = 0 Then
        sTit = " PLANILLA DE PROVISION DE PAGO A PROVEEDORES "
    Else
        sTit = " PLANILLA DE PAGO A PROVEEDORES "
    End If
End If
If Opt(1).value Then
   sTit = " PLANILLA DE PROVISIONES PENDIENTES DE PAGO "
End If
If Opt(2).value Then
   sTit = " PROVISIONES DE PROVEEDORES ATENDIDAS "
End If
If Opt(3).value Then
   sTit = " PAGOS DE PROVEEDORES ENTREGADAS "
End If
If nLin > 60 Then
   If P > 0 Then lsImpre = lsImpre & oImpresora.gPrnSaltoPagina
   P = P + 1
   Linea lsImpre, gsInstCmac & Space(42) & gdFecSis & " - " & Format(Time, "hh:mm:ss")
   Linea lsImpre, Space(82) & "Pag. " & Format(P, "000")
   Linea lsImpre, BON & Centra(sTit, gnColPage)
   Linea lsImpre, Centra(" M O N E D A   " & IIf(Not lMN, "E X T R A N J E R A ", "N A C I O N A L "), gnColPage)
   Linea lsImpre, Centra("( DEL " & txtFechaDel & " AL " & txtFechaAl & ")", gnColPage) & BOFF
   Linea lsImpre, CON
   nLin = 8
   Linea lsImpre, " =====================================================================================================================================================" & IIf(lMN, "", String(15, "="))
   Linea lsImpre, " Item   Fecha de    Comprobante       Proveedor                      RUC         Fec Prov      Detalle                      Importe " & IIf(lMN, "", "     Importe") & "  Fecha Pago    Cta.Cont"
   Linea lsImpre, "        " & ImpreFormat(lvProvis.ColumnHeaders(3), 7, 0) & "    Tpo    Número                                                                                              M.N.  " & IIf(lMN, "", "       M.E.")
   Linea lsImpre, " -----------------------------------------------------------------------------------------------------------------------------------------------------" & IIf(lMN, "", String(15, "-")) & COFF
End If
CabeceraRepo = lsImpre
End Function

'Private Sub GeneraListado(pnOpt As Integer)
'Dim sCond As String
'Dim nItem As Integer
'Dim lvItem As ListItem
'Dim nImporte As Currency, nTipCambio As Currency
'Dim sTpoDoc  As String
'Dim sCondPag As String
'
'sCondPag = ""
'Select Case pnOpt
'   Case 0
'     If chk1.value = 0 Then
'        sCond = " and h.cOpeDocMetodo <> '2' and Ent.cmovnro is null "
'     Else
'        sCond = " and h.cOpeDocMetodo <> '2' and Ent.cmovnro is not null "
'     End If
'      sTpoDoc = "Provisionados"
'   Case 1
'      sCond = " and (NOT EXISTS (SELECT h.nMovNro FROM  MovRef h Inner Join Mov Mh On h.nmovnro = Mh.nMovNro WHERE h.nMovNroRef = a.nMovNro And Mh.nMovFlag <> 1) " _
'            & "      OR  EXISTS (SELECT h.nMovNro " _
'            & "                  FROM  MovRef h JOIN MOV MX on MX.nMovNro = h.nMovNro  " _
'            & "                  WHERE h.nMovNroRef = a.nMovNro AND LEFT(MX.cMovNro,8)>'" & Format(txtFechaAl, "yyyymmdd") & "') ) " _
'            & " and c.nMovImporte < 0 and h.cOpeDocMetodo <> '2' "
'
'
'      sTpoDoc = "Pendientes"
'   Case 2
'      sCond = " and c.nMovImporte > 0 and b.nDocTpo IN (" & TpoDocOrdenPago & "," & TpoDocCheque & "," & TpoDocCarta & ") "
'      sTpoDoc = "Emitidos"
'      sCondPag = " and LEFT(mEnt.cMovNro,8) BETWEEN '" & Format(CDate(txtFechaDel), gsFormatoMovFecha) & "' and '" & Format(CDate(txtFechaAl), gsFormatoMovFecha) & "' "
'   Case 3
'      sCond = " and c.nMovImporte > 0 and b.nDocTpo IN (" & TpoDocOrdenPago & "," & TpoDocCheque & "," & TpoDocCarta & ") " _
'            & " " _
'            & " "
'      sTpoDoc = "Entregados"
'End Select
'
'Dim lsCtaProvis
'If lMN Then
'    lsCtaProvis = "'25160209','25160209'"
'Else
'    lsCtaProvis = "'25260209','25260209'"
'End If
'
'sSql = "SELECT DISTINCT b.dDocFecha, Ent.cMovNro cMovEntrega, g.cDocAbrev, b.nDocTpo cDocTpo, b.cDocNro, e.cPersNombre cNomPers, a.cMovDesc, d.cPersCod cObjetoCod, " _
'     & "  Ruc=isnull((select cPersIDnro from persid i where cPersIDTpo='2' and e.cperscod=i.cperscod),''), " _
'     & "       a.cMovNro as cFecProvi, c.nMovImporte * -1 as nDocImporte " & IIf(lMN, "", ", isnull(ME.nMovMEImporte,0) * -1 as nDocMEImporte ") _
'     & ", (Select mdd.cDocNro from movdoc mdd where mdd.nmovnro = a.nmovnro And mdd.nDocTpo = " & TpoDocVoucherEgreso & ") Voucher, ISNULL(ENT.CMOVNRO,'') as cMovPago " _
'     & "FROM   Mov a " _
'     & "             JOIN MovDoc b ON b.nMovNro = a.nMovNro " _
'     & "             JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(lMN, "", "LEFT JOIN MovME ME ON (ME.nMovNro=c.nMovNro and ME.nMovItem=c.nMovItem)") _
'     & "             JOIN MovGasto d ON d.nMovNro = a.nMovNro " _
'     & "             JOIN Persona e ON e.cPersCod = d.cPersCod " _
'     & "        LEFT JOIN OpeDoc h ON h.nDocTpo = b.nDocTpo and h.cOpeCod = '" & gsOpeCod & "' LEFT JOIN Documento g ON g.nDocTpo = b.nDocTpo " _
'     & "        LEFT JOIN ( SELECT  max(mEnt.cMovNro) cMovNro, rEnt.nMovNroRef " _
'     & "                    FROM    MovRef rEnt " _
'     & "                            JOIN Mov mEnt ON mEnt.nMovNro = rEnt.nMovNro " _
'     & "                    WHERE mEnt.nMovFlag = 0 " & sCondPag & " GRoup by rEnt.nMovNroRef " _
'     & "                  ) Ent ON Ent.nMovNroRef = a.nMovNro " _
'     & "WHERE  a.nMovEstado = '" & gMovEstContabMovContable & "' and a.nMovFlag <> '" & gMovFlagEliminado & "' " _
'     & IIf(pnOpt = 3, " and not Ent.cMovNro is NULL ", " and not h.nDocTpo is NULL and LEFT(a.cMovNro,8) BETWEEN '" & Format(CDate(txtFechaDel), gsFormatoMovFecha) & "' and '" & Format(CDate(txtFechaAl), gsFormatoMovFecha) & "' ") _
'     & "       and c.cCtaContCod IN (" & lsCtaProvis & ") " _
'     & "       " & sCond _
'     & "       ORDER BY a.cMovNro, e.cPersNombre "
'
'Set rs = oCon.CargaRecordSet(sSql)
'lSalir = False
'If RSVacio(rs) Then
'   MsgBox "No existen Comprobantes " & sTpoDoc, vbInformation, "¡Aviso!"
'   lSalir = True
'   Exit Sub
'End If
'lvProvis.ListItems.Clear
'prg.Visible = True
'prg.Min = 0
'prg.Max = rs.RecordCount
'Do While Not rs.EOF
'   prg.value = rs.Bookmark
'   Status.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
'   nItem = nItem + 1
'   Set lvItem = lvProvis.ListItems.Add(, , Format(nItem, "000"))
'   lvItem.SmallIcon = 1
'   lvItem.SubItems(1) = Mid(rs!cDocAbrev & Space(3), 1, 3) & " " & rs!cDocNro
'   If pnOpt = 3 Then
'      lvItem.SubItems(2) = GetFechaMov(rs!cMovEntrega, True)
'   Else
'      lvItem.SubItems(2) = rs!dDocFecha
'   End If
'   lvItem.SubItems(3) = rs!cNomPers
'   lvItem.SubItems(4) = rs!RUC
'   lvItem.SubItems(5) = rs!cMovDesc
'   nImporte = rs!nDocImporte
'   If Not lMN Then
'      nImporte = rs!nDocMEImporte
'      If nImporte = 0 Then
'         nImporte = Round(rs!nDocImporte / gnTipCambio, 2)
'      End If
'      If LTrim(RTrim(rs!cdoctpo)) = "7" Then
'          lvItem.SubItems(13) = Format(Val(nImporte), gsFormatoNumeroView)
'      Else
'          lvItem.SubItems(13) = Format(nImporte * IIf(Val(rs!nDocImporte) > 0, 1, -1), gsFormatoNumeroView)
'      End If
'   End If
'
'   If LTrim(RTrim(rs!cdoctpo)) = "7" Then
'       lvItem.SubItems(6) = Format(Val(rs!nDocImporte), gsFormatoNumeroView)
'       lvItem.SubItems(7) = rs!cObjetoCod
'   Else
'       lvItem.SubItems(6) = Format(Val(rs!nDocImporte * IIf(Val(rs!nDocImporte) > 0, 1, -1)), gsFormatoNumeroView)
'       lvItem.SubItems(7) = rs!cObjetoCod
'   End If
'
''  lvItem.SubItems(8) = rs!cMovNro
'   If rs!cFecProvi <> "" Then
''        lvItem.SubItems(8) = Mid(rs!cFecProvi, 7, 2) & "/" & Mid(rs!cFecProvi, 5, 2) & "/" & Left(rs!cFecProvi, 4)
'        lvItem.SubItems(8) = rs!cFecProvi
'   Else
'        lvItem.SubItems(8) = ""
'   End If
'   lvItem.SubItems(9) = rs!cdoctpo
'   lvItem.SubItems(10) = rs!cDocNro
'   lvItem.SubItems(11) = rs!Voucher & ""
'   If rs!cMovPago <> "" Then
'        lvItem.SubItems(12) = Mid(rs!cMovPago, 7, 2) & "/" & Mid(rs!cMovPago, 5, 2) & "/" & Left(rs!cMovPago, 4) & " - " & Right(rs!cMovPago, 4)
'   Else
'        lvItem.SubItems(12) = ""
'   End If
'   rs.MoveNext
'Loop
'rs.Close
'lvProvis.ListItems(1).Selected = True
'txtMovDesc = lvProvis.ListItems(1).SubItems(5)
'prg.Visible = False
'End Sub

Private Sub GeneraListado(pnOpt As Integer)
Dim sCond As String
Dim nItem As Integer
Dim lvItem As ListItem
Dim nImporte As Currency, nTipCambio As Currency
Dim sTpoDoc  As String
Dim sCondPag As String

sCondPag = ""
Select Case pnOpt
   Case 0
     If chk1.value = 0 Then '' si check cancelados no esta seleccionado
        sCond = " and h.cOpeDocMetodo <> '2' and Ent.cmovnro is null "
     Else
        sCond = " and h.cOpeDocMetodo <> '2' and Ent.cmovnro is not null "
     End If
      sTpoDoc = "Provisionados"
   Case 1
      sCond = " and (NOT EXISTS (SELECT h.nMovNro FROM  MovRef h Inner Join Mov Mh On h.nmovnro = Mh.nMovNro WHERE h.nMovNroRef = a.nMovNro And Mh.nMovFlag <> 1) " _
            & "      OR  EXISTS (SELECT h.nMovNro " _
            & "                  FROM  MovRef h JOIN MOV MX on MX.nMovNro = h.nMovNro  " _
            & "                  WHERE h.nMovNroRef = a.nMovNro AND LEFT(MX.cMovNro,8)>'" & Format(txtFechaAl, "yyyymmdd") & "') ) " _
            & " and c.nMovImporte < 0 and h.cOpeDocMetodo <> '2' "
      
            
      sTpoDoc = "Pendientes"
   Case 2
      sCond = " and c.nMovImporte > 0 and b.nDocTpo IN (" & TpoDocOrdenPago & "," & TpoDocCheque & "," & TpoDocCarta & "," & TpoDocNotaAbono & ") "
      sTpoDoc = "Emitidos"
      sCondPag = " and LEFT(mEnt.cMovNro,8) BETWEEN '" & Format(CDate(txtFechaDel), gsFormatoMovFecha) & "' and '" & Format(CDate(txtFechaAl), gsFormatoMovFecha) & "' "
   Case 3
      sCond = " and c.nMovImporte > 0 and b.nDocTpo IN (" & TpoDocOrdenPago & "," & TpoDocCheque & "," & TpoDocCarta & "," & TpoDocNotaAbono & ") " _
            & " " _
            & " "
      sTpoDoc = "Entregados"
End Select

Dim lsCtaProvis
If lMN Then
    lsCtaProvis = "'25160201','25160202','251601'"
Else
    lsCtaProvis = "'25260201','25260202','252601'"
End If

'*** PEAC 20110607
'sSql = "SELECT DISTINCT b.dDocFecha, Ent.cMovNro cMovEntrega, g.cDocAbrev, b.nDocTpo cDocTpo, b.cDocNro, e.cPersNombre cNomPers, a.cMovDesc, d.cPersCod cObjetoCod, " _
     & "  Ruc=isnull((select cPersIDnro from persid i where cPersIDTpo='2' and e.cperscod=i.cperscod),''), " _
     & "       a.cMovNro as cFecProvi, c.nMovImporte * -1 as nDocImporte " & IIf(lMN, "", ", isnull(ME.nMovMEImporte,0) * -1 as nDocMEImporte ") _
     & ", (Select mdd.cDocNro from movdoc mdd where mdd.nmovnro = a.nmovnro And mdd.nDocTpo = " & TpoDocVoucherEgreso & ") Voucher, ISNULL(ENT.CMOVNRO,'') as cMovPago " _
     & "FROM   Mov a " _
     & "             JOIN MovDoc b ON b.nMovNro = a.nMovNro " _
     & "             JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(lMN, "", "LEFT JOIN MovME ME ON (ME.nMovNro=c.nMovNro and ME.nMovItem=c.nMovItem)") _
     & "             JOIN MovGasto d ON d.nMovNro = a.nMovNro " _
     & "             JOIN Persona e ON e.cPersCod = d.cPersCod " _
     & "        LEFT JOIN OpeDoc h ON h.nDocTpo = b.nDocTpo and h.cOpeCod = '" & gsOpeCod & "' LEFT JOIN Documento g ON g.nDocTpo = b.nDocTpo " _
     & "        LEFT JOIN ( SELECT  max(mEnt.cMovNro) cMovNro, rEnt.nMovNroRef " _
     & "                    FROM    MovRef rEnt " _
     & "                            JOIN Mov mEnt ON mEnt.nMovNro = rEnt.nMovNro " _
     & "                    WHERE mEnt.nMovFlag = 0 " & sCondPag & " GRoup by rEnt.nMovNroRef " _
     & "                  ) Ent ON Ent.nMovNroRef = a.nMovNro " _
     & "WHERE  a.nMovEstado = '" & gMovEstContabMovContable & "' and a.nMovFlag <> '" & gMovFlagEliminado & "' " _
     & IIf(pnOpt = 3, " and not Ent.cMovNro is NULL ", " and not h.nDocTpo is NULL and LEFT(a.cMovNro,8) BETWEEN '" & Format(CDate(txtFechaDel), gsFormatoMovFecha) & "' and '" & Format(CDate(txtFechaAl), gsFormatoMovFecha) & "' ") _
     & "       and c.cCtaContCod IN (" & lsCtaProvis & ") " _
     & "       " & sCond _
     & "       ORDER BY a.cMovNro, e.cPersNombre "

sSql = "exec stp_sel_ObtieneProvisionProveedores '" & IIf(lMN, "1", "2") & "','" & gsOpeCod & "'," & pnOpt & ",'" & IIf(Len(sCondPag) > 0, "1", "0") & "','" & Format(CDate(txtFechaDel), gsFormatoMovFecha) & "','" & Format(CDate(txtFechaAl), gsFormatoMovFecha) & "','" & IIf(chk1.value = 0, "0", "1") & "'"
'*** FIN PEAC

Set rs = oCon.CargaRecordSet(sSql)
lSalir = False
If RSVacio(rs) Then
   MsgBox "No existen Comprobantes " & sTpoDoc, vbInformation, "¡Aviso!"
   lSalir = True
   Exit Sub
End If
lvProvis.ListItems.Clear
prg.Visible = True
prg.Min = 0
prg.Max = rs.RecordCount
Do While Not rs.EOF

   prg.value = rs.Bookmark
   Status.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
   nItem = nItem + 1
   Set lvItem = lvProvis.ListItems.Add(, , Format(nItem, "000"))
   
   lvItem.SmallIcon = 1
   lvItem.SubItems(1) = Mid(rs!cDocAbrev & Space(3), 1, 3) & " " & rs!cDocNro
   
   'lvitem.ListSubItems.Count
   
   If pnOpt = 3 Then
      lvItem.SubItems(2) = GetFechaMov(rs!cMovEntrega, True)
   Else
      lvItem.SubItems(2) = rs!dDocFecha
   End If
   lvItem.SubItems(3) = rs!cNomPers
   lvItem.SubItems(4) = rs!RUC
   lvItem.SubItems(5) = rs!cMovDesc
   nImporte = rs!nDocImporte
   If Not lMN Then
      nImporte = rs!nDocMEImporte
      If nImporte = 0 Then
         nImporte = Round(rs!nDocImporte / gnTipCambio, 2)
      End If
      If LTrim(RTrim(rs!cdoctpo)) = "7" Then
          lvItem.SubItems(13) = Format(Val(nImporte), gsFormatoNumeroView)
      Else
          lvItem.SubItems(13) = Format(nImporte * IIf(Val(rs!nDocImporte) > 0, 1, -1), gsFormatoNumeroView)
      End If
   End If
   
   
   If LTrim(RTrim(rs!cdoctpo)) = "7" Then
       lvItem.SubItems(6) = Format(Val(rs!nDocImporte), gsFormatoNumeroView)
       lvItem.SubItems(7) = rs!cObjetoCod
   Else
       lvItem.SubItems(6) = Format(Val(rs!nDocImporte * IIf(Val(rs!nDocImporte) > 0, 1, -1)), gsFormatoNumeroView)
       lvItem.SubItems(7) = rs!cObjetoCod
   End If

'  lvItem.SubItems(8) = rs!cMovNro
   If rs!cFecProvi <> "" Then
'        lvItem.SubItems(8) = Mid(rs!cFecProvi, 7, 2) & "/" & Mid(rs!cFecProvi, 5, 2) & "/" & Left(rs!cFecProvi, 4)
        lvItem.SubItems(8) = rs!cFecProvi
   Else
        lvItem.SubItems(8) = ""
   End If
   lvItem.SubItems(9) = rs!cdoctpo
   lvItem.SubItems(10) = rs!cDocNro
   lvItem.SubItems(11) = rs!Voucher & ""
   If rs!cMovPago <> "" Then
        lvItem.SubItems(12) = Mid(rs!cMovPago, 7, 2) & "/" & Mid(rs!cMovPago, 5, 2) & "/" & Left(rs!cMovPago, 4) & " - " & Right(rs!cMovPago, 4)
   Else
        lvItem.SubItems(12) = ""
   End If
   
   '*** PEAC 20110609
   If pnOpt = 1 And lMN Then
        lvItem.SubItems(13) = rs!cCtaContCod
   End If
   '*** FIN PEAC
   
   rs.MoveNext
Loop
rs.Close
lvProvis.ListItems(1).Selected = True
txtMovDesc = lvProvis.ListItems(1).SubItems(5)
prg.Visible = False
End Sub

Sub saldodiario()
Dim sql As String
Dim lSalirA As Boolean
Dim lSalirB As Boolean
Dim lnSaldoProv As Currency
Dim lnSaldoCanc As Currency
Dim lnSaldo As Currency
Dim sCondPag As String

Dim lsSaldoProv As String
Dim lsSaldoCanc As String
Dim lsSaldo As String


sCondPag = ""
Dim lsCtaProvis As String

If lMN Then
    lsCtaProvis = "'25160109','25160209'"
Else
    lsCtaProvis = "'25260109','25260209'"
End If

sql = "SELECT  IsNull(sum(c.nMovImporte),0) * -1 as lnSaldoProv " & IIf(lMN, "", ", isnull(sum(ME.nMovMEImporte),0) * -1 as lnSaldoProvME ") _
     & "FROM   Mov a " _
     & "             JOIN MovDoc b ON b.nMovNro = a.nMovNro " _
     & "             JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(lMN, "", "LEFT JOIN MovME ME ON (ME.nMovNro=c.nMovNro and ME.nMovItem=c.nMovItem)") _
     & "             JOIN MovGasto d ON d.nMovNro = a.nMovNro " _
     & "             JOIN Persona e ON e.cPersCod = d.cPersCod " _
     & "        LEFT JOIN OpeDoc h ON h.nDocTpo = b.nDocTpo and h.cOpeCod = '" & gsOpeCod & "' LEFT JOIN Documento g ON g.nDocTpo = b.nDocTpo " _
     & "        LEFT JOIN ( SELECT  max(mEnt.cMovNro) cMovNro, rEnt.nMovNroRef " _
     & "                    FROM    MovRef rEnt " _
     & "                            JOIN Mov mEnt ON mEnt.nMovNro = rEnt.nMovNro " _
     & "                    WHERE mEnt.nMovFlag = 0 " & sCondPag & " GRoup by rEnt.nMovNroRef " _
     & "                  ) Ent ON Ent.nMovNroRef = a.nMovNro " _
     & "WHERE  a.nMovEstado = '" & gMovEstContabMovContable & "' and a.nMovFlag <> '" & gMovFlagEliminado & "' " _
     & " and not h.nDocTpo is NULL and LEFT(a.cMovNro,8) BETWEEN '" & Format(CDate(txtFechaDel), gsFormatoMovFecha) & "' and '" & Format(CDate(txtFechaAl), gsFormatoMovFecha) & "' " _
     & "       and c.cCtaContCod IN (" & lsCtaProvis & ") " _
     & "       and h.cOpeDocMetodo <> '2' and Ent.cmovnro is null"
Set rs = oCon.CargaRecordSet(sql)
lSalirA = False
If RSVacio(rs) Then
   MsgBox "No existen Comprobantes " & vbInformation, "¡Aviso!"
   lSalirA = True
   Exit Sub
End If
If lMN Then
    lnSaldoProv = rs!lnSaldoProv
Else
    lnSaldoProv = rs!lnsaldoProvME
End If
rs.Close

sql = "SELECT  sum(c.nMovImporte) * -1 as lnSaldocanc " & IIf(lMN, "", ", isnull(sum(ME.nMovMEImporte),0) * -1 as lnSaldocancME ") _
     & "FROM   Mov a " _
     & "             JOIN MovDoc b ON b.nMovNro = a.nMovNro " _
     & "             JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(lMN, "", "LEFT JOIN MovME ME ON (ME.nMovNro=c.nMovNro and ME.nMovItem=c.nMovItem)") _
     & "             JOIN MovGasto d ON d.nMovNro = a.nMovNro " _
     & "             JOIN Persona e ON e.cPersCod = d.cPersCod " _
     & "        LEFT JOIN OpeDoc h ON h.nDocTpo = b.nDocTpo and h.cOpeCod = '" & gsOpeCod & "' LEFT JOIN Documento g ON g.nDocTpo = b.nDocTpo " _
     & "        LEFT JOIN ( SELECT  max(mEnt.cMovNro) cMovNro, rEnt.nMovNroRef " _
     & "                    FROM    MovRef rEnt " _
     & "                            JOIN Mov mEnt ON mEnt.nMovNro = rEnt.nMovNro " _
     & "                    WHERE mEnt.nMovFlag = 0 " & sCondPag & " GRoup by rEnt.nMovNroRef " _
     & "                  ) Ent ON Ent.nMovNroRef = a.nMovNro " _
     & "WHERE  a.nMovEstado = '" & gMovEstContabMovContable & "' and a.nMovFlag <> '" & gMovFlagEliminado & "' " _
     & " and not h.nDocTpo is NULL and LEFT(a.cMovNro,8) BETWEEN '" & Format(CDate(txtFechaDel), gsFormatoMovFecha) & "' and '" & Format(CDate(txtFechaAl), gsFormatoMovFecha) & "' " _
     & "       and c.cCtaContCod IN (" & lsCtaProvis & ") " _
     & "       and h.cOpeDocMetodo <> '2' and Ent.cmovnro is not null"
Set rs = oCon.CargaRecordSet(sql)
lSalirA = False
If RSVacio(rs) Then
   MsgBox "No existen Comprobantes " & vbInformation, "¡Aviso!"
   lSalirA = True
   Exit Sub
End If
If lMN Then
    lnSaldoCanc = rs!lnSaldoCanc
Else
    lnSaldoCanc = rs!lnsaldocancME
End If
rs.Close

lnSaldo = lnSaldoProv - lnSaldoCanc
lsSaldoProv = Str(lnSaldoProv)
lsSaldoCanc = Str(lnSaldoCanc)
lsSaldo = Str(lnSaldo)

Dim sTexto As String
Dim oImp As NContImprimir
Dim lsFechadel As String
Dim lsFechaal As String

lsFechadel = Format(CDate(txtFechaDel), gsFormatoMovFecha)
lsFechaal = Format(CDate(txtFechaAl), gsFormatoMovFecha)

Set oImp = New NContImprimir
Me.Enabled = False

sTexto = oImp.ImprimeSaldoProvision(lsSaldoProv, lsSaldoCanc, lsSaldo, lsFechadel, lsFechaal)
'sTexto = oimp.ImprimeMayorCta(txtCtaCod, txtCtaDesc, CDate(txtFechaDel), CDate(txtFechaAl), nVal(txtImporte), cboFiltro, gnLinPage)
EnviaPrevio sTexto, "PROVISION DE PAGO A PROVEEDORES", gnLinPage, False
Me.Enabled = True
Set oImp = Nothing


    
End Sub

Private Sub cmdAsientos_Click()
Dim N As Integer
Dim lOk As Boolean
Dim lsImpre As String
Dim sMovs As String
Dim oContImp As NContImprimir

If lvProvis.ListItems.Count = 0 Then
   MsgBox "No existen elementos que Imprimir...!", vbInformation, "Error"
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
      sMovs = sMovs & ",'" & lvProvis.ListItems(N).SubItems(8) & "'"
   End If
Next
If sMovs <> "" Then
    sMovs = Mid(sMovs, 2, Len(sMovs))
    lsImpre = ImprimeAsientosContables(sMovs, prg, Status, "( DEL " & txtFechaDel & " AL " & txtFechaAl & ")")
    EnviaPrevio lsImpre, "ASIENTOS DE PROVISION DE PAGO A PROVEEDORES", gnLinPage, False
End If
prg.Visible = False
lvProvis.SetFocus
End Sub

Private Sub cmdImprimir_Click()
Dim N As Integer
Dim nLin As Integer, P As Integer
Dim nTot As Currency
Dim nTotME As Currency
Dim lOk As Boolean
Dim lsImpre As String

If lvProvis.ListItems.Count = 0 Then
   MsgBox "No existen elementos que Imprimir...!", vbInformation, "Error"
   Exit Sub
End If
nLin = 66
nTot = 0
prg.Min = 0
prg.Max = lvProvis.ListItems.Count
prg.Visible = True
Me.Enabled = False
CON = PrnSet("C+")
COFF = PrnSet("C-")
BON = PrnSet("B+")
BOFF = PrnSet("B-")
For N = 1 To lvProvis.ListItems.Count
   lOk = True
   If OptS(1).value Then
      lOk = IIf(lvProvis.ListItems(N).Selected, True, False)
   End If
   prg.value = N
   Status.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
   If lOk Then
      Linea lsImpre, CabeceraRepo(nLin, P), 0
      With lvProvis.ListItems(N)
        
         Linea lsImpre, CON & " " & Format(N, "000") & "  " & .SubItems(2) & " " & Mid(.SubItems(1) & Space(17), 1, 17) & " " & Mid(.SubItems(3) & Space(25), 1, 25) & " " & Mid(.SubItems(4) & Space(12), 1, 12) & " " & Left(.SubItems(8), 8) & " " & Mid(Replace(.SubItems(5), Chr(13) & oImpresora.gPrnSaltoLinea, " ") & Space(25), 1, 25) & " " & Right(Space(14) & .SubItems(6), 14), 0
         If Not lMN Then
            Linea lsImpre, " " & Right(Space(14) & .SubItems(13), 14), 0
         End If
         Linea lsImpre, Space(3) & .SubItems(12), 0
         
         If lMN Then '*** PEAC 20110613
            Linea lsImpre, IIf(Len(.SubItems(12)) = 0, Space(17), .SubItems(12)) & Space(3) & .SubItems(13), 0
         End If
         Linea lsImpre, COFF
         nTot = nTot + Val(Format(.SubItems(6), gsFormatoNumeroDato))
         If Not lMN Then
            nTotME = nTotME + Val(Format(.SubItems(13), gsFormatoNumeroDato))
         End If
      End With
   End If
   nLin = nLin + 1
Next
Linea lsImpre, CON & " ===================================================================================================================================================" & IIf(Not lMN, String(15, "="), "")
Linea lsImpre, BON & Space(107) & "TOTAL   " & Right(Space(14) & gcMN & " " & Format(nTot, gsFormatoNumeroView), 18) & IIf(Not lMN, Right(Space(14) & gcME & " " & Format(nTotME, gsFormatoNumeroView), 15), "") & BOFF & COFF & oImpresora.gPrnSaltoLinea
EnviaPrevio lsImpre, "Planilla de Provisiones", gnLinPage, False
prg.Visible = False
Me.Enabled = True
lvProvis.SetFocus
End Sub

Private Sub cmdProcesar_Click()
Me.lvProvis.ListItems.Clear
GeneraListado IIf(Opt(0).value, 0, IIf(Opt(1).value, 1, IIf(Opt(3).value, 3, 2)))
End Sub

Private Sub cmdRUC_Click()
Dim lsRuc As String
Dim I  As Integer
lsRuc = txtNroRuc.Text
                If txtNroRuc.Text = "" Then
                    MsgBox "El Numero de documento no puede estar Vacio", vbInformation, "El documento No puede estar Vacio"
                    txtNroRuc.SetFocus
                    Exit Sub
                End If
                If IsNumeric(txtNroRuc.Text) = False Then
                    MsgBox "El Numero de Ruc No es Numerico", vbInformation, "Numero de Ruc No es Numerico"
                    txtNroRuc.SetFocus
                    Exit Sub
                End If
                If Len(txtNroRuc.Text) < 8 Then
                    MsgBox "El Numero de Ruc Ingresado es menor Que 8 Digitos", vbInformation, "Numero de Ruc es Menor que 8 Digitos"
                    txtNroRuc.SetFocus
                    Exit Sub
                End If
                
                If Len(txtNroRuc.Text) > 11 Then
                    MsgBox "El Numero de Ruc Ingresado es mayor Que 11 Digitos", vbInformation, "El Numero de Ruc es mayor que 11 digitos"
                    txtNroRuc.SetFocus
                    Exit Sub
                End If
For I = 1 To lvProvis.ListItems.Count
        If Val(lvProvis.ListItems(I).SubItems(4)) = Val(Trim(lsRuc)) Then
           lvProvis.ListItems(I).Selected = True
           lvProvis.ListItems(I).Top = lvProvis.ListItems(I).Selected
           lvProvis.Refresh
           txtNroRuc.Text = ""
           
           Exit Sub
        Else
            lvProvis.ListItems(I).Selected = False
        End If
Next
End Sub

Private Sub cmdSaldo_Click()

saldodiario
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
If gsSimbolo = gcME Then
   lMN = False
   lvProvis.ColumnHeaders.Add , , "Importe M.E."
   lvProvis.ColumnHeaders(6).Text = "Importe M.N."
   lvProvis.ColumnHeaders(12).Alignment = lvwColumnRight
End If
Set oCon = New DConecta
oCon.AbreConexion
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
   txtMovDesc = lvProvis.ListItems(nPos).SubItems(5)
End If
End Sub

Private Sub lvProvis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nPos As Variant
If lvProvis.ListItems.Count > 0 Then
   nPos = lvProvis.SelectedItem.Index
   txtMovDesc = lvProvis.ListItems(nPos).SubItems(5)
End If
End Sub

Private Sub Opt_Click(Index As Integer)
lvProvis.ColumnHeaders(3).Text = "Emisión"
If Index = 3 Then
   lvProvis.ColumnHeaders(3).Text = "Entrega"
End If
End Sub

