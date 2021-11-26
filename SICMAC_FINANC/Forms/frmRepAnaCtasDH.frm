VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepAnaCtasDH 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisis de Movimientos"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   Icon            =   "frmRepAnaCtasDH.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Fecha al "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   6015
      TabIndex        =   10
      Top             =   60
      Width           =   2580
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   285
         Left            =   135
         TabIndex        =   11
         Top             =   270
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFF 
         Height          =   285
         Left            =   1335
         TabIndex        =   14
         Top             =   270
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9840
      TabIndex        =   9
      Top             =   5505
      Width           =   1125
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   8670
      TabIndex        =   6
      Top             =   5505
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Pendientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   90
      TabIndex        =   4
      Top             =   60
      Width           =   5940
      Begin VB.TextBox txtOpeDes 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1305
         TabIndex        =   5
         Top             =   240
         Width           =   4455
      End
      Begin Sicmact.TxtBuscar txtOpeCod 
         Height          =   345
         Left            =   60
         TabIndex        =   15
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         Appearance      =   0
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
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   9240
      TabIndex        =   1
      Top             =   60
      Width           =   1710
      Begin VB.OptionButton optMoneda 
         Caption         =   "Soles"
         Height          =   345
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "Dolares"
         Height          =   345
         Index           =   1
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton CmdGenerar 
      Caption         =   "&Generar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7500
      TabIndex        =   0
      Top             =   5505
      Width           =   1125
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Top             =   5550
      Visible         =   0   'False
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   503
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg 
      Height          =   4545
      Left            =   60
      TabIndex        =   8
      Top             =   870
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   8017
      _Version        =   393216
      Cols            =   8
      BackColorBkg    =   14737632
      AllowBigSelection=   0   'False
      TextStyleFixed  =   4
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   390
      Left            =   1200
      TabIndex        =   12
      Top             =   3750
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   688
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmRepAnaCtasDH.frx":030A
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
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   240
      Left            =   165
      SizeMode        =   1  'Stretch
      TabIndex        =   13
      Top             =   5895
      Visible         =   0   'False
      Width           =   3210
   End
End
Attribute VB_Name = "frmRepAnaCtasDH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSql     As String
Dim rs       As ADODB.Recordset
Dim pnMoneda As Integer
Dim lsCuentas() As String

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub FormatoFlex()
fg.RowHeight(-1) = 285
fg.ColWidth(0) = 300
fg.ColWidth(1) = 1100
fg.ColWidth(2) = 5000
fg.ColWidth(3) = 0
fg.ColWidth(4) = 1200
fg.ColWidth(5) = 1200
fg.ColWidth(6) = 3000
fg.ColWidth(7) = 1500
fg.ColAlignment(4) = 7
fg.ColAlignment(5) = 7
fg.ColAlignment(1) = 1
fg.ColAlignment(2) = 1
fg.TextMatrix(0, 0) = "#"
fg.TextMatrix(0, 1) = "Fecha"
fg.TextMatrix(0, 2) = "Descripción"
fg.TextMatrix(0, 3) = "CtaCont"
fg.TextMatrix(0, 4) = "Debe"
fg.TextMatrix(0, 5) = "Haber"
fg.TextMatrix(0, 6) = "Persona"
fg.TextMatrix(0, 7) = "Documento"

End Sub

Private Function GeneraRepPendSubsidios(psOpeCod As String, pnMoneda As Integer, pdFecha As Date) As ADODB.Recordset
'SELECT m.cMovNro, m.cMovDesc, '" & sCta & "' cCtaContCod, mc.nMovImporte, mc.nMovImporte nSaldo, ISNULL(p.cNomPers,'') cPersona, md.cDocNro
    sSql = "SELECT pd.cMovNro, pla.cPlaInsDes cMovDesc, cc.cCtaContCod, pd.nMonto nMovImporte, ISNULL(mRend.nSaldo, pd.nMonto) nSaldo, 'PLA' cDocAbrev, '' cDocTpo, pla.cPlaCod cDocNro, Convert( varchar(10), pd.dPlaInsCod) dDocFecha, " _
         & "     p.cNomPers cPersona, '00' + p.cCodPers cCodPers " _
         & "FROM planilladetalle pd JOIN planillains pla ON pla.dPlaInsCod = pd.dPlaInsCod " _
         & "     JOIN conceptocta cc ON cc.cConcepCod = pd.cConcepCod " _
         & "     JOIN persona p ON p.cCodPers = pd.cCodPers " _
         & "     LEFT JOIN (SELECT m.cMovNro, cMovNroRef FROM MovRef mr JOIN Mov m ON m.cMovNro = mr.cMovNro WHERE m.cMovEstado = '0' and m.cMovFlag <> 'X') ref ON ref.cMovNroRef = pd.cMovNro " _
         & "     LEFT JOIN MovPendientesRend mRend ON mRend.cMovNro = pd.cMovNro, OpeCta oc " _
         & "WHERE not pd.cMovNro is NULL and LEFT(pd.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and cc.cOpeCod like '622601' and oc.cOpeCod = '" & psOpeCod & "' and cc.cCtaContCod = LEFT(oc.cCtaContCod,2)+'" & pnMoneda & "'+SubString(oc.cCtaContCod,4,22) " _
         & "and pd.cPlaCod = 'PE007' and ( ref.cMovNro is NULL or mRend.nSaldo <> 0 ) " _
         & "UNION ALL " _
         & "SELECT pd.cMovNro, pla.cPlaInsDes cMovDesc, cc.cCtaContCod, pd.nMonto nMovImporte, ISNULL(mRend.nSaldo, pd.nMonto) nSaldo, 'PLA' cDocAbrev, '' cDocTpo, pla.cPlaCod cDocNro, Convert( varchar(10), pd.dPlaInsCod) dDocFecha, " _
         & "     p.cNomPers cPersona, '00' + p.cCodPers cCodPers " _
         & "FROM planilladetalle pd JOIN planillains pla ON pla.dPlaInsCod = pd.dPlaInsCod " _
         & "     JOIN conceptocta cc ON cc.cConcepCod = pd.cConcepCod " _
         & "     JOIN persona p ON p.cCodPers = pd.cCodPers " _
         & "     LEFT JOIN (SELECT m.cMovNro, cMovNroRef FROM MovRef mr JOIN Mov m ON m.cMovNro = mr.cMovNro WHERE m.cMovEstado = '0' and m.cMovFlag <> 'X') ref ON ref.cMovNroRef = pd.cMovNro " _
         & "     LEFT JOIN MovPendientesRend mRend ON mRend.cMovNro = pd.cMovNro, OpeCta oc " _
         & "WHERE not pd.cMovNro is NULL and LEFT(pd.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and cc.cOpeCod like '622602' and oc.cOpeCod = '" & psOpeCod & "' and cc.cCtaContCod = LEFT(oc.cCtaContCod,2)+'" & pnMoneda & "'+SubString(oc.cCtaContCod,4,22) " _
         & "and pd.cPlaCod = 'PE012' and ( ref.cMovNro is NULL or mRend.nSaldo <> 0 ) " _
         & "ORDER BY cc.cCtaContCod, pd.cMovNro "
'Set GeneraRepPendSubsidios = CargaRecord(sSql)
Exit Function
End Function

Private Sub GeneraReporte()
Dim sCtaCod As String
Dim nRow As Integer
Dim nTot As Currency
Dim nSdo As Currency
CmdGenerar.Enabled = False
MousePointer = 11
fg.MousePointer = 11
   nTot = 0
   nSdo = 0
   prg.Visible = True
   Select Case txtOpeCod
       Case gOpePendOpeAgencias
            'Set rs = GeneraRepAnalisisPendientesInterAgencias(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendRendirCuent
            'Set rs = GeneraRepEntregaARendir(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendOrdendePago
            'Set rs = GeneraRepOrdenPago(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendSobrantCaja
            'Set rs = GeneraRepAnalisisPendientesSobrante(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendFaltantCaja
            'Set rs = GeneraRepAnalisisPendientesFaltante(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendOPCertifica
            'Set rs = GeneraRepAnalisisPendientesOPCertificada(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendPagoSubsidi
            'Set rs = GeneraRepPendSubsidios(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendCobraLiquid
            'Set rs = GeneraRepAnalisisPendientesCobServicios(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendOtrasOpeLiqPas
            'Set rs = GeneraRepAnalisisPendOtraOpeLiqCajaGeneral(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendProvPagoProv
            'Set rs = GeneraRepProvisionPagoProveedor(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendRecursHuman
            Set rs = GeneraRepAnalisisPendientesRRHH(txtOpeCod, pnMoneda, txtFecha, Me.mskFF.Text)
       Case Else
            'Set rs = GeneraRepOtrasPendientesContables(txtOpeCod, pnMoneda, txtFecha, "D")
   End Select
   fg.Rows = 2
   EliminaRow fg, 1
If Not rs Is Nothing Then
If rs.State = adStateOpen Then
   If Not rs.EOF Then
      sCtaCod = ""
      prg.Max = rs.RecordCount
      Do While Not rs.EOF
         prg.value = rs.Bookmark
         If sCtaCod <> rs!cCtaContCod Then
            If sCtaCod <> "" Then
               AdicionaRow fg
               nRow = fg.row
               BackColorFg fg, "&H00E0E0E0"
               fg.TextMatrix(nRow, 2) = "TOTALES : "
               fg.TextMatrix(nRow, 4) = Format(nTot, gsFormatoNumeroView)
               fg.TextMatrix(nRow, 5) = Format(nSdo, gsFormatoNumeroView)
               nTot = 0
               nSdo = 0
            End If
            AdicionaRow fg
            nRow = fg.row
            BackColorFg fg, "&H00FFFFC0"
            fg.TextMatrix(nRow, 1) = "Cuenta : "
            fg.TextMatrix(nRow, 2) = rs!cCtaContCod & " - " '& CuentaNombre(rs!cCtaContCod)
            sCtaCod = rs!cCtaContCod
         End If
         AdicionaRow fg
         nRow = fg.row
         fg.TextMatrix(nRow, 1) = Left(rs!cMovNro, 8) & "-" & Mid(rs!cMovNro, 19, 24)
         fg.TextMatrix(nRow, 2) = rs!cMovDesc
         fg.TextMatrix(nRow, 3) = rs!cCtaContCod
         fg.TextMatrix(nRow, 4) = Format(rs!nMovImporte, gsFormatoNumeroView)
         fg.TextMatrix(nRow, 5) = Format(rs!nSaldo, gsFormatoNumeroView)
         fg.TextMatrix(nRow, 6) = PstaNombre(rs!cPErsona & "", False)
         'fg.TextMatrix(nRow, 7) = rs!cDocAbrev & " " & rs!cDocNro
         fg.TextMatrix(nRow, 7) = rs!cDocNro & ""
         nTot = nTot + rs!nMovImporte
         nSdo = nSdo + rs!nSaldo
         rs.MoveNext
      Loop
      AdicionaRow fg
      nRow = fg.row
      BackColorFg fg, "&H00E0E0E0"
      fg.TextMatrix(nRow, 2) = "TOTALES : "
      fg.TextMatrix(nRow, 4) = Format(nTot, gsFormatoNumeroView)
      fg.TextMatrix(nRow, 5) = Format(nSdo, gsFormatoNumeroView)
      prg.value = 0
      prg.Visible = False
   End If
End If
End If
RSClose rs
MousePointer = 0
fg.MousePointer = 0
CmdGenerar.Enabled = True
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaManCtaPend
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Genero el Reporte de Pendientes Detallado : " & txtOpeDes.Text & "hasta la Fecha " & txtFecha.Text
            Set objPista = Nothing
            '*******
End Sub

Private Function GeneraRepOrdenPago(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As Recordset
'Dim prs  As ADODB.Recordset
'Dim sDH  As String
'Dim sCta As String
'Dim lsCadSer As String
'Dim lnI As Integer
'Dim sql As String
'Dim rsAge As ADODB.Recordset
'Set rsAge = New ADODB.Recordset
'Dim rsR As ADODB.Recordset
'Set rsR = New ADODB.Recordset
'
'On Error GoTo GeneraRepOrdenPagoErr
'Set rs = CargaOpeCta(psOpeCod, , , True)
'
'
'sql = " Select cValor From TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
'rsAge.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'lnI = 0
'ReDim Preserve lsCuentas(15, 2)
'sSql = ""
'While Not rsAge.EOF
'    If AbreConeccion(Right(rsAge!cValor, 2)) Then
'        If pnMoneda = 1 Then
'            sql = " Select cValorVar from Varsistema Where cCodProd = 'AHO' And cNomVar = 'cCtaFonFijMN'"
'        Else
'            sql = " Select cValorVar from Varsistema Where cCodProd = 'AHO' And cNomVar = 'cCtaFonFijME'"
'        End If
'
'        If rsR.State = 1 Then rsR.Close
'        rsR.Open sql, dbCmactN, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'        If Not (rsR.EOF And rs.BOF) Then
'
'            If Right(rsAge!cValor, 2) = Left(rsR!cValorVar, 2) Then
'                lnI = lnI + 1
'                lsCuentas(lnI, 1) = Trim(rsR!cValorVar)
'                lsCuentas(lnI, 2) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
'
'                If sSql <> "" Then sSql = sSql & " UNION "
'
'                If pnMoneda = 1 Then
'                   sCta = "21110502" & Right(rsAge!cValor, 2)
'                Else
'                   sCta = "21210502" & Right(rsAge!cValor, 2)
'                End If
'
'                sSql = sSql & " SELECT m.cMovNro, m.cMovDesc, '" & sCta & "' cCtaContCod, Case Substring(mc.cCtaContCod,3,1) When '2' Then Abs(me.nMovMEImporte) Else Abs(mc.nMovImporte) End nMovImporte, mc.nMovImporte nSaldo, ISNULL(p.cNomPers,'') cPersona, 'OP' cDocAbrev,  md.cDocNro " _
'                     & " FROM   mov m JOIN movdoc md ON md.cMovNro = m.cMovNro " _
'                     & "             JOIN movcta mc ON mc.cMovNro = m.cMovNro " _
'                     & "        LEFT JOIN MovObj mo ON mo.cMovNro =  mc.cMovNro and mo.cObjetoCod LIKE '00%' " _
'                     & "        LEFT JOIN Persona p ON '00'+p.cCodPers = mo.cObjetoCod " _
'                     & "        LEFT JOIN MovME ME On mc.cMovNro = ME.cMovNro And mc.cMovItem = ME.cMovItem " _
'                     & "        LEFT JOIN " & lsCuentas(lnI, 2) & "VistaOrdenPago OPG On convert(Int,cDocNro) = OPG.nNumOp And OPG.dFecOP <= '" & Format(pdFecha, gsFormatoFecha) & " 23:59:59' And OPG.cEstOP = '4'  And OPG.cCodCta = '" & lsCuentas(lnI, 1) & "'" _
'                     & " WHERE LEFT(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.cMovEstado = '0' And" _
'                     & " ((Not m.cMovFlag in ('X','E','N')) Or (m.cMovFlag = 'E' And NOT EXISTS (SELECT MR.cMovNro FROM MovRef MR Inner Join Mov MRM On MRM.cMovNro = MR.cMovNro WHERE MR.cMovNroRef = m.cMovnro And LEFT(MRM.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' And MRM.cMovFlag in ('X','E','N') ) ) )" _
'                     & " and md.cDocTpo = '" & gcDocTpoOPago & "' " _
'                     & " and mc.cCtaContCod LIKE '__" & pnMoneda & "%' " _
'                     & " and mc.nMovImporte < 0 And OPG.nNumOp Is Null And convert(Int,cDocNro) Not In (Select Convert(Int,cDocNro) from movpendientesdet Where bExcluyeDoc = 1) And mc.cCtaContCod = '" & sCta & "' "
'            End If
'
'        End If
'    End If
'
'    rsAge.MoveNext
'Wend
'    sSql = sSql & " ORDER BY cCtaContCod, m.cmovnro "
'
'
'If Not rs.EOF Then
'   sCta = rs!cCtaContCod
'End If
'RSClose rs
'
' Set GeneraRepOrdenPago = CargaRecord(sSql)
'Exit Function
'GeneraRepOrdenPagoErr:
'   MsgBox Err.Description, vbInformation, "¡Aviso!"
'End Function
'
'Private Function GeneraRepProvisionPagoProveedor(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
''SELECT m.cMovNro, m.cMovDesc, '" & sCta & "' cCtaContCod, Case Substring(mc.cCtaContCod,3,1) When '2' Then Abs(me.nMovMEImporte) Else Abs(mc.nMovImporte) End nMovImporte, mc.nMovImporte nSaldo, ISNULL(p.cNomPers,'') cPersona, 'OP' cDocAbrev,  md.cDocNro
'
'sSql = "SELECT a.cMovNro, a.cMovDesc, c.cCtaContCod, ISNULL(me.nMovMEImporte,c.nMovImporte) nMovImporte, ISNULL(me.nMovMEImporte,c.nMovImporte) nSaldo, b.dDocFecha, e.cNomPers cPersona, d.cObjetoCod cCodPers, Doc.cDocAbrev, b.cDocNro, b.cDocTpo " _
'     & "FROM   Mov a JOIN MovDoc b ON b.cMovNro = a.cMovNro JOIN Documento Doc ON Doc.cDocTpo = b.cDocTpo " _
'     & "             JOIN MovCta c ON c.cMovNro = a.cMovNro LEFT JOIN MovMe me ON me.cMovNro = c.cMovNro and me.cMovItem = c.cMovItem " _
'     & "             JOIN MovObj d ON d.cMovNro = c.cMovNro and d.cMovItem = c.cMovItem " _
'     & "        LEFT JOIN (SELECT h.cMovNro, h.cMovNroRef FROM MovRef h JOIN Mov M ON m.cMovNro = h.cMovNro " _
'     & "                   WHERE LEFT(h.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.cMovFlag NOT IN ('X','E','N') " _
'     & "                  ) ref ON  ref.cMovNroRef = a.cMovNro " _
'     & "        LEFT JOIN (SELECT mr.cMovNro, mr.cMovNroRef, mc.cCtaContCod FROM Mov m1 JOIN MovRef mr ON mr.cMovNroRef = m1.cMovNro " _
'     & "                      JOIN MovCta mc ON mc.cMovNro = m1.cMovNro " _
'     & "                   WHERE m1.cMovFlag NOT IN ('X','E','N') and m1.cMovEstado = '0' " _
'     & "                  ) RefP ON RefP.cMovNro = a.cMovNro and RefP.cCtaContCod = c.cCtaContCod " _
'     & "       ,dbPersona.dbo.Persona e, dbComunes.dbo.OpeCta f, " _
'     & "        OpeDoc h " _
'     & "WHERE  LEFT(a.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and a.cMovEstado = '0' and a.cMovFlag NOT IN ('X','E','N') and f.cOpeCod = '" & psOpeCod & "' and h.cOpeCod = f.cOpeCod and " _
'     & "       c.cCtaContCod = LEFT(f.cCtaContCod,2) + '" & pnMoneda & "'+SubString(f.cCtaContCod,4,22) and " _
'     & "       h.cOpeDocMetodo NOT IN ('2')  and cOpeDocEstado not in('12') and ref.cMovNro IS NULL and " _
'     & "       SUBSTRING(d.cObjetoCod, 3, 10) = e.cCodPers And h.cDocTpo = b.cDocTpo " _
'     & "ORDER BY c.cCtaContCod, a.cMovNro"
'
'Set GeneraRepProvisionPagoProveedor = CargaRecord(sSql)

End Function


Private Function GeneraRepAnalisisPendOtraOpeLiqCajaGeneral(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
''A RENDIR CUENTA CON CTACONT 29180706
''-------------------------------------
'
'Dim prs    As ADODB.Recordset
'Dim rsAge  As ADODB.Recordset
'Dim prsDat As ADODB.Recordset
'
'Dim sCta As String
'Dim nCol As Integer
'
'On Error GoTo GeneraRepAnalisisPendOtraOpeLiqCajaGeneralErr
'
'Set prsDat = New ADODB.Recordset
'Set prs = CargaOpeCta(psOpeCod, "H", "0", True)
'If Not prs.EOF Then
'   sCta = prs!cCtaContCod
'End If
'sCta = Left(sCta, 2) & pnMoneda & Mid(sCta, 4, 22)
'
'sSql = "SELECT cValor FROM TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
'Set rsAge = CargaRecord(sSql)
'
'ReDim Preserve lsCuentas(15, 2)
'sSql = ""
'While Not rsAge.EOF
'    If AbreConeccion(Right(rsAge!cValor, 2)) Then
'        sSql = "SELECT convert(varchar(8),td.dFecTran,112) + convert(varchar(20),td.nNumTran) cMovNro, ISNULL(oo.cGlosa,'') cMovdesc, " & sCta & " cCtaContCod, " _
'             & "   td.nMonTran nMovImporte, td.nMonTran nSaldo, '' cDocAbrev, '' cDocTpo, td.cNumDoc cDocNro, Convert(VarChar(10), td.dFecTran,103) dDocFecha, " _
'             & "      ISNULL(p.cNomPers,'') cPersona, '00' + p.cCodPers cCodPers " _
'             & "FROM TranDiariaConsol td JOIN TransRef tr ON tr.cNroTranRef = td.nNumTran " _
'             & "     JOIN OtrasOpe oo ON oo.nNumTran = td.nNumTran " _
'             & "     LEFT JOIN Persona p ON p.cCodPers = oo.cCodPers " _
'             & "WHERE datediff(d,td.dFecTran,'" & Format(pdFecha, gsFormatoFecha) & "') >= 0 and SubString(td.cCodCta,6,1) = '" & pnMoneda & "' and td.cCodOpe in (SELECT cCodOpe FROM OpeCuentaN WHERE cCodCnt LIKE '29_80706') " _
'             & "  and td.cFlag is NULL and (cNroTran is NULL or LEFT(cNroTran,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "' ) " _
'             & "ORDER BY td.dFecTran "
'        Set prs = CargaRecordRemoto(sSql)
'        AdicionaRecordSet prsDat, prs
'        RSClose prs
'        CierraConeccion
'    End If
'    rsAge.MoveNext
'Wend
'RSClose rsAge
'
'sSql = "SELECT mc.cMovNro , m.cMovDesc, mc.cCtaContCod, mc.nMovImporte, mRend.nSaldo, ISNULL(d.cDocAbrev,'') cDocAbrev, ISNULL(md.cDocTpo,'') cDocTpo, ISNULL(md.cDocNro,'') cDocNro, ISNULL(md.dDocFecha,'') dDocFecha, " _
'     & "       ISNULL(p.cNomPers,'') cPersona, ISNULL(mop.cObjetoCod,'') cCodPers " _
'     & "       " _
'     & "FROM mov m join MovCta mc on m.cMovNro = mc.cMovNro " _
'     & "      LEFT JOIN MovMe  me ON me.cMovNro = mc.cMovNro and me.cMovItem = mc.cMovItem " _
'     & "      LEFT JOIN MovDoc md ON md.cMovNro = m.cMovNro " _
'     & "      LEFT JOIN dbComunes.dbo.Documento d ON d.cDocTpo = md.cDocTpo " _
'     & "      LEFT JOIN MovObj moP ON moP.cMovNro = mc.cMovNro and moP.cMovItem = mc.cMovItem and moP.cObjetoCod LIKE '00%' " _
'     & "      LEFT JOIN dbPersona.dbo.Persona P ON '00'+p.cCodPers = moP.cObjetoCod " _
'     & "           JOIN MovPendientesRend mRend ON mRend.cMovNro = m.cMovNro and mRend.cCtaContCod = mc.cCtaContCod " _
'     & "WHERE LEFT(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.cMovEstado= '0' and m.cMovFlag <> 'X' " _
'     & "      and nMovImporte  <  0 and mc.cCtaContCod =  '" & sCta & "' " _
'     & "     and mRend.nSaldo <> 0 " _
'     & "ORDER BY mc.cCtaContCod, m.cMovNro "
'Set prs = CargaRecord(sSql)
'AdicionaRecordSet prsDat, prs
'
'If Not prsDat Is Nothing Then
'    If Not prsDat.State = adStateClosed Then
'        prsDat.MoveFirst
'    End If
'End If
'If sSql = "" Then Exit Function
'Set GeneraRepAnalisisPendOtraOpeLiqCajaGeneral = prsDat
'Exit Function
'GeneraRepAnalisisPendOtraOpeLiqCajaGeneralErr:
'    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function


Private Function GeneraRepOtrasPendientesContables(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date, psTpoCta As String) As ADODB.Recordset
'    sSql = "SELECT mc.cMovNro, m.cMovDesc, mc.cCtaContCod, ISNULL(me.nMovMEImporte,mc.nMovImporte) nMovImporte, mRend.nSaldo, ISNULL(d.cDocAbrev,'') cDocAbrev, ISNULL(md.cDocTpo,'') cDocTpo, ISNULL(md.cDocNro,'') cDocNro, ISNULL(md.dDocFecha,'') dDocFecha, " _
'         & "       ISNULL(p.cNomPers,'') cPersona, ISNULL(mop.cObjetoCod,'') cCodPers " _
'         & "       " _
'         & "FROM mov m join MovCta mc on m.cMovNro = mc.cMovNro " _
'         & "      LEFT JOIN MovMe  me ON me.cMovNro = mc.cMovNro and me.cMovItem = mc.cMovItem " _
'         & "      LEFT JOIN MovDoc md ON md.cMovNro = m.cMovNro LEFT JOIN Documento d ON d.cDocTpo = md.cDocTpo " _
'         & "      LEFT JOIN MovObj moP ON moP.cMovNro = mc.cMovNro and moP.cMovItem = mc.cMovItem and moP.cObjetoCod LIKE '00%' " _
'         & "      LEFT JOIN Persona P ON '00'+p.cCodPers = moP.cObjetoCod " _
'         & "           JOIN MovPendientesRend mRend ON mRend.cMovNro = m.cMovNro and mRend.cCtaContCod = mc.cCtaContCod, OpeCta oc " _
'         & "WHERE LEFT(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.cMovEstado= '0' and m.cMovFlag <> 'X' " _
'         & "      and nMovImporte * CASE WHEN LEFT(mc.cCtaContCod,1) = '1' THEN 1 ELSE -1 END > 0 and oc.cOpeCod = '" & psOpeCod & "' and mc.cCtaContCod LIKE LEFT(oc.cCtaContCod,2) + '" & pnMoneda & "'+SubString(oc.cCtaContCod,4,22) + '%' " _
'         & "      and mRend.nSaldo <> 0 " _
'         & "ORDER BY mc.cCtaContCod, m.cMovNro "
'Set GeneraRepOtrasPendientesContables = CargaRecord(sSql)

End Function


'Private Function GeneraRepAnalisisPendientes(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As Recordset
'
'On Error GoTo GeneraRepAnalisisPendientesErr
'   sSql = "SELECT mc.cMovNro, m.cMovDesc, mc.cCtaContCod, mc.nMovImporte, mc.nMovImporte + ISNULL(Ref.nMontoSust,0) as nSaldo, '' cPersona, ISNULL(md.cDocNro,'') cDocNro " _
'        & "FROM mov m JOIN MovCta mc on m.cmovnro = mc.cmovnro " _
'        & "      LEFT JOIN MovDoc md ON md.cMovNro = m.cMovNro " _
'        & "      LEFT JOIN MovRef mr ON mr.cMovNro = m.cMovNro " _
'        & "      LEFT JOIN (SELECT mr1.cMovNroRef, SUM(nMovImporte) nMontoSust " _
'        & "                 FROM MovRef mr1 JOIN Mov m1 ON m1.cMovNro = mr1.cMovNro " _
'        & "                                 JOIN MovCta mc1 ON mc1.cMovNro = m1.cMovNro " _
'        & "                 WHERE LEFT(m1.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m1.cMovEstado = '0' and not m1.cMovFlag IN ('X','E','N') and mc1.cCtaContCod IN (" & sCta & ")" _
'        & "                 GROUP BY mr1.cMovNroRef " _
'        & "               ) Ref ON Ref.cMovNroRef = mc.cMovNro " _
'        & "WHERE LEFT(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.cMovEstado= '0' and m.cMovFlag <> 'X' and mc.cCtaContCod IN (" & sCta & ")" _
'        & "      and nMovImporte > 0 and mr.cMovNro iS NULL " _
'        & "      and mc.nMovImporte + ISNULL(Ref.nMontoSust,0) <> 0 " _
'        & "ORDER BY mc.cCtaContCod, m.cMovNro "
'
'   Set GeneraRepAnalisisPendientes = CargaRecord(sSql)
'End If
'Exit Function
'GeneraRepAnalisisPendientesErr:
'   MsgBox Err.Description, vbInformation, "¡Aviso!"
'End Function

Private Function GeneraRepAnalisisPendientesInterAgencias(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As Recordset
'Dim prs  As ADODB.Recordset
'Dim sDH  As String
'Dim sCta As String
'Dim rsAge As ADODB.Recordset
'Set rsAge = New ADODB.Recordset
'Dim sql As String
'Dim lnI As Integer
'
'On Error GoTo GeneraRepAnalisisPendientesErr
'Set prs = CargaOpeCta(psOpeCod, , "0", True)
'If Not prs.EOF Then
'   sCta = MuestraListaRecordSet(prs, 0)
'   RSClose prs
'
'sql = " Select cValor From TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
'rsAge.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'lnI = 0
'ReDim Preserve lsCuentas(15, 2)
'sSql = ""
'While Not rsAge.EOF
'    If AbreConeccion(Right(rsAge!cValor, 2)) Then
'        lnI = lnI + 1
'        lsCuentas(lnI, 1) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
'        lsCuentas(lnI, 2) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
'        If sSql <> "" Then sSql = sSql & " UNION "
'        sSql = sSql & " Select dFecTran,cMovNro, '" & gsCodAgeN & " ' + cMovDesc cMovDesc,cCtaContCod,nMovImporte,nSaldo,cPersona,cDocNro From " & lsCuentas(lnI, 2) & "VistaMovAgePendiente TA " _
'             & " Where TA.dFecTran <= '" & Format(pdFecha, gsFormatoFecha) & " 23:59:59' And cCtaContCod Like '__" & pnMoneda & "%' And TA.dFecTran > '05/01/2002'"
'    End If
'
'    rsAge.MoveNext
'Wend
'If sSql <> "" Then sSql = sSql & " UNION "
'sSql = sSql & " Select '" & Format(pdFecha, gsFormatoFecha) & "' dFecTran, TA.cMovNro, '" & gsCodAgeN & " ' + TA.cMovDesc cMovDesc, mc.cCtaContCod, nMovImporte, nMovImporte nSaldo, '' cPersona, '' cDocNro " _
'     & " From Mov TA " _
'     & " Inner Join MovCta MC On TA.cMovNro = MC.cMovNro Where TA.cMovNro = '20020228184513070000JPNZ' And TA.cMovNro Not In (Select cMovNro From MovPendientesDet Where bExcluyeMov = 1) "
'
'sSql = sSql & " ORDER BY cCtaContCod, cDocNro"
'
'   Set GeneraRepAnalisisPendientesInterAgencias = CargaRecord(sSql)
'End If
'Exit Function
'GeneraRepAnalisisPendientesErr:
'   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepAnalisisPendientesFaltante(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
'Dim prs  As ADODB.Recordset
'Dim sDH  As String
'Dim sCta As String
'Dim rsAge As ADODB.Recordset
'Set rsAge = New ADODB.Recordset
'Dim sql As String
'Dim lnI As Integer
'Set rs = New ADODB.Recordset
'rs.CursorLocation = adUseServer
'
'On Error GoTo GeneraRepAnalisisPendientesErr
'Set prs = CargaOpeCta(psOpeCod, , "0", True)
'If Not prs.EOF Then
'   sCta = MuestraListaRecordSet(prs, 0)
'   RSClose prs
'
'sql = " Select cValor From TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
'rsAge.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'lnI = 0
'ReDim Preserve lsCuentas(15, 2)
'sSql = ""
'While Not rsAge.EOF
'    If AbreConeccion(Right(rsAge!cValor, 2)) Then
'        lnI = lnI + 1
'        lsCuentas(lnI, 1) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01", True))
'        lsCuentas(lnI, 2) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
'        If sSql <> "" Then sSql = sSql & " UNION "
'        sSql = sSql & " Select dFecTran, cMovNro, '" & gsCodAgeN & " ' + cMovDesc cMovDesc,cCtaContCod,nMovImporte,nSaldo,cPersona,'' cDocAbrev, cDocNro From " & lsCuentas(lnI, 2) & "Vistasobrantesfaltantes V" _
'             & " Where dFecTran < '" & Format(pdFecha, gsFormatoFecha) & "' And cCtaContCod Like '19" & pnMoneda & "80202%'" _
'             & " And Not Exists (Select nNumTran From " & lsCuentas(lnI, 2) & "VistaSobFalRef TA Where TA.nNumTran = V.cNroTranRef And dFecTran < '" & Format(pdFecha, gsFormatoFecha) & "' And cFlag Is Null)"
'    End If
'
'    rsAge.MoveNext
'Wend
'    sSql = sSql & " Order By dFecTran"
'    rs.CursorLocation = adUseClient
'    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'    Set GeneraRepAnalisisPendientesFaltante = rs
'End If
'Exit Function
'GeneraRepAnalisisPendientesErr:
'   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepAnalisisPendientesOPCertificada(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
'Dim prs  As ADODB.Recordset
'Dim sDH  As String
'Dim sCta As String
'Dim rsAge As ADODB.Recordset
'Set rsAge = New ADODB.Recordset
'Dim sql As String
'Dim lnI As Integer
'Set rs = New ADODB.Recordset
'rs.CursorLocation = adUseServer
'
'On Error GoTo GeneraRepAnalisisPendientesErr
'
'sql = " Select cValor From TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
'rsAge.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'lnI = 0
'ReDim Preserve lsCuentas(15, 2)
'sSql = ""
'While Not rsAge.EOF
'    If AbreConeccion(Right(rsAge!cValor, 2)) Then
'        lnI = lnI + 1
'        lsCuentas(lnI, 1) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01", True))
'        lsCuentas(lnI, 2) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
'        If sSql <> "" Then sSql = sSql & " UNION "
'        sSql = sSql & " Select dFecTran,cCodCta,cMovNro, '" & gsCodAgeN & " ' + cMovDesc  cMovDesc,cCtaContCod,nMovImporte,nSaldo,cPersona,'OP' cDocAbrev, cDocNro From " & lsCuentas(lnI, 2) & "VistaOPCert TA Where TA.dFectran < '" & Format(pdFecha, gsFormatoFecha) & " 23:59:59' And Substring(TA.cCodCta,6,1) = '" & pnMoneda & "'" _
'             & " And Not Exists (Select nNumTran From " & lsCuentas(lnI, 2) & "VistaOPCertRef TAR Where TAR.dFectran < '" & Format(pdFecha, gsFormatoFecha) & " 23:59:59' And TAR.cCodCta = TA.cCodCta And Substring(TAR.cCodCta,6,1) = '" & pnMoneda & "'" _
'             & " And TAR.cNumDoc = TA.cDocNro)  And TA.cDocNro Not In (Select cDocNro From MovPendientesDet Where bExcluyeDoc = 1)"
'    End If
'
'    rsAge.MoveNext
'Wend
'sSql = sSql & " Order By dFecTran"
'rs.CursorLocation = adUseClient
'rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'Set GeneraRepAnalisisPendientesOPCertificada = rs
'
'Exit Function
'GeneraRepAnalisisPendientesErr:
'   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepAnalisisPendientesSobrante(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
'Dim prs  As ADODB.Recordset
'Dim sDH  As String
'Dim sCta As String
'Dim rsAge As ADODB.Recordset
'Set rsAge = New ADODB.Recordset
'Dim sql As String
'Dim lnI As Integer
'Set rs = New ADODB.Recordset
'rs.CursorLocation = adUseServer
'
'On Error GoTo GeneraRepAnalisisPendientesErr
'Set prs = CargaOpeCta(psOpeCod, , "0", True)
'If Not prs.EOF Then
'   sCta = MuestraListaRecordSet(prs, 0)
'   RSClose prs
'
'sql = " Select cValor From TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
'rsAge.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'
'
'lnI = 0
'ReDim Preserve lsCuentas(15, 2)
'sSql = ""
'While Not rsAge.EOF
'    If AbreConeccion(Right(rsAge!cValor, 2)) Then
'        lnI = lnI + 1
'        lsCuentas(lnI, 1) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01", True))
'        lsCuentas(lnI, 2) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
'        If sSql <> "" Then sSql = sSql & " UNION "
'        sSql = sSql & " Select dFecTran, cMovNro, '" & gsCodAgeN & " ' + cMovDesc cMovDesc, REPLACE(cCtaContCod,'AG','" & Right(gsCodAgeN, 2) & "') cCtaContCod,nMovImporte,nSaldo,cPersona,'' cDocAbrev, cDocNro From " & lsCuentas(lnI, 2) & "Vistasobrantesfaltantes V" _
'             & " Where dFecTran < '" & Format(pdFecha, gsFormatoFecha) & "' And cCtaContCod Like '29" & pnMoneda & "20102%'" _
'             & " And Not Exists (Select nNumTran From " & lsCuentas(lnI, 2) & "VistaSobFalRef TA Where TA.nNumTran = V.cNroTranRef And dFecTran < '" & Format(pdFecha, gsFormatoFecha) & "' And cFlag Is Null)"
'    End If
'
'    rsAge.MoveNext
'Wend
'    sSql = sSql & " Order By cCtaContCod"
'    rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'    Set GeneraRepAnalisisPendientesSobrante = rs
'End If
'Exit Function
'GeneraRepAnalisisPendientesErr:
'   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepAnalisisPendientesRRHH(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFechaI As Date, ByVal pdFechaF As Date) As ADODB.Recordset
    Dim prs  As ADODB.Recordset
    Dim sDH  As String
    Dim sCta As String
    Dim rsAge As ADODB.Recordset
    Set rsAge = New ADODB.Recordset
    Dim sql As String
    Dim lnI As Integer
    Set rs = New ADODB.Recordset
    Dim lsAgeConsol As String
    rs.CursorLocation = adUseServer
    Dim prsDat As ADODB.Recordset
    Set prsDat = New ADODB.Recordset
    
    Dim oConRem As DConecta
    Set oConRem = New DConecta
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    On Error GoTo GeneraRepAnalisisPendientesErr
    Set prs = CargaOpeCta(psOpeCod, , "0", True)
    If Not prs.EOF Then
       sCta = MuestraListaRecordSet(prs, 0)
       RSClose prs
    
    oCon.AbreConexion
    oConRem.AbreConexion 'Remota "07"
    sql = " Select cValor From dbComunes..TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
    Set rsAge = oConRem.CargaRecordSet(sql)
    
    sCta = "29" & pnMoneda & "8070401"
    
    oConRem.CierraConexion
    
    lnI = 0
    ReDim Preserve lsCuentas(rsAge.RecordCount, 2)
    sSql = ""
    While Not rsAge.EOF
        If oConRem.AbreConexion Then 'Remota(Right(rsAge!cValor, 2))
            gsCodAgeN = rsAge!cvalor
            lnI = lnI + 1
            lsAgeConsol = "11207"
            'lsCuentas(lnI, 1) = Trim(oCon.GetCadenaConexionEnlazado(Right(lsAgeConsol, 2), "01"))
'            lsCuentas(lnI, 2) = Trim(oCon.GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
            If sSql <> "" Then sSql = sSql & " UNION "
            sSql = " Select distinct dFecTran, cMovNro, '" & gsCodAgeN & " ' + cCodCta + ' ' + cMovDesc + ' ' + RTRIM(cNomOpe) cMovDesc, '" & sCta & "' cCtaContCod,   " _
                 & " Case cDebeHab When 'D' Then abs(nMovImporte) Else 0 End nMovImporte," _
                 & " Case cDebeHab When 'H' Then abs(nMovImporte) Else 0 End nSaldo, cNomPers cPersona, ISNULL(cNumDoc,'') + ' - ' + str(cDocNro) cDocNro From VistaPendientesRRHH Where dFecTran Between '" & Format(pdFechaI, gsFormatoFecha) & "' And '" & Format(pdFechaF + 1, gsFormatoFecha) & "' And Substring(cCodCta,6,1) = '" & pnMoneda & "' Order by dFecTran"
            Set rs = oConRem.CargaRecordSet(sSql)
            
            AdicionaRecordSet prsDat, rs
        
            rs.Close
        End If
        rsAge.MoveNext
    Wend
        
    sSql = " Select Convert(Datetime,Left(M.cMovNro,8)) dFecTran, M.cMovNro, M.cMovDesc + ' ' + OPT.cOpeDesc cMovDesc, MC.cCtaContCod," _
         & " Case When MC.nMovImporte > 0 Then Abs(MC.nMovImporte) Else 0 End nMovImporte," _
         & " Case When MC.nMovImporte < 0 Then Abs(MC.nMovImporte) Else 0 End nSaldo," _
         & " PE.cPersNombre cPersona, ' ' cDocNro From MovCta MC" _
         & " Inner Join Mov M On M.nMovNro = MC.nMovNro" _
         & " Left  Join MovGasto MO On MO.nMovNro = M.nMovNro" _
         & " Left  Join Persona PE On MO.cPersCod = PE.cPersCod" _
         & " Left  Join MovDoc MD On MD.nMovNro = M.nMovNro" _
         & " Inner Join OpeTpo OPT On OPT.cOpeCod = M.cOpeCod" _
         & " Where Left(M.cMovNro,8) Between  '" & Format(pdFechaI, gsFormatoMovFecha) & "' And '" & Format(pdFechaF + 1, gsFormatoMovFecha) & "' And MC.cCtaContCod = '29" & pnMoneda & "8070401'" _
         & " And nMovFlag = " & MovFlag.gMovFlagVigente & " And M.cOpeCod <> '701107' "
        sSql = sSql & " Order By dFecTran Desc"
        Set rs = oCon.CargaRecordSet(sSql)
        
        AdicionaRecordSet prsDat, rs
        
        prsDat.Sort = "dFectran DESC"
        Set GeneraRepAnalisisPendientesRRHH = prsDat
    End If
    Exit Function
GeneraRepAnalisisPendientesErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepEntregaARendir(ByVal nCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
'On Error GoTo GeneraRepEntregaARendirErr
'sSql = "SELECT b.dDocFecha, g.cDocAbrev, b.cDocTpo, b.cDocNro, e.cNomPers, e.cTipPers, a.cMovDesc, d.cObjetoCod, " _
'     & "       a.cMovNro, c.cCtaContCod, c.nMovImporte * -1 as nDocImporte " _
'     & "FROM   Mov a JOIN MovDoc b ON b.cMovNro = a.cMovNro " _
'     & "             JOIN MovCta c ON c.cMovNro = a.cMovNro " _
'     & "             JOIN MovObj d ON d.cMovNro = c.cMovNro and d.cMovItem = c.cMovItem " _
'     & "        LEFT JOIN (SELECT h.cMovNro, h.cMovNroRef FROM MovRef h JOIN Mov M ON m.cMovNro = h.cMovNro " _
'     & "                   WHERE m.cMovNro < '" & Format(CDate(txtFecha) + 1, gsFormatoMovFecha) & "' and m.cMovFlag NOT IN ('X','E','N') " _
'     & "                  ) ref ON  ref.cMovNroRef = a.cMovNro " _
'     & "       ,dbPersona.dbo.Persona e, dbComunes.dbo.OpeCta f,  dbComunes.dbo.Documento g,  dbComunes.dbo.OpeDoc h " _
'     & "WHERE  a.cMovEstado = '0' and a.cMovFlag NOT IN ('X','E','N') and f.cOpeCod = '501301' and f.cOpeCtaDH = 'D' and h.cOpeCod = f.cOpeCod and " _
'     & "       c.cCtaContCod = f.cCtaContCod and " _
'     & "       h.cOpeDocMetodo NOT IN ('2')  and cOpeDocEstado not in('12') and ref.cMovNro IS NULL and " _
'     & "       SUBSTRING(d.cObjetoCod, 3, 10) = e.cCodPers And h.cDocTpo = b.cDocTpo And g.cDocTpo = b.cDocTpo " _
'     & "       AND a.cMovNro < '" & Format(CDate(txtFecha) + 1, gsFormatoMovFecha) & "' ORDER BY b.dDocFecha, a.cMovNro"
'Set GeneraRepEntregaARendir = CargaRecord(sSql)
'
'GeneraRepEntregaARendirErr:
'    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Sub cmdGenerar_Click()
   Call GeneraReporte
End Sub

Private Sub cmdImprimir_Click()
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    
    
    lsArchivoN = App.path & "\Spooler\" & Format(CDate(Me.txtFecha.Text) & " " & Time, "yyyymmddhhmmss") & gsCodUser & ".xlsx"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       
       Call GeneraReporteE
    
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       CargaArchivo lsArchivoN, App.path & "\SPOOLER"
       
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaManCtaPend
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Genero el Reporte de Pendientes Detallado : " & txtOpeDes.Text & "hasta la Fecha " & txtFecha.Text
            Set objPista = Nothing
            '*******
    End If
       
       
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oOpe As New DOperacion
    txtFecha = gdFecSis
    mskFF = gdFecSis
    FormatoFlex
    pnMoneda = 1
    txtOpeCod.rs = oOpe.CargaOpeTpo(Left(gOpePendOpeAgencias, 4), True, , 0, 2)
    CentraForm Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CierraConexion
End Sub


Private Sub mskFF_GotFocus()
    mskFF.SelStart = 0
    mskFF.SelLength = 50
End Sub

Private Sub mskFF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If optMoneda(0).value = True Then
          optMoneda(0).SetFocus
       Else
          optMoneda(1).SetFocus
       End If
    End If
End Sub

Private Sub optMoneda_Click(Index As Integer)
pnMoneda = Index + 1
End Sub

Private Sub optMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   CmdGenerar.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
    txtFecha.SelStart = 0
    txtFecha.SelLength = 50
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.mskFF.SetFocus
End If
End Sub

Private Sub txtOpeCod_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If ValidaOperacion(txtOpeCod) Then
      txtFecha.SetFocus
   End If
End If
End Sub

Private Function ValidaOperacion(psOpeCod As String) As Boolean
Dim prs As ADODB.Recordset
ValidaOperacion = False
   Set prs = CargaOpeTpo(psOpeCod)
   If prs.EOF Then
      RSClose prs
      MsgBox "Operación de Pendiente no Existe", vbInformation, "¡Aviso!"
      Exit Function
   Else
      txtOpeDes = prs!cOpeDesc
   End If
   RSClose prs
ValidaOperacion = True
End Function

Private Sub GeneraReporteE()
    Dim I As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lnAcum As Currency
    Dim VSQL As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Me.MousePointer = 11
    
    For I = 0 To Me.fg.Rows - 1
        lnAcum = 0
        For j = 1 To Me.fg.Cols - 1
            xlHoja1.Cells(I + 1, j + 1) = Me.fg.TextMatrix(I, j)
            If I > 1 And j > 1 Then
                If Me.fg.TextMatrix(I, j) <> "" Then
                    'lnAcum = lnAcum + CCur(Me.fg.TextMatrix(I, J))
                End If
            End If
        Next j
        If I > 1 Then
            VSQL = Format(lnAcum, "#,##0.00")  ' "=SUMA(" & Trim(ExcelColumnaString(3)) & Trim(I + 1) & ":" & Trim(ExcelColumnaString(Me.fg.Cols)) & Trim(I + 1) & ")"
            'xlHoja1.Cells(I + 1, Me.fg.Cols + 1).Formula = VSQL
            xlHoja1.Cells(I + 1, Me.fg.Cols + 1) = VSQL
        End If
    Next I
    
    xlHoja1.Range("E1:F" & Trim(Str(Me.fg.Rows))).NumberFormat = "#,##0.00"
    xlHoja1.Range("A1:A" & Trim(Str(Me.fg.Rows))).Font.Bold = True
    xlHoja1.Range("1:1").Font.Bold = True
    
    xlHoja1.Cells.Select
     
    With xlHoja1.Cells.Font
        .Name = "Arial"
        .Size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    
    xlHoja1.Range("E1:F" & Trim(Str(Me.fg.Rows))).Font.Bold = True
    xlHoja1.Range("A1:A" & Trim(Str(Me.fg.Rows))).Font.Bold = True
    xlHoja1.Range("1:1").Font.Bold = True
    
    xlHoja1.Cells.EntireColumn.AutoFit
     
    xlHoja1.Range("A1:I" & Me.fg.Rows).Select
    xlHoja1.Range("A1:I" & Me.fg.Rows).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A1:I" & Me.fg.Rows).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlHoja1.Range("A1:I" & Me.fg.Rows).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:I" & Me.fg.Rows).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:I" & Me.fg.Rows).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:I" & Me.fg.Rows).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:I" & Me.fg.Rows).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:I" & Me.fg.Rows).Borders(xlInsideHorizontal)
    '    .LineStyle = xlContinuous
    '    .Weight = xlThin
    '    .ColorIndex = xlAutomatic
    End With
    
    With xlHoja1.PageSetup
        
        .CenterHeader = "&""Arial,Negrita""&16" & UCase(Me.txtOpeDes.Text)
        .RightHeader = "&P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
'        .LeftMargin = Application.InchesToPoints(0)
'        .RightMargin = Application.InchesToPoints(0)
'        .TopMargin = Application.InchesToPoints(0.55)
'        .BottomMargin = Application.InchesToPoints(0)
'        .HeaderMargin = Application.InchesToPoints(0.19)
'        .FooterMargin = Application.InchesToPoints(0)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
'       .PrintQuality = 600
        .CenterHorizontally = True
        .CenterVertically = True
        .Orientation = xlLandscape
        .Draft = False
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = False
        .FitToPagesWide = 1
    End With
    Me.MousePointer = 0
End Sub

Private Sub txtOpeCod_EmiteDatos()
    txtOpeDes = txtOpeCod.psDescripcion
    If txtOpeCod <> "" Then
        txtFecha.SetFocus
    End If
End Sub
