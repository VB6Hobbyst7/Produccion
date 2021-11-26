VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepAnaCtas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Analisis de Cuentas Contables: Reporte"
   ClientHeight    =   6000
   ClientLeft      =   2070
   ClientTop       =   2835
   ClientWidth     =   10995
   Icon            =   "frmRepAnaCtas.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFecha 
      Caption         =   "&Utilizar fecha de Reporte"
      Height          =   255
      Left            =   30
      TabIndex        =   14
      Top             =   5550
      Width           =   2745
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
      Left            =   7455
      TabIndex        =   4
      Top             =   5475
      Width           =   1125
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
      Left            =   8610
      TabIndex        =   12
      Top             =   30
      Width           =   1710
      Begin VB.OptionButton optMoneda 
         Caption         =   "Dolares"
         Height          =   345
         Index           =   1
         Left            =   870
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   765
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "Soles"
         Height          =   345
         Index           =   0
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   765
      End
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
      Left            =   60
      TabIndex        =   11
      Top             =   30
      Width           =   5940
      Begin Sicmact.TxtBuscar txtOpeCod 
         Height          =   345
         Left            =   120
         TabIndex        =   13
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
      Begin VB.TextBox txtOpeDes 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   1365
         TabIndex        =   0
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   8640
      TabIndex        =   6
      Top             =   5475
      Width           =   1125
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   285
      Left            =   15
      TabIndex        =   10
      Top             =   5535
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
      Left            =   30
      TabIndex        =   5
      Top             =   840
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
   Begin VB.CommandButton cmdsalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9810
      TabIndex        =   7
      Top             =   5475
      Width           =   1125
   End
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
      Left            =   5985
      TabIndex        =   8
      Top             =   30
      Width           =   1395
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   165
         TabIndex        =   1
         Top             =   240
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   390
      Left            =   1170
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   688
      _Version        =   393217
      TextRTF         =   $"frmRepAnaCtas.frx":030A
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
Attribute VB_Name = "frmRepAnaCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sSql     As String
Dim rs       As ADODB.Recordset
Dim pnMoneda As Integer
Dim lsCuentas() As String
Dim oCon As DConecta
Dim oAna As NAnalisisCtas
Dim oOpe As DOperacion
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
fg.TextMatrix(0, 4) = "Importe"
fg.TextMatrix(0, 5) = "Saldo"
fg.TextMatrix(0, 6) = "Persona"
fg.TextMatrix(0, 7) = "Documento"

End Sub

Private Function GeneraRepPendSubsidios(psOpeCod As String, pnMoneda As Integer, pdFecha As Date) As ADODB.Recordset
Dim prs As ADODB.Recordset
    Set oCon = New DConecta
    oCon.AbreConexion
    sSql = "SELECT mp.cMovNro, pla.cPlaInsDes cMovDesc, cc.cCtaContCod, pd.nMonto nMovImporte, ISNULL(mRend.nSaldo, ISNULL(pd.nMonto+ref.nPago ,pd.nMonto)) nSaldo, " _
     & "           'PLA' cDocAbrev, 0 nDocTpo, pla.cPlanillaCod cDocNro, Convert( varchar(10), pd.cRRHHPeriodo) dDocFecha, " _
     & "       p.cPersNombre cPersona, p.cPersCod, pla.nMovNroEst nMovNro " _
     & "FROM RHPlanilladetCon pd JOIN RHPlanilla pla ON pla.cRRHHPeriodo = pd.cRRHHPeriodo AND pla.cPlanillaCod = pd.cPlanillaCod " _
     & "     JOIN Mov mp ON mp.nMovNro = pla.nMovNroEst " _
     & "     JOIN RHConceptoCta cc ON cc.cRHConceptoCod = pd.cRHConceptoCod " _
     & "     JOIN Persona p ON p.cPersCod = pd.cPersCod " _
     & "     LEFT JOIN (SELECT m.cMovNro, m.nMovNro, nMovNroRef, SUM(mc.nMovImporte) as nPago FROM MovRef mr JOIN Mov m ON m.nMovNro = mr.nMovNro JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN OpeCta oc ON oc.cCtaContCod = mc.cCtaContCod " _
     & "                WHERE m.nMovEstado = '" & gMovEstContabMovContable & "' and m.nMovFlag <> '" & gMovFlagEliminado & "' and oc.cOpeCod = '" & psOpeCod & "' " _
     & "                GROUP BY m.cMovNro, m.nMovNro, nMovNroRef) ref ON ref.nMovNroRef = pla.nMovNroEst " _
     & "     LEFT JOIN MovPendientesRend mRend ON mRend.nMovNro = pla.nMovNroEst, OpeCta oc " _
     & "WHERE not pla.nMovNroEst is NULL and LEFT(mp.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and cc.cOpeCod like '62260_' " _
     & "      and oc.cOpeCod like '" & psOpeCod & "' and cc.cCtaContCod = LEFT(oc.cCtaContCod,2)+'" & pnMoneda & "'+SubString(oc.cCtaContCod,4,22) and cc.cOpeCod = mp.cOpeCod " _
     & "      and pd.cPlanillaCod IN ('" & gsRHPlanillaSubsidio & "','" & gsRHPlanillaSubsidioEnfermedad & "') " _
     & " and ( ISNULL(mRend.nSaldo, ISNULL(pd.nMonto+ref.nPago ,pd.nMonto)) <> 0 or mRend.nSaldo <> 0 ) ORDER BY cc.cCtaContCod, mp.cMovNro " _
'Set prs = oAna.GetSubsidiosPendientes(psOpeCod, pdFecha, "", psOpeCod)


'Set prs = oAna.GetSubsidiosPendientes(psOpeCod, pdFecha, "", psOpeCod)
Set GeneraRepPendSubsidios = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
Exit Function
End Function

Private Sub GeneraReporte()
Dim sCtaCod As String
Dim nRow As Integer
Dim nTot As Currency
Dim nSdo As Currency
Dim oCta As DCtaCont
Set oCta = New DCtaCont
If ValidaFecha(txtFecha) <> "" Then
    MsgBox "Fecha no valida", vbInformation, "¡Aviso!"
    Exit Sub
End If
cmdGenerar.Enabled = False
MousePointer = 11
fg.MousePointer = 11
   nTot = 0
   nSdo = 0
   prg.Visible = True
   Select Case txtOpeCod
       Case gOpePendOpeAgencias
            Set rs = GeneraRepAnalisisPendientesInterAgencias(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendRendirCuent
            Set rs = GeneraRepEntregaARendir(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendOrdendePago
            Set rs = GeneraRepOrdenPago(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendSobrantCaja
            Set rs = GeneraRepAnalisisPendientesSobrante(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendFaltantCaja
            Set rs = GeneraRepAnalisisPendientesFaltante(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendOPCertifica
            Set rs = GeneraRepAnalisisPendientesOPCertificada(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendPagoSubsidi
            Set rs = GeneraRepPendSubsidios(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendCobraLiquid
            Set rs = GeneraRepAnalisisPendientesCobServicios(txtOpeCod, pnMoneda, txtFecha)
       Case gOpePendProvPagoProv
            Set rs = GeneraRepProvisionPagoProveedor(txtOpeCod, pnMoneda, txtFecha)
       Case Else
            Set rs = GeneraRepAnalisisPendOtraOpeLiquidar(txtOpeCod, pnMoneda, txtFecha)
'            Set rs = GeneraRepOtrasPendientesContables(txtOpecod, pnMoneda, txtFecha, "D")
   End Select
   fg.Rows = 2
   EliminaRow fg, 1
If Not rs Is Nothing Then

If rs.State = adStateOpen Then
   If Not rs.EOF Then
      sCtaCod = ""
      prg.Max = rs.RecordCount
      rs.MoveFirst
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
            fg.TextMatrix(nRow, 2) = rs!cCtaContCod & " - " & oCta.GetCtaContDesc(rs!cCtaContCod)
            sCtaCod = rs!cCtaContCod
         End If
         AdicionaRow fg
         nRow = fg.row
         fg.TextMatrix(nRow, 1) = Left(rs!cMovNro, 8) & "-" & Mid(rs!cMovNro, 9, 6) & Right(rs!cMovNro, 4)
         fg.TextMatrix(nRow, 2) = rs!cMovDesc
         fg.TextMatrix(nRow, 3) = rs!cCtaContCod
         fg.TextMatrix(nRow, 4) = Format(rs!nMovImporte, gsFormatoNumeroView)
         fg.TextMatrix(nRow, 5) = Format(rs!nSaldo, gsFormatoNumeroView)
         fg.TextMatrix(nRow, 6) = IIf(IsNull(rs!cPErsona), "", rs!cPErsona)
         fg.TextMatrix(nRow, 7) = rs!cDocAbrev & " " & rs!cDocNro
         'fg.TextMatrix(nRow, 7) = rs!cDocNro
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
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaManCtaPend
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Se Genero el Reporte de Pendientes : " & txtOpeDes.Text & "hasta la Fecha " & txtFecha.Text
            Set objPista = Nothing
            '*******
   Else
        MsgBox "No se encontraron datos Pendientes para Reportar", vbInformation, "¡Aviso!"
   End If
End If
End If
RSClose rs
MousePointer = 0
fg.MousePointer = 0
cmdGenerar.Enabled = True
End Sub

Private Function GeneraRepOrdenPago(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As Recordset
Dim prs  As ADODB.Recordset
Dim sDH  As String
Dim sCta As String
Dim lsCadSer As String
Dim lnI As Integer
Dim sql As String

Dim prsDat As ADODB.Recordset
Dim rsAge As ADODB.Recordset
Dim rsR As ADODB.Recordset
Set prsDat = New ADODB.Recordset

Dim oConR As New DConecta
Dim Ope As New DOperacion
On Error GoTo GeneraRepOrdenPagoErr

Set prs = CargaOpeCta(psOpeCod, , "0", True)
If Not prs.EOF Then
   sCta = MuestraListaRecordSet(prs, 0)
End If
RSClose prs

Dim oAge As DActualizaDatosArea
Set oAge = New DActualizaDatosArea
Set rsAge = oAge.GetAgencias(, False, True)
Set oAge = Nothing

lnI = 0
ReDim Preserve lsCuentas(15, 2)
sSql = ""
Set oCon = New DConecta
If gbBitCentral Then
    oCon.AbreConexion

    sSql = " SELECT A.cMovNro, A.cMovDesc, C.cCtaContCod, ISNULL(ME.nMovMEImporte, C.nMovImporte) * -1 As nMovImporte, ISNULL(ME.nMovMEImporte, C.nMovImporte) * -1 As nSaldo, IsNull(E.cPersNombre,'') cPersona, G.cDocAbrev, B.cDocNro cDocNro" _
         & " FROM Mov A JOIN MovDoc B ON B.nMovNro = A.nMovNro " _
         & "     JOIN MovCta C ON c.nMovNro = a.nMovNro LEFT JOIN MovMe ME ON ME.nMovNro = C.nMovNro And ME.nMovItem = C.nMovItem " _
         & "     JOIN MovGasto D ON D.nMovNro = A.nMovNro " _
         & "     JOIN Persona E ON E.cPersCod = D.cPersCod " _
         & "     JOIN Documento G ON G.nDocTpo = B.nDocTpo " _
         & " WHERE LEFT(A.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' And A.nMovEstado = " & gMovEstContabMovContable & "  and A.nMovFlag <> " & gMovFlagEliminado & "  And B.nDoctpo = " & TpoDocOrdenPago & " " _
         & " And C.cCtaContCod LIKE '21" & Trim(Str(pnMoneda)) & "10502%'" _
         & " And C.nMovImporte < 0 " _
         & " And EXISTS (SELECT H.nMovNro FROM MovRef H WHERE H.nMovNro = A.nMovNro) " _
         & " And NOT EXISTS (SELECT MP.nMovNro FROM MovRef Pag JOIN Mov MP on MP.nMovNro = Pag.nMovNro " _
         & "     WHERE LEFT(MP.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' And MP.nMovFlag <> " & gMovFlagEliminado & "  And MP.cOpeCod IN ('" & gAhoRetFondoFijo & "','" & gAhoRetFondoFijoCanje & " ')) " _
         & " ORDER BY C.cCtaContCod, a.cMovNro"
                     
    Set prs = oCon.CargaRecordSet(sSql)
    RecordSetAdiciona prsDat, prs
    
    oConR.CierraConexion
    Set oConR = Nothing
    
    sSql = " SELECT m.cMovNro, m.cMovDesc,mc.cCtaContCod, nMovMEImporte nMovImporte, nMovImporte dSaldo, '' cPersona, 'OP' cDocAbrev , '' cDocNro  FROM MOV m" _
        & " Inner Join MovCta mc On mc.nMovNro = m.nMovNro" _
        & " Left  Join MovME me On me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
        & " WHERE m.CMOVNRO = '200104041806581120700JPNZ' And m.cMovNro Not in (Select cMovNro From MovPendientesDet Where bExcluyeMov = 1) And Substring(mc.cCtaContCod,3,1) = '" & pnMoneda & "' " _
        & " ORDER BY cCtaContCod, m.cmovnro "
    
    Set prs = oCon.CargaRecordSet(sSql)
    If Not (prs.EOF And prs.BOF) Then
        RecordSetAdiciona prsDat, prs
    End If
    oCon.CierraConexion
    Set oCon = Nothing
    
    RSClose rs
    If Not (prsDat.EOF And prsDat.EOF) Then
        prsDat.MoveFirst
    End If
    Set GeneraRepOrdenPago = prsDat
'******************************
Else
    oCon.AbreConexion
    While Not rsAge.EOF
        If oConR.AbreConexion Then 'Remota(Right(rsAge!Codigo, 2), False)
            If pnMoneda = 1 Then
                sql = " Select cValorVar from Varsistema Where cCodProd = 'AHO' And cNomVar = 'cCtaFonFijMN'"
            Else
                sql = " Select cValorVar from Varsistema Where cCodProd = 'AHO' And cNomVar = 'cCtaFonFijME'"
            End If
            Set rsR = oConR.CargaRecordSet(sql)
            If Not (rsR.EOF) Then
                If Right(rsAge!Codigo, 2) = Left(rsR!cValorVar, 2) Then
                    lnI = lnI + 1
                    lsCuentas(lnI, 1) = Trim(rsR!cValorVar)
                    'lsCuentas(lnI, 2) = oConR.StringServidorRemoto(Right(rsAge!Codigo, 2), "01", True)
                    
                    If pnMoneda = 1 Then
                       sCta = "21110502" & Right(rsAge!Codigo, 2)
                    Else
                       sCta = "21210502" & Right(rsAge!Codigo, 2)
                    End If
                    
                    
                    sSql = " SELECT m.cMovNro, m.cMovDesc, '" & sCta & "' cCtaContCod, Case Substring(mc.cCtaContCod,3,1) When '2' Then Abs(me.nMovMEImporte) Else Abs(mc.nMovImporte) End nMovImporte, mc.nMovImporte nSaldo, ISNULL(p.cPersNombre,'') cPersona, 'OP' cDocAbrev,  md.cDocNro " _
                         & " FROM  Mov m JOIN MovDoc md   ON md.nMovNro = m.nMovNro " _
                         & "             JOIN MovCta mc   ON mc.nMovNro = m.nMovNro " _
                         & "        LEFT JOIN MovGasto mo ON mo.nMovNro =  m.nMovNro LEFT JOIN MovArendir ma ON ma.nMovNro = m.nMovNro " _
                         & "        LEFT JOIN Persona p   ON (p.cPersCod = mo.cPersCod or p.cPersCod = ma.cPersCod) " _
                         & "        LEFT JOIN MovME ME On mc.nMovNro = ME.nMovNro And mc.nMovItem = ME.nMovItem " _
                         & "        LEFT JOIN " & lsCuentas(lnI, 2) & "VistaOrdenPago OPG On convert(Int,cDocNro) = OPG.nNumOp And datediff(d,OPG.dFecOP,'" & Format(pdFecha, gsFormatoFecha) & "') >= 0 And OPG.cEstOP = '4'  And OPG.cCodCta = '" & lsCuentas(lnI, 1) & "'" _
                         & " WHERE m.cMovNro Not In ('200012141909141120700LJCR','200101241902431120700LJCR','200103301001291120700LJCR','200106281733051120700CLMA','200207121659071120700LJCR','200208221123121120100LJCR','200303121719141120100LJCR','200305071207371120100LJCR')  And LEFT(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.nMovEstado = '10' And" _
                         & " ((Not m.nMovFlag In ('1','2','3')) Or (m.nMovFlag = '2' And NOT EXISTS (SELECT MR.nMovNro FROM MovRef MR Inner Join Mov MRM On MRM.nMovNro = MR.nMovNro WHERE MR.nMovNroRef = m.nMovnro And LEFT(MRM.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' And MRM.nMovFlag in ('1','2','3') ) ) )" _
                         & " and md.nDocTpo = '" & TpoDocOrdenPago & "' " _
                         & " and mc.cCtaContCod LIKE '__" & pnMoneda & "%' " _
                         & " and mc.nMovImporte < 0 And OPG.nNumOp Is Null And convert(Int,cDocNro) Not In (Select Convert(Int,cDocNro) from MovPendientesdet Where bExcluyeDoc = 1) And mc.cCtaContCod = '" & sCta & "' " _
                         & " ORDER BY cCtaContCod, m.cmovnro "
                         
                    Set prs = oCon.CargaRecordSet(sSql)
                    RecordSetAdiciona prsDat, prs
                End If
            End If
        End If
        oConR.CierraConexion
        Set oConR = Nothing
        rsAge.MoveNext
    Wend
   ' sSql = " SELECT m.cMovNro, m.cMovDesc,mc.cCtaContCod, nMovMEImporte nMovImporte, nMovImporte dSaldo, '' cPersona, 'OP' cDocAbrev , '' cDocNro  FROM MOV m" _
        & " Inner Join MovCta mc On mc.nMovNro = m.nMovNro" _
        & " Left  Join MovME me On me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
        & " WHERE m.CMOVNRO = '200104041806581120700JPNZ' And m.cMovNro Not in (Select cMovNro From MovPendientesDet Where bExcluyeMov = 1) And Substring(mc.cCtaContCod,3,1) = '" & pnMoneda & "' " _
        & " ORDER BY cCtaContCod, m.cmovnro "
    
    'Set prs = oCon.CargaRecordSet(sSql)
    'RecordSetAdiciona prsDat, prs
    oCon.CierraConexion
    Set oCon = Nothing
    
    RSClose rs
    prsDat.MoveFirst
    Set GeneraRepOrdenPago = prsDat
End If

Exit Function
GeneraRepOrdenPagoErr:
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepProvisionPagoProveedor(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
Dim oDCaja As New DCajaGeneral
Dim oOpe As New DOperacion
Dim lsCtaB  As String
Dim lsCtaS  As String
Dim lsDocs  As String
lsCtaB = oOpe.EmiteOpeCta(OpeCGOpeProvPago, "D")
lsCtaS = oOpe.EmiteOpeCta(OpeCGOpeProvPago, "D", 1)
If pnMoneda = 2 Then
    lsCtaB = Left(lsCtaB, 2) & pnMoneda & Mid(lsCtaB, 4, 22)
    lsCtaS = Left(lsCtaS, 2) & pnMoneda & Mid(lsCtaS, 4, 22)
End If

Set rs = oOpe.CargaOpeDoc(OpeCGOpeProvPago, OpeDocMetDigitado)
lsDocs = RSMuestraLista(rs, 1)

sSql = "SELECT md.dDocFecha, ISNULL(doc.cDocAbrev,'') cDocAbrev, ISNULL(md.nDocTpo,0) nDocTpo, ISNULL(md.cDocNro,'') cDocNro, ISNULL(Prov.cPersNombre,'') cPersona, M.cMovDesc, ISNULL(Prov.cPersCod,'') cPersCod, " _
     & "       m.cMovNro, m.nMovNro, mc.cCtaContCod, ISNULL(me.nMovMeImporte,mc.nMovImporte) * -1 as nMovImporte, ISNULL(me.nMovMeImporte,mc.nMovImporte) * -1 as nMovImporte, ISNULL(Pend.nSaldo - CASE WHEN RefP.nImporte <> 0 THEN ISNULL(me.nMovMeImporte,mc.nMovImporte) END,ISNULL(me.nMovMeImporte,mc.nMovImporte) * -1) as nSaldo " _
     & "FROM   Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro " _
     & "             JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN MovMe ME ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
     & "             JOIN MovGasto mg ON mg.nMovNro = m.nMovNro " _
     & "              LEFT JOIN (SELECT m1.cMovNro, mr.nMovNro, mr.nMovNroRef, mc.cCtaContCod, mc.nMovImporte nImporte " _
     & "                   FROM Mov m1 JOIN MovRef mr ON mr.nMovNro = m1.nMovNro " _
     & "                          JOIN MovCta mc ON mc.nMovNro = m1.nMovNro and mc.cCtaContCod IN ('" & lsCtaB & "','" & lsCtaS & "') " _
     & "                   WHERE LEFT(m1.cMovNro,8) >= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m1.nMovFlag NOT IN (1,2,3,5) " _
     & "                     and m1.nMovEstado = 10 " _
     & "            ) RefPago ON RefPago.nMovNroRef = m.nMovNro and RefPago.cCtaContCod = mc.cCtaContCod " _
     & "             LEFT JOIN Persona Prov  ON Prov.cPersCod = mg.cPersCod " _
     & "        LEFT JOIN (SELECT mr.nMovNro, mr.nMovNroRef, mc.cCtaContCod, mc.nMovImporte nImporte FROM Mov m1 JOIN MovRef mr ON mr.nMovNro = m1.nMovNro " _
     & "                          JOIN MovCta mc ON mc.nMovNro = m1.nMovNro " _
     & "                   WHERE LEFT(m1.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and ( m1.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagDeExtorno & "," & gMovFlagModificado & ") or ( m1.nMovFlag = 2 and NOT EXISTS(Select cMovNro FROM Mov MExt JOIN MovRef mrExt ON mrExt.nMovNro = mExt.nMovNro WHERE mrExt.nMovNroRef = m1.nMovNro and mExt.nMovFlag = 3 and LEFT(mExt.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' ) ) ) and m1.nMovEstado = " & gMovEstContabMovContable & " and mc.cCtaContCod IN ('" & lsCtaB & "','" & lsCtaS & "') " _
     & "                  ) RefP ON RefP.nMovNroRef = m.nMovNro and RefP.cCtaContCod = mc.cCtaContCod LEFT JOIN MovPendientesRend Pend ON Pend.nMovNro = mc.nMovNro and Pend.cCtaContCod = mc.cCtaContCod " _
     & "             JOIN Documento Doc ON Doc.nDocTpo = md.nDocTpo " _
     & "WHERE  m.nMovEstado = " & gMovEstContabMovContable & " and m.nMovFlag NOT IN ('" & gMovFlagEliminado & "','2','3','" & gMovFlagModificado & "') " _
     & "       and ((refP.nMovNroRef IS NULL and Pend.nSaldo is NULL) or (refP.nMovNroRef IS NULL and Pend.nSaldo <> 0 ) or (refP.nMovNroRef IS NULL and Pend.nSaldo = 0 and not RefPago.nMovNro is NULL) ) and " _
     & "       LEFT(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' " _
     & "       and md.nDocTpo in (" & lsDocs & ") and mc.cCtaContCod IN ('" & lsCtaB & "','" & lsCtaS & "') " _
     & "ORDER BY mc.cCtaContCod, m.cMovNro"

Set oCon = New DConecta
oCon.AbreConexion
Set GeneraRepProvisionPagoProveedor = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function


Private Function GeneraRepAnalisisPendOtraOpeLiquidar(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
Dim prsCta As ADODB.Recordset
Dim prs    As ADODB.Recordset
Dim rsAge  As ADODB.Recordset
Dim prsDat As ADODB.Recordset
Dim oAge As DActualizaDatosArea

Dim sCta As String
Dim nCol As Integer

On Error GoTo GeneraRepAnalisisPendOtraOpeLiqCajaGeneralErr
Set oCon = New DConecta
Set prsDat = New ADODB.Recordset
Set prsCta = CargaOpeCta(psOpeCod, "D", "0", True)
If Not prsCta.EOF Then
Do While Not prsCta.EOF
    sCta = prsCta!cCtaContCod
    sCta = Left(sCta, 2) & pnMoneda & Mid(sCta, 4, 22)
    If gbBitCentral Then
        sSql = "SELECT m.cMovNro, m.cMovDesc, '" & sCta & "' cCtaContCod, " _
          & "      ABS(ov.nMovImporte) nMovImporte, ABS(ov.nMovImporte) nSaldo, d.cDocAbrev, md.nDocTpo cDocTpo, md.cDocNro, md.dDocFecha, " _
          & "      ISNULL(PB.cPersNombre, POV.cPersNombre) cPersona,  ISNULL(TB.cPersCodIF,MG.cPersCod) cCodPers " _
          & "From Mov M " _
          & "     JOIN (SELECT Distinct cOpeCod FROM OpeCtaNeg oc " _
          & "           WHERE '" & sCta & "' LIKE REPLACE(oc.cCtaContCod,'M','" & pnMoneda & "') + '%') oc ON oc.cOpeCod = M.cOpeCod  " _
          & "     LEFT JOIN MovDoc md ON md.nMovNro = m.nMovNro LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
          & "     LEFT JOIN MovTransferBco TB ON M.nMovNro = TB.nMovNro LEFT JOIN MovCap MC ON MC.nMovNro = M.nMovNro And substring(MC.cCtaCod,9,1) = " & pnMoneda & " LEFT JOIN Persona  PB ON PB.cPersCod = TB.cPersCodIF " _
          & "     LEFT JOIN MovOpeVarias   OV ON M.nMovNro = OV.nMovNro and OV.nMoneda = " & pnMoneda & " LEFT JOIN MovGasto MG ON mg.nMovNro = M.nMovNro LEFT JOIN Persona POV ON POV.cPersCod = MG.cPersCod " _
          & "WHERE M.nMovNro NOT IN (Select MR.nMovNroRef From Mov M1 JOIN MovRef MR ON M1.nMovNro = MR.nMovNro " _
          & "  And M1.nMovFlag = " & gMovFlagVigente & " ) And M.nMovFlag = " & gMovFlagVigente & " and (not tb.nMovNro is NULL or not ov.nMovNro is NULL) "
          
        oCon.AbreConexion
        Set prs = oCon.CargaRecordSet(sSql)
        RecordSetAdiciona prsDat, prs
        RSClose prs
        oCon.CierraConexion
    Else
        Set oAge = New DActualizaDatosArea
        Set rsAge = oAge.GetAgencias(, False)
        Set oAge = Nothing

        sSql = ""
        While Not rsAge.EOF
            If oCon.AbreConexion Then 'Remota(Right(rsAge!Codigo, 2), False)
                sSql = "SELECT convert(varchar(8),td.dFecTran,112) + convert(varchar(20),td.nNumTran) cMovNro, ISNULL(oo.cGlosa,'') cMovdesc, " & sCta & " cCtaContCod, " _
                     & "   td.nMonTran nMovImporte, td.nMonTran nSaldo, '' cDocAbrev, '' cDocTpo, td.cNumDoc cDocNro, Convert(VarChar(10), td.dFecTran,103) dDocFecha, " _
                     & "      ISNULL(p.cNomPers,'') cPersona, p.cCodPers cCodPers " _
                     & "FROM TranDiariaConsol td JOIN TransRef tr ON tr.cNroTranRef = td.nNumTran " _
                     & "     JOIN OtrasOpe oo ON oo.nNumTran = td.nNumTran " _
                     & "     LEFT JOIN Persona p ON p.cCodPers = oo.cCodPers " _
                     & "WHERE convert(varchar(8),td.dFecTran,112) = '" & Format(pdFecha, gsFormatoMovFecha) & "' and SubString(td.cCodCta,6,1) = '" & pnMoneda & "' and td.cCodOpe in (SELECT cCodOpe FROM OpeCuentaN WHERE cCodCnt LIKE '" & sCta & "') " _
                     & "  and td.cFlag is NULL and (cNroTran is NULL or LEFT(cNroTran,8) > '" & Format(pdFecha, gsFormatoMovFecha) & "' ) " _
                     & "ORDER BY td.dFecTran "
                Set prs = oCon.CargaRecordSet(sSql)
                RecordSetAdiciona prsDat, prs
                RSClose prs
                oCon.CierraConexion
            End If
            rsAge.MoveNext
        Wend
        RSClose rsAge
    End If
    Set oCon = New DConecta
    oCon.AbreConexion
    sSql = "SELECT m.cMovNro , m.cMovDesc, mc.cCtaContCod, mc.nMovImporte, mRend.nSaldo, ISNULL(d.cDocAbrev,'') cDocAbrev, ISNULL(md.nDocTpo,'') nDocTpo, ISNULL(md.cDocNro,'') cDocNro, ISNULL(md.dDocFecha,'') dDocFecha, " _
         & "       ISNULL(p.cPersNombre,'') cPersona, ISNULL(mop.cPersCod,'') cPersCod " _
         & "       " _
         & "FROM mov m join MovCta mc on m.nMovNro = mc.nMovNro " _
         & "      LEFT JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "      LEFT JOIN MovDoc md ON md.nMovNro = m.nMovNro " _
         & "      LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
         & "      LEFT JOIN MovGasto moP ON moP.nMovNro = mc.nMovNro " _
         & "      LEFT JOIN Persona P ON p.cPersCod = moP.cPersCod " _
         & "           JOIN MovPendientesRend mRend ON mRend.nMovNro = m.nMovNro and mRend.cCtaContCod = mc.cCtaContCod " _
         & "WHERE LEFT(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.nMovEstado= '10' and m.nMovFlag <> '1' " _
         & "      and nMovImporte  * CASE WHEN LEFT(mc.cCtaContCod,1) = '1' THEN 1 ELSE -1 END > 0  and mc.cCtaContCod =  '" & sCta & "' " _
         & "     and mRend.nSaldo <> 0 " _
         & "ORDER BY mc.cCtaContCod, m.cMovNro "
    Set prs = oCon.CargaRecordSet(sSql)
    RecordSetAdiciona prsDat, prs
    If Not prsDat Is Nothing Then
        If Not prsDat.State = adStateClosed Then
           If Not prsDat.EOF Then
              prsDat.MoveFirst
           End If
        End If
    End If
    oCon.CierraConexion
    If sSql = "" Then Exit Function
    Set GeneraRepAnalisisPendOtraOpeLiquidar = prsDat
    prsCta.MoveNext
Loop
Else
    MsgBox "No se definió Cuenta de Pendiente para Analizar", vbInformation, "¡Aviso!"
End If
Set oCon = Nothing

Exit Function
GeneraRepAnalisisPendOtraOpeLiqCajaGeneralErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function


Private Function GeneraRepOtrasPendientesContables(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date, psTpoCta As String) As ADODB.Recordset
Set oCon = New DConecta
oCon.AbreConexion
    sSql = "SELECT ISNULL(d.cDocAbrev,'') cDocAbrev, ISNULL(md.cDocNro,'') cDocNro, ISNULL(p.cPersNombre,'') cPersona, " _
         & "       ISNULL(md.dDocFecha,'') dDocFecha, ISNULL(" & IIf(pnMoneda = 2, " me.nMovMEImporte", " mc.nMovImporte") & ",0) * CASE WHEN LEFT(mc.cCtaContCod,1) = '1' THEN 1 ELSE -1 END nMovImporte, m.cMovDesc, " _
         & "       ISNULL(moP.cPersCod,'') cPersCod, m.nMovNro, ISNULL(md.nDocTpo,'') nDocTpo, mRend.nSaldo, m.cMovNro, mc.cCtaContCod " _
         & "FROM Mov m join MovCta mc on m.nMovNro = mc.nMovNro " _
         & "      LEFT JOIN MovMe  me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
         & "      LEFT JOIN MovDoc md ON md.nMovNro = m.nMovNro LEFT JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
         & "      LEFT JOIN MovGasto moP ON moP.nMovNro = m.nMovNro " _
         & "      LEFT JOIN Persona P ON p.cPersCod = moP.cPersCod " _
         & "           JOIN MovPendientesRend mRend ON mRend.nMovNro = m.nMovNro and mRend.cCtaContCod = mc.cCtaContCod, OpeCta oc " _
         & "WHERE LEFT(m.cMovNro,8) <= '" & Format(pdFecha, gsFormatoMovFecha) & "' and m.nMovEstado in ('" & gMovEstContabMovContable & "','" & gMovEstContabPendiente & "') and m.nMovFlag <> '" & gMovFlagEliminado & "' " _
         & "      and nMovImporte * CASE WHEN LEFT(mc.cCtaContCod,1) = '1' THEN 1 ELSE -1 END > 0 and oc.cOpeCod = '" & psOpeCod & "' and mc.cCtaContCod LIKE LEFT(oc.cCtaContCod,2) + '" & pnMoneda & "'+SubString(oc.cCtaContCod,4,22) + '%' " _
         & "      and mRend.nSaldo <> 0 " _
         & "ORDER BY mc.cCtaContCod, m.cMovNro "

Set GeneraRepOtrasPendientesContables = oCon.CargaRecordSet(sSql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Private Function GeneraRepAnalisisPendientesInterAgencias(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As Recordset
Dim prs  As ADODB.Recordset
Dim sDH  As String
Dim sCta As String
Dim rsAge As ADODB.Recordset
Dim rsDat As ADODB.Recordset

Set rsAge = New ADODB.Recordset
Set rsDat = New ADODB.Recordset

Dim sql As String
Dim lnI As Integer
Dim oAge As DActualizaDatosArea
On Error GoTo GeneraRepAnalisisPendientesErr
Set oCon = New DConecta
Set prs = CargaOpeCta(psOpeCod, , "0", True)
If Not prs.EOF Then
   sCta = MuestraListaRecordSet(prs, 0)
   RSClose prs
   If Not gbBitCentral Then
        Set oAge = New DActualizaDatosArea
        Set rsAge = oAge.GetAgencias(, False)
        Set oAge = Nothing
        lnI = 0
        ReDim Preserve lsCuentas(15, 2)
        sSql = ""
        While Not rsAge.EOF
            If oCon.AbreConexion Then 'Remota(Right(rsAge!Codigo, 2), False)
                lnI = lnI + 1
                lsCuentas(lnI, 1) = ""
                lsCuentas(lnI, 2) = ""
                
                sSql = " Select dFecTran, cMovNro, '" & gsCodAgeN & " ' + cMovDesc cMovDesc, cCtaContCod, nMovImporte, nSaldo, cPersona, ' ' cDocAbrev, cDocNro " _
                     & " From " & lsCuentas(lnI, 2) & "VistaMovAgePendiente TA " _
                     & " Where convert(varchar(8),TA.dFecTran,112) = '" & Format(pdFecha, gsFormatoMovFecha) & "' And cCtaContCod Like '__" & pnMoneda & "%' And TA.dFecTran > '05/01/2002' " _
                     & " ORDER BY cCtaContCod, cDocNro"
                Set rs = oCon.CargaRecordSet(sSql, adLockReadOnly)
                RecordSetAdiciona rsDat, rs
                RSClose rs
                     
            End If
            oCon.CierraConexion
            rsAge.MoveNext
        Wend
    End If
    'Operaciones de InterAgencias Anteriores
    oCon.AbreConexion
    sSql = " Select '" & Format(pdFecha, gsFormatoFecha) & "' dFecTran, TA.cMovNro, TA.cMovDesc cMovDesc, mc.cCtaContCod, nMovImporte, nMovImporte nSaldo, ' ' cPersona, ' ' cDocAbrev, 0 cDocNro " _
         & " From Mov TA " _
         & " Inner Join MovCta MC On TA.nMovNro = MC.nMovNro " _
         & " Where TA.cMovNro = '200202281845131120700JPNZ' And TA.nMovNro Not In (Select nMovNro From MovPendientesDet Where bExcluyeMov = 1 ) " _
         & " and Not Exists (SELECT mr.nMovNro FROM MovRef mr JOIN Mov m ON m.nMovNro = mr.nMovNro WHERE LEFT(m.cMovNro,8) < '" & Format(pdFecha, gsFormatoMovFecha) & "' and mr.nMovNroRef = TA.nMovNro ) ORDER BY cCtaContCod, cDocNro"
    Set rs = oCon.CargaRecordSet(sSql)
    RecordSetAdiciona rsDat, rs
    RSClose rs
    oCon.CierraConexion
    Set oCon = Nothing
    
    Set GeneraRepAnalisisPendientesInterAgencias = rsDat
End If
Exit Function
GeneraRepAnalisisPendientesErr:
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepAnalisisPendientesFaltante(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
Dim prs  As ADODB.Recordset
Dim sDH  As String
Dim sCta As String
Dim rsAge As ADODB.Recordset
Set rsAge = New ADODB.Recordset
Dim sql As String
Dim lnI As Integer
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseServer
Dim rsDat As ADODB.Recordset
Set rsDat = New ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Dim oConAdd As DConecta
Set oConAdd = New DConecta
Dim oDGeneral As nTipoCambio
Set oDGeneral = New nTipoCambio
Dim lnTipoCamb As Double

On Error GoTo GeneraRepAnalisisPendientesErr

Set prs = CargaOpeCta(psOpeCod, , "0", True)
   
If gbBitCentral Then
    sCta = MuestraListaRecordSet(prs, 0)
    RSClose prs
    oCon.AbreConexion
            
    lnTipoCamb = oDGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
   sSql = " Select dbo.GetFechaMov(M.cMovNro,103) dFecTran, cMovNro, SUbstring(M.cMovNro,15,5) + ' ' + M.cMovDesc cMovDesc,  " & sCta & " cCtaContCod, MSF.nMovImporte, MSF.nMovImporte nSaldo,PE.cPersNombre cPersona, '' cDocAbrev, '' cDocNro" _
        & " From Mov M" _
        & " Inner Join MovOpeVarias MSF On M.nMovNro = MSF.nMovNro" _
        & " Left  Join RRHH RH On RH.cUser = Right(M.cMovNro,4)" _
        & " Left  Join Persona PE On PE.cPersCod = RH.cPersCod" _
        & "    Where  (" _
        & "             (M.cOpeCod = '901020' And MSF.nMovImporte < -1 And MSF.nMoneda = 1)" _
        & "          Or (M.cOpeCod = '901020' And MSF.nMovImporte < (-1 / " & IIf(lnTipoCamb = 0, 1, lnTipoCamb) & ") And MSF.nMoneda = 2) )" _
        & "    And M.nMovFlag = 0 And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' And MSF.nMoneda = " & pnMoneda & "" _
        & "    And Not Exists (Select MR.nMovNro From Mov MR" _
        & "                             Inner Join MovRef MRR On MR.nMovNro = MRR.nMovNro" _
        & "                             Where MR.nMovFlag = 0 And MRR.nMovNroRef = M.nMovNro And MR.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "')" _
        & " Order By  M.cMovNro Desc"
        
        Set rs = oCon.CargaRecordSet(sSql)
        Set GeneraRepAnalisisPendientesFaltante = rs
Else
    If Not prs.EOF Then
        sCta = MuestraListaRecordSet(prs, 0)
        RSClose prs
        oCon.AbreConexion 'Remota "07"
        sql = " Select cValor From DbComunes..TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
        Set rsAge = oCon.CargaRecordSet(sql)
        
        lnI = 0
        ReDim Preserve lsCuentas(15, 2)
        sSql = ""
        rsAge.MoveFirst
        While Not rsAge.EOF
            If oConAdd.AbreConexion Then 'Remota(Right(rsAge!cValor, 2))
                lnI = lnI + 1
                'lsCuentas(lnI, 1) = Trim(oConAdd.GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "03"))
                lsCuentas(lnI, 2) = "" 'Trim(oConAdd.GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
                If sSql <> "" Then sSql = sSql & " UNION "
                sSql = " Select dFecTran, cMovNro, '" & gsCodAgeN & " ' + cMovDesc cMovDesc, cCtaContCod, nMovImporte,nSaldo,cPersona, '' cDocAbrev, cDocNro From " & lsCuentas(lnI, 2) & "Vistasobrantesfaltantes V" _
                     & " Where dFecTran < '" & Format(pdFecha, gsFormatoFecha) & "' And cCtaContCod Like '19" & pnMoneda & "80202%'" _
                     & " And Not Exists (Select nNumTran From " & lsCuentas(lnI, 2) & "VistaSobFalRef TA Where TA.nNumTran = V.cNroTranRef And dFecTran < '" & Format(pdFecha, gsFormatoFecha) & "' And cFlag Is Null) And cDocNro Not In (Select cDocNro from " & lsCuentas(lnI, 1) & "movpendientesDet Where bExcluyeDoc = 1 )"
                
                Set rs = oConAdd.CargaRecordSet(sSql)
                If Not (rs.EOF And rs.BOF) Then
                    RecordSetAdiciona rsDat, rs
                End If
            End If
            rs.Close
            rsAge.MoveNext
        Wend
        Set GeneraRepAnalisisPendientesFaltante = rs
    End If
End If
Exit Function
GeneraRepAnalisisPendientesErr:
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepAnalisisPendientesOPCertificada(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
Dim prs  As ADODB.Recordset
Dim sDH  As String
Dim sCta As String
Dim rsAge As ADODB.Recordset
Set rsAge = New ADODB.Recordset
Dim sql As String
Dim lnI As Integer
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseServer

On Error GoTo GeneraRepAnalisisPendientesErr

sql = " Select cValor From TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
'rsAge.Open sql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

lnI = 0
ReDim Preserve lsCuentas(15, 2)
sSql = ""
While Not rsAge.EOF
  '  If AbreConeccion(Right(rsAge!cValor, 2)) Then
        lnI = lnI + 1
        'lsCuentas(lnI, 1) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01", True))
        'lsCuentas(lnI, 2) = Trim(GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
        If sSql <> "" Then sSql = sSql & " UNION "
        sSql = sSql & " Select dFecTran,cCodCta,cMovNro, '" & gsCodAgeN & " ' + cMovDesc  cMovDesc,cCtaContCod,nMovImporte,nSaldo,cPersona,'OP' cDocAbrev, cDocNro From " & lsCuentas(lnI, 2) & "VistaOPCert TA Where TA.dFectran < '" & Format(pdFecha, gsFormatoFecha) & " 23:59:59' And Substring(TA.cCodCta,6,1) = '" & pnMoneda & "'" _
             & " And Not Exists (Select nNumTran From " & lsCuentas(lnI, 2) & "VistaOPCertRef TAR Where TAR.dFectran < '" & Format(pdFecha, gsFormatoFecha) & " 23:59:59' And TAR.cCodCta = TA.cCodCta And Substring(TAR.cCodCta,6,1) = '" & pnMoneda & "'" _
             & " And TAR.cNumDoc = TA.cDocNro)  And TA.cDocNro Not In (Select cDocNro From MovPendientesDet Where bExcluyeDoc = 1)"
  '  End If
    
    rsAge.MoveNext
Wend
sSql = sSql & " Order By dFecTran"
rs.CursorLocation = adUseClient
'rs.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText

Set GeneraRepAnalisisPendientesOPCertificada = rs

Exit Function
GeneraRepAnalisisPendientesErr:
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepAnalisisPendientesSobrante(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
Dim prs  As ADODB.Recordset
Dim sDH  As String
Dim sCta As String
Dim rsAge As ADODB.Recordset
Set rsAge = New ADODB.Recordset
Dim sql As String
Dim lnI As Integer
Set rs = New ADODB.Recordset
Dim rsDat As ADODB.Recordset
Set rsDat = New ADODB.Recordset
Dim oCon As DConecta
Dim oConAdd As DConecta
Set oCon = New DConecta
Set oConAdd = New DConecta
Dim oDGeneral As nTipoCambio
Set oDGeneral = New nTipoCambio
Dim lnTipoCamb As Double

rs.CursorLocation = adUseServer

On Error GoTo GeneraRepAnalisisPendientesErr

Set prs = CargaOpeCta(psOpeCod, , "0", True)
    
If gbBitCentral Then
    sCta = MuestraListaRecordSet(prs, 0)
    RSClose prs
    oCon.AbreConexion
    
    lnTipoCamb = oDGeneral.EmiteTipoCambio(pdFecha, TCFijoMes)
    sSql = " Select dbo.GetFechaMov(M.cMovNro,103) dFecTran, cMovNro, SUbstring(M.cMovNro,15,5) + ' ' + M.cMovDesc cMovDesc, " & sCta & " cCtaContCod, MSF.nMovImporte, MSF.nMovImporte nSaldo,PE.cPersNombre cPersona, '' cDocAbrev, '' cDocNro" _
        & " From Mov M" _
        & " Inner Join MovOpeVarias MSF On M.nMovNro = MSF.nMovNro" _
        & " Left  Join RRHH RH On RH.cUser = Right(M.cMovNro,4)" _
        & " Left  Join Persona PE On PE.cPersCod = RH.cPersCod" _
        & "    Where  (" _
        & "             (M.cOpeCod = '901014' And MSF.nMovImporte > 5 And MSF.nMoneda = 1)" _
        & "          Or (M.cOpeCod = '901014' And MSF.nMovImporte > (5 / " & IIf(lnTipoCamb = 0, 1, lnTipoCamb) & ") And MSF.nMoneda = 2))" _
        & "    And M.nMovFlag = 0 And M.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "' And MSF.nMoneda = " & pnMoneda & "" _
        & "    And Not Exists (Select MR.nMovNro From Mov MR" _
        & "                             Inner Join MovRef MRR On MR.nMovNro = MRR.nMovNro" _
        & "                             Where MR.nMovFlag = " & gMovFlagEliminado & " And MRR.nMovNroRef = M.nMovNro And MR.cMovNro < '" & Format(DateAdd("d", 1, pdFecha), gsFormatoMovFecha) & "')" _
        & " Order By  M.cMovNro Desc"
            
        Set rs = oCon.CargaRecordSet(sSql)
        Set GeneraRepAnalisisPendientesSobrante = rs

Else
    oCon.AbreConexion 'Remota "07"
    
    If Not prs.EOF Then
        sCta = MuestraListaRecordSet(prs, 0)
        RSClose prs
           
        sql = " Select cValor From TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
        Set rsAge = oCon.CargaRecordSet(sql)
        
        lnI = 0
        ReDim Preserve lsCuentas(15, 2)
        sSql = ""
        While Not rsAge.EOF
            If oConAdd.AbreConexion Then 'Remota(Right(rsAge!cValor, 2))
                lnI = lnI + 1
'                lsCuentas(lnI, 1) = Trim(oConAdd.GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01", True))
'                lsCuentas(lnI, 2) = Trim(oConAdd.GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
        
                sSql = " Select dFecTran, cMovNro, '" & gsCodAgeN & " ' + cMovDesc cMovDesc, REPLACE(cCtaContCod,'AG','" & Right(gsCodAgeN, 2) & "') cCtaContCod,nMovImporte,nSaldo,cPersona, '' cDocAbrev, cDocNro From Vistasobrantesfaltantes V" _
                     & " Where dFecTran < '" & Format(pdFecha, gsFormatoFecha) & "' And cCtaContCod Like '29" & pnMoneda & "20102%'" _
                     & " And Not Exists (Select nNumTran From VistaSobFalRef TA Where TA.nNumTran = V.cNroTranRef And dFecTran < '" & Format(pdFecha, gsFormatoFecha) & "' And cFlag Is Null) And cDocNro Not In (Select cDocNro from [128.107.2.3].dbAdmin.dbo.movpendientesDet Where bExcluyeDoc = 1 )"
                Set rs = oConAdd.CargaRecordSet(sSql)
                RecordSetAdiciona rsDat, rs
                RSClose rs
            End If
            
            rsAge.MoveNext
        Wend
        rsDat.Sort = "cCtaContCod Desc"
        Set GeneraRepAnalisisPendientesSobrante = rsDat
    End If
End If
Exit Function
GeneraRepAnalisisPendientesErr:
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepAnalisisPendientesCobServicios(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
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
    Dim oCon As DConecta
    Dim oConAdd As DConecta
    Set oCon = New DConecta
    Set oConAdd = New DConecta
    
    On Error GoTo GeneraRepAnalisisPendientesErr
    
    If gbBitCentral Then
        Set GeneraRepAnalisisPendientesCobServicios = Nothing
    Else
        oCon.AbreConexion 'Remota "07"
        Set prs = CargaOpeCta(psOpeCod, , "0", True)
        If Not prs.EOF Then
           sCta = MuestraListaRecordSet(prs, 0)
           RSClose prs
           
            sql = " Select cValor From TablaCod TC where cValor Like '112%' And cValor Not Like '1129%' And cCodTab Like '47__'"
            Set rsAge = oCon.CargaRecordSet(sql)
            
            sCta = "211501"
            lnI = 0
            ReDim Preserve lsCuentas(15, 2)
            sSql = ""
            While Not rsAge.EOF
                If oConAdd.AbreConexion Then 'Remota(Right(rsAge!cValor, 2))
                    lnI = lnI + 1
                    'lsAgeConsol = ReadVarSisCon("AHO", "cAgeConsolSer", dbCmactN)
'                    lsCuentas(lnI, 1) = Trim(oConAdd.GetCadenaConexionEnlazado(Right(lsAgeConsol, 2), "01"))
'                    lsCuentas(lnI, 2) = Trim(oConAdd.GetCadenaConexionEnlazado(Right(rsAge!cValor, 2), "01"))
                    If sSql <> "" Then sSql = sSql & " UNION "
                    sSql = sSql & " Select V.dFecTran, V.cMovNro, '" & gsCodAgeN & " ' + V.cMovDesc cMovDesc, '" & sCta & "' cCtaContCod, nMonTran nMovImporte, nMonTran nSaldo,cPersona, '' cDocAbrev, cDocNro From " & lsCuentas(lnI, 2) & "VistaServ V" _
                         & " Where V.dFecTran > '2002/05/02' And V.dFecTran < '" & Format(pdFecha, gsFormatoFecha) & "' And V.cCodCta Like '_____" & pnMoneda & "%'" _
                         & " And Not Exists (Select nNumTran From " & lsCuentas(lnI, 1) & "VistaServRef VR Where VR.nNumTran = V.NumRef And dFecTran <= '" & Format(DateAdd("d", 1, pdFecha), gsFormatoFecha) & "')"
                End If
                rsAge.MoveNext
            Wend
            sSql = sSql & " order by V.cMovDesc, V.dFecTran Desc"
            Set rs = oCon.CargaRecordSet(sSql)
            Set GeneraRepAnalisisPendientesCobServicios = rs
        End If
    End If
    
    Exit Function
GeneraRepAnalisisPendientesErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Function

Private Function GeneraRepEntregaARendir(ByVal psOpeCod As String, ByVal pnMoneda As Integer, ByVal pdFecha As Date) As ADODB.Recordset
On Error GoTo GeneraRepEntregaARendirErr

Dim lsCtaCod As String
Dim lsCtaContDev As String
Dim lsTipoDoc As String
Dim lsObjARendir As String
Set rs = CargaOpeCta(psOpeCod, "H")
If rs.EOF Then
    RSClose rs
    MsgBox "No se definió Cuenta Contable para Analizar Pendiente", vbInformation, "¡Aviso!"
    Exit Function
End If

lsCtaCod = Left(rs!cCtaContCod, 2) & pnMoneda & Mid(rs!cCtaContCod, 4, 22)
lsCtaContDev = ""

Dim oARend As New NARendir
Set GeneraRepEntregaARendir = oARend.ARendirPendientesTotal(Val(pnMoneda), lsCtaCod, pdFecha)
Set oARend = Nothing

Exit Function
GeneraRepEntregaARendirErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
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
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Se Imprimio el Reporte de Pendientes : " & txtOpeDes.Text & "hasta la Fecha " & txtFecha.Text
            Set objPista = Nothing
            '*******
    End If
       
       
End Sub

Private Sub cmdOpeExa_Click()
Dim prs As ADODB.Recordset
Set prs = CargaOpeTpo("7013__")
If prs.EOF Then
   RSClose prs
   Exit Sub
End If
'frmDescObjeto.Inicio prs, "", 2
'If frmDescObjeto.lOk Then
'   txtOpeCod = gaObj(0, 0, 0)
'   txtOpeDes = gaObj(0, 1, 0)
   
   'If txtOpeCod = "701301" Then
      'Me.mskFecIni.Enabled = True
   '   Me.mskFecIni.SetFocus
   'Else
      'Me.mskFecIni.Enabled = False
      txtFecha.SetFocus
   'End If
   
'End If
RSClose prs
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oOpe As New DOperacion
    CentraForm Me
    txtFecha = gdFecSis
    FormatoFlex
    pnMoneda = 1
    txtOpeCod.rs = oOpe.CargaOpeTpo(Left(gOpePendOpeAgencias, 4), True, , 0, 2)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    CierraConexion
End Sub

Private Sub optMoneda_Click(Index As Integer)
pnMoneda = Index + 1
End Sub

Private Sub optMoneda_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub

Private Sub txtFecha_GotFocus()
    txtFecha.SelStart = 0
    txtFecha.SelLength = 50
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If optMoneda(0).value = True Then
      optMoneda(0).SetFocus
   Else
      optMoneda(1).SetFocus
   End If
End If
End Sub

Private Sub txtOpeCod_EmiteDatos()
txtOpeDes = txtOpeCod.psDescripcion
If txtOpeCod <> "" And txtFecha.Visible Then
    txtFecha.SetFocus
End If
End Sub

Private Function ValidaOperacion(psOpeCod As String) As Boolean
End Function

Private Sub GeneraReporteE()
    Dim i As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lnAcum As Currency
    
    Dim sTipoGara As String
    Dim sTipoCred As String
    
    On Error GoTo ErrPendientes
    
    Me.MousePointer = 11
    xlHoja1.Range("A1:D" & Trim(Str(Me.fg.Rows + 4))).NumberFormat = "@"
    For i = 0 To Me.fg.Rows - 1
        lnAcum = 0
        For j = 1 To Me.fg.Cols - 1
            xlHoja1.Cells(i + 4, j + 1) = Me.fg.TextMatrix(i, j)
        Next j
    Next i
    
    xlHoja1.Range("E4:F" & Trim(Str(Me.fg.Rows + 4))).NumberFormat = "#,##0.00"
    xlHoja1.Range("A4:A" & Trim(Str(Me.fg.Rows + 4))).Font.Bold = True
    xlHoja1.Range("4:4").Font.Bold = True
    
'    xlHoja1.Cells.Select
     
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
    
    xlHoja1.Range("E4:F" & Trim(Str(Me.fg.Rows + 4))).Font.Bold = True
    xlHoja1.Range("4:4").Font.Bold = True
    
    xlHoja1.Range("A1:A1").ColumnWidth = 2
    xlHoja1.Range("B1:B1").ColumnWidth = 14
    xlHoja1.Range("C1:C1").ColumnWidth = 70
    xlHoja1.Range("D1:D1").ColumnWidth = 8
    xlHoja1.Range("E1:E1").ColumnWidth = 12
    xlHoja1.Range("F1:F1").ColumnWidth = 12
    xlHoja1.Range("G1:G1").ColumnWidth = 40
    xlHoja1.Range("H1:H1").ColumnWidth = 14
     
    xlHoja1.Range("A4:H" & Me.fg.Rows + 2).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A4:H" & Me.fg.Rows + 2).Borders(xlDiagonalUp).LineStyle = xlNone
    With xlHoja1.Range("A4:H" & Me.fg.Rows + 2).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A4:H" & Me.fg.Rows + 3).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A4:H" & Me.fg.Rows + 3).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A4:H" & Me.fg.Rows + 3).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A4:H" & Me.fg.Rows + 3).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A4:H" & Me.fg.Rows + 3).Borders(xlInsideHorizontal)
    '    .LineStyle = xlContinuous
    '    .Weight = xlThin
    '    .ColorIndex = xlAutomatic
    End With
    
    xlHoja1.Cells(1, 1) = gsNomCmac
    xlHoja1.Cells(2, 1) = Trim(UCase(Me.txtOpeDes.Text)) & " AL " & txtFecha
    xlHoja1.Range("A2:H2").Font.Bold = True
    xlHoja1.Range("A2:H2").Font.Size = 14
    xlHoja1.Range("A2:H2").Merge
    xlHoja1.Range("A2:H2").HorizontalAlignment = vbCenter
    
    With xlHoja1.PageSetup
        .RightHeader = "&P"
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
'        .CenterVertically = True
        .Orientation = xlLandscape
        .PaperSize = xlPaperLetter
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 65
    End With
    Me.MousePointer = 0
Exit Sub
ErrPendientes:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
    
End Sub

