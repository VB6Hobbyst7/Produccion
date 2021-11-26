VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFoncodesNotas 
   Caption         =   "Foncodes"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   Icon            =   "frmFoncodesNotas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   0
      Width           =   5115
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   345
         Left            =   3720
         TabIndex        =   2
         Top             =   180
         Width           =   1275
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   210
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "al"
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
         Height          =   315
         Left            =   2100
         TabIndex        =   14
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Fechas"
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
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   5820
      TabIndex        =   7
      Top             =   4200
      Width           =   1275
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   4170
      Width           =   1275
   End
   Begin Sicmact.FlexEdit fg 
      Height          =   2955
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   5212
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Codigo-Agencia-Desembolsos-Pagos"
      EncabezadosAnchos=   "400-800-3000-1200-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-R-R"
      FormatosEdit    =   "0-0-0-2-2"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.CommandButton cmdAbonar 
      Caption         =   "&Abonar"
      Height          =   345
      Left            =   3240
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar"
      Height          =   345
      Left            =   4530
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label lblTotal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "TOTALES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2820
      TabIndex        =   12
      Top             =   3690
      Width           =   945
   End
   Begin VB.Label txtHaber 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
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
      Left            =   4260
      TabIndex        =   11
      Top             =   3690
      Width           =   1290
   End
   Begin VB.Label txtDebe 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "0.00"
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
      Left            =   5520
      TabIndex        =   10
      Top             =   3690
      Width           =   1290
   End
   Begin VB.Label lblTotalC 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2610
      TabIndex        =   13
      Top             =   3630
      Width           =   4275
   End
End
Attribute VB_Name = "frmFoncodesNotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oAge As DActualizaDatosArea
Dim rs   As ADODB.Recordset
Dim sSql As String
Dim N    As Integer
Dim oCon As DConecta
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub cmdAbonar_Click()
On Error GoTo Abonarerr
If nVal(txtDebe) = 0 Then
    MsgBox "Falta Obtener información de las Agencias", vbInformation, "¡Aviso!"
    Exit Sub
End If
If gbBitCentral Then
    gsOpeCod = gCGPagProvAbonoFoncodesCent
Else
    gsOpeCod = gCGPagProvAbonoFoncodesDist
End If
gnDocTpo = TpoDocNotaAbono
gsGlosa = "ABONO POR COBRANZA CREDITOS LINEA FONCODES : " & txtFecha
GrabaOperacionFoncodes nVal(txtDebe)
Exit Sub
Abonarerr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCargar_Click()
If nVal(txtHaber) = 0 Then
    MsgBox "Falta Obtener información de las Agencias", vbInformation, "¡Aviso!"
    Exit Sub
End If

'Nota de Cargo
If gbBitCentral Then
    gsOpeCod = gCGPagProvCargoFoncodesCent
Else
    gsOpeCod = gCGPagProvCargoFoncodesDist
End If
gsGlosa = "CARGO POR DESEMBOLSO CREDITOS LINEA FONCODES : " & txtFecha
gnDocTpo = TpoDocNotaCargo

GrabaOperacionFoncodes nVal(txtHaber)

End Sub

Private Sub GrabaOperacionFoncodes(pnImporte As Currency)
Dim lsDocNRo     As String
Dim lsFechaDoc   As String
Dim lsDocumento  As String
Dim lsPersNombre As String
Dim lsUbigeo     As String
Dim lsCuentaAho  As String
Dim lsPersDireccion As String
Dim lsMovNroNegocio As String
Dim oMov As New DMov
Dim oContImp As NContImprimir
Dim oCon As New DConecta

Dim oDocRec As NDocRec
Dim oConst  As New NConstSistemas
lsCuentaAho = oConst.LeeConstSistema(73)
Set oConst = Nothing
If gbBitCentral Then
    If oMov.BuscarMov(Format(gdFecSis, gsFormatoMovFecha), " cOpeCod = '" & gsOpeCod & "' ") Then
        MsgBox "Operación ya fue realizada. Consultar Estado de Cuenta de FONCODES", vbInformation, "¡Aviso!"
        Exit Sub
    End If
Else
    If oCon.AbreConexion Then 'Remota(Left(lsCuentaAho, 2))
        sSql = "SELECT nNumTran FROM TransAho WHERE cCodOpe = '" & gsOpeCod & "' and convert(varchar(8), dFectran, 112) = '" & Format(gdFecSis, gsFormatoMovFecha) & "' and cFlag is NULL "
        Set rs = oCon.CargaRecordSet(sSql)
        If Not rs.EOF Then
            MsgBox "Operación ya fue realizada. Consultar Estado de Cuenta de FONCODES", vbInformation, "¡Aviso!"
            RSClose rs
            Exit Sub
        End If
    Else
        MsgBox "No se puede establecer conexión con Agencia " & Left(lsCuentaAho, 2), vbInformation, "¡Aviso!"
        Exit Sub
    End If
    oCon.CierraConexion
End If

frmNotaCargoAbono.Inicio gnDocTpo, pnImporte, gdFecSis, gsGlosa, gsOpeCod, , , lsPersNombre, lsCuentaAho
If frmNotaCargoAbono.vbOk Then
    lsDocNRo = frmNotaCargoAbono.NroNotaCA
    lsFechaDoc = frmNotaCargoAbono.FechaNotaCA
    gsGlosa = frmNotaCargoAbono.Glosa
    lsDocumento = frmNotaCargoAbono.NotaCargoAbono
    lsPersNombre = frmNotaCargoAbono.PersNombre
    lsPersDireccion = frmNotaCargoAbono.PersDireccion
    lsUbigeo = frmNotaCargoAbono.PersUbigeo
    lsCuentaAho = frmNotaCargoAbono.CuentaAhoNro
Else
    Exit Sub
End If

If MsgBox("Desea Grabar Operación ?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    Set oDocRec = New NDocRec
    gsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    'lsDocNRo = oDocRec.GetNroNotaCargoAbono(gnDocTpo)
    oMov.InsertaNotaAbonoCargo gnDocTpo, lsDocNRo, gNCNARegistrado, gNCFoncodes, pnImporte
    oMov.InsertaNotaAbonoCargoEst gnDocTpo, lsDocNRo, gNCNARegistrado, gsMovNro
    oMov.InsertaRegDocCuenta gnDocTpo, lsDocNRo, lsCuentaAho

    Dim oCapta As NCapMovimientos
    Dim lnSaldo As Double
    Set oCapta = New NCapMovimientos
    If gbBitCentral Then
        lsMovNroNegocio = oMov.GeneraMovNro(, , , gsMovNro)
        If gnDocTpo = TpoDocNotaAbono Then
            lnSaldo = oCapta.CapAbonoCuentaAho(lsCuentaAho, pnImporte, gCGArendirAtencAbonoCent, lsMovNroNegocio, gsGlosa, gnDocTpo, lsDocNRo, , , , , , , , , , True, False, , , oMov.GetConexion, False)
        Else
            lnSaldo = oCapta.CapCargoCuentaAho(lsCuentaAho, pnImporte, gCGArendirAtencCargoCent, lsMovNroNegocio, gsGlosa, gnDocTpo, lsDocNRo, , , , , , , , , , True, "", , oMov.GetConexion, False)
        End If
    Else
        Dim oDis As New NRHProcesosCierre
        
        If oCon.AbreConexion Then 'Remota(Left(lsCuentaAho, 2))
            oCon.BeginTrans
            If gnDocTpo = TpoDocNotaAbono Then
                lnSaldo = oDis.Abono(lsCuentaAho, pnImporte, gsOpeCod, gsOpeCod, "112" & Left(lsCuentaAho, 2), Right(gsMovNro, 4), lsDocNRo, "ABONO : " & gsGlosa, oCon, GetFechaMov(gsMovNro, True))
            Else
                lnSaldo = oDis.Cargo(lsCuentaAho, pnImporte, gsOpeCod, "112" & Left(lsCuentaAho, 2), Right(gsMovNro, 4), lsDocNRo, "CARGO : " & gsGlosa, oCon, GetFechaMov(gsMovNro, True))
            End If
            oCon.CommitTrans
        End If
        oCon.CierraConexion
        Set oCon = Nothing
        Set oDis = Nothing
    End If
    Set oContImp = New NContImprimir
    'lsDocumento = oContImp.ImprimeNotaAbono(GetFechaMov(gsMovNro, True), pnImporte, gsGlosa, lsCuentaAho, lsPersNombre)
    lsDocumento = oContImp.ImprimeNotaCargoAbono(lsDocNRo, gsGlosa, CCur(pnImporte), _
                    lsPersNombre, lsPersDireccion, lsUbigeo, gdFecSis, Mid(gsOpeCod, 3, 1), lsCuentaAho, gnDocTpo, gsNomAge, gsCodUser)
    EnviaPrevio lsDocumento, "FONCODES: Operaciones", gnLinPage, True
    
    lsDocumento = oDis.ImprimeBoletaCad(gdFecSis, "FONCODES:", IIf(gnDocTpo = TpoDocNotaCargo, "RETIRO", "DEPOSITO") & " FONCODES" & "*Nro." & lsDocNRo, "", pnImporte, lsPersNombre, lsCuentaAho, "", CCur(lnSaldo), 0, IIf(gnDocTpo = TpoDocNotaCargo, "Nota Cargo", "Nota Abono"), 0, lnSaldo, False, True, , , , True, , , , False, gsNomAge) & oImpresora.gPrnSaltoPagina

    Dim lsDocVentanilla As String
    Dim lbimp  As Boolean
    Dim oPlant As dPlantilla
    Dim oNPlant As NPlantilla
    Dim oPrevio As clsPrevioFinan
    
    Set oPrevio = New clsPrevioFinan
    Set oPlant = New dPlantilla
    Set oNPlant = New NPlantilla
    
    'Lee Recibos Anteriores
    lsDocVentanilla = oNPlant.GetPlantillaDoc("RecCajero") & lsDocumento
    If MsgBox(" ¿ Desea Imprimir Comprobante(s) de Ventanilla ? ", vbQuestion + vbYesNo, "¡Confirmación") = vbYes Then
        lbimp = True
        Do While lbimp
            oPrevio.ShowImpreSpool lsDocVentanilla, False, 22
            If MsgBox("Desea Reimprimir Comprobante de Ventanilla??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbimp = False
            End If
        Loop
        oPlant.GrabaPlantilla "RecCajero", "Documento de Ventanilla, para impresiones en Batch", ""
    Else
        oPlant.GrabaPlantilla "RecCajero", "Documento de Ventanilla, para impresiones en Batch", lsDocVentanilla
    End If
End If
End Sub

Private Sub cmdImprimir_Click()
Dim lsArchivo As String
Dim lsRuta    As String

If fg.TextMatrix(1, 3) = "" Then
    MsgBox "Falta Obtener información de las Agencias", vbInformation, "¡Aviso!"
    Exit Sub
End If
lsRuta = App.path & "\Spooler\"
lsArchivo = lsRuta & "FONCODES" & "_" & Left(Format(txtFecha, gsFormatoMovFecha), 6) & ".xls"

ExcelBegin lsArchivo, xlAplicacion, xlLibro, True
ExcelAddHoja Replace(txtFecha2, "/", "-"), xlLibro, xlHoja1, True
CabeceraExcel
For N = 1 To fg.Rows - 1
    xlHoja1.Cells(N + 7, 2) = fg.TextMatrix(N, 1)
    xlHoja1.Cells(N + 7, 3) = fg.TextMatrix(N, 2)
    xlHoja1.Cells(N + 7, 4) = fg.TextMatrix(N, 3)
    xlHoja1.Cells(N + 7, 5) = fg.TextMatrix(N, 4)
Next
ExcelCuadro xlHoja1, 2, 8, 5, N + 6
xlHoja1.Cells(N + 7, 3) = "TOTALES"
xlHoja1.Cells(N + 7, 4) = txtHaber
xlHoja1.Cells(N + 7, 5) = txtDebe
xlHoja1.Range(xlHoja1.Cells(N + 7, 4), xlHoja1.Cells(N + 7, 5)).HorizontalAlignment = xlHAlignRight
xlHoja1.Range(xlHoja1.Cells(N + 7, 2), xlHoja1.Cells(N + 7, 5)).Font.Bold = True
ExcelCuadro xlHoja1, 2, N + 7, 5, N + 7
ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
CargaArchivo lsArchivo, lsRuta
End Sub

Private Sub CabeceraExcel()
Dim nCol As Integer
Dim sCol As String
xlHoja1.PageSetup.Zoom = 80
xlHoja1.Cells(1, 1) = gsNomCmac
'xlHoja1.Cells(2, 6) = "Código : " & Left(gsCodAge, 3)
xlHoja1.Cells(3, 1) = "FONCODES: CUADRO RESUMEN DE OPERACIONES"
xlHoja1.Cells(4, 2) = "Del " & txtFecha & " al " & txtFecha2
xlHoja1.Cells(5, 2) = "( En Nuevos Soles )"

xlHoja1.Range("A3:G3").Merge
xlHoja1.Range("A4:G4").Merge
xlHoja1.Range("A5:G5").Merge
xlHoja1.Range("A3:G5").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("A3:G3").Font.Size = 13

xlHoja1.Cells(7, 2) = "Código"
xlHoja1.Cells(7, 3) = "Agencia"
xlHoja1.Cells(7, 4) = "Desembolsos"
xlHoja1.Cells(7, 5) = "Pagos"
ExcelCuadro xlHoja1, 2, 7, 5, 7

xlHoja1.Range("B1:G7").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("B7:G7").Font.Bold = True
xlHoja1.Range("B1:B1").ColumnWidth = 8
xlHoja1.Range("C1:C1").ColumnWidth = 30
xlHoja1.Range("D1:D1").ColumnWidth = 13
xlHoja1.Range("E1:E1").ColumnWidth = 13

End Sub

Private Sub cmdProcesar_Click()
Dim lsSrvNegocio As String

If fg.TextMatrix(1, 1) = "" Then
    Exit Sub
End If
Set oCon = New DConecta
oCon.AbreConexion
'lsSrvNegocio = oCon.GetCadenaConexionEnlazado("07", "01")
sSql = "SELECT aaa.cAgeCod, cAgeDescripcion, nDesembolso, nPago + ISNULL(nJudicial,0) nPago " _
     & "FROM (SELECT a.cAgeCod, a.cAgeDescripcion, ISNULL(sum(IsNull(nMontoDesembN,0)),0) nDesembolso, " _
     & "      ISNULL(Sum(IsNull(nCapPag,0) + IsNull(nIntPag,0) + IsNull(nMoraPag,0) + IsNull(nMontoRefinan,0)),0) nPago " _
     & "      FROM Agencias a left JOIN ColocEstadDiaCred e ON e.cCodAge = a.cAgeCod and convert(varchar(8), dEstad, 112) BETWEEN '" & Format(txtFecha, gsFormatoMovFecha) & "' and '" & Format(txtFecha2, gsFormatoMovFecha) & "' " _
     & "           And Left(cLineaCred,2) = '04' " _
     & "      GROUP BY a.cAgeCod, a.cAgeDescripcion " _
     & "     ) aaa LEFT JOIN " _
     & "     (Select  Substring(MCD.cCtaCod,4,2) cAgeCod, Sum(MCD.nMonto) nJudicial From ColocRecup CR" _
     & "      Inner Join MovColDet MCD On CR.cCtaCod = MCD.cCtaCod" _
     & "      Inner Join Mov M On M.nMovNro = MCD.nMovNro And nPrdConceptoCod = 3000" _
     & "      Inner Join Colocaciones C On CR.cCtaCod = C.cCtaCod" _
     & "      Where nMovFlag <> 1 And cMovNro Between '" & Format(txtFecha, gsFormatoMovFecha) & "' and '" & Format(txtFecha2, gsFormatoMovFecha) & "' And Left(cLineaCred,2) = '04'" _
     & "      Group By Substring(MCD.cCtaCod,4,2)) bbb ON bbb.cAgeCod = aaa.cAgeCod " _
     & "ORDER BY aaa.cAgeCod "
Set fg.Recordset = oCon.CargaRecordSet(sSql)
Me.txtDebe = Format(fg.SumaRow(4), gsFormatoNumeroView)
Me.txtHaber = Format(fg.SumaRow(3), gsFormatoNumeroView)
cmdAbonar.Visible = True
cmdCargar.Visible = True
If fg.Enabled Then
    fg.SetFocus
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Set oAge = New DActualizaDatosArea
Set rs = oAge.GetAgencias(, False)
Me.txtFecha = gdFecSis - 1
Me.txtFecha2 = gdFecSis - 1

N = 0
Do While Not rs.EOF
    N = N + 1
    fg.AdicionaFila
    fg.TextMatrix(N, 1) = rs!Codigo
    fg.TextMatrix(N, 2) = rs!Descripcion
    rs.MoveNext
Loop
RSClose rs
Set oAge = Nothing
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha) = True Then
       txtFecha2.SetFocus
    End If
End If
End Sub

Private Sub txtFecha2_GotFocus()
fEnfoque txtFecha2
End Sub

Private Sub txtFecha2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtFecha2) = True Then
       cmdProcesar.SetFocus
    End If
End If
End Sub


