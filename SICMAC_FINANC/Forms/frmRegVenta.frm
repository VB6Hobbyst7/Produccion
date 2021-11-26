VERSION 5.00
Begin VB.Form frmRegVenta 
   Caption         =   "Registro de Ventas"
   ClientHeight    =   6330
   ClientLeft      =   345
   ClientTop       =   2190
   ClientWidth     =   11130
   Icon            =   "frmRegVenta.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.FlexEdit fgReg 
      Height          =   4815
      Left            =   150
      TabIndex        =   18
      Top             =   870
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8493
      Cols0           =   23
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmRegVenta.frx":030A
      EncabezadosAnchos=   "350-1100-450-600-1000-2000-3000-1300-3000-1100-1100-1100-1100-1100-0-0-0-0-0-0-1200-1200-1200"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-12-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-L-L-L-L-L-R-R-R-R-R-C-C-C-L-C-C-R-L-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-2-2-2-0-0-0-0-5-0-2-1-1"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   345
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame Frame3 
      Caption         =   "&Periodo"
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
      Height          =   735
      Left            =   150
      TabIndex        =   14
      Top             =   60
      Width           =   5340
      Begin VB.CommandButton cmdVer 
         Caption         =   "&Ver"
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
         Left            =   4050
         TabIndex        =   2
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3060
         MaxLength       =   4
         TabIndex        =   1
         Top             =   270
         Width           =   855
      End
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmRegVenta.frx":03D7
         Left            =   570
         List            =   "frmRegVenta.frx":03FF
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Año :"
         Height          =   195
         Left            =   2610
         TabIndex        =   16
         Top             =   330
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mes :"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   315
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consolidar de ..."
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
      Height          =   735
      Left            =   5610
      TabIndex        =   12
      Top             =   60
      Visible         =   0   'False
      Width           =   5475
      Begin VB.CommandButton cmdServicio 
         Caption         =   "Se&rvicios"
         Enabled         =   0   'False
         Height          =   405
         Left            =   4050
         TabIndex        =   6
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdBienes 
         Caption         =   "Subasta &Bienes"
         Enabled         =   0   'False
         Height          =   405
         Left            =   2730
         TabIndex        =   5
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdJoyas 
         Caption         =   "Subasta &Joyas"
         Height          =   405
         Left            =   1410
         TabIndex        =   4
         Top             =   240
         Width           =   1305
      End
      Begin VB.CommandButton cmdCustodia 
         Caption         =   "&Custodia"
         Height          =   405
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9540
      TabIndex        =   11
      Top             =   5850
      Width           =   1305
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
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
      Left            =   8070
      TabIndex        =   10
      Top             =   5850
      Width           =   1425
   End
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Width           =   4335
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   2880
         TabIndex        =   9
         Top             =   180
         Width           =   1395
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   345
         Left            =   1470
         TabIndex        =   8
         Top             =   195
         Width           =   1395
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   180
         Width           =   1395
      End
   End
   Begin VB.OLE OleExcel 
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   405
      Left            =   4560
      SizeMode        =   1  'Stretch
      TabIndex        =   17
      Top             =   5790
      Visible         =   0   'False
      Width           =   825
   End
End
Attribute VB_Name = "frmRegVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql  As String
Dim rs    As New ADODB.Recordset
Dim sDocs As String
Dim dFecha1 As Date
Dim dFecha2 As Date
Dim nTasaIGV As Currency

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim oReg As DRegVenta

Dim lbReporteFormatoLima As Boolean

'YIHU20152002-ERS181-2014**************
Dim nMoneda As Integer
Dim psTipoAccion As String 'NAGL 20170805
Public Sub setnMoneda(N As Integer)
    nMoneda = N
End Sub
'END YIHU***

Public Sub FormatoRegVenta(pbFormatoLima As Boolean)
    lbReporteFormatoLima = pbFormatoLima
End Sub

Private Sub AsignaValores(nItem As Long, prs As ADODB.Recordset)
   fgReg.TextMatrix(nItem, 1) = Format(prs!dDocFecha, "dd/mm/yyyy")
   fgReg.TextMatrix(nItem, 2) = Format(prs!nDocTpo, "00")
   fgReg.TextMatrix(nItem, 3) = Format(Mid(prs!cDocNroNew, 1, 4), "0000") 'NAGL ERS 012-2017 Se cambio de Long.Serie de 3 a 4 Dig y el Campo.
    fgReg.TextMatrix(nItem, 4) = Mid(Trim(prs!cDocNroNew), 5, 20) 'NAGL ERS 012-2017 Se cambio de Long.NroRest de 8 a 7 Dig y el Campo. 'Se quito el Format de long 7 JIPR 20180721
   fgReg.TextMatrix(nItem, 5) = IIf(IsNull(prs!cRuc), "", prs!cRuc)
   fgReg.TextMatrix(nItem, 6) = IIf(IsNull(prs!cPersNombre), "", prs!cPersNombre)
   fgReg.TextMatrix(nItem, 7) = prs!cCtaCod
   fgReg.TextMatrix(nItem, 8) = prs!cDescrip
   If prs!nIGV <> 0 Then
      fgReg.TextMatrix(nItem, 9) = Format(prs!nVVenta, gsFormatoNumeroView)
   Else
      fgReg.TextMatrix(nItem, 10) = Format(prs!nVVenta, gsFormatoNumeroView)
   End If
   fgReg.TextMatrix(nItem, 11) = Format(prs!nIGV, gsFormatoNumeroView)
   fgReg.TextMatrix(nItem, 12) = Format(prs!nOtrImp, gsFormatoNumeroView)
   fgReg.TextMatrix(nItem, 13) = Format(prs!nPVenta, gsFormatoNumeroView)
   fgReg.TextMatrix(nItem, 14) = prs!cOpeTpo
   fgReg.TextMatrix(nItem, 15) = Format(rs!dDocFecha, gsFormatoFechaHoraView)
   fgReg.TextMatrix(nItem, 16) = rs!nDocTpo
   fgReg.TextMatrix(nItem, 17) = IIf(IsNull(rs!cDocNroRefe), "", rs!cDocNroRefe)
   fgReg.TextMatrix(nItem, 18) = Format(IIf(IsNull(rs!dDocRefeFec), "", rs!dDocRefeFec), "dd/mm/yyyy")
   fgReg.TextMatrix(nItem, 19) = IIf(prs!cTipoDoc = 2, 6, prs!cTipoDoc)
   
      'YIHU20152002-ERS181-2014**************
   
   
   If IsNull(prs!nTipoCambio) Then
      fgReg.TextMatrix(nItem, 21) = Format(0, "##,###,##0.000")
   Else
         fgReg.TextMatrix(nItem, 21) = Format(prs!nTipoCambio, "##,###,##0.000")
   End If
   
   If IsNull(prs!nMoneda) Then
      fgReg.TextMatrix(nItem, 22) = "MN"
   Else
      If prs!nMoneda <> 1 Then
        fgReg.TextMatrix(nItem, 22) = "ME"
        If prs!nTipoCambio <> 0 Then
            fgReg.TextMatrix(nItem, 20) = Format(prs!nPVenta / prs!nTipoCambio, "##,###,##0.000")
        Else
            fgReg.TextMatrix(nItem, 20) = Format(0, "##,###,##0.000")
        End If
      Else
        fgReg.TextMatrix(nItem, 22) = "MN"
      End If
   End If
   'END YIHU *****************************
   
   
End Sub

Private Function DefineFechas() As Boolean
DefineFechas = False
If txtAnio = "" Then
   MsgBox "Falta definir año de proceso", vbInformation, "!Aviso!"
   Exit Function
End If
If cboMes.ListIndex = -1 Then
   MsgBox "Falta definir mes de proceso", vbInformation, "!Aviso!"
   Exit Function
End If
dFecha1 = CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & Format(txtAnio, "0000"))
dFecha2 = DateAdd("m", 1, dFecha1) - 1
End Function

Private Function GetDocRegistro(nTipo As Integer) As String
Dim K As Integer
Dim sTexto As String
sTexto = ""
For K = 1 To fgReg.Rows - 1
   If fgReg.TextMatrix(K, 11) = Format(nTipo, "#") Then
      sTexto = sTexto & "'" & Trim(fgReg.TextMatrix(K, 2) & fgReg.TextMatrix(K, 3)) & "',"
   End If
Next
If Len(sTexto) > 1 Then
   GetDocRegistro = Mid(sTexto, 1, Len(sTexto) - 1)
Else
   GetDocRegistro = sTexto
End If
End Function

Private Sub cboMes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtAnio.SetFocus
End If
End Sub

Private Sub cmdAgregar_Click()
Dim nItem As Long
glAceptar = False
psTipoAccion = "A"
frmRegVentaDet.inicio True, nTasaIGV, nMoneda, psTipoAccion 'NAGL 20170808 Agregó psTipoAccion
If glAceptar Then
   'DefineFechas 'YIHU20152002-ERS181-2014.
   'Set rs = oReg.CargaRegistro(gnDocTpo, gsDocNro, gdFecha, gdFecSis)
'comentado por  YIHU20150220-ERS181-2014
'fgReg.AdicionaFila
   'nItem = fgReg.row
   'AsignaValores nItem, rs
    'fin YIHU
'**************NAGL 20170805
fgReg.Rows = 2
fgReg.EliminaFila 1
Set rs = oReg.CargaRegistro(0, "", dFecha1, dFecha2, gsCodArea & gsCodAge)
If Not (rs.BOF And rs.EOF) Then
     Do While Not rs.EOF
       fgReg.AdicionaFila
       nItem = fgReg.Row
       AsignaValores nItem, rs
       rs.MoveNext
    Loop
End If
   RSClose rs
End If '**********FIN NAGL
End Sub

Private Sub cmdCustodia_Click()
Dim rsA  As ADODB.Recordset
Dim rsV  As ADODB.Recordset
Dim nIGV As Currency
Dim oCon As DConecta
Dim oAge As DActualizaDatosArea
Dim oConL As DConecta
On Error GoTo ErrCustodia

Set oCon = New DConecta
Set oAge = New DActualizaDatosArea
MousePointer = 11
sDocs = GetDocRegistro(1)
DefineFechas


Set oConL = New DConecta
oConL.AbreConexion

'Set rs = oAge.GetAgencias(, False)
'Do While Not rs.EOF
   If oCon.AbreConexion() Then
'      sSql = "SELECT tp.cCodCta, tp.dFecha, tp.nMontoTran, tp.nNumDoc, pc.cCodPers " _
'         & "FROM TransPrenda tp JOIN (SELECT cCodCta, MIN(cCodPers) as cCodPers FROM perscuenta WHERE crelacta = 'TI' GROUP BY cCodCta)  pc ON pc.cCodCta = tp.cCodCta " _
'         & " WHERE ISNULL(tp.cFlag,'') <> 'X' and cCodTran in ('036000','036001','036100') and dFecha Between '" & Format(dFecha1, gsFormatoFecha) & "' and '" & Format(dFecha2 + 1, gsFormatoFecha) & "' " & IIf(sDocs <> "", "and not RTRIM(nNumDoc) IN (" & sDocs & ")", "") _
'         & "   "

'       sSql = " select mc.cctacod,m.copecod,convert(datetime,substring(cmovnro,1,8),103) fecha," & _
'              " mc.nmonto,p.cPerscod from mov m inner join movcol mc on m.nmovnro=mc.nmovnro " & _
'              " JOIN (     SELECT cCtacod, MIN(cPerscod) as cPerscod " & _
'              "                 From productopersona " & _
'              "                 Where nprdpersrelac = 20 " & _
'              "                 GROUP BY cCtacod" & _
'              "              )  p ON p.cCtacod = mc.cCtacod" & _
'              " Where IsNull(mc.nFlag, 0) <> 1" & _
'              " and m.copecod in ('121900','121600')" & _
'             " and substring(cmovnro,1,8) Between '" & Format(dFecha1, "yyyymmdd") & "' and '" & Format(dFecha2 + 1, "yyyymmdd") & "' " & IIf(sDocs <> "", "and not RTRIM(p.cdocnro) IN (" & sDocs & ")", "")
' Modificado por ENCU 07/12/2004
 
 sSql = " select mc.cctacod ,convert(datetime,substring(cmovnro,1,8),103) dfecha,mc.nmonto,md.ndoctpo,  m.copecod ," & _
        " Case WHEN md.cDocNro ='' then right(mc.cCtaCod,10)" & _
        " else  md.cdocnro end cDocNRo, " & _
        " IsNull((Select cPersCod from ColocPigRecupVta V where V.cCtaCod = mc.cctacod), p.cPerscod ) cPerscod, m.cmovdesc" & _
              " from mov m inner join movcol mc on m.nmovnro=mc.nmovnro " & _
              " left join movdoc md on m.nmovnro=md.nmovnro " & _
              " JOIN (     SELECT cCtacod, MIN(cPerscod) as cPerscod " & _
              "                 From productopersona " & _
              "                 Where nprdpersrelac = 20 " & _
              "                 GROUP BY cCtacod" & _
              "              )  p ON p.cCtacod = mc.cCtacod" & _
              " Where IsNull(mc.nFlag, 0) <> 1" & _
              " and ((m.copecod in ('122000') And nDocTpo In (1,23)) Or (m.copecod in ('121900') And nDocTpo In (23))   )   " & _
              " and convert(datetime,substring(cmovnro,1,8),103) Between '" & Format(dFecha1, "yyyymmdd") & "' and '" & Format(dFecha2 + 1, "yyyymmdd") & "' " & IIf(sDocs <> "", "and not RTRIM(p.cdocnro) IN (" & sDocs & ")", "") & _
              " ORDER BY mc.cctacod ,p.cperscod"
              'Modificacion ENCU insert la linea order by
              '" and m.copecod in ('121900','121600')" & _

      Set rsA = oCon.CargaRecordSet(sSql)
      Do While Not rsA.EOF
         nIGV = rsA!nMonto - Round(rsA!nMonto / (1 + nTasaIGV), 2)
         'sSql = "SELECT nDocTpo FROM RegVenta WHERE cOpeTpo = 1 and nDocTpo = 3 and cDocNro = '" & Trim(rsA!cdocnro) & "' and dDocFecha = '" & Format(rsA!Fecha, gsFormatoFechaHora) & "' "
         'Modificado por ENCU 07/12/2004
         sSql = "SELECT nDocTpo,nVVenta,nPVenta FROM RegVenta WHERE cOpeTpo = 1 and nDocTpo in (3,23) and cDocNro = '" & Trim(rsA!cDocNro) & "' and dDocFecha = '" & Format(rsA!dFecha, gsFormatoFechaHora) & "' "
         Set rsV = oConL.CargaRecordSet(sSql)
        
         'Modifico ENCU
         'If rsV.EOF Then
         'Genero estas lineas para Boleta con tipo de operacion custodia
         If rsV.EOF Then
           If rsA!cOpeCod = 121900 Then
           'If rsA!nDocTpo = 3 Then
            'sSql = "INSERT RegVenta (cOpeTpo,nDocTpo,cDocNro, dDocFecha, cPersCod, cCtaCod, cDescrip, nVVenta, nIGV, nPVenta) " _
            '     & "       VALUES   ('1','03','','" & Format(rsA!Fecha, gsFormatoFechaHora) & "', '" & gsCodCMAC & rsA!cPersCod & "','" & rsA!cCtaCod & "','PAGO DE CUSTODIA'," & rsA!nMonto - nIGV & "," & nIGV & "," & rsA!nMonto & ")"
            'Modificado por ENCU 07/12/2004
             sSql = "INSERT RegVenta (cOpeTpo,nDocTpo,cDocNro, dDocFecha, cPersCod, cCtaCod, cDescrip, nVVenta, nIGV, nPVenta) " _
                 & "       VALUES   ('1','" & rsA!nDocTpo & "','" & Trim(rsA!cDocNro) & "','" & Format(rsA!dFecha, gsFormatoFechaHora) & "', '" & rsA!cPersCod & "','" & rsA!cCtaCod & "','" & Replace(Replace(rsA!cMovDesc, oImpresora.gPrnSaltoLinea, ""), Chr(13), " ") & "'," & rsA!nMonto - nIGV & "," & nIGV & "," & rsA!nMonto & ")"
            oConL.Ejecutar sSql
           Else
          'Genero estas lineas para documento Boleta (1% de la venta) y Operacion Remate
              If rsA!cOpeCod = 122000 And rsA!nDocTpo = 3 Then
               sSql = "INSERT RegVenta (cOpeTpo,nDocTpo,cDocNro, dDocFecha, cPersCod, cCtaCod, cDescrip, nVVenta, nIGV, nPVenta) " _
                 & "       VALUES   ('1','" & rsA!nDocTpo & "','" & Trim(rsA!cDocNro) & "','" & Format(rsA!dFecha, gsFormatoFechaHora) & "', '" & rsA!cPersCod & "','" & rsA!cCtaCod & "','" & Replace(Replace(rsA!cMovDesc, oImpresora.gPrnSaltoLinea, ""), Chr(13), " ") & "'," & rsA!nMonto * 0.01 - nIGV * 0.01 & "," & nIGV * 0.01 & "," & rsA!nMonto * 0.01 & ")"
               oConL.Ejecutar sSql
           'si es Venta de remate y el documento una poliza (99% de la venta)
              Else
               If rsA!cOpeCod = 122000 And rsA!nDocTpo = 23 Then
                sSql = "INSERT RegVenta (cOpeTpo,nDocTpo,cDocNro, dDocFecha, cPersCod, cCtaCod, cDescrip, nVVenta, nIGV, nPVenta) " _
                 & "       VALUES   ('1','" & rsA!nDocTpo & "','" & Trim(rsA!cDocNro) & "','" & Format(rsA!dFecha, gsFormatoFechaHora) & "', '" & rsA!cPersCod & "','" & rsA!cCtaCod & "','" & Replace(Replace(rsA!cMovDesc, oImpresora.gPrnSaltoLinea, ""), Chr(13), " ") & "'," & rsA!nMonto - nIGV & "," & nIGV & "," & rsA!nMonto & ")"
                oConL.Ejecutar sSql
               End If
              End If
           End If
           ' Hasta esta linea se modifico el codigo para el registro de Venta
         Else 'Actualiza el monto de la poliza solo para las agencias por gitu
            If Mid(rsA!cCtaCod, 4, 2) <> "01" Then
                sSql = "Update RegVenta Set nVVenta = " & rsV!nVVenta + rsA!nMonto & ", nPVenta = " & rsV!nPVenta + rsA!nMonto & " WHERE cOpeTpo = 1 and nDocTpo in (3,23) and cDocNro = '" & Trim(rsA!cDocNro) & "' and dDocFecha = '" & Format(rsA!dFecha, gsFormatoFechaHora) & "' "
                oConL.Ejecutar sSql
            End If
         End If
                    
         RSClose rsV
         rsA.MoveNext
      Loop
   End If
 '  oCon.CierraConexion
 '  rs.MoveNext
'Loop
'RSClose rs
cmdVer_Click
MousePointer = 0
oConL.CierraConexion
Set oConL = Nothing
Set oCon = Nothing
Set oAge = Nothing
Exit Sub
ErrCustodia:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   MousePointer = 0
End Sub


Private Sub cmdEliminar_Click()
Dim nItem As Integer
Dim ldFecIni As Date, ldFecFin As Date
Dim lcValida As String
Dim lsMovNro As String
Dim oCont As New NContFunciones
'Set oContFunc = New NContFunciones

If fgReg.TextMatrix(1, 0) = "" Then
    Exit Sub
End If

lcValida = "1" ' 0= realiza validacion 1= no valida
'---------------

If Not oCont.PermiteBorrarRegPorOpe(gsOpeCod, ldFecIni, ldFecFin) Then
    lcValida = "0"
Else
    If Not (gdFecSis >= Format(ldFecIni, "dd/mm/yyyy") And gdFecSis <= Format(ldFecFin, "dd/mm/yyyy")) Then
        lcValida = "0"
    End If
End If

nItem = fgReg.Row
If lcValida = "0" Then
    If Not oCont.PermiteModificarAsiento(Format(CDate(fgReg.TextMatrix(nItem, 1)), gsFormatoMovFecha), False) Then
       MsgBox "No se puede Eliminar un registro que pertenece a un mes cerrado", vbInformation, "¡Aviso!"
        Set oCont = Nothing
       Exit Sub
    End If
End If
Set oCont = Nothing

'PASI20161108 ERS0532016****************
If oReg.EsDocAutorizado(Trim(fgReg.TextMatrix(fgReg.Row, 3) & fgReg.TextMatrix(fgReg.Row, 4))) Then
    MsgBox "El documento no puede ser eliminado por esta opción. Consulte con el Dpto. de TI.", vbInformation, "Aviso"
    Exit Sub
End If
'PASI END*******************************

If MsgBox(" ¿ Seguro de Eliminar Documento ? ", vbQuestion + vbYesNo + vbDefaultButton2, "!Confirmación!") = vbNo Then
   Exit Sub
End If
'nItem = fgReg.Row
'oReg.EliminaVenta fgReg.TextMatrix(nItem, 13), fgReg.TextMatrix(nItem, 2) & fgReg.TextMatrix(nItem, 3), CDate(fgReg.TextMatrix(nItem, 12))

lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

oReg.EliminaVenta fgReg.TextMatrix(nItem, 2), fgReg.TextMatrix(nItem, 3) & fgReg.TextMatrix(nItem, 4), CDate(fgReg.TextMatrix(nItem, 1)), fgReg.TextMatrix(nItem, 14), lsMovNro
fgReg.EliminaFila nItem
End Sub

Private Sub CabeceraExcel()
xlHoja1.PageSetup.Zoom = 75
'ALPA 20090925***********************************************************
'xlHoja1.Cells(1, 2) = gsNomCmac
xlHoja1.Cells(3, 8) = "REGISTRO DE VENTAS E INGRESO"
xlHoja1.Cells(5, 2) = "EJERCICIO : "
xlHoja1.Cells(5, 6) = ":"
xlHoja1.Cells(5, 7) = cboMes & " " & txtAnio
xlHoja1.Cells(6, 2) = "RUC"
xlHoja1.Cells(6, 6) = ":"
xlHoja1.Cells(6, 7) = "20103845328"
xlHoja1.Cells(7, 2) = "APELLIDOS Y NOMBRES, DENOMINACION SOCIAL"
xlHoja1.Cells(7, 6) = ":"
xlHoja1.Cells(7, 7) = "CMAC MAYNAS SA"
xlHoja1.Range("A4:T7").HorizontalAlignment = xlHAlignLeft
xlHoja1.Range("A3:N3").HorizontalAlignment = xlHAlignCenter
'xlHoja1.Range("A4:T7").VerticalAlignment = xlHAlignLeft
'xlHoja1.Cells(6, 2) = "COMPROBANTE"
'xlHoja1.Cells(7, 2) = "FECHA"
'xlHoja1.Cells(7, 3) = "TIPO"
'xlHoja1.Cells(7, 4) = "SERIE"
'xlHoja1.Cells(7, 5) = "NUMERO"
'xlHoja1.Cells(6, 6) = "RUC"
'xlHoja1.Cells(6, 7) = "CLIENTE"
'xlHoja1.Cells(6, 8) = "NRO"
'xlHoja1.Cells(7, 8) = "CONTRATO"
'xlHoja1.Cells(6, 9) = "DESCRIPCION"
'xlHoja1.Cells(6, 10) = "VENTAS"
'xlHoja1.Cells(7, 10) = "GRABADAS"
'xlHoja1.Cells(6, 11) = "VENTAS"
'xlHoja1.Cells(7, 11) = "EXONER."
'xlHoja1.Cells(6, 12) = "I.G.V."
'xlHoja1.Cells(6, 13) = "OTROS"
'xlHoja1.Cells(7, 13) = "TRIBUTOS"
'xlHoja1.Cells(6, 14) = "PRECIO"
'xlHoja1.Cells(7, 14) = "VENTA"

'xlHoja1.Cells(1, 2) = gsNomCmac
'xlHoja1.Cells(3, 2) = "REGISTRO DE VENTAS"
'xlHoja1.Cells(4, 2) = "MES : " & cboMes & " " & txtAnio
'xlHoja1.Cells(6, 2) = "COMPROBANTE"
xlHoja1.Cells(11, 2) = "NUMERO CORRELATIVO DEL REGISTRO O CODIGO UNICO DE LA OPERACION"
xlHoja1.Cells(11, 3) = "FECHA DE EMISION DEL COMPROBANTE DE PAGO O DOCUMENTO"
xlHoja1.Cells(11, 4) = "FECHA DE VENCIMIENTO Y/O PAGO"
'ALPA 20100112 *****************************************************
'xlHoja1.Cells(11, 5) = "TIPO (TABLA 10)"
xlHoja1.Cells(11, 5) = "TIPO (TABLA 02)"
'*******************************************************************
            xlHoja1.Cells(9, 5) = "COMPROBANTE DE PAGO O DOCUMENTO"
            
xlHoja1.Cells(11, 6) = "SERIE"
            
'---NUEVO
xlHoja1.Cells(11, 7) = "NUMERO"
xlHoja1.Cells(11, 8) = "TIPO (TABLA 10)"
        xlHoja1.Cells(10, 8) = "DOCUMENTO DE IDENTIDAD"
        xlHoja1.Cells(9, 8) = "INFORMACION DEL CLIENTE"
xlHoja1.Cells(11, 9) = "NUMERO"
xlHoja1.Cells(11, 10) = "DENOMINACION O RAZON SOCIAL"
        xlHoja1.Cells(10, 10) = "APELLIDOS Y NOMBRES"
xlHoja1.Cells(11, 11) = "BASE IMPONIBLE DE LA OPERACION GRAVADA"
xlHoja1.Cells(11, 12) = "EXONERADA"
        xlHoja1.Cells(9, 12) = "IMPORTE TOTAL DE LA OPERACION EXONERADA O INEFACTA"
xlHoja1.Cells(11, 13) = "INAFECTA"
xlHoja1.Cells(11, 14) = "IGV Y/O IPM"
xlHoja1.Cells(11, 15) = "OTROS TRIBUTOS Y CARGOS QUE NO FORMAN PARTE DE LA BASE IMPONIBLE"
xlHoja1.Cells(11, 16) = "IMPORTE TOTAL DEL COMPROBANTE DE PAGO"
xlHoja1.Cells(11, 17) = "TIPO DE CAMBIO"
xlHoja1.Cells(11, 18) = "FECHA"
        xlHoja1.Cells(9, 18) = "REFERENCIA DEL COMPROBANTE DE PAGO O DOCUMENTO ORIGINAL QUE SE MODIFICA"
xlHoja1.Cells(11, 19) = "TIPO (TABLE 10)"
xlHoja1.Cells(11, 20) = "SERIE"
xlHoja1.Cells(11, 21) = "Nº DEL COMPROBANTE DE PAGO O DOCUMENTO"
xlHoja1.Range("A9:U11").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("A9:U11").VerticalAlignment = xlHAlignJustify
xlHoja1.Range("A9:U11").Font.Size = 8
xlHoja1.Range("B9:B11").Merge
xlHoja1.Range("C9:C11").Merge
xlHoja1.Range("D9:D11").Merge
xlHoja1.Range("E9:G9").Merge
xlHoja1.Range("E10:E11").Merge
xlHoja1.Range("F10:F11").Merge
xlHoja1.Range("G10:G11").Merge
xlHoja1.Range("H9:J9").Merge
xlHoja1.Range("H10:I10").Merge
xlHoja1.Range("K9:K11").Merge
xlHoja1.Range("L9:M9").Merge
xlHoja1.Range("L10:L11").Merge
xlHoja1.Range("M10:M11").Merge
xlHoja1.Range("N9:N11").Merge
xlHoja1.Range("O9:O11").Merge
xlHoja1.Range("P9:P11").Merge
xlHoja1.Range("Q9:Q11").Merge
xlHoja1.Range("R9:U10").Merge

xlHoja1.Range("B3:N3").Merge
xlHoja1.Range("B4:N4").Merge

xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(11, 21)).Borders.LineStyle = 1

xlHoja1.Range("B3:N3").Font.Size = 12
xlHoja1.Range("A1:N7").Font.Bold = True

'xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, 14)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
'xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, 14)).Borders(xlInsideVertical).LineStyle = xlContinuous

xlHoja1.Range("B6:E6").Merge
'xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("A1:A1").ColumnWidth = 1
xlHoja1.Range("B1:B1").ColumnWidth = 10
xlHoja1.Range("C1:C1").ColumnWidth = 5
xlHoja1.Range("D1:D1").ColumnWidth = 7
xlHoja1.Range("E1:E1").ColumnWidth = 12
xlHoja1.Range("F1:F1").ColumnWidth = 15
xlHoja1.Range("G1:G1").ColumnWidth = 40
xlHoja1.Range("H1:H1").ColumnWidth = 18
xlHoja1.Range("I1:I1").ColumnWidth = 40

xlHoja1.Range("C1:H1").EntireColumn.NumberFormat = "@"
xlHoja1.Range("J1:N1").EntireColumn.NumberFormat = "#,##0.00;-#,##0.00"
'xlHoja1.Range(xlHoja1.Cells(8, 3), xlHoja1.Cells(fgReg.Rows + 8, 5)).HorizontalAlignment = xlHAlignCenter
'xlHoja1.Range(xlHoja1.Cells(8, 9), xlHoja1.Cells(fgReg.Rows + 8, 7)).HorizontalAlignment = xlHAlignCenter
'******************************************************************************************
            
            
'--- ANTIGUO
'xlHoja1.Cells(11, 6) = "NUMERO"
'xlHoja1.Cells(11, 7) = "TIPO (TABLA 10)"
'        xlHoja1.Cells(10, 7) = "DOCUMENTO DE IDENTIDAD"
'        xlHoja1.Cells(9, 7) = "INFORMACION DEL CLIENTE"
'xlHoja1.Cells(11, 8) = "NUMERO"
'xlHoja1.Cells(11, 9) = "DENOMINACION O RAZON SOCIAL"
'        xlHoja1.Cells(10, 9) = "APELLIDOS Y NOMBRES"
'xlHoja1.Cells(11, 10) = "BASE IMPONIBLE DE LA OPERACION GRAVADA"
'xlHoja1.Cells(11, 11) = "EXONERADA"
'        xlHoja1.Cells(9, 11) = "IMPORTE TOTAL DE LA OPERACION EXONERADA O INEFACTA"
'xlHoja1.Cells(11, 12) = "INAFECTA"
'xlHoja1.Cells(11, 13) = "IGV Y/O IPM"
'xlHoja1.Cells(11, 14) = "OTROS TRIBUTOS Y CARGOS QUE NO FORMAN PARTE DE LA BASE IMPONIBLE"
'xlHoja1.Cells(11, 15) = "IMPORTE TOTAL DEL COMPROBANTE DE PAGO"
'xlHoja1.Cells(11, 16) = "TIPO DE CAMBIO"
'xlHoja1.Cells(11, 17) = "FECHA"
'        xlHoja1.Cells(9, 17) = "REFERENCIA DEL COMPROBANTE DE PAGO O DOCUMENTO ORIGINAL QUE SE MODIFICA"
'xlHoja1.Cells(11, 18) = "TIPO (TABLE 10)"
'xlHoja1.Cells(11, 19) = "SERIE"
'xlHoja1.Cells(11, 20) = "Nº DEL COMPROBANTE DE PAGO O DOCUMENTO"
'xlHoja1.Range("A9:T11").HorizontalAlignment = xlHAlignCenter
'xlHoja1.Range("A9:T11").VerticalAlignment = xlHAlignJustify
'xlHoja1.Range("A9:T11").Font.Size = 8
'xlHoja1.Range("B9:B11").Merge
'xlHoja1.Range("C9:C11").Merge
'xlHoja1.Range("D9:D11").Merge
'xlHoja1.Range("E9:F9").Merge
'xlHoja1.Range("E10:E11").Merge
'xlHoja1.Range("F10:F11").Merge
'xlHoja1.Range("G9:I9").Merge
'xlHoja1.Range("G10:H10").Merge
'xlHoja1.Range("J9:J11").Merge
'xlHoja1.Range("K9:L9").Merge
'xlHoja1.Range("K10:K11").Merge
'xlHoja1.Range("L10:L11").Merge
'xlHoja1.Range("M9:M11").Merge
'xlHoja1.Range("N9:N11").Merge
'xlHoja1.Range("O9:O11").Merge
'xlHoja1.Range("P9:P11").Merge
'xlHoja1.Range("Q9:T10").Merge
'
'xlHoja1.Range("B3:N3").Merge
'xlHoja1.Range("B4:N4").Merge
'
'xlHoja1.Range(xlHoja1.Cells(9, 2), xlHoja1.Cells(11, 20)).Borders.LineStyle = 1
'
'xlHoja1.Range("B3:N3").Font.Size = 12
'xlHoja1.Range("A1:N7").Font.Bold = True
'
''xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, 14)).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
''xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(7, 14)).Borders(xlInsideVertical).LineStyle = xlContinuous
'
'xlHoja1.Range("B6:E6").Merge
''xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 5)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'xlHoja1.Range("A1:A1").ColumnWidth = 1
'xlHoja1.Range("B1:B1").ColumnWidth = 10
'xlHoja1.Range("C1:C1").ColumnWidth = 5
'xlHoja1.Range("D1:D1").ColumnWidth = 7
'xlHoja1.Range("E1:E1").ColumnWidth = 12
'xlHoja1.Range("F1:F1").ColumnWidth = 15
'xlHoja1.Range("G1:G1").ColumnWidth = 40
'xlHoja1.Range("H1:H1").ColumnWidth = 18
'xlHoja1.Range("I1:I1").ColumnWidth = 40
'
'xlHoja1.Range("C1:H1").EntireColumn.NumberFormat = "@"
'xlHoja1.Range("J1:N1").EntireColumn.NumberFormat = "#,##0.00;-#,##0.00"
''xlHoja1.Range(xlHoja1.Cells(8, 3), xlHoja1.Cells(fgReg.Rows + 8, 5)).HorizontalAlignment = xlHAlignCenter
''xlHoja1.Range(xlHoja1.Cells(8, 9), xlHoja1.Cells(fgReg.Rows + 8, 7)).HorizontalAlignment = xlHAlignCenter
''******************************************************************************************
End Sub

Private Sub cmdImprimir_Click()
Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim N           As Long
Dim oBarra      As New clsProgressBar

If fgReg.TextMatrix(1, 0) = "" Then
   MsgBox "No exiten datos para Imprimir", vbInformation, "Aviso"
   Exit Sub
End If

On Error GoTo ErrImprime
   MousePointer = 11
   oBarra.ShowForm Me
   oBarra.Max = fgReg.Rows - 1
   oBarra.CaptionSyle = eCap_CaptionPercent
   oBarra.Progress 0, "Registro de Ventas", "Creando Archivo Excel", "", vbBlue
   lsArchivo = App.path & "\Spooler\RVENTA" & "_" & txtAnio & Format(Time, "hhmmss") & gsCodUser & ".xls"
   lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
   If lbLibroOpen Then
      Set xlHoja1 = xlLibro.Worksheets(1)
      ExcelAddHoja cboMes, xlLibro, xlHoja1
      CabeceraExcel
      For N = 1 To fgReg.Rows - 1
         oBarra.Progress N, "Registro de Ventas", "", "Generando... ", vbBlue
'          xlHoja1.Cells(N + 11, 2) = fgReg.TextMatrix(N, 1)
'          xlHoja1.Cells(N + 11, 3) = fgReg.TextMatrix(N, 2)
'          xlHoja1.Cells(N + 11, 4) = fgReg.TextMatrix(N, 3)
'          xlHoja1.Cells(N + 11, 5) = "'" & fgReg.TextMatrix(N, 4)
'          xlHoja1.Cells(N + 11, 6) = fgReg.TextMatrix(N, 5)
'          xlHoja1.Cells(N + 11, 7) = fgReg.TextMatrix(N, 6)
'          xlHoja1.Cells(N + 11, 8) = fgReg.TextMatrix(N, 7)
'          xlHoja1.Cells(N + 11, 9) = fgReg.TextMatrix(N, 8)
'          xlHoja1.Cells(N + 11, 10) = fgReg.TextMatrix(N, 9)
'          xlHoja1.Cells(N + 11, 11) = fgReg.TextMatrix(N, 10)
'          xlHoja1.Cells(N + 11, 12) = fgReg.TextMatrix(N, 11)
'          xlHoja1.Cells(N + 11, 13) = fgReg.TextMatrix(N, 12)
'          xlHoja1.Cells(N + 11, 14) = fgReg.TextMatrix(N, 13)
          xlHoja1.Cells(N + 11, 2) = N
          xlHoja1.Cells(N + 11, 3) = fgReg.TextMatrix(N, 1)
          xlHoja1.Cells(N + 11, 4) = fgReg.TextMatrix(N, 1)
          xlHoja1.Cells(N + 11, 5) = fgReg.TextMatrix(N, 2)
          '*** PEAC 20110405
          xlHoja1.Cells(N + 11, 6) = "'" & IIf(Len(Trim(fgReg.TextMatrix(N, 3))) > 0, fgReg.TextMatrix(N, 3), "")
          
          'xlHoja1.Cells(N + 11, 6) = "'" & fgReg.TextMatrix(N, 4)
          'xlHoja1.Cells(N + 11, 6) = "'" & IIf(Len(Trim(fgReg.TextMatrix(N, 3))) > 0, fgReg.TextMatrix(N, 3) & "-", "") & fgReg.TextMatrix(N, 4)
          
          xlHoja1.Cells(N + 11, 7) = "'" & fgReg.TextMatrix(N, 4)
          
          '*** FIN PEAC
'          If Trim(fgReg.TextMatrix(N, 6)) = "MAPFRE PERU" Then
'            MsgBox "sdsfsd"
'          End If
          
          xlHoja1.Cells(N + 11, 8) = fgReg.TextMatrix(N, 19)
          xlHoja1.Cells(N + 11, 9) = fgReg.TextMatrix(N, 5)
          
          xlHoja1.Cells(N + 11, 10) = IIf(Trim(fgReg.TextMatrix(N, 8)) = "A N U L A D O", fgReg.TextMatrix(N, 8), fgReg.TextMatrix(N, 6))
          
          '*****************NAGL 20190121********************
          xlHoja1.Cells(N + 11, 11) = fgReg.TextMatrix(N, 9)
          xlHoja1.Cells(N + 11, 14) = fgReg.TextMatrix(N, 11)
          xlHoja1.Cells(N + 11, 16) = fgReg.TextMatrix(N, 13)
          '********************END NAGL**********************
          
          xlHoja1.Cells(N + 11, 12) = fgReg.TextMatrix(N, 10)
          'xlHoja1.Cells(N + 11, 16) = fgReg.TextMatrix(N, 10)'Comentado by NAGL 20190121
          xlHoja1.Cells(N + 11, 18) = fgReg.TextMatrix(N, 18)
          If Len(Trim(fgReg.TextMatrix(N, 17))) > 0 Then
            xlHoja1.Cells(N + 11, 19) = "01"
          End If
          xlHoja1.Cells(N + 11, 20) = "'" & Format(Mid(fgReg.TextMatrix(N, 17), 1, 3), "000")
          xlHoja1.Cells(N + 11, 21) = "'" & Format(Mid(fgReg.TextMatrix(N, 17), 4, 12), "000000000")
          
      Next
      N = N + 11
      xlHoja1.Range("B" & 12 & ":U" & N).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
      xlHoja1.Range("B" & 12 & ":U" & N).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic
      xlHoja1.Range("B" & 12 & ":U" & N).Borders(xlInsideVertical).LineStyle = xlContinuous
      xlHoja1.Cells(N, 12) = "TOTALES"
      xlHoja1.Range("K" & N).Formula = "=SUM(K12:K" & N - 1 & ")"
      xlHoja1.Range("L" & N).Formula = "=SUM(L12:L" & N - 1 & ")"
      xlHoja1.Range("M" & N).Formula = "=SUM(M12:M" & N - 1 & ")"
      xlHoja1.Range("N" & N).Formula = "=SUM(N12:N" & N - 1 & ")"
      xlHoja1.Range("O" & N).Formula = "=SUM(O12:O" & N - 1 & ")"
      xlHoja1.Range("P" & N).Formula = "=SUM(P12:P" & N - 1 & ")"
'      xlHoja1.Range("O" & N).Formula = "=SUM(O12:O" & N - 1 & ")"
      xlHoja1.Range("K" & N & ":P" & N).Font.Bold = True
      
      
      With xlHoja1.PageSetup
          .LeftHeader = ""
          .CenterHeader = ""
          .RightHeader = ""
          .LeftFooter = ""
          .CenterFooter = ""
          .RightFooter = ""
'          .LeftMargin = Application.InchesToPoints(0)
'          .RightMargin = Application.InchesToPoints(0)
'          .TopMargin = Application.InchesToPoints(0)
'          .BottomMargin = Application.InchesToPoints(0)
'          .HeaderMargin = Application.InchesToPoints(0)
'          .FooterMargin = Application.InchesToPoints(0)
          .PrintHeadings = False
          .PrintGridlines = False
          .PrintComments = xlPrintNoComments
          .CenterHorizontally = False
          .CenterVertically = False
          .Orientation = xlLandscape
          .Draft = False
'          .PaperSize = xlPaperLetter
          .FirstPageNumber = xlAutomatic
          .Order = xlDownThenOver
          .BlackAndWhite = False
          .Zoom = 60
      End With
      
      'OleExcel.Class = "ExcelWorkSheet"
      'ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
      'OleExcel.SourceDoc = lsArchivo
      'OleExcel.Verb = 1
      'OleExcel.Action = 1
      'OleExcel.DoVerb -1 'Comentado by NAGL 20170718
      
      ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
      oBarra.CloseForm Me
      CargaArchivo lsArchivo, App.path & "\SPOOLER\"
      Set oBarra = Nothing
      Set xlLibro = Nothing
      Set xlHoja1 = Nothing 'NAGL 20170718
      
   End If
   'oBarra.CloseForm Me
   'Set oBarra = Nothing 'Comentado by NAGL 20170718
   MousePointer = 0
Exit Sub
ErrImprime:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   MousePointer = 0
End Sub

Private Sub cmdJoyas_Click()
Dim rsA  As New ADODB.Recordset
Dim rsV  As ADODB.Recordset
Dim nIGV As Currency
Dim oCon As DConecta
Dim oAge As DActualizaDatosArea
Dim oConL As DConecta
On Error GoTo ErrCustodia

Set oCon = New DConecta
Set oAge = New DActualizaDatosArea
MousePointer = 11
sDocs = GetDocRegistro(2)
DefineFechas

Set oConL = New DConecta
oConL.AbreConexion

'Me.fgReg.Clear
Set rs = oAge.GetAgencias(, False)
Do While Not rs.EOF
   If oCon.AbreConexion Then
      'sSql = "SELECT tp.cCodCta, tp.dFecha, tp.nMontoTran, tp.nNumDoc, tp.cCodPers, cp.mdescLote FROM transprenda tp JOIN CredPrenda cp ON cp.cCodCta = tp.cCodCta WHERE ISNULL(tp.cFlag,'') <> 'X' and cCodTran in ('122800') and dFecha Between '" & Format(dFecha1, gsFormatoFecha) & "' and '" & Format(dFecha2 + 1, gsFormatoFecha) & "' " & IIf(sDocs <> "", "and not RTRIM(nNumDoc) IN (" & sDocs & ")", "")
      sSql = " SELECT mc.cCtacod, convert(datetime,substring(cmovnro,1,8),103)dfecha, mc.nMonto,md.ndoctpo, IsNull(md.cdocnro,right(mc.cCtaCod,10)) nDocnro, IsNull((Select cPersCod from ColocPigRecupVta V where V.cCtaCod = mc.cctacod), p.cPerscod ) cPerscod , m.cmovdesc" & _
             " FROM mov m inner join movcol mc on m.nmovnro=mc.nmovnro " & _
             " left join movdoc md on m.nmovnro=md.nmovnro inner join productopersona p on mc.cctacod=p.cctacod" & _
             " WHERE ISNULL(mc.nFlag,0) <> 1 and m.copecod in ('122800') And nDocTpo In (3,1) " & _
             " and convert(datetime,substring(cmovnro,1,8),103) Between '" & Format(dFecha1, "yyyymmdd") & "' and '" & Format(dFecha2 + 1, "yyyymmdd") & "' " & IIf(sDocs <> "", "and not RTRIM(p.cdocnro) IN (" & sDocs & ")", "") & _
             " ORDER BY p.cperscod"
      'ENCU insert la linea ORDER BY
      Set rsA = oCon.CargaRecordSet(sSql)
      Do While Not rsA.EOF
         'nIGV = rsA!nMontoTran - Round(rsA!nMontoTran / (1 + nTasaIGV), 2)
         nIGV = rsA!nMonto - Round(rsA!nMonto / (1 + nTasaIGV), 2)
         
        'sSql = "SELECT nDocTpo FROM RegVenta WHERE cOpeTpo = 2 and nDocTpo = 3 and cDocNro = '" & Trim(rsA!nNumDoc) & "' and dDocFecha = '" & Format(rsA!dFecha, gsFormatoFechaHora) & "' "
         sSql = "SELECT nDocTpo FROM RegVenta WHERE cOpeTpo = 2 and nDocTpo = 3 and cDocNro = '" & Trim(rsA!ndocnro) & "' and dDocFecha = '" & Format(rsA!dFecha, gsFormatoFechaHora) & "' "
         Set rsV = oConL.CargaRecordSet(sSql)
         If rsV.EOF Then
'            sSql = "INSERT RegVenta (cOpeTpo,nDocTpo,cDocNro, dDocFecha, cPersCod, cCtaCod, cDescrip, nVVenta, nIGV, nPVenta) " _
'                 & "       VALUES   ('2','03','" & Trim(rsA!nNumDoc) & "','" & Format(rsA!dFecha, gsFormatoFechaHora) & "', '" & gsCodCMAC & rsA!cCodpers & "','" & rsA!cCodCta & "','" & Replace(Replace(rsA!mDescLote, oImpresora.gPrnSaltoLinea, ""), Chr(13), " ") & "'," & rsA!nMontoTran - nIGV & "," & nIGV & "," & rsA!nMontoTran & ")"
            sSql = "INSERT RegVenta (cOpeTpo,nDocTpo,cDocNro, dDocFecha, cPersCod, cCtaCod, cDescrip, nVVenta, nIGV, nPVenta) " _
                 & "       VALUES   ('2','03','" & Trim(rsA!ndocnro) & "','" & Format(rsA!dFecha, gsFormatoFechaHora) & "', '" & rsA!cPersCod & "','" & rsA!cCtaCod & "','" & Replace(Replace(rsA!cMovDesc, oImpresora.gPrnSaltoLinea, ""), Chr(13), " ") & "'," & rsA!nMonto - nIGV & "," & nIGV & "," & rsA!nMonto & ")"

            oConL.Ejecutar sSql
         End If
         RSClose rsV
         rsA.MoveNext
      Loop
   End If
   oCon.CierraConexion
   rs.MoveNext
Loop
RSClose rs
oConL.CierraConexion
Set oConL = Nothing
Set oCon = Nothing
Set oAge = Nothing
cmdVer_Click
MousePointer = 0
Exit Sub
ErrCustodia:
  MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
End Sub

Private Sub cmdModificar_Click()

' YIHU20150220-ERS1812014

Dim cMoneda As String
cMoneda = fgReg.TextMatrix(fgReg.Row, 22)
psTipoAccion = "M" 'NAGL 20170805

If nMoneda = 1 And cMoneda = "ME" Then
    MsgBox "La Venta se registró en moneda Extranjera, para modificar ir a la opción MONEDA EXTRANJERA", vbInformation, "Aviso"
    Exit Sub
End If
If nMoneda = 2 And cMoneda = "MN" Then
    MsgBox "La Venta se registró en moneda Nacional, para modificar ir a la opción MONEDA NACIONAL", vbInformation, "Aviso"
    Exit Sub
End If

'END YIHU ***********************

Dim nItem As Long
If fgReg.TextMatrix(1, 0) = "" Then
    Exit Sub
End If

'PASI20161108 ERS0532016****************
If oReg.EsDocAutorizado(Trim(fgReg.TextMatrix(fgReg.Row, 3) & fgReg.TextMatrix(fgReg.Row, 4))) Then
    MsgBox "El documento no puede ser editado por esta opción. Consulte con el Dpto. de TI.", vbInformation, "Aviso"
    Exit Sub
End If
'PASI END*******************************

glAceptar = False
gsDocNro = Trim(fgReg.TextMatrix(fgReg.Row, 3) & fgReg.TextMatrix(fgReg.Row, 4))
gnDocTpo = fgReg.TextMatrix(fgReg.Row, 2)
gdFecha = CDate(fgReg.TextMatrix(fgReg.Row, 15))
frmRegVentaDet.inicio False, nTasaIGV, nMoneda, psTipoAccion  'NAGL 20170807 AGREGÓ psTipoAccion
DefineFechas 'YIHU20150220

If glAceptar Then
   'Set rs = oReg.CargaRegistro(gnDocTpo, gsDocNro, gdFecha, gdFecSis)
   fgReg.Rows = 2: fgReg.EliminaFila 1
   Set rs = oReg.CargaRegistro(0, "", dFecha1, dFecha2, gsCodArea & gsCodAge) 'NAGL 20170805
   If Not rs.EOF Then  'GIPO 02/12/2016
      Do While Not rs.EOF
           fgReg.AdicionaFila
           nItem = fgReg.Row
           AsignaValores nItem, rs
           rs.MoveNext
       Loop 'NAGL 20170805
    End If
End If
RSClose rs
gdFecha = gdFecSis
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdVer_Click()
Dim nItem As Long
DefineFechas
fgReg.Rows = 2
fgReg.EliminaFila 1
Set rs = oReg.CargaRegistro(0, "", dFecha1, dFecha2, gsCodArea & gsCodAge) 'NAGL 20170804
If Not (rs.BOF And rs.EOF) Then
    Do While Not rs.EOF
       fgReg.AdicionaFila
       nItem = fgReg.Row
       AsignaValores nItem, rs
       rs.MoveNext
    Loop
Else
    MsgBox "No existen Registros de Venta en el periodo seleccionado.", vbOKOnly + vbInformation, "Atención"
End If
RSClose rs
End Sub

Private Sub Form_Load()
Dim sCtaIGV As String
CentraForm Me
frmReportes.Enabled = False
Set oReg = New DRegVenta
Dim oOpe As New DOperacion
Set rs = oOpe.CargaOpeCta(gsOpeCod)
If Not rs.EOF Then
   sCtaIGV = rs!cCtaContCod
End If
Set oOpe = Nothing
Dim oImp As New DImpuesto
Set rs = oImp.CargaImpuesto(sCtaIGV)
Set oImp = Nothing
If rs.EOF Then
   MsgBox "No se definio tasa de IGV. Consultar con Sistemas!", vbInformation, "Aviso"
   Exit Sub
End If
nTasaIGV = rs!nImpTasa / 100
txtAnio = Year(gdFecSis)
cboMes.ListIndex = Month(gdFecSis) - 1
If gbBitCentral Then
    'Frame1.Visible = True
    cmdBienes.Visible = False
    cmdServicio.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmReportes.Enabled = True
Set oReg = Nothing
End Sub


Private Sub txtAnio_GotFocus()
fEnfoque txtAnio
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If Not ValidaAnio(txtAnio) Then
      Exit Sub
   End If
   cmdVer.SetFocus
End If
End Sub

Private Sub txtAnio_Validate(Cancel As Boolean)
   If Not ValidaAnio(Val(txtAnio)) Then
      Cancel = True
   End If
End Sub

