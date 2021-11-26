VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmRepProvCaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Pago de Proveedores"
   ClientHeight    =   6675
   ClientLeft      =   210
   ClientTop       =   1500
   ClientWidth     =   11265
   Icon            =   "frmRepProvCaja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMontos 
      Caption         =   "    Montos"
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
      Height          =   645
      Left            =   90
      TabIndex        =   17
      Top             =   5610
      Width           =   3795
      Begin VB.CheckBox chkMontos 
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   15
         Width           =   210
      End
      Begin SICMACT.EditMoney txtMontoIni 
         Height          =   315
         Left            =   435
         TabIndex        =   19
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtMontoFin 
         Height          =   315
         Left            =   2250
         TabIndex        =   20
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Al"
         Height          =   225
         Left            =   1890
         TabIndex        =   22
         Top             =   315
         Width           =   255
      End
      Begin VB.Label Label9 
         Caption         =   "Del"
         Height          =   225
         Left            =   105
         TabIndex        =   21
         Top             =   315
         Width           =   345
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgProv 
      Height          =   4680
      Left            =   120
      TabIndex        =   15
      Top             =   870
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   8255
      _Version        =   393216
      Cols            =   8
      RowHeightMin    =   300
      FocusRect       =   0
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
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
      Left            =   3975
      TabIndex        =   3
      ToolTipText     =   "Procesa según Rango de Fechas y Opciones"
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Frame Frame2 
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
      Height          =   705
      Left            =   3435
      TabIndex        =   10
      Top             =   75
      Width           =   7665
      Begin VB.CheckBox Chkproveedor 
         Caption         =   "PROVEEDOR"
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
         Height          =   225
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   1725
      End
      Begin VB.OptionButton Opt 
         Caption         =   "&Girados"
         Height          =   285
         Index           =   0
         Left            =   5550
         TabIndex        =   13
         Top             =   255
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Pen&dientes"
         Height          =   285
         Index           =   1
         Left            =   6480
         TabIndex        =   2
         Top             =   270
         Width           =   1095
      End
      Begin SICMACT.TxtBuscar txtCodProv 
         Height          =   345
         Left            =   90
         TabIndex        =   16
         Top             =   240
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
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
         TipoBusPers     =   2
      End
      Begin VB.Label lblNomPers 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1710
         TabIndex        =   14
         Top             =   240
         Width           =   3750
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
      Height          =   705
      Left            =   105
      TabIndex        =   7
      Top             =   75
      Width           =   3300
      Begin MSComCtl2.DTPicker txtFechaDel 
         Height          =   315
         Left            =   420
         TabIndex        =   0
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   78905345
         CurrentDate     =   36509
      End
      Begin MSComCtl2.DTPicker txtFechaAl 
         Height          =   315
         Left            =   1935
         TabIndex        =   1
         Top             =   225
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         Format          =   78905345
         CurrentDate     =   36509
      End
      Begin VB.Label Label2 
         Caption         =   "Al"
         Height          =   225
         Left            =   1740
         TabIndex        =   9
         Top             =   270
         Width           =   225
      End
      Begin VB.Label Label1 
         Caption         =   "Del"
         Height          =   225
         Left            =   105
         TabIndex        =   8
         Top             =   270
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   8190
      TabIndex        =   4
      ToolTipText     =   "Imprime Planilla de Provisión"
      Top             =   5768
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   9690
      TabIndex        =   5
      Top             =   5768
      Width           =   1335
   End
   Begin MSComctlLib.ImageList imgRec 
      Left            =   5055
      Top             =   4575
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
            Picture         =   "frmRepProvCaja.frx":030A
            Key             =   "recibo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   285
      Left            =   2055
      TabIndex        =   11
      Top             =   6360
      Visible         =   0   'False
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   6300
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtf 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      _Version        =   393217
      TextRTF         =   $"frmRepProvCaja.frx":0404
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmRepProvCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim sSql  As String
Dim lSalir As Boolean
Dim lMN As Boolean
Private Type Reporte
    lsFechaPago As String
    lsFechaEmision As String
    lsFechaprov As String
    lsTipo As String
    lsNumDoc As String
    lsGlosa As String
    lsTipoDocPago As String
    lsDocPago As String
    lnSubTotal As Currency
    lnTotal As Currency
    lsPersona As String
    'By Capi 28122007
    lsAgencia As String
End Type
Dim lsReporte() As Reporte
Dim lsTexto As String

Private Sub CabeceraRepo(ByRef nLin As Integer, ByRef P As Integer)
Dim sTit As String
If Opt(0).value Then
   sTit = " PLANILLA DE PROVISIONES DE PAGOS GIRADAS "
End If
If Opt(1).value Then
   sTit = " PLANILLA DE PROVISIONES PENDIENTES DE PAGO "
End If
If nLin > 60 Then
   
   Linea lsTexto, oImpresora.gPrnCondensadaON + CabeRepo(gsNomCmac, gsNomAge, "", "MONEDA " & IIf(gsSimbolo = gcME, "EXTRANJERA ", "NACIONAL"), Format(gdFecSis, gsFormatoFechaView), sTit, "( DEL " & txtFechaDel & " AL " & txtFechaAl & ")", "", "", P, 150) & oImpresora.gPrnCondensadaOFF, 0
   Linea lsTexto, oImpresora.gPrnBoldON + ImpreFormat("PROVEEDOR : " & Trim(lblNomPers), 50) & oImpresora.gPrnBoldOFF
   nLin = 10
   If Me.Opt(0).value Then
        Linea lsTexto, oImpresora.gPrnCondensadaON + oImpresora.gPrnBoldON + IIf(Me.Chkproveedor.value = 0, String(200, "-"), String(155, "-"))
        Linea lsTexto, ImpreFormat("FECHA PAGO", 10) & ImpreFormat("PROVISION", 10) & ImpreFormat("EMISION", 10) & ImpreFormat("DOC. REFERENCIA", 15) & ImpreFormat("DOC. PAGO", 15) & ImpreFormat("DESCRIPCION", 40) & ImpreFormat("SUBTOTAL", 12) & ImpreFormat("TOTAL", 12) & IIf(Me.Chkproveedor.value = 0, ImpreFormat("PROVEEDOR", 30), "") & ImpreFormat("AGENCIAS", 8) 'By Capi 28122007 se adiciono agencia
        Linea lsTexto, IIf(Me.Chkproveedor.value = 0, String(200, "-"), String(155, "-")) & oImpresora.gPrnBoldOFF
   Else
        Linea lsTexto, oImpresora.gPrnCondensadaON + oImpresora.gPrnBoldON + IIf(Me.Chkproveedor.value = 0, String(200, "-"), String(155, "-"))
        Linea lsTexto, ImpreFormat("PROVISION", 10) & ImpreFormat("EMISION", 10) & ImpreFormat("TIPO", 6) & ImpreFormat("DOCUMENTO", 15) & ImpreFormat("DESCRIPCION", 40) & ImpreFormat("TOTAL", 12) & IIf(Me.Chkproveedor.value = 0, ImpreFormat("PROVEEDOR", 30), "") & ImpreFormat("AGENCIA", 4) 'By Capi 28122007 se adiciono agencia
        Linea lsTexto, IIf(Me.Chkproveedor.value = 0, String(200, "-"), String(155, "-")) & oImpresora.gPrnBoldOFF
   End If
End If
End Sub

Private Sub GeneraListado(pnOpt As Integer)
Dim sCond As String
Dim nItem As Integer, N As Integer
Dim lvItem As ListItem
Dim nImporte As Currency, nTipCambio As Currency
Dim lsMovNro As String
Dim lsCondPers As String
Dim lsCtaCont  As String
Dim sAgeBen As String
Dim rs1 As ADODB.Recordset
Dim rs0 As ADODB.Recordset
Dim sSQL1 As String




Dim oOpe As New DOperacion



Set rs = oOpe.CargaOpeCta(gsOpeCod)
Do While Not rs.EOF
    lsCtaCont = lsCtaCont & "'" & Left(rs!cCtaContCod, 2) & IIf(gsSimbolo = gcME, "2", "1") & Mid(rs!cCtaContCod, 4, 22) & "',"
    rs.MoveNext
Loop
If lsCtaCont <> "" Then
    lsCtaCont = Left(lsCtaCont, Len(lsCtaCont) - 1)
End If
Set oOpe = Nothing
If lsCtaCont = "" Then
    MsgBox "Falta Definir Cuentas Contables en Definicion de Reporte", vbInformation, "¡Aviso!"
    Exit Sub
End If
If Me.Opt(0).value Then
    sCond = " AND c.nMovImporte > 0 and EXISTS (SELECT i.nMovNro FROM  MovRef i WHERE i.nMovNro = a.nMovNro)   "
Else
    sCond = " AND c.nMovImporte < 0 and NOT EXISTS (SELECT i.nMovNro FROM  MovRef i WHERE i.nMovNroRef = a.nMovNro)   "
End If

If Chkproveedor.value = 1 Then
    lsCondPers = " AND d.cPersCod='" & Trim(txtCodProv.Tag) & "'   "
Else
    lsCondPers = ""
End If

'By Capi 28122007 se adiciono campo para determinar la agencia



sSQL1 = "SELECT  a.nMovNro,b.dDocFecha, g.cDocAbrev, g.cDocDesc, b.nDocTpo, b.cDocNro, e.cPersNombre, " _
    & " a.cMovDesc, d.cPersCod, a.cMovNro, ABS(c.nMovImporte) as nDocImporte, ISNULL(j.nMovMEImporte,0) as nDocImporteME, " _
    & " DOC.MovDocRef, DOC.GlosaDocRef, DOC.NroDocRef, DOC.FechaDocRef, DOC.TipoDocRef , DOC.ImporteDocRefMN, DOC.ImporteDocRefME, " _
    & " ISNULL(DOCPAGO.nDocTpo,'') , ISNULL(DOCPAGO.cDocAbrev,'') AS TipDocPago ,ISNULL(DOCPAGO.cDocNro,'') as NumDocPago,Doc.Agencia  " _
    & " FROM    Mov a " _
    & "         JOIN    MovCta  c   ON c.nMovnro=a.nMovNro LEFT JOIN MOVME j   ON j.nMovnro=c.nMovNro and c.nMovItem = j.nMovItem " _
    & "    LEFT JOIN    MovDoc  b   ON b.nMovNro=a.nMovnro  " _
    & "         JOIN    MovGasto d   ON d.nMovNro=a.nMovNro " _
    & "         JOIN    Persona e   ON e.cPersCod = d.cPersCod " _
    & "         JOIN    Documento g     ON g.nDocTpo = b.nDocTpo " _
    & "         LEFT JOIN    (SELECT  MR.cAgeCodBen Agencia,MR.nMOVNRO AS MovPago, M.CMOVDESC AS GlosaDocRef, M.CMOVNRO as MovDocRef, MD.cDocNro as NroDocRef, " _
    & "                 MD.dDocFecha as FechaDocRef, D.cDocAbrev as TipoDocRef,  ABS(MC.nMovImporte)*-1 AS ImporteDocRefMN, ISNULL(Me.NMOVMEIMPORTE, 0)*-1 As ImporteDocRefME " _
    & "                 FROM    MOV M JOIN MOVREF MR ON M.nMOVNRO=MR.nMOVNROREF JOIN MOVCTA MC ON MC.nMOVNRO=M.nMOVNRO " _
    & "                         LEFT JOIN MOVME ME ON ME.nMOVNRO=MC.nMOVNRO AND ME.nMovItem=MC.nMovItem JOIN MOVDOC MD ON M.nMOVNRO=MD.nMOVNRO " _
    & "                         JOIN Documento D ON D.nDocTpo = MD.nDocTpo " _
    & "                 WHERE   MC.NMOVIMPORTE<0 and mc.cCtaContCod in (" & lsCtaCont & ") AND M.nMOVFLAG NOT IN('1','2','3')) DOC ON  DOC.MovPago = a.nMovNro " _
    & "         LEFT JOIN ( SELECT  MD1.nMovNro, MD1.cDocNro, MD1.nDocTpo, D1.cDocAbrev  " _
    & "                     FROM    MOVDOC MD1 JOIN Documento D1 ON D1.nDocTpo = MD1.nDocTpo WHERE md1.nDocTpo = " & TpoDocVoucherEgreso & " ) AS DOCPAGO " _
    & "         ON DOCPAGO.nMovNro= A.nMOVNRO " _
    & " WHERE   a.nMovEstado = '10' AND a.nMovFlag NOT IN('1','2','3') and c.cCtaContCod in (" & lsCtaCont & ") " _
    & "         AND SubString(a.cMovNro,1,8) BETWEEN '" & Format(txtFechaDel, "yyyymmdd") & "' and '" & Format(txtFechaAl, "yyyymmdd") & "' " & IIf(Me.chkMontos.value = 1, " And ABS(c.nMovImporte) Between " & Me.txtMontoIni.value & " And " & Me.txtMontoFin.value & " ", "") & " " _
    & sCond & lsCondPers & " ORDER BY a.nMovNro,a.cMovNro "

'and b.nDocTpo IN (" & TpoDocOrdenPago & "," & TpoDocCarta & "," & TpoDocCheque & "," & TpoDocNotaAbono & ")
Dim oCon0 As New DConecta
oCon0.AbreConexion
'oCon.CommandTimeout 7800
Set rs0 = oCon0.CargaRecordSet(sSQL1)
'By Capi 28122007 para actualizar la agencia(s) beneficiadas

Do While Not rs0.EOF
    sSql = "Select Distinct Right(cCTaContCod,2) Agencia From MovRef A Inner Join MovCta B On A.nMovNroRef=B.nMovNro Where Left(cCtaContCod,2) in (44,45) And a.nmovnro= " & rs0!nmovnro & ""
    Dim oCon1 As New DConecta
    oCon1.AbreConexion
    Set rs1 = oCon1.CargaRecordSet(sSql)
    sAgeBen = "_"
    Do While Not rs1.EOF
        sAgeBen = sAgeBen & rs1!Agencia & "_"
        rs1.MoveNext
    Loop
    sSql = "Update MovRef Set cAgeCodBen='" & sAgeBen & "' " & "Where nMovNro= " & rs0!nmovnro & ""
    Dim oCon2 As New DConecta
    oCon2.AbreConexion
    oCon2.Ejecutar (sSql)
  
    rs0.MoveNext
Loop



Dim oCon As New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sSQL1)

'End By

prg.Visible = True
prg.Min = 0
ReDim lsReporte(0)
If Not RSVacio(rs) Then
    prg.Max = rs.RecordCount
    lsMovNro = Trim(rs!cMovNro)
    nItem = 0
    Do While Not rs.EOF
        prg.value = rs.Bookmark
        Status.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
        If nItem = 0 Then 'Cabecera Pago de Documentos
            N = N + 1
            ReDim Preserve lsReporte(N)
            lsReporte(N).lsFechaPago = Format(rs!dDocFecha, "dd/mm/yyyy")
            lsReporte(N).lsFechaEmision = IIf(Me.Opt(0).value, "", Format(rs!dDocFecha, "dd/mm/yyyy"))
            lsReporte(N).lsFechaprov = IIf(Me.Opt(0).value, "", Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4))
            lsReporte(N).lsTipo = IIf(Trim(rs!cDocAbrev) + Trim(rs!cDocNro) = Trim(rs!TipDocPago) + Trim(rs!NumDocPago), "", Trim(rs!cDocAbrev))
            lsReporte(N).lsNumDoc = IIf(Trim(rs!cDocAbrev) + Trim(rs!cDocNro) = Trim(rs!TipDocPago) + Trim(rs!NumDocPago), "", Trim(rs!cDocNro))
            lsReporte(N).lsGlosa = Trim(rs!cMovDesc)
            lsReporte(N).lsTipoDocPago = Trim(rs!TipDocPago)
            lsReporte(N).lsDocPago = Trim(rs!NumDocPago)
            lsReporte(N).lnSubTotal = 0
            lsReporte(N).lnTotal = Format(IIf(gsSimbolo = gcME, rs!nDocImporteME, rs!nDocImporte), "#,#0.00")
            lsReporte(N).lsPersona = IIf(Me.Chkproveedor.value = 0, PstaNombre(rs!cPersNombre, False), "")
            'By capi 28122007
            lsReporte(N).lsAgencia = IIf(IsNull(rs!Agencia), "", rs!Agencia)
            
        End If
        nItem = nItem + 1
        If Not IsNull(rs!FechaDocRef) Then
            N = N + 1
            ReDim Preserve lsReporte(N)
            lsReporte(N).lsFechaPago = ""  'Str(nItem)
            lsReporte(N).lsFechaEmision = IIf(IsNull(rs!FechaDocRef), "", Format(rs!FechaDocRef, "dd/mm/yyyy"))
            lsReporte(N).lsFechaprov = IIf(IsNull(rs!MovDocRef), "", Mid(rs!MovDocRef, 7, 2) & "/" & Mid(rs!MovDocRef, 5, 2) & "/" & Mid(rs!MovDocRef, 1, 4))
            lsReporte(N).lsTipo = IIf(IsNull(rs!TipoDocRef), "", Trim(rs!TipoDocRef))
            lsReporte(N).lsNumDoc = IIf(IsNull(rs!NroDocRef), "", Trim(rs!NroDocRef))
            lsReporte(N).lsGlosa = IIf(IsNull(rs!GlosaDocRef), "", Trim(rs!GlosaDocRef))
            lsReporte(N).lnSubTotal = IIf(IsNull(rs!ImporteDocRefME), 0, Format(IIf(gsSimbolo = gcME, rs!ImporteDocRefME, rs!ImporteDocRefMN), "#,#0.00"))
            lsReporte(N).lnTotal = 0
            lsReporte(N).lsPersona = ""
            lsReporte(N).lsAgencia = rs!Agencia
        End If
        
        rs.MoveNext
        If Not rs.EOF Then
            If lsMovNro <> Trim(rs!cMovNro) Then
                lsMovNro = Trim(rs!cMovNro)
                nItem = 0
            End If
        End If
    Loop
End If
RSClose rs
oCon.CierraConexion
prg.Visible = False
End Sub
Private Sub MuestraDatos()
Dim I As Integer
Dim N As Integer
Dim lnTotal As Currency
CabeceraGrid
If UBound(lsReporte) > 0 Then
    lnTotal = 0
    fgProv.BackColorSel = vbBlue
    'fgProv.Rows = UBound(lsReporte) + 1
    For I = 1 To UBound(lsReporte)
        AdicionaRow fgProv
        N = fgProv.Row
        fgProv.TextMatrix(N, 1) = lsReporte(I).lsFechaPago
        fgProv.TextMatrix(N, 2) = lsReporte(I).lsFechaprov
        fgProv.TextMatrix(N, 3) = lsReporte(I).lsFechaEmision
        fgProv.TextMatrix(N, 4) = Trim(lsReporte(I).lsTipo)
        fgProv.TextMatrix(N, 5) = Trim(lsReporte(I).lsNumDoc)
        fgProv.TextMatrix(N, 6) = Trim(lsReporte(I).lsTipoDocPago)
        fgProv.TextMatrix(N, 7) = Trim(lsReporte(I).lsDocPago)
        fgProv.TextMatrix(N, 8) = Trim(lsReporte(I).lsGlosa)
        fgProv.TextMatrix(N, 9) = IIf(lsReporte(I).lnSubTotal = 0, "", Format(lsReporte(I).lnSubTotal, "#,#0.00"))
        fgProv.TextMatrix(N, 10) = IIf(lsReporte(I).lnTotal = 0, "", Format(lsReporte(I).lnTotal, "#,#0.00"))
        lnTotal = lnTotal + lsReporte(I).lnTotal
        fgProv.TextMatrix(N, 11) = Trim(lsReporte(I).lsPersona)
        'By Capi 28122007
        fgProv.TextMatrix(N, 12) = Trim(lsReporte(I).lsAgencia)
        If lsReporte(I).lnTotal > 0 And Opt(0).value Then
            BackColorFg fgProv, "&H00E0E0E0", True
        End If
    Next I
    If Chkproveedor.value = 0 Then
        AdicionaRow fgProv
        N = fgProv.Row
        fgProv.TextMatrix(N, 8) = IIf(Opt(0).value = True, "TOTAL GIRADO :", "TOTAL PENDIENTE :")
        fgProv.TextMatrix(N, 10) = Format(lnTotal, "#,#0.00")
        BackColorFg fgProv, vbYellow, True
    End If
    fgProv.Row = 1
    fgProv.Col = 1
    fgProv.SetFocus
Else
    MsgBox "No existen datos para Reporte", vbInformation, "Aviso"
End If
End Sub

Private Sub chkProveedor_Click()
If Me.Chkproveedor.value = 1 Then
    txtCodProv.SetFocus
Else
    Me.txtCodProv.Tag = ""
    Me.lblNomPers = ""
End If
End Sub

Private Sub cmdImprimir_Click()
Dim N As Integer
Dim nLin As Integer, P As Integer
Dim nTot As Currency
Dim lnTotalPagado As Currency

Dim lOk As Boolean
If Me.fgProv.TextMatrix(1, 0) = "" Then
   MsgBox "No existen elementos que Imprimir...!", vbInformation, "Error"
   Exit Sub
End If
nLin = 66
rtf = ""
nTot = 0
prg.Min = 0
prg.Max = Me.fgProv.Rows
prg.Visible = True
Me.Enabled = False
lsTexto = ""
lnTotalPagado = 0
For N = 1 To Me.fgProv.Rows - 1
   lOk = True
   prg.value = N
   Status.Panels(1).Text = "Proceso " & Format(prg.value * 100 / prg.Max, gsFormatoNumeroView) & "%"
   If lOk Then
        CabeceraRepo nLin, P
        If Me.Opt(0).value Then
            lnTotalPagado = lnTotalPagado + CCur(IIf(fgProv.TextMatrix(N, 9) = "", "0", fgProv.TextMatrix(N, 9)))
            Linea lsTexto, IIf(fgProv.TextMatrix(N, 10) = "", oImpresora.gPrnBoldOFF, oImpresora.gPrnBoldON) + ImpreFormat(fgProv.TextMatrix(N, 1), 10) & ImpreFormat(fgProv.TextMatrix(N, 2), 10) & ImpreFormat(fgProv.TextMatrix(N, 3), 10) & ImpreFormat(fgProv.TextMatrix(N, 4) & "-" & fgProv.TextMatrix(N, 5), 15) & ImpreFormat(fgProv.TextMatrix(N, 6) & IIf(fgProv.TextMatrix(N, 6) = "", "", "-") & fgProv.TextMatrix(N, 7), 15) & ImpreFormat(fgProv.TextMatrix(N, 8), 35) & ImpreFormat(CCur(IIf(fgProv.TextMatrix(N, 9) = "", "0", fgProv.TextMatrix(N, 9))), 10) & ImpreFormat(CCur(IIf(fgProv.TextMatrix(N, 10) = "", "0", fgProv.TextMatrix(N, 10))), 10) & IIf(Me.Chkproveedor.value = 0, ImpreFormat(fgProv.TextMatrix(N, 11), 35, 5), "") & ImpreFormat(fgProv.TextMatrix(N, 12), 25) & oImpresora.gPrnBoldOFF 'By Capi 28122007
        Else
            Linea lsTexto, ImpreFormat(fgProv.TextMatrix(N, 2), 10) & ImpreFormat(fgProv.TextMatrix(N, 3), 10) & ImpreFormat(fgProv.TextMatrix(N, 4), 6) & ImpreFormat(fgProv.TextMatrix(N, 5), 15) & ImpreFormat(fgProv.TextMatrix(N, 8), 70) & ImpreFormat(CCur(IIf(fgProv.TextMatrix(N, 10) = "", "0", fgProv.TextMatrix(N, 10))), 12) & IIf(Me.Chkproveedor.value = 0, ImpreFormat(fgProv.TextMatrix(N, 11), 40, 5), "")
        End If
   End If
   nLin = nLin + 1
Next
Linea lsTexto, ImpreFormat(String(155, "="), 155)
Linea lsTexto, ImpreFormat("", 97) & ImpreFormat("TOTAL GENERAL :", 15) & ImpreFormat(lnTotalPagado, 15, 2, True)
EnviaPrevio lsTexto, "REPORTE DE PROVEEDORES", gnLinPage, True
prg.Visible = False
Me.Enabled = True

End Sub

Public Function Linea(psVarImpre As String, psTexto As String, Optional pnLineas As Integer = 1, Optional ByRef pnLinCnt As Integer = 0) As String
Dim K As Integer
psVarImpre = psVarImpre & psTexto
For K = 1 To pnLineas
   psVarImpre = psVarImpre & oImpresora.gPrnSaltoLinea
   pnLinCnt = pnLinCnt + 1
Next
End Function

Private Sub cmdProcesar_Click()

If Me.Chkproveedor.value = 1 Then
    If Me.txtCodProv.Tag = "" Then
        MsgBox "Ingrese Proveedor a buscar", vbInformation, "Aviso"
        txtCodProv.SetFocus
        Exit Sub
    End If
End If
Screen.MousePointer = 11
Me.Status.Panels(1).Text = "Espere por favor...."
cmdProcesar.Enabled = False
GeneraListado IIf(Opt(0).value = True, 0, Opt(1).value)
MuestraDatos
Screen.MousePointer = 0
cmdProcesar.Enabled = True
Me.Status.Panels(1).Text = "Reporte Generado"
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
Dim lvItem As ListItem, nItem As Integer
Dim rsPer As New ADODB.Recordset
Dim oConec As DConecta
Set oConec = New DConecta

oConec.AbreConexion
lSalir = False
CentraForm Me
txtFechaDel = gdFecSis
txtFechaAl = gdFecSis
gsSimbolo = gcMN
If frmReportes.optMoneda(1).value Then
   gsSimbolo = gcME
End If
CabeceraGrid
txtCodProv.TipoBusqueda = BuscaPersona
txtCodProv.TipoBusPers = BusPersDocumentoRuc

End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim oConec As DConecta
Set oConec = New DConecta

oConec.CierraConexion
End Sub
Private Sub CabeceraGrid()
Dim I As Integer
fgProv.Cols = 13
fgProv.Rows = 2
fgProv.Clear
For I = 0 To fgProv.Cols - 1
    fgProv.Row = 0
    fgProv.Col = I
    fgProv.CellTextStyle = 0
    fgProv.CellAlignment = 4
    fgProv.CellFontBold = True
    Select Case I
        Case 1
            fgProv.Text = IIf(Me.Opt(0).value, "Fecha Pago", "Emisión")
        Case 2
            fgProv.Text = "Provisión"
        Case 3
            fgProv.Text = "Emisión"
        Case 4
            fgProv.Text = "Doc"
        Case 5
            fgProv.Text = "Referencia"
        Case 6
            fgProv.Text = "Doc"
        Case 7
            fgProv.Text = "Pago"
        Case 8
            fgProv.Text = "GLOSA"
        Case 9
            fgProv.Text = "SUB TOTAL"
        Case 10
            fgProv.Text = "TOTAL"
        Case 11
            fgProv.Text = "PROVEEDOR"
        Case 12
            fgProv.Text = "AGENCIA"
        Case Else
            fgProv.Text = ""
    End Select
Next I
fgProv.ColAlignment(0) = 0
fgProv.ColAlignment(1) = 7
fgProv.ColAlignment(2) = 1
fgProv.ColAlignment(3) = 1
fgProv.ColAlignment(4) = 4
fgProv.ColAlignment(5) = 4
fgProv.ColAlignment(6) = 4
fgProv.ColAlignment(7) = 4
fgProv.ColAlignment(8) = 0
fgProv.ColAlignment(9) = 7
fgProv.ColAlignment(10) = 7
fgProv.ColAlignment(11) = 0

fgProv.ColWidth(0) = 350
fgProv.ColWidth(1) = IIf(Me.Opt(0).value, 1200, 0)
fgProv.ColWidth(2) = 1100
fgProv.ColWidth(3) = 1100
fgProv.ColWidth(4) = 500
fgProv.ColWidth(5) = 1300
fgProv.ColWidth(6) = IIf(Me.Opt(0).value, 500, 0)
fgProv.ColWidth(7) = IIf(Me.Opt(0).value, 1300, 0)
fgProv.ColWidth(8) = IIf(Me.Opt(0).value, 4500, 5000)
fgProv.ColWidth(9) = IIf(Me.Opt(0).value, 1200, 0)
fgProv.ColWidth(10) = 1200
If Me.Chkproveedor.value = 0 Then
    fgProv.ColWidth(11) = 4000
Else
    fgProv.ColWidth(11) = 0
End If
End Sub
Private Sub Opt_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdProcesar.SetFocus
End If

End Sub

Private Sub txtCodProv_EmiteDatos()
lblNomPers = txtCodProv.psDescripcion
txtCodProv.Tag = txtCodProv.psCodigoPersona
If lblNomPers <> "" Then
   txtFechaDel.SetFocus
End If
End Sub

Private Sub txtFechaAl_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtCodProv.SetFocus
End If
End Sub

Private Sub txtFechaDel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtFechaAl.SetFocus
End If
End Sub
