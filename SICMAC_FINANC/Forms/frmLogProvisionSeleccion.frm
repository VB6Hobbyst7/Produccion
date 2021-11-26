VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmLogProvisionSeleccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Logística: "
   ClientHeight    =   5280
   ClientLeft      =   705
   ClientTop       =   2205
   ClientWidth     =   10050
   Icon            =   "frmLogProvisionSeleccion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   405
      Left            =   1755
      TabIndex        =   7
      Top             =   4665
      Width           =   1410
   End
   Begin VB.CheckBox chkMismoPer 
      Caption         =   "Mismo Periodo"
      Height          =   315
      Left            =   150
      TabIndex        =   6
      Top             =   4725
      Width           =   1515
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Provisionar"
      Height          =   345
      Left            =   7350
      TabIndex        =   1
      Top             =   4830
      Width           =   1185
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   8580
      TabIndex        =   2
      Top             =   4830
      Width           =   1185
   End
   Begin VB.TextBox txtMovNro 
      Height          =   315
      Left            =   5715
      TabIndex        =   3
      Top             =   4785
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
      Left            =   8475
      TabIndex        =   4
      Top             =   4410
      Width           =   1290
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fg 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedRows       =   2
      BackColorBkg    =   -2147483643
      AllowBigSelection=   0   'False
      TextStyleFixed  =   3
      FocusRect       =   2
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
      _Band(0).Cols   =   11
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
      Left            =   7290
      TabIndex        =   5
      Top             =   4470
      Width           =   885
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   315
      Left            =   7020
      Top             =   4410
      Width           =   2745
   End
End
Attribute VB_Name = "frmLogProvisionSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql  As String
Dim rs    As ADODB.Recordset
Dim lSalir As Boolean
Dim sCtaProvis As String
Dim sDocTpoOC  As String
Dim sDocDesc   As String
Dim txtImporte As Currency

Dim lbBienes   As Boolean

Dim lbReporteFechas As Boolean

Public Sub Inicio(pbBienes As Boolean)
    lbBienes = pbBienes
    Me.Show 1
End Sub

Private Sub FormatoOCompra()
fg.TextMatrix(0, 0) = " "
fg.TextMatrix(1, 0) = " "
fg.TextMatrix(0, 1) = "Documento"
fg.TextMatrix(0, 2) = "Documento"
fg.TextMatrix(0, 3) = "Documento"
fg.TextMatrix(1, 1) = "Tipo"
fg.TextMatrix(1, 2) = "Número"
fg.TextMatrix(1, 3) = "Fecha"
fg.TextMatrix(0, 4) = "Proveedor"
fg.TextMatrix(1, 4) = "Proveedor"
fg.TextMatrix(0, 5) = "Importe"
fg.TextMatrix(1, 5) = "Importe"
fg.TextMatrix(0, 6) = "Observaciones"
fg.TextMatrix(1, 6) = "Observaciones"
fg.TextMatrix(1, 7) = "cMovNro"
fg.TextMatrix(1, 8) = "nImporte"

fg.TextMatrix(1, 9) = "Saldo"
fg.TextMatrix(1, 10) = "Estado"
fg.RowHeight(-1) = 285
fg.ColWidth(0) = 385
fg.ColWidth(1) = 500
fg.ColWidth(2) = 1200
fg.ColWidth(3) = 1100
fg.ColWidth(4) = 3200
fg.ColWidth(5) = 1200
fg.ColWidth(6) = 3770
fg.ColWidth(7) = 0
fg.ColWidth(8) = 0
fg.ColWidth(9) = 1200
fg.ColWidth(10) = 1700

fg.MergeCells = flexMergeRestrictColumns
fg.MergeCol(0) = True
fg.MergeCol(1) = True
fg.MergeCol(2) = True
fg.MergeCol(3) = True
fg.MergeCol(4) = True
fg.MergeCol(5) = True
fg.MergeCol(6) = True

fg.MergeRow(0) = True
fg.MergeRow(1) = True
fg.RowHeight(-1) = 285
fg.RowHeight(0) = 200
fg.RowHeight(1) = 200
fg.ColAlignmentFixed(-1) = flexAlignCenterCenter
fg.ColAlignment(1) = flexAlignCenterCenter
fg.ColAlignment(3) = flexAlignCenterCenter
fg.ColAlignment(6) = flexAlignLeftCenter
End Sub

Private Sub GetOCPendientes()
Dim lsCadSerAdm As String
Dim nItem As Integer
Dim nTot  As Currency
Dim oCon As DConecta
Set oCon = New DConecta
oCon.AbreConexion

sSql = " SELECT DISTINCT b.dDocFecha, g.cDocAbrev, b.nDocTpo, b.cDocNro," _
     & " cNomPers = (" _
     & "             SELECT cPersNombre + space(100) + PE.cPersCod" _
     & "             FROM Persona PE WHERE PE.cPersCod = dd.cPersCod)," _
     & "     dd.cPersCod, a.cMovDesc, a.cMovNro, a.nMovNro, a.nMovEstado, a.nMovFlag, c.cCtaContCod," _
     & "     " & IIf(Mid(gsOpeCod, 3, 1) = 1, " c.nMovImporte", "me.nMovMEImporte") & " * -1 as nDocImporte,  ISNULL(ref.nMontoA,0) * -1 nMontoA" _
     & " FROM Mov a JOIN MovDoc b ON b.nMovNro = a.nMovNro" _
     & "            JOIN Documento g ON g.nDocTpo = b.nDocTpo" _
     & "            JOIN MovCta c ON c.nMovNro = a.nMovNro " & IIf(Mid(gsOpeCod, 3, 1) = 1, "", "JOIN MovME me ON me.nMovNro = c.nMovNro and me.nMovItem = c.nMovItem") _
     & "            JOIN MovGasto dd ON dd.nMovNro = c.nMovNro" _
     & "       LEFT JOIN (" _
     & "             SELECT h.nMovNroRef," & IIf(Mid(gsOpeCod, 3, 1) = 1, " SUM(nMovImporte)", "SUM(nMovMEImporte)") & " nMontoA " _
     & "             FROM MovRef h JOIN Mov m ON m.nMovNro = h.nMovNro" _
     & "                           JOIN MovCta mc ON mc.nMovNro = h.nMovNro " & IIf(Mid(gsOpeCod, 3, 1) = 1, "", "JOIN MovME me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem") _
     & "             WHERE m.nMovEstado = '" & gMovEstContabMovContable & "' and m.nMovFlag NOT IN ('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "') And mc.nMovImporte < 0 " _
     & "             GROUP BY h.nMovNroRef" _
     & "            ) ref ON ref.nMovNroRef = a.nMovNro" _
     & " WHERE  " & IIf(Me.chkMismoPer.value = 1, " a.cMovNro Like '" & Format(gdFecSis, "yyyy") & "%'  And ", "") & "  a.nMovEstado IN (" & gMovEstPresupAceptado & "," & gMovEstLogIngBienAceptado & ") and a.nMovFlag NOT IN (" & gMovFlagEliminado & "," & gMovFlagExtornado & "," & gMovFlagExtornado & ", " & gMovFlagModificado & " ) " _
     & " And c.cCtaContCod = '" & sCtaProvis & "' And b.nDocTpo = '" & sDocTpoOC & "'" _
     & " AND " & IIf(Mid(gsOpeCod, 3, 1) = 1, " c.nMovImporte", "me.nMovMEImporte") & " - ISNULL(ref.nMontoA,0) <> 0 " _
     & " ORDER BY b.cDocNro"

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
   fg.TextMatrix(nItem, 3) = rs!dDocFecha
   If Not IsNull(rs!cNomPers) Then
      fg.TextMatrix(nItem, 4) = PstaNombre(Trim(Mid(rs!cNomPers, 1, Len(rs!cNomPers) - 50)), True)
   End If
   fg.TextMatrix(nItem, 5) = Format(rs!nDocImporte, gsFormatoNumeroView)
   fg.TextMatrix(nItem, 6) = rs!cMovDesc
   fg.TextMatrix(nItem, 7) = rs!nMovNro
   fg.TextMatrix(nItem, 8) = Right(rs!cNomPers, 13) & "" ' rs!cBSCod 'CODIGO PERSONA
   fg.TextMatrix(nItem, 9) = Format(rs!nDocImporte - rs!nMontoA, gsFormatoNumeroView)
   If rs!nMovEstado = gMovEstPresupAceptado And Not rs!nMovFlag = gMovFlagExtornado Then
      fg.TextMatrix(nItem, 10) = "Aprobado"
   ElseIf rs!nMovEstado = gMovEstPresupPendiente And Not rs!nMovFlag = gMovFlagExtornado Then
      fg.TextMatrix(nItem, 10) = "Pendiente"
      fg.Col = 10
      fg.CellBackColor = "&H00C0C0FF"
   ElseIf rs!nMovEstado = gMovEstPresupRechazado Then
      fg.TextMatrix(nItem, 10) = "RECHAZADO"
      fg.Col = 10
      fg.CellBackColor = "&H0080FF80"
   Else
      If rs!nMovFlag = gMovFlagExtornado Or rs!nMovFlag = gMovFlagEliminado Then
         fg.TextMatrix(nItem, 10) = "ELIMINADO"
         fg.Col = 10
         fg.CellBackColor = "&H0080FF80"
      End If
   End If
   nTot = nTot + rs!nDocImporte
   rs.MoveNext
Loop
RSClose rs
txtTot = Format(nTot, gsFormatoNumeroView)
fg.Row = 2
fg.Col = 1
End Sub

Private Sub CmdAceptar_Click()
Dim N As Integer
Dim sMovAnt As String
Dim sMovNro As String
Dim nCont   As Integer
Dim sCta    As String
Dim nSaldo  As Currency
On Error GoTo ErrSub

If fg.TextMatrix(2, 1) = "" Then
   Exit Sub
End If
If fg.Row = 0 Then
    MsgBox "Seleccione Documento a provisionar", vbInformation, "¡Aviso!"
    Exit Sub
End If
If MsgBox(" ¿ Seguro que desea Provisionar " & sDocDesc & " Nro. " & fg.TextMatrix(fg.Row, 2) & " ? ", vbQuestion + vbYesNo, "Confirmación") = vbNo Then
   Exit Sub
End If
gnMovNro = fg.TextMatrix(fg.Row, 7)
gsGlosa = fg.TextMatrix(fg.Row, 6)
gsPersNombre = fg.TextMatrix(fg.Row, 4)
gnDocTpo = sDocTpoOC
gsDocNro = fg.TextMatrix(fg.Row, 2)
gdFecha = fg.TextMatrix(fg.Row, 3)

gnSaldo = nVal(fg.TextMatrix(fg.Row, 9))
nSaldo = gnSaldo
'ALPA 20090303*****************************************************************
'frmLogProvisionPago.Inicio True, lbBienes, fg.TextMatrix(fg.Row, 8)
frmLogProvisionPago.Inicio True, lbBienes, fg.TextMatrix(fg.Row, 8), , , , , True
'*******************************************************************************
If frmLogProvisionPago.lOk Then
   If gnSaldo <= 0 Then
      EliminaRow fg, fg.Row, 2
      ActualizaTot nSaldo * -1
   Else
      fg.TextMatrix(fg.Row, 9) = Format(gnSaldo, gsFormatoNumeroView)
      ActualizaTot (nSaldo - gnSaldo) * -1
   End If
End If
Exit Sub
ErrSub:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdActualizar_Click()
    FormatoOCompra
    GetOCPendientes
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If lSalir Then
   RSClose rs
   Unload Me
End If
End Sub

Private Sub Form_Load()
    Dim lnColorBien As Double
    Dim lnColorServ As Double
    Dim oDoc As DOperacion
    Set oDoc = New DOperacion
    CentraForm Me
    
    lnColorBien = "&H00F0FFFF"
    lnColorServ = "&H00FFFFC0"
    If lbBienes Then
       sDocDesc = "Orden de Compra"
       fg.BackColor = lnColorBien
    Else
       sDocDesc = "Orden de Servicio"
       fg.BackColor = lnColorServ
    End If
    Me.Caption = gsOpeDesc
    lSalir = False
    Set rs = oDoc.CargaOpeCta(gsOpeCod, "H")
    If rs.EOF And rs.BOF Then
       MsgBox "Cuenta Contable de Provisión no fue asignada a Operación." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Sistemas", vbInformation, "¡Aviso!"
       lSalir = True
       Exit Sub
    End If
    sCtaProvis = rs!cCtaContCod
    
    Set rs = oDoc.CargaOpeDoc(gsOpeCod, , OpeDocMetAutogenerado)
    If rs.EOF Then
       MsgBox "No se asignó Tipo de Documento " & sDocDesc & " a Operación", vbInformation, "¡Aviso!"
       lSalir = True
       Exit Sub
    End If
    sDocTpoOC = rs!nDocTpo
    RSClose rs

    FormatoOCompra

End Sub

Private Sub ActualizaTot(pnMonto As Currency)
    txtTot = Format(nVal(txtTot) + pnMonto, gsFormatoNumeroView)
End Sub

