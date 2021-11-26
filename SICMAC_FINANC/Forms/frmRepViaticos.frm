VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepViaticos 
   Caption         =   "Reporte de Viáticos"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   Icon            =   "frmRepViaticos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   405
      Left            =   2580
      TabIndex        =   9
      Top             =   1950
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   405
      Left            =   1080
      TabIndex        =   8
      Top             =   1950
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Reporte"
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
      Left            =   120
      TabIndex        =   3
      Top             =   930
      Width           =   4875
      Begin VB.OptionButton OptTipo 
         Caption         =   "Detalle por Concepto"
         Height          =   375
         Index           =   1
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   210
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Resumen por Persona"
         Height          =   375
         Index           =   0
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   210
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   765
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4170
      Begin MSMask.MaskEdBox txtFechaDel 
         Height          =   345
         Left            =   660
         TabIndex        =   1
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox txtFechaAl 
         Height          =   345
         Left            =   2760
         TabIndex        =   2
         Top             =   300
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DEL"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   390
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "AL"
         Height          =   195
         Left            =   2370
         TabIndex        =   6
         Top             =   390
         Width           =   195
      End
   End
End
Attribute VB_Name = "frmRepViaticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsArchivo As String
Dim lbExcel As Boolean
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim rs      As ADODB.Recordset
Dim oCon    As DConecta

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If ValidaFecha(txtFechaDel) <> "" Then
   MsgBox "Fecha Inicial no válida", vbInformation, "¡Aviso!"
   txtFechaDel.SetFocus
   Exit Function
End If
If ValidaFecha(txtFechaAl) <> "" Then
   MsgBox "Fecha Final no válida", vbInformation, "¡Aviso!"
   txtFechaAl.SetFocus
End If
ValidaDatos = True
End Function

Private Sub cmdGenerar_Click()
Dim N As Integer
On Error GoTo GeneraEstadError
MousePointer = 11
If ValidaDatos() Then
   lsArchivo = App.path & "\SPOOLER\" & "RPViaticos_" & Format(txtFechaAl, "yyyymmdd") & "_" & IIf(Mid(gsOpeCod, 3, 1) = gMonedaNacional, "MN", "ME") & ".XLS"
   lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
   If Not lbExcel Then
      MousePointer = 0
      Exit Sub
   End If

    ExcelAddHoja IIf(OptTipo(0).value, "Por Persona", "Por Concepto"), xlLibro, xlHoja1, False
    xlHoja1.PageSetup.Orientation = xlPortrait
    xlHoja1.PageSetup.CenterVertically = True
    xlHoja1.PageSetup.CenterHorizontally = True

   If OptTipo(0).value = True Then
       CabeceraReporte 0
       Set rs = CargaDatosViaticos(0)
       Do While Not rs.EOF
         xlHoja1.Cells(rs.Bookmark + 5, 1) = rs!cPersNombre
         xlHoja1.Cells(rs.Bookmark + 5, 2) = rs!cCtaContCod
         xlHoja1.Cells(rs.Bookmark + 5, 3) = rs!cCtaContDesc
         xlHoja1.Cells(rs.Bookmark + 5, 4) = rs!Gasto
         rs.MoveNext
       Loop
       xlHoja1.Cells(rs.RecordCount + 6, 3) = "TOTAL"
       xlHoja1.Range(xlHoja1.Cells(rs.RecordCount + 6, 4), xlHoja1.Cells(rs.RecordCount + 6, 4)).Formula = "=SUM(D6:D" & rs.RecordCount + 5 & ")"
       xlHoja1.Range(xlHoja1.Cells(rs.RecordCount + 6, 3), xlHoja1.Cells(rs.RecordCount + 6, 4)).Font.Bold = True
       xlHoja1.Range(xlHoja1.Cells(6, 4), xlHoja1.Cells(rs.RecordCount + 6, 4)).NumberFormat = "#,##0.00"
       ExcelCuadro xlHoja1, 1, rs.RecordCount + 6, 4, rs.RecordCount + 6, False, False
   Else
       CabeceraReporte 1
       Set rs = CargaDatosViaticos(1)
       N = 6
       Do While Not rs.EOF
         N = N + 1
         gsMovNro = rs!cMovNro
         xlHoja1.Cells(N, 1) = rs!cDocNro
         xlHoja1.Cells(N, 2) = rs!cPersNombre
         xlHoja1.Cells(N, 3) = rs!cRHCargoDesc
         xlHoja1.Cells(N, 4) = Replace(Replace(rs!cMovDesc, Chr(10), " "), Chr(13), " ")
         xlHoja1.Cells(N, 5) = rs!nMovViaticosDias
         xlHoja1.Cells(N, 14) = rs!cCtaContCod
         Do While rs!cMovNro = gsMovNro
            Select Case rs!cObjetoCod
               Case "2001"  'Alimentacion y Hospedaje
                  xlHoja1.Cells(N, 6) = rs!nMontoSol
                  xlHoja1.Cells(N, 7) = rs!nMontoSust
               Case "2004"  'MOVILIDAD TERRESTRE
                  xlHoja1.Cells(N, 8) = rs!nMontoSol
                  xlHoja1.Cells(N, 9) = rs!nMontoSust
               Case "2003"  'Movilidad Interna
                  xlHoja1.Cells(N, 10) = rs!nMontoSol
                  xlHoja1.Cells(N, 11) = rs!nMontoSust
               Case Else    'OTROS
                  xlHoja1.Cells(N, 12) = xlHoja1.Cells(N, 12) + rs!nMontoSol
                  xlHoja1.Cells(N, 13) = xlHoja1.Cells(N, 13) + rs!nMontoSust
            End Select
            rs.MoveNext
            If rs.EOF Then
               Exit Do
            End If
         Loop
       Loop
       ExcelCuadro xlHoja1, 1, 7, 14, N, True, False
       N = N + 1
       ExcelCuadro xlHoja1, 1, N, 14, N, True, False
       xlHoja1.Cells(N, 4) = "TOTALES"
       xlHoja1.Range("E5:E" & N).HorizontalAlignment = xlCenter
       xlHoja1.Range("F" & N & ":F" & N).Formula = "=SUM(F7:F" & N - 1 & ")"
       xlHoja1.Range("G" & N & ":G" & N).Formula = "=SUM(G7:G" & N - 1 & ")"
       xlHoja1.Range("H" & N & ":H" & N).Formula = "=SUM(H7:H" & N - 1 & ")"
       xlHoja1.Range("I" & N & ":I" & N).Formula = "=SUM(I7:I" & N - 1 & ")"
       xlHoja1.Range("J" & N & ":J" & N).Formula = "=SUM(J7:J" & N - 1 & ")"
       xlHoja1.Range("K" & N & ":K" & N).Formula = "=SUM(K7:K" & N - 1 & ")"
       xlHoja1.Range("L" & N & ":L" & N).Formula = "=SUM(L7:L" & N - 1 & ")"
       xlHoja1.Range("M" & N & ":M" & N).Formula = "=SUM(M7:M" & N - 1 & ")"
       xlHoja1.Range("A" & N & ":N" & N).Font.Bold = True
       xlHoja1.Range("F7:M" & N).NumberFormat = "#,##0.00"
   End If
   ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, True
   CargaArchivo lsArchivo, App.path & "\SPOOLER\"
   MousePointer = 0
   MsgBox "Reporte generado satisfactoriamente.", vbInformation, "¡Aviso!"
End If
Exit Sub
GeneraEstadError:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
   ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
   MousePointer = 0
End Sub

Private Sub CabeceraReporte(pnTipo As Integer)
xlHoja1.Cells(1, 1) = gsNomCmac
xlHoja1.Cells(3, 1) = "DEL " & txtFechaDel & " AL " & txtFechaAl
If pnTipo = 0 Then
   xlHoja1.Cells(2, 1) = "REPORTE DE VIATICOS POR EMPLEADO"
   xlHoja1.Cells(5, 1) = "PERSONA"
   xlHoja1.Cells(5, 2) = "CTA CONTABLE"
   xlHoja1.Cells(5, 3) = "DESCRIPCION"
   xlHoja1.Cells(5, 4) = "IMPORTE"
   xlHoja1.Range("A5:A5").ColumnWidth = 35
   xlHoja1.Range("B5:B5").ColumnWidth = 18
   xlHoja1.Range("C5:C5").ColumnWidth = 35
   xlHoja1.Range("D5:D5").ColumnWidth = 15
   xlHoja1.Range("A2:D3").HorizontalAlignment = xlCenter
   xlHoja1.Range("A3:D3").Font.Size = 14
   xlHoja1.Range("A1:D5").Font.Bold = True
   xlHoja1.Range("A2:D2").Merge True
   xlHoja1.Range("A3:D3").Merge True
   ExcelCuadro xlHoja1, 1, 5, 4, 5, True
Else
   xlHoja1.Cells(2, 1) = "REPORTE DE VIATICOS POR EMPLEADO DETALLADO POR CONCEPTO"
   xlHoja1.Cells(5, 1) = "NroDoc"
   xlHoja1.Cells(5, 2) = "Persona"
   xlHoja1.Cells(5, 3) = "Cargo"
   xlHoja1.Cells(5, 4) = "Motivo del Viaje"
   xlHoja1.Cells(5, 5) = "Días"
   xlHoja1.Cells(5, 6) = "Alimentacion y Hospedaje"
   xlHoja1.Cells(5, 8) = "Transporte Terrestre"
   xlHoja1.Cells(5, 10) = "Movilidad Interna"
   xlHoja1.Cells(5, 12) = "Otros"
   xlHoja1.Cells(5, 14) = "Cta Contable"
   xlHoja1.Cells(6, 6) = "Asignado"
   xlHoja1.Cells(6, 7) = "Gasto"
   xlHoja1.Cells(6, 8) = "Asignado"
   xlHoja1.Cells(6, 9) = "Gasto"
   xlHoja1.Cells(6, 10) = "Asignado"
   xlHoja1.Cells(6, 11) = "Gasto"
   xlHoja1.Cells(6, 12) = "Asignado"
   xlHoja1.Cells(6, 13) = "Gasto"
   xlHoja1.Range("A2:K2").Merge True
   xlHoja1.Range("A3:K3").Merge True
   xlHoja1.Range("F5:G5").Merge True
   xlHoja1.Range("H5:I5").Merge True
   xlHoja1.Range("J5:K5").Merge True
   xlHoja1.Range("L5:M5").Merge True
   
   
   xlHoja1.Range("A5:A6").MergeCells = True
   xlHoja1.Range("B5:B6").MergeCells = True
   xlHoja1.Range("C5:C6").MergeCells = True
   xlHoja1.Range("D5:D6").MergeCells = True
   xlHoja1.Range("E5:E6").MergeCells = True
   xlHoja1.Range("N5:N6").MergeCells = True
   
   xlHoja1.Range("B5:B5").ColumnWidth = 28
   xlHoja1.Range("C5:C5").ColumnWidth = 25
   xlHoja1.Range("D5:D5").ColumnWidth = 30
   xlHoja1.Range("E5:E5").ColumnWidth = 8

   xlHoja1.Range("A2:N6").HorizontalAlignment = xlCenter
   xlHoja1.Range("A5:N6").VerticalAlignment = xlCenter
   xlHoja1.Range("A1:N6").Font.Bold = True
   ExcelCuadro xlHoja1, 1, 5, 14, 6, True, True
End If
End Sub


Private Function CargaDatosViaticos(pnTipo As Integer) As ADODB.Recordset
Dim sSql As String
If pnTipo = 0 Then
   sSql = "SELECT p.cPersCod, p.cPersNombre, left(Sust.cCtaContCod,8) cCtaContCod, c.cCtaContDesc, sum(sust.nMovImporte) Gasto " _
        & "FROM mov m join movarendir ma on m.nmovnro = ma.nmovnro " _
        & " left join Persona p ON p.cPersCod = ma.cPersCod " _
        & " JOIN (SELECT mrA.nMovNroRef FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro " _
        & "         JOIN MovRef mrA ON mrA.nMovNro = m.nMovNro " _
        & "       WHERE copecod like '40__3%' and mc.cCtaContCod LIKE '19_106' " _
        & "      ) Ate ON Ate.nMovNroRef = m.nMovNro " _
        & " Left Join (SELECT mr.nMovNroRef, left(mc.cCtaContCod,8) cCtaContCod, sum(mc.nMovImporte) nMovImporte " _
        & "        FROM mov m JOIN movcta mc on mc.nMovNro = m.nMovNro " _
        & "        JOIN movref mr on mr.nmovnro = m.nMovNro " _
        & "        WHERE not copecod like '40__[356]%' and "
        
   sSql = sSql & "mc.cCtaContCod LIKE '4%' and m.nMovFlag = 0 and m.nMovEstado = 10 " _
        & "        GROUP BY mr.nMovNroRef, left(mc.cCtaContCod,8) " _
        & "       ) sust ON sust.nMovNroRef = m.nMovNro " _
        & "       join ctacont c on c.cCtaContCod = sust.cCtaContCod " _
        & "   WHERE LEFT(m.cmovnro,8) between '" & Format(txtFechaDel, gsFormatoFecha) & "' and '" & Format(txtFechaAl, gsFormatoFecha) & "' and ma.cTpoArendir = '2' and m.nMovFlag = 0 " _
        & "     and SubString(m.cOpeCod,3,1) = " & Mid(gsOpeCod, 3, 1) _
        & "   GROUP BY p.cPersCod, p.cPersNombre, left(Sust.cCtaContCod,8), c.cCtaContDesc " _
        & "   ORDER BY p.cPersNombre, left(Sust.cCtaContCod,8) "
Else

   sSql = ""
   sSql = sSql & " SELECT m.nMovNro, m.cMovNro, md.cDocNro, p.cPersNombre, m.cMovDesc, via.nMovViaticosDias, ate.nImporteAte, ISNULL(via.cObjetoCod,'OTROS') cObjetoCod, ISNULL(via.nMontoSol,0) nMontoSol, ISNULL(via.nMontoSust,0) nMontoSust, via.cCtaContCod "
   sSql = sSql & "  ,cRHCargoDesc = (SELECT rhct.cRHCargoDescripcion"
   sSql = sSql & "                  FROM rhcargos rhc JOIN rhcargostabla rhct on rhc.cRHCargoCod = rhct.cRHCargoCod"
   sSql = sSql & "                   and dRHCargoFecha = (SELECT Max(dRHCargoFecha) FROM rhCargos rhc1 WHERE rhc1.cPersCod = rhc.cPersCod )"
   sSql = sSql & "                   and rhc.cPersCod = p.cPersCod"
   sSql = sSql & "                 )"
   sSql = sSql & " FROM mov m join movdoc md on md.nMovNro = m.nMovNro"
   sSql = sSql & " join movarendir ma on m.nmovnro = ma.nmovnro"
   sSql = sSql & " left join Persona p ON p.cPersCod = ma.cPersCod"
   sSql = sSql & " JOIN (SELECT mrA.nMovNroRef, MAx(m.cMovNro) cMovNro, Max(m.nMovNro) nMovNro, Sum(nMovImporte) nImporteAte"
   sSql = sSql & "       FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro"
   sSql = sSql & "    JOIN MovRef mrA ON mrA.nMovNro = m.nMovNro"
   sSql = sSql & "       WHERE copecod like '40_23%' and mc.cCtaContCod LIKE '19_106'"
   sSql = sSql & "       GROUP BY mrA.nMovNroRef"
   sSql = sSql & "      ) Ate ON Ate.nMovNroRef = m.nMovNro"
   sSql = sSql & " join (SELECT ISNULL(sol.nViaticoMovNro, sust.nMovNroRef) nViaticoMovNro, ISNULL(sol.cObjetoCod, sust.cObjetoCod) cObjetoCod,"
   sSql = sSql & "              Max(sol.nMovViaticosDias) nMovViaticosDias, Max(sust.cCtaContCod) cCtaContCod,"
   sSql = sSql & "              Sum(Sol.nMovImporte) nMontoSol, Sum(sust.nMovImporte) nMontoSust"
   sSql = sSql & "       FROM ( Select nViaticoMovNro, mov.cObjetoCod, sum(nMovViaticosDias) nMovViaticosDias, Sum(nMovImporte) nMovImporte"
   sSql = sSql & "              FROM movviaticos mv JOIN MovObjViaticos mov on mov.nMovNro = mv.nMovNro"
   sSql = sSql & "              GROUP BY nViaticoMovNro, mov.cObjetoCod"
   sSql = sSql & "            ) Sol FULL OUTER JOIN"
   sSql = sSql & "            (SELECT mr.nMovNroRef, ISNULL(mo.cObjetoCod,'Otros') cObjetoCod, mc.cCtaContCod, SUM(mc.nMovImporte) nMovImporte"
   sSql = sSql & "             FROM mov m JOIN movcta mc on mc.nMovNro = m.nMovNro"
   sSql = sSql & "                   join movref mr on mr.nmovnro = m.nMovNro"
   sSql = sSql & "               left join movobj mo on mo.nmovnro = mc.nmovnro and mo.nMovItem = mc.nMovItem and Not mo.cObjetoCod IN ('13', '90')"
   sSql = sSql & "             WHERE not copecod like '40_23%' and"
   sSql = sSql & "                 ( mc.cCtaContCod not in ('191106','29180706','192106','29280706') or ( cOpeCod LIKE '40_2[56]%' and mc.cCtaContCod like '29_80706') )"
   sSql = sSql & "                   and m.nMovFlag = 0 and m.nMovEstado = 10"
   sSql = sSql & "             GROUP BY mr.nMovNroRef, ISNULL(mo.cObjetoCod,'Otros'), mc.cCtaContCod"
   sSql = sSql & "            ) sust ON sust.nMovNroRef = Sol.nViaticoMovNro and sust.cObjetoCod = sol.cObjetoCod"
   sSql = sSql & "       GROUP BY ISNULL(sol.nViaticoMovNro, sust.nMovNroRef), ISNULL(sol.cObjetoCod, sust.cObjetoCod), sust.cCtaContCod"
   sSql = sSql & "       ) Via ON Via.nViaticoMovNro = m.nMovNro"
   sSql = sSql & " WHERE m.cOpeCod LIKE '__" & Mid(gsOpeCod, 3, 1) & "%' and m.cmovnro like '2003%' and ma.cTpoArendir = '2' and m.nMovFlag = 0"
   sSql = sSql & "       and md.nDocTpo = 62"
   sSql = sSql & " ORDER BY md.cDocNro, cObjetoCod"


End If
Set CargaDatosViaticos = oCon.CargaRecordSet(sSql)
End Function


Private Sub Form_Load()
CentraForm Me
Set oCon = New DConecta
oCon.AbreConexion
End Sub

Private Sub OptTipo_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
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
   If Me.OptTipo(0).value = True Then
      OptTipo(0).SetFocus
   Else
      OptTipo(1).SetFocus
   End If
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

