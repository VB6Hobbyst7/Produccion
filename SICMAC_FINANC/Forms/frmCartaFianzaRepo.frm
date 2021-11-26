VERSION 5.00
Begin VB.Form frmCartaFianzaRepo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cartas Fianza: Reportes"
   ClientHeight    =   2505
   ClientLeft      =   3630
   ClientTop       =   3105
   ClientWidth     =   4935
   Icon            =   "frmCartaFianzaRepo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   2430
      Begin VB.OptionButton optSeleccion 
         Caption         =   "Cartas Fianzas entregadas"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   465
         Width           =   2250
      End
      Begin VB.OptionButton optSeleccion 
         Caption         =   "Cartas Fianzas Ingresadas "
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Value           =   -1  'True
         Width           =   2250
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1920
      Width           =   1395
   End
   Begin VB.CommandButton cmdgenerar 
      Caption         =   "&Generar "
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   1395
   End
End
Attribute VB_Name = "frmCartaFianzaRepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsObjeto() As String
Dim lsTipoDoc As String
Dim lsCtaCont As String
Dim lnPaginas  As Long

Private Sub cmdGenerar_Click()
rtf.Text = ""
If Me.optSeleccion(0).Value = True Then
    ReporteCartas True, lsCtaCont
Else
    ReporteCartas False, lsCtaCont
End If

frmPrevio.Previo rtf, "Reporte de a Rendir a Cuenta", True, 66
Barra.Value = 0
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
AbreConexion
CentraSdi Me
txtDesde = gdFecSis
txtHasta = gdFecSis
CargaVariablesOperacion gcOpeCod
End Sub
Private Sub CargaVariablesOperacion(lsCodOpe As String)
Dim sql As String
Dim rs As New ADODB.Recordset
Dim I As Integer

sql = "SELECT * FROM " & gcCentralCom & "OpeDoc WHERE COPECOD='" & lsCodOpe & "'"
rs.Open sql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
If Not RSVacio(rs) Then
    lsTipoDoc = rs!cDocTpo
End If
rs.Close
Set rs = Nothing
ReDim lsObjeto(1)
I = 0
sql = "SELECT * FROM " & gcCentralCom & "OpeObj WHERE COPECOD='" & lsCodOpe & "'"
rs.Open sql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
Do While Not rs.EOF
    I = I + 1
    ReDim Preserve lsObjeto(I)
    lsObjeto(I) = Trim(rs!cObjetoCod)
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing

sql = "SELECT * FROM " & gcCentralCom & "OpeCta WHERE COPECOD='" & lsCodOpe & "'"
rs.Open sql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
If Not RSVacio(rs) Then
    lsCtaCont = Trim(rs!cCtaContCod)
End If
rs.Close
Set rs = Nothing
End Sub
Private Function EmiteCuentaContable(lsCodOpe As String) As String
Dim sql As String
Dim rs As New ADODB.Recordset

sql = "SELECT * FROM " & gcCentralCom & "OpeCta WHERE COPECOD='" & lsCodOpe & "'"
rs.Open sql, dbCmact, adOpenStatic, adLockOptimistic, adCmdText
If Not RSVacio(rs) Then
    EmiteCuentaContable = Trim(rs!cCtaContCod)
End If
rs.Close
Set rs = Nothing

End Function

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Sub txtDesde_GotFocus()
fEnfoque txtDesde
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtDesde) = True Then
        Me.txtHasta.SetFocus
    End If
End If
End Sub
Private Sub txtHasta_GotFocus()
fEnfoque txtHasta
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If ValFecha(txtHasta) = True Then
        Me.cmdGenerar.SetFocus
    End If
End If
End Sub
Private Sub ReporteCartas(lbIngreso As Boolean, lsCodCtaCont As String)
Dim lnPaginas As Long
Dim lnLineas As Long
Dim I As Integer
Dim Total As Currency
Dim sql As String
Dim rs As New ADODB.Recordset
Dim lsFechaIni As String
Dim lsFechaFin As String

rtf.Text = ""

lsFechaIni = Format(Me.txtDesde, "yyyymmdd")
lsFechaFin = Format(Me.txtHasta, "yyyymmdd")

lnPaginas = 1
EncabezadoReporte Str(lnPaginas), lbIngreso
lnLineas = 6
If lbIngreso = True Then
    If Mid(lsCodCtaCont, 3, 1) = "1" Then
        sql = " SELECT MD.CDOCNRO, CONVERT(CHAR(50),M.CMOVDESC) DESCRIPCION , O.COBJETODESC, P.CNOMPERS, " _
            & " MD.DDOCFECHA, MC.NMOVIMPORTE,MD.DDOCFECHA, CONVERT(datetime,MOTR.CMOVOTROVARIABLE) AS FECHAVENC,MCO.COBJETOCOD,M.CMOVNRO " _
            & " FROM MOVCTA MC INNER JOIN MOV M ON M.CMOVNRO=MC.CMOVNRO " _
            & " INNER JOIN MOVDOC MD ON MD.CMOVNRO=M.CMOVNRO " _
            & " INNER JOIN MOVCABOBJ MCO ON M.CMOVNRO=MCO.CMOVNRO " _
            & " INNER JOIN " & gcCentralCom & "Objeto O ON O.COBJETOCOD=MCO.COBJETOCOD " _
            & " INNER JOIN MOVOTROS MOTR ON MOTR.CMOVNRO=M.CMOVNRO " _
            & " INNER JOIN MOVCABOBJ MCO1 ON M.CMOVNRO=MCO1.CMOVNRO " _
            & " INNER JOIN " & gcCentralPers & "Persona P ON P.CCODPERS=SUBSTRING(MCO1.COBJETOCOD,3,10) " _
            & " WHERE MC.NMOVIMPORTE>0 and MC.CCTACONTCOD='" & Trim(lsCodCtaCont) & "' and M.CMOVESTADO='0' AND M.CMOVFLAG NOT  IN('X','E','N') " _
            & " AND NOT EXISTS (SELECT MR.CMOVNRO FROM MOVREF MR WHERE MR.CMOVNROREF=M.CMOVNRO) " _
            & " AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "'"
    Else
        sql = " SELECT MD.CDOCNRO, CONVERT(CHAR(30),M.CMOVDESC) DESCRIPCION , O.COBJETODESC, P.CNOMPERS, " _
            & " MD.DDOCFECHA, MME.NMOVMEIMPORTE NMOVIMPORTE,MD.DDOCFECHA, CONVERT(DATETIME,MOTR.CMOVOTROVARIABLE) AS FECHAVENC,MCO.COBJETOCOD,M.CMOVNRO " _
            & " FROM MOVCTA MC INNER JOIN MOV M ON M.CMOVNRO=MC.CMOVNRO " _
            & " INNER JOIN MOVDOC MD ON MD.CMOVNRO=M.CMOVNRO " _
            & " INNER JOIN MOVCABOBJ MCO ON M.CMOVNRO=MCO.CMOVNRO  " _
            & " INNER JOIN " & gcCentralCom & "Objeto O ON O.COBJETOCOD=MCO.COBJETOCOD " _
            & " INNER JOIN MOVOTROS MOTR ON MOTR.CMOVNRO=M.CMOVNRO " _
            & " INNER JOIN MOVCABOBJ MCO1 ON M.CMOVNRO=MCO1.CMOVNRO " _
            & " INNER JOIN " & gcCentralPers & "Persona P ON P.CCODPERS=SUBSTRING(MCO1.COBJETOCOD,3,10) " _
            & " INNER JOIN MOVME MME ON MME.CMOVNRO=MC.CMOVNRO AND MC.CMOVITEM=MME.CMOVITEM " _
            & " WHERE MME.NMOVMEIMPORTE>0 and MC.CCTACONTCOD='" & Trim(lsCodCtaCont) & "' and M.CMOVESTADO='0' AND M.CMOVFLAG NOT  IN('X','E','N') " _
            & " AND NOT EXISTS (SELECT MR.CMOVNRO FROM MOVREF MR WHERE MR.CMOVNROREF=M.CMOVNRO) " _
            & " AND SUBSTRING(M.CMOVNRO,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "'"
    End If
Else
    If Mid(lsCodCtaCont, 3, 1) = "1" Then
            sql = "SELECT MD.CDOCNRO, CONVERT(CHAR(30),M.CMOVDESC) DESCRIPCION , O.COBJETODESC, P.CNOMPERS, " _
            & "MC.NMOVIMPORTE,MD.DDOCFECHA, MOTR.CMOVOTROVARIABLE AS FECHAVENC,MCO.COBJETOCOD,M.CMOVNRO, " _
            & "MD.DDOCFECHA, MR.CMOVNRO as cMovNroRef " _
            & "FROM MOVCTA MC INNER JOIN MOV M ON M.CMOVNRO=MC.CMOVNRO " _
            & "INNER JOIN MOVDOC MD ON MD.CMOVNRO=M.CMOVNRO " _
            & "INNER JOIN MOVCABOBJ MCO ON M.CMOVNRO=MCO.CMOVNRO " _
            & "INNER JOIN " & gcCentralCom & "Objeto O ON O.COBJETOCOD=MCO.COBJETOCOD " _
            & "INNER JOIN MOVOTROS MOTR ON MOTR.CMOVNRO=M.CMOVNRO " _
            & "INNER JOIN MOVCABOBJ MCO1 ON M.CMOVNRO=MCO1.CMOVNRO " _
            & "INNER JOIN " & gcCentralPers & "Persona P ON P.CCODPERS=SUBSTRING(MCO1.COBJETOCOD,3,10) " _
            & "INNER JOIN MOVREF MR ON  MR.CMOVNROREF=M.CMOVNRO " _
            & "WHERE MC.NMOVIMPORTE>0 and MC.CCTACONTCOD='" & Trim(lsCodCtaCont) & "' AND M.CMOVESTADO='0' AND M.CMOVFLAG NOT  IN('X','E','N') AND " _
            & "EXISTS (SELECT MR.CMOVNRO FROM MOVREF MR WHERE MR.CMOVNROREF=M.CMOVNRO) " _
            & "AND SUBSTRING(MR.CMOVNROREF,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "'"
    Else
        sql = "SELECT MD.CDOCNRO, CONVERT(CHAR(30),M.CMOVDESC) DESCRIPCION , O.COBJETODESC, P.CNOMPERS, " _
            & "MD.DDOCFECHA, MME.NMOVMEIMPORTE NMOVIMPORTE ,MD.DDOCFECHA, MOTR.CMOVOTROVARIABLE AS FECHAVENC,MCO.COBJETOCOD,M.CMOVNRO, MR.CMOVNRO as cMovNroRef " _
            & "FROM MOVCTA MC INNER JOIN MOV M ON M.CMOVNRO=MC.CMOVNRO " _
            & "INNER JOIN MOVDOC MD ON MD.CMOVNRO=M.CMOVNRO " _
            & "INNER JOIN MOVCABOBJ MCO ON M.CMOVNRO=MCO.CMOVNRO " _
            & "INNER JOIN " & gcCentralCom & "Objeto O ON O.COBJETOCOD=MCO.COBJETOCOD " _
            & "INNER JOIN MOVOTROS MOTR ON MOTR.CMOVNRO=M.CMOVNRO " _
            & "INNER JOIN MOVCABOBJ MCO1 ON M.CMOVNRO=MCO1.CMOVNRO " _
            & "INNER JOIN " & gcCentralPers & "Persona P ON P.CCODPERS=SUBSTRING(MCO1.COBJETOCOD,3,10) " _
            & "INNER JOIN MOVME MME ON MME.CMOVNRO=MC.CMOVNRO AND MC.CMOVITEM=MME.CMOVITEM " _
            & "INNER JOIN MOVREF MR ON  MR.CMOVNROREF=M.CMOVNRO " _
            & "WHERE MME.NMOVMEIMPORTE>0 and MC.CCTACONTCOD='" & Trim(lsCodCtaCont) & "' AND M.CMOVESTADO='0' AND M.CMOVFLAG NOT  IN('X','E','N') " _
            & "AND EXISTS (SELECT MR.CMOVNRO FROM MOVREF MR WHERE MR.CMOVNROREF=M.CMOVNRO) " _
            & " AND SUBSTRING(MR.CMOVNROREF,1,8) BETWEEN '" & lsFechaIni & "' AND '" & lsFechaFin & "'"
    End If
End If
    
rs.Open sql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
Do While Not rs.EOF
        If lbIngreso = True Then
            rtf.Text = rtf.Text & PrnSet("MI", 5) & ImpreFormat(Trim(rs!cDocNro), 22, 0) & ImpreFormat(Trim(PstaNombre(rs!cNomPers)), 25) & ImpreFormat(Trim(EliminaEnters(rs!DESCRIPCION)), 25) & ImpreFormat(rs!cObjetoDesc, 18) & ImpreFormat(rs!nMovImporte, 12) & _
                    ImpreFormat(Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4), 12) & ImpreFormat(Format(rs!dDocFecha, "dd/mm/yyyy"), 12) & ImpreFormat(Format(rs!FECHAVENC, "dd/mm/yyyy"), 12) & Chr(10)
        Else
            rtf.Text = rtf.Text & PrnSet("MI", 5) & ImpreFormat(Trim(rs!cDocNro), 22, 0) & ImpreFormat(Trim(PstaNombre(rs!cNomPers)), 25) & ImpreFormat(Trim(EliminaEnters(rs!DESCRIPCION)), 25) & ImpreFormat(rs!cObjetoDesc, 18) & ImpreFormat(rs!nMovImporte, 12) & _
                    ImpreFormat(Mid(rs!cMovNro, 7, 2) & "/" & Mid(rs!cMovNro, 5, 2) & "/" & Mid(rs!cMovNro, 1, 4), 12) & ImpreFormat(Format(rs!FECHAVENC, "dd/mm/yyyy"), 12) & ImpreFormat(Mid(rs!cMovNroRef, 7, 2) & "/" & Mid(rs!cMovNroRef, 5, 2) & "/" & Mid(rs!cMovNroRef, 1, 4), 12) & Chr(10)
        End If
        Total = Total + rs!nMovImporte
        lnLineas = lnLineas + 1
        If lnLineas > 60 Then
            lnPaginas = lnPaginas + 1
            EncabezadoReporte Str(lnPaginas), lbIngreso
            lnLineas = 6
        End If
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
rtf.Text = rtf.Text & PrnSet("MI", 5) & String(150, "-") & Chr(10)
If lbIngreso = True Then
    rtf.Text = rtf.Text & PrnSet("B+") & ImpreFormat("TOTAL :", 12, 95) & ImpreFormat(Total, 12) & PrnSet("B-") & Chr(10)
Else
    rtf.Text = rtf.Text & PrnSet("B+") & ImpreFormat("TOTAL :", 12, 72) & ImpreFormat(Total, 12) & PrnSet("B-") & Chr(10)
End If
End Sub
Private Sub EncabezadoReporte(lsNumPagina As String, lbIngreso As Boolean)
Dim lsTitulo As String
If lbIngreso = True Then
    lsTitulo = "REPORTE DE CARTAS FIANZAS INGRESADAS"
Else
    lsTitulo = "REPORTE DE CARTAS FIANZAS ENTREGADAS"
End If
rtf.Text = rtf.Text & PrnSet("MI", 5) & CabeRepo("", "", 150, "CAJA GENERAL-" & IIf(Mid(gcOpeCod, 3, 1) = "1", "SOLES", "DOLARES"), lsTitulo, " DESDE " & txtDesde & " AL " & txtHasta, "", lsNumPagina) & Chr(10)
rtf.Text = rtf.Text & PrnSet("MI", 5) & String(150, "-") & Chr(10)
If lbIngreso = True Then
    rtf.Text = rtf.Text & PrnSet("MI", 5) & ImpreFormat("No CARTA FIANZA", 22, 0) & ImpreFormat("PROVEEDOR", 25) & ImpreFormat("DESCRIPCION", 25) & ImpreFormat("ENTIDAD FINANCIERA", 25) & ImpreFormat("MONTO", 6) & ImpreFormat("FECHA ING.", 12) & ImpreFormat("FECHA EMIS.", 12) & ImpreFormat("FECHA VENC.", 12) & Chr(10)
Else
    rtf.Text = rtf.Text & PrnSet("MI", 5) & ImpreFormat("No CARTA FIANZA", 22, 0) & ImpreFormat("PROVEEDOR", 25) & ImpreFormat("DESCRIPCION", 25) & ImpreFormat("ENTIDAD FINANCIERA", 25) & ImpreFormat("MONTO", 6) & ImpreFormat("FECHA ING.", 12) & ImpreFormat("FECHA VENC.", 12) & ImpreFormat("FEC.ENTREG.", 12) & Chr(10)
End If
rtf.Text = rtf.Text & PrnSet("MI", 5) & String(150, "-") & Chr(10)
End Sub
