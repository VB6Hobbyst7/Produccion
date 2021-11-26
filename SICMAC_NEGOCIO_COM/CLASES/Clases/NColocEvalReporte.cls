VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "nColocEvalReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event ShowProgress()
Public Event CloseProgress()
Public Event Progress(pnValor As Long, pnTotal As Long)

Dim ServerCons As String
Dim RCD As nRcdReportes


Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet  'Micro
Dim fs As Scripting.FileSystemObject

Public Sub Genera_Reporte178101(ByVal psServer As String, ByVal lsFecha As String, ByVal nChkMicro As Byte, _
ByVal nChkComer As Byte, ByVal nChkConsu As Byte, ByVal nChkHipo As Byte)
Dim Sql As String


Dim rs As New ADODB.Recordset
Dim RsM As New ADODB.Recordset
Dim fs As New Scripting.FileSystemObject
Dim lnFila As Integer
Dim lnCol As Integer
Dim lnTotal As Integer, i As Integer
Dim Y1  As Integer, Y2 As Integer
Dim lbExisteHoja As Boolean
Dim lsNomHoja As String
Dim lsTotalSaldoCap As String
Dim lsCadena  As String
Dim lsMoneda As String * 1
Dim m As Integer
Dim rsFuente As New ADODB.Recordset
Dim Conecta As DConecta
Dim ldFechaFinMes As Date
Dim lsArchivo As String
Dim lbExcel As Boolean

ldFechaFinMes = GetUltimaFechaCierre
Set Conecta = New DConecta
    RaiseEvent ShowProgress
    lsCadena = "  AND SUBSTRING(AD.cCtaCod,6,1) IN ("

    If nChkComer = 1 Then
         lsCadena = lsCadena & "'1',"
    End If
    If nChkMicro = 1 Then
        lsCadena = lsCadena & "'2',"
    End If
    If nChkConsu = 1 Then
        lsCadena = lsCadena & "'3',"
    End If
    If nChkHipo = 1 Then
        lsCadena = lsCadena & "'4',"
    End If
    RaiseEvent Progress(1, 10)
    If InStr(Len(lsCadena) - 2, lsCadena, "',", vbTextCompare) <> 0 Then
        lsCadena = Mid(lsCadena, 1, Len(lsCadena) - 1)
    End If
    lsCadena = lsCadena & ")"
    If nChkMicro = 0 And nChkComer = 0 And nChkConsu = 0 And nChkHipo = 0 Then
        lsCadena = ""
    End If
    lsNomHoja = Format(lsFecha, "ddmmyyyy") & "-" _
                & IIf(nChkMicro = 1, "Micro" & "-", "") _
                & IIf(nChkComer = 1, "Comer" & "-", "") _
                & IIf(nChkConsu = 1, "Consu" & "-", "") _
                & IIf(nChkHipo = 1, "Hipo", "")

    lsNomHoja = IIf(Right(lsNomHoja, 1) = "-", Mid(lsNomHoja, 1, Len(lsNomHoja) - 1), lsNomHoja)
    RaiseEvent Progress(3, 10)
        lsArchivo = "CredEval" & Format(ldFechaFinMes, "ddmmyyyy") & ".XLS"
        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
            Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
        Else
            Set xlLibro = xlAplicacion.Workbooks.Add
        End If
        'lsArchivo = True
        lbExisteHoja = False
        For Each xlHoja1 In xlLibro.Worksheets
            If xlHoja1.Name = lsNomHoja Then
                xlHoja1.Activate
                xlHoja1.Range("A1", "AZ10000") = ""
                lbExisteHoja = True
                Exit For
            End If
        Next
        If lbExisteHoja = False Then
            Set xlHoja1 = xlLibro.Worksheets.Add
            xlHoja1.Name = lsNomHoja
        End If

        lnFila = 2
        xlHoja1.PageSetup.Zoom = 80
        xlHoja1.PageSetup.Orientation = xlLandscape
        xlAplicacion.Range("A1:R1000").Font.Size = 9

        xlHoja1.Range("A1").ColumnWidth = 12
        xlHoja1.Range("B1").ColumnWidth = 22
        xlHoja1.Range("C1").ColumnWidth = 45
        xlHoja1.Range("D1").ColumnWidth = 30
        xlHoja1.Range("E1").ColumnWidth = 6
        xlHoja1.Range("F1").ColumnWidth = 12
        xlHoja1.Range("G1").ColumnWidth = 10
        xlHoja1.Range("H1").ColumnWidth = 10
        xlHoja1.Range("I1").ColumnWidth = 12
        xlHoja1.Range("J1").ColumnWidth = 30
        xlHoja1.Range("K1").ColumnWidth = 25


        xlHoja1.Cells(lnFila, 1) = gsNomCmac
        xlHoja1.Cells(lnFila, 7) = ""
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
        lnFila = lnFila + 2
        xlHoja1.Cells(lnFila, 4) = "REPORTE DE CREDITOS EVALUADOS AL " & lsFecha
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
        lnFila = lnFila + 2
        Y1 = lnFila
        xlHoja1.Cells(lnFila, 1) = "Desembolso"
        xlHoja1.Cells(lnFila, 2) = "N° Credito"
        xlHoja1.Cells(lnFila, 3) = "Cliente"
        xlHoja1.Cells(lnFila, 4) = "Linea de Credito"
        'xlHoja1.Cells(lnFila, 5) = "Tipo de Credito"
        xlHoja1.Cells(lnFila, 5) = "D/M"
        xlHoja1.Cells(lnFila, 6) = "Saldo Capital"
        xlHoja1.Cells(lnFila, 7) = "Cal."
        xlHoja1.Cells(lnFila, 8) = "Ult.Cal"
        xlHoja1.Cells(lnFila, 9) = "Documento"
        xlHoja1.Cells(lnFila, 10) = "Direccion"
        xlHoja1.Cells(lnFila, 11) = "Zona"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 12)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 12)).HorizontalAlignment = xlCenter
        CuadroExcel 1, Y1, 11, Y1


        'sql = "SELECT   P.CNOMPERS, ISNULL(P.cNudoci,'') AS cNudoci, ISNULL(P.cNudotr,'') AS cNudotr  , AP.cCalGen, AD.cCtaCod, AD.DFECCAL, AD.NSALDOCAP, AD.CCALAUD, " _
            & "         AD.NDIASATRASO, C.dFecultDesemb, C.cCodLinCred , LC.cDesLinCred, T.cValor, T.cNomTab , " _
            & "         FI.cDireccion as cDirFI, FI.cCodZon as cZonFI, ZFI.cDesZon as cDesZonFI " _
            & " FROM    " & gcServerAudit & "AUDPERS AP " _
            & "         JOIN " & gcCentralPers & "Persona    P   ON P.CCODPERS =AP.CCODPERS " _
            & "         JOIN " & gcServerAudit & "AUDPERSDET AD  ON AD.CCODPERS=AP.CCODPERS " _
            & "         JOIN " & gsServerConsol & "CREDITOCONSOL C ON C.cCtaCod = AD.cCtaCod JOIN " & gcCentralCom & "LineaCredito LC ON LC.CCODLINCRED = C.CCODLINCRED " _
            & "         JOIN " & gcCentralCom & "TablaCod T     ON T.CVALOR = SUBSTRING(AD.cCtaCod,6,1) " _
            & "         LEFT JOIN " & gsServerConsol & "FUENTEINGRESOConsol FI ON FI.cNumFuente = C.cNumFuente " _
            & "         LEFT JOIN " & gcCentralCom & "ZONAS ZFI ON FI.cCodZon = ZFI.cCodZon " _
            & " WHERE   AD.dFecCal = (  SELECT MAX(dFecCal) " _
            & "                         FROM   " & gcServerAudit & "AUDPERSDET AD1 " _
            & "                         WHERE  dFecCal='" & Format(lsFecha, "mm/dd/yyyy") & "' AND AD1.CCODCTA=AD.cCtaCod) " _
            & lsCadena & " AND SUBSTRING(T.CCODTAB,1,2)='03' AND AD.CESTADO IN ('1'" & IIf(chkCancelados.Value = 1, ",'0'", "") & ")" _
            & "     ORDER BY T.CCODTAB, P.CNOMPERS, AD.DFECCAL "

        'SQL = "SELECT   P.CNOMPERS, ISNULL(P.cNudoci,'') AS cNudoci, ISNULL(P.cNudotr,'') AS cNudotr  , AP.cCalGen, AD.cCtaCod, AD.DFECCAL, AD.NSALDOCAP, AD.CCALAUD, " _
        '    & "         AD.NDIASATRASO, T.cValor, T.cNomTab " _
        '    & " FROM    " & gcServerAudit & "AUDPERS AP " _
        '    & "         JOIN " & gcCentralPers & "Persona    P   ON P.CCODPERS =AP.CCODPERS " _
        '    & "         JOIN " & gcServerAudit & "AUDPERSDET AD  ON AD.CCODPERS=AP.CCODPERS JOIN " & gcCentralCom & "TablaCod T     ON T.CVALOR = SUBSTRING(AD.cCtaCod,6,1) " _
        '    & " WHERE   AD.dFecCal = '" & Format(lsFecha, "mm/dd/yyyy") & "' " _
        '    & lsCadena & " AND SUBSTRING(T.CCODTAB,1,2)='03' AND AD.CESTADO IN ('1'" & IIf(chkCancelados.value = 1, ",'0'", "") & ")" _
        '    & "     ORDER BY T.CCODTAB, P.CNOMPERS, AD.DFECCAL "

        Sql = " SELECT DISTINCT  P.cPersNombre,"
        Sql = Sql & " ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=1) ,'') AS cNudoci,"
        Sql = Sql & " ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=2) ,'') AS cNudotr,"
        Sql = Sql & " AP.cEvalCalif cCalGen, AD.cCtaCod, AD.dEval DFECCAL, CS.NSALDOCAP,"
        Sql = Sql & " CS.NDIASATRASO, SUBSTRING(AD.cCtaCod,9,1) cValor, cNomTab  = (case when SUBSTRING(AD.cCtaCod,9,1)='1' then 'Soles' else 'Dolares' end),"
        Sql = Sql & " AD.cEvalCalifDet cCalAud"
        Sql = Sql & " FROM ColocEvalCalif AP"
        Sql = Sql & " JOIN Persona P  ON P.cPersCod =AP.cPersCod"
        Sql = Sql & " JOIN ColocEvalCalifDetalle AD  ON AD.cPersCod=AP.cPersCod"
        Sql = Sql & " JOIN ColocacSaldo CS ON CS.cCtaCod=AD.cCtaCOd"
        Sql = Sql & " WHERE "
        Sql = Sql & " datediff(day,AD.dEval,'" & Format(lsFecha, "mm/dd/yyyy") & "')  = 0 and "
        Sql = Sql & " datediff(day,CS.dFecha,'" & Format(lsFecha, "mm/dd/yyyy") & "')  = 0  "
        Sql = Sql & lsCadena
        'SQL = SQL & " --AD.cEvalESTADO IN ('1'" & IIf(chkCancelados.value = 1, ",'0'", "") & ")"
        Sql = Sql & " ORDER BY SUBSTRING(AD.cCtaCod,9,1), P.cPersNombre, AD.dEval"

        Conecta.AbreConexion
            Set rs = Conecta.CargaRecordSet(Sql)
        Conecta.CierraConexion

        i = 0
        lsMoneda = ""
        lnTotal = rs.RecordCount
        m = 0
        RaiseEvent Progress(6, 10)
        If Not RSVacio(rs) Then
            Do While Not rs.EOF
                If lsMoneda <> Trim(rs!cValor) And i <> 0 Then
                    lsTotalSaldoCap = lsTotalSaldoCap & ":" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
                    Y2 = lnFila
                    CuadroExcel 1, Y1, 11, Y2
                    lnFila = lnFila + 1
                    xlHoja1.Cells(lnFila, 3) = "TOTAL"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=sum(" & lsTotalSaldoCap & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
                    CuadroExcel 1, lnFila, 11, lnFila
                    lsTotalSaldoCap = ""
                    m = 0
                End If
                If lsMoneda <> Trim(rs!cValor) Then
                    lnFila = lnFila + 2
                    xlHoja1.Cells(lnFila, 1) = Trim(rs!cNomTab) & " " & IIf(nChkMicro = 1, "-" & "Micro", "") _
                            & IIf(nChkComer = 1, "-" & "Comer", "") _
                            & IIf(nChkConsu = 1, "-" & "Consu", "")

                    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
                    Y1 = lnFila + 1
                    lsMoneda = Trim(rs!cValor)
                End If

                i = i + 1
                m = m + 1
                lnFila = lnFila + 1

                'SQL = "Select C.dFecultDesemb, C.cCodLinCred , LC.cDesLinCred, T.cValor, T.cNomTab, " _
                '    & "  FI.cDireccion as cDirFI, FI.cCodZon as cZonFI, ZFI.cDesZon as cDesZonFI " _
                '    & "  FROM " & gsServerConsol & "CREDITOCONSOL C " _
                '    & " JOIN " & gcCentralCom & "LineaCredito LC ON LC.CCODLINCRED = C.CCODLINCRED " _
                '    & " JOIN " & gcCentralCom & "TablaCod T     ON T.CVALOR = SUBSTRING(C.cCtaCod,6,1) " _
                '    & " LEFT JOIN " & gsServerConsol & "FUENTEINGRESOConsol FI ON FI.cNumFuente = C.cNumFuente " _
                '    & " LEFT JOIN " & gcCentralCom & "ZONAS ZFI ON FI.cCodZon = ZFI.cCodZon " _
                '    & " WHERE C.cCtaCod ='" & rs!cCtaCod & "' AND SUBSTRING(T.CCODTAB,1,2)='03' "

                Sql = " Select C.dFecApr, C.nSaldoCap , C.nDiasAtraso, C.dFecVig,"
                Sql = Sql & " LineaCred = (Select cDescripcion FROM ColocLineaCredito where cLineaCred = C.cLineaCred ),"
                Sql = Sql & "                    FI.cDireccion as cDirFI, FI.cCodZon as cZonFI, ZFI.cUbiGeoDescripcion as cDesZonFI"
                Sql = Sql & " FROM " & psServer & "CREDITOCONSOL C"
                Sql = Sql & " LEFT JOIN " & psServer & "FUENTEINGRESOConsol FI ON FI.cNumFuente = C.cNumFuente"
                Sql = Sql & " LEFT JOIN UBICACIONgEOGRAFICA ZFI ON FI.cCodZon = ZFI.cUbiGeoCod"
                Sql = Sql & " WHERE C.cCtaCod ='" & rs!cCtaCod & "' "

                'rsFuente.CursorLocation = adUseClient
                'rsFuente.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                'rsFuente.ActiveConnection = Nothing
                Conecta.AbreConexion
                    Set rsFuente = Conecta.CargaRecordSet(Sql)
                Conecta.CierraConexion

                xlHoja1.Cells(lnFila, 1) = "'" & Format(rsFuente!dFecApr, "dd/mm/yyyy")
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 2) = rs!cCtaCod
                xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 3) = PstaNombre(rs!cpersnombre, False)
                xlHoja1.Cells(lnFila, 4) = Trim(rsFuente!LineaCred)
                xlHoja1.Cells(lnFila, 5) = rs!nDiasAtraso
                xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 6) = Format(rs!nSaldoCap, "#,#0.00")
                If m = 1 Then
                    lsTotalSaldoCap = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
                End If
                xlHoja1.Cells(lnFila, 7) = rs!cCalAud
                xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 8) = rs!cCalGen
                xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 9) = IIf(Trim(rs!cNudoci) = "", Trim(rs!cNudotr), Trim(rs!cNudoci))
                xlHoja1.Cells(lnFila, 10) = Trim(rsFuente!cDirFI)
                xlHoja1.Cells(lnFila, 11) = Trim(rsFuente!cDesZonFI)

                rsFuente.Close
                Set rsFuente = Nothing

                'Me.barraProgreso.value = Int((I / lnTotal) * 100)
                'Me.EstatuBarICC.Panels(2).Text = "Reporte en " & Trim(rs!cNomTab) & " : " & Format((I / lnTotal * 100), "#0.00") & "%"
                rs.MoveNext
            Loop
            RaiseEvent Progress(8, 10)
            lsTotalSaldoCap = lsTotalSaldoCap & ":" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
            Y2 = lnFila
            CuadroExcel 1, Y1, 11, Y2
            lnFila = lnFila + 1
            xlHoja1.Cells(lnFila, 3) = "TOTAL"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=sum(" & lsTotalSaldoCap & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
            CuadroExcel 1, lnFila, 11, lnFila
            lsTotalSaldoCap = ""
            m = 0
        End If
        rs.Close
        Set rs = Nothing
    '    rsM.MoveNext
    'Loop
'End If
xlHoja1.PageSetup.Zoom = 75
RaiseEvent Progress(9, 10)
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
RaiseEvent Progress(10, 10)
MsgBox "Se genero el archivo en " & App.path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
lbExcel = False
RaiseEvent CloseProgress
End Sub

Public Sub Genera_Reporte178102(ByVal lsFecha As String, ByVal nChkMicro As Byte, _
ByVal nChkComer As Byte, ByVal nChkConsu As Byte, ByVal nChkHipo As Byte)
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim RsM As New ADODB.Recordset
Dim fs As New Scripting.FileSystemObject
Dim lnFila As Integer
Dim lnCol As Integer
Dim lnTotal As Integer, i As Integer
Dim Y1  As Integer, Y2 As Integer
Dim lbExisteHoja As Boolean
Dim lsNomHoja As String
Dim lsTotalSaldoCap As String
Dim lsCadena  As String
Dim lsMoneda As String * 1
Dim m As Integer
Dim ldFechaFinMes As Date
Dim rsFuente As New ADODB.Recordset
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim Conecta As DConecta
Set Conecta = New DConecta
Dim RCD As nRcdReportes
Dim Server As String
Set RCD = New nRcdReportes
    Server = RCD.GetServerConsol
Set RCD = Nothing

ldFechaFinMes = GetUltimaFechaCierre
RaiseEvent ShowProgress
    lsCadena = "  AND SUBSTRING(ARPD.cCtaCod,6,1) IN ("
        If nChkComer = 1 Then
             lsCadena = lsCadena & "'1',"
        End If
        If nChkMicro = 1 Then
            lsCadena = lsCadena & "'2',"
        End If
        If nChkConsu = 1 Then
             lsCadena = lsCadena & "'3',"
        End If
        If nChkHipo = 1 Then
            lsCadena = lsCadena & "'4',"
        End If
    If InStr(Len(lsCadena) - 2, lsCadena, "',", vbTextCompare) <> 0 Then
        lsCadena = Mid(lsCadena, 1, Len(lsCadena) - 1)
    End If
    RaiseEvent Progress(1, 10)
    lsCadena = lsCadena & ")"
    If nChkMicro = 0 And nChkComer = 0 And nChkConsu = 0 And nChkHipo = 0 Then
        lsCadena = ""
    End If
    lsNomHoja = Format(lsFecha, "ddmmyyyy") & "-" _
                & IIf(nChkMicro = 1, "Micro" & "-", "") _
                & IIf(nChkComer = 1, "Comer" & "-", "") _
                & IIf(nChkConsu = 1, "Cons" & "-", "") & IIf(nChkHipo = 1, "Hipo", "")

    lsNomHoja = IIf(Right(lsNomHoja, 1) = "-", Mid(lsNomHoja, 1, Len(lsNomHoja) - 1), lsNomHoja)

        lsArchivo = "CredRevEval" & Format(ldFechaFinMes, "mmyyyy") & ".XLS"
        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
            Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
        Else
            Set xlLibro = xlAplicacion.Workbooks.Add
        End If
        lbExcel = True
        lbExisteHoja = False
        For Each xlHoja1 In xlLibro.Worksheets
            If xlHoja1.Name = lsNomHoja Then
                xlHoja1.Activate
                xlHoja1.Range("A1", "AZ10000") = ""
                lbExisteHoja = True
                Exit For
            End If
        Next
        RaiseEvent Progress(3, 10)
        If lbExisteHoja = False Then
            Set xlHoja1 = xlLibro.Worksheets.Add
            xlHoja1.Name = lsNomHoja
        End If

        lnFila = 2
        xlHoja1.PageSetup.Zoom = 75
        xlHoja1.PageSetup.Orientation = xlLandscape
        xlAplicacion.Range("A1:R1000").Font.Size = 9

        xlHoja1.Range("A1").ColumnWidth = 12
        xlHoja1.Range("B1").ColumnWidth = 12
        xlHoja1.Range("C1").ColumnWidth = 40
        xlHoja1.Range("D1").ColumnWidth = 30
        xlHoja1.Range("E1").ColumnWidth = 6
        xlHoja1.Range("F1").ColumnWidth = 12
        xlHoja1.Range("G1").ColumnWidth = 12
        xlHoja1.Range("H1").ColumnWidth = 12
        xlHoja1.Range("I1").ColumnWidth = 12
        xlHoja1.Range("J1").ColumnWidth = 30
        xlHoja1.Range("K1").ColumnWidth = 15

        xlHoja1.Cells(lnFila, 1) = gsNomCmac
        xlHoja1.Cells(lnFila, 7) = gsNomAge & " -  Auditoria Interna"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
        lnFila = lnFila + 2
        xlHoja1.Cells(lnFila, 4) = "REPORTE DE REVISION DE CREDITOS EVALUADOS AL " & lsFecha
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
        lnFila = lnFila + 2
        Y1 = lnFila
        xlHoja1.Cells(lnFila, 1) = "Desembolso"
        xlHoja1.Cells(lnFila, 2) = "N° Credito"
        xlHoja1.Cells(lnFila, 3) = "Cliente"
        xlHoja1.Cells(lnFila, 4) = "Linea de Credito"
        'xlHoja1.Cells(lnFila, 5) = "Tipo de Credito"
        xlHoja1.Cells(lnFila, 5) = "D/M"
        xlHoja1.Cells(lnFila, 6) = "Saldo Capital"
        xlHoja1.Cells(lnFila, 7) = "Calific."
        xlHoja1.Cells(lnFila, 8) = "Rev.Calif."
        xlHoja1.Cells(lnFila, 9) = "Documento"
        xlHoja1.Cells(lnFila, 10) = "Direccion"
        xlHoja1.Cells(lnFila, 11) = "Zona"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 12)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 12)).HorizontalAlignment = xlCenter
        CuadroExcel 1, Y1, 11, Y1



        'SQL = "SELECT P.CNOMPERS, ISNULL(P.cNudoci,'') AS cNudoci, ISNULL(P.cNudotr,'') AS cNudotr, " _
        '    & " ARPD.cCalRev, ARPD.cCtaCod, ARPD.dFecCal,APD.cCalAud " _
        '    & " FROM  " & gcServerAudit & "AudRevPersDet ARPD " _
        '    & "      JOIN " & gcCentralPers & "Persona P ON P.CCODPERS =ARPD.CCODPERS " _
        '    & "      LEFT JOIN " & gcServerAudit & "AudPersDet APD  ON APD.cCodPers = ARPD.cCodPers " _
        '    & "           AND APD.cCtaCod = ARPD.cCtaCod  AND APD.dFecCal = ARPD.dFecCal " _
        '    & " WHERE ARPD.dFecCal = '" & Format(lsFecha, "mm/dd/yyyy") & "' " _
        '    & lsCadena _
        '    & " ORDER BY Substring(ARPD.cCtaCod,6,1),  P.cNomPers "

        Sql = " SELECT DISTINCT P.cPersNombre,"
        Sql = Sql & " ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=1) ,'') AS cNudoci, "
        Sql = Sql & " ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=2) ,'') AS cNudotr, "
        Sql = Sql & " ARPD.nEvaltipo cCalRev, ARPD.cCtaCod, ARPD.dEval as  dFecCal , APD.cEvalCalifDet cCalAud "
        Sql = Sql & " FROM  ColocEvalCalifDetalle ARPD "
        Sql = Sql & " JOIN Persona P ON P.cPersCod =ARPD.cPersCod "
        Sql = Sql & " LEFT JOIN ColocEvalCalifDetalle APD  ON APD.cPersCod = ARPD.cPersCod and "
        Sql = Sql & "                   APD.cCtaCod = ARPD.cCtaCod And APD.dEval = ARPD.dEval "
        Sql = Sql & "  WHERE ARPD.dEval = '" & Format(lsFecha, "mm/dd/yyyy") & "' "
        Sql = Sql & lsCadena
        'sql = sql & " ORDER BY Substring(ARPD.cCtaCod,9,1),  P.cPersNombre"


        Conecta.AbreConexion
        Set rs = Conecta.CargaRecordSet(Sql)
        Conecta.CierraConexion

        i = 0
        lsMoneda = ""
        lnTotal = rs.RecordCount
        m = 0
        RaiseEvent Progress(5, 10)
        If Not RSVacio(rs) Then
            Do While Not rs.EOF
                If lsMoneda <> Mid(rs!cCtaCod, 9, 1) And i <> 0 Then
                    lsTotalSaldoCap = lsTotalSaldoCap & ":" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
                    Y2 = lnFila
                    CuadroExcel 1, Y1, 11, Y2
                    lnFila = lnFila + 1
                    xlHoja1.Cells(lnFila, 3) = "TOTAL"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=sum(" & lsTotalSaldoCap & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
                    CuadroExcel 1, lnFila, 11, lnFila
                    lsTotalSaldoCap = ""
                    m = 0
                End If
                If lsMoneda <> Mid(rs!cCtaCod, 9, 1) Then
                    lnFila = lnFila + 2
                    xlHoja1.Cells(lnFila, 1) = IIf(Mid(rs!cCtaCod, 9, 1) = "1", "SOLES", "DOLARES") & " " & IIf(nChkMicro = 1, "-" & "Micro", "") _
                            & IIf(nChkComer = 1, "-" & "Comer", "") _
                            & IIf(nChkConsu = 1, "-" & "Consu", "")

                    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
                    Y1 = lnFila + 1
                    lsMoneda = Mid(rs!cCtaCod, 9, 1)
                End If

                i = i + 1
                m = m + 1
                lnFila = lnFila + 1

                'SQL = " Select C.nSaldoCap , C.nDiasAtraso, C.dfecUltDesemb, " _
                '    & " LineaCred = (Select cDesLinCred FROM dbcomunes.dbo.LineaCredito where cCodLinCred = C.cCodLinCred ), " _
                '    & " FI.cDireccion as cDirFI, FI.cCodZon as cZonFI, ZFI.cDesZon as cDesZonFI " _
                '    & " FROM " & gsServerConsol & "CREDITOCONSOL C " _
                '    & "      LEFT JOIN " & gsServerConsol & "FUENTEINGRESOConsol FI ON FI.cNumFuente = C.cNumFuente " _
                '    & "      LEFT JOIN " & gcCentralCom & "ZONAS ZFI ON FI.cCodZon = ZFI.cCodZon " _
                '    & " WHERE C.cCtaCod ='" & rs!cCtaCod & "' "

                Sql = " Select C.nSaldoCap , C.nDiasAtraso, C.dFecVig,"
                Sql = Sql & " LineaCred = (Select cDescripcion FROM ColocLineaCredito where cLineaCred = C.cLineaCred ),"
                Sql = Sql & "                    FI.cDireccion as cDirFI, FI.cCodZon as cZonFI, ZFI.cUbiGeoDescripcion as cDesZonFI"
                Sql = Sql & " FROM " & Server & "CREDITOCONSOL C"
                Sql = Sql & " LEFT JOIN " & Server & "FUENTEINGRESOConsol FI ON FI.cNumFuente = C.cNumFuente"
                Sql = Sql & " LEFT JOIN UBICACIONgEOGRAFICA ZFI ON FI.cCodZon = ZFI.cUbiGeoCod"
                Sql = Sql & " WHERE C.cCtaCod ='" & rs!cCtaCod & "' "


                Conecta.AbreConexion
                Set rsFuente = Conecta.CargaRecordSet(Sql)
                Conecta.CierraConexion
                If rsFuente.BOF And rsFuente.EOF Then
                    MsgBox "No existen creditos con evaluacion Revisada", vbInformation, "Aviso"
                End If

                xlHoja1.Cells(lnFila, 1) = Format(rsFuente!dFecVig, "dd/mm/yyyy")
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 2) = rs!cCtaCod
                xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 3) = PstaNombre(rs!cpersnombre, False)
                xlHoja1.Cells(lnFila, 4) = Trim(rsFuente!LineaCred)
                xlHoja1.Cells(lnFila, 5) = rsFuente!nDiasAtraso
                xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 6) = Format(rsFuente!nSaldoCap, "#,#0.00")
                If m = 1 Then
                    lsTotalSaldoCap = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
                End If
                xlHoja1.Cells(lnFila, 7) = rs!cCalAud
                xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 8) = rs!cCalRev

                'xlHoja1.Cells(lnFila, 8) = rs!cCalAud
                xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 9) = IIf(Trim(rs!cNudoci) = "", Trim(rs!cNudotr), Trim(rs!cNudoci))
                xlHoja1.Cells(lnFila, 10) = rsFuente!cDirFI
                xlHoja1.Cells(lnFila, 11) = rsFuente!cDesZonFI
                rsFuente.Close
                Set rsFuente = Nothing
                rs.MoveNext
            Loop
            RaiseEvent Progress(8, 10)
            lsTotalSaldoCap = lsTotalSaldoCap & ":" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
            Y2 = lnFila
            CuadroExcel 1, Y1, 11, Y2
            lnFila = lnFila + 1
            xlHoja1.Cells(lnFila, 3) = "TOTAL"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=sum(" & lsTotalSaldoCap & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
            CuadroExcel 1, lnFila, 11, lnFila
            lsTotalSaldoCap = ""
            m = 0
        End If
        rs.Close
        RaiseEvent Progress(9, 10)
        Set rs = Nothing

xlHoja1.PageSetup.Zoom = 75

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
lbExcel = False
RaiseEvent Progress(10, 10)
RaiseEvent CloseProgress
Set Conecta = Nothing
MsgBox App.path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
End Sub
Public Sub Genera_Reporte178103(ByVal lsFecha As String, ByVal nChkMicro As Byte, _
ByVal nChkComer As Byte, ByVal nChkConsu As Byte, ByVal psCodCmat As String, ByVal nChkHipo As Byte)
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim RsM As New ADODB.Recordset
Dim fs As New Scripting.FileSystemObject
Dim lnFila As Integer
Dim lnCol As Integer
Dim lnTotal As Integer, i As Integer
Dim Y1  As Integer, Y2 As Integer
Dim lbExisteHoja As Boolean
Dim lsNomHoja As String
Dim lsTotalSaldoCap As String
Dim lsCadena  As String
Dim lsMoneda As String * 1
Dim m As Integer
Dim ldFechaFinMes As Date
Dim lsArchivo  As String
Dim lbExcel As Boolean
Dim Conecta As DConecta
Set Conecta = New DConecta
ldFechaFinMes = GetUltimaFechaCierre

RaiseEvent ShowProgress
    lsCadena = "  AND SUBSTRING(AR.cCtaCod,6,1) IN ("

    If nChkComer = 1 Then
            lsCadena = lsCadena & "'1',"
    End If
    If nChkMicro = 1 Then
        lsCadena = lsCadena & "'2',"
    End If
    If nChkConsu = 1 Then
        lsCadena = lsCadena & "'3',"
    End If
    If nChkHipo = 1 Then
        lsCadena = lsCadena & "'4',"
    End If

    'If nChkComer = 1 Then
    '    If nChkComer = 1 Then
    '         lsCadena = lsCadena & "'1',"
    '    End If
    '    If nChkMicro = 1 Then
    '        lsCadena = lsCadena & "'2',"
    '    End If
    '    lsCadena = lsCadena & "'3'"
    'End If
    If InStr(Len(lsCadena) - 2, lsCadena, "',", vbTextCompare) <> 0 Then
        lsCadena = Mid(lsCadena, 1, Len(lsCadena) - 1)
    End If
    lsCadena = lsCadena & ")"
    If nChkMicro = 0 And nChkComer = 0 And nChkConsu = 0 Then
        lsCadena = ""
    End If
    lsNomHoja = Format(lsFecha, "ddmmyyyy") & "-" _
                & IIf(nChkMicro = 1, "Micro" & "-", "") _
                & IIf(nChkComer = 1, "Comer" & "-", "") _
                & IIf(nChkConsu = 1, "Consu", "")

    lsNomHoja = IIf(Right(lsNomHoja, 1) = "-", Mid(lsNomHoja, 1, Len(lsNomHoja) - 1), lsNomHoja)
    RaiseEvent Progress(1, 10)
        lsArchivo = "CredRevEvalVigentes" & Format(ldFechaFinMes, "mmyyyy") & ".XLS"
        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
            Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
        Else
            Set xlLibro = xlAplicacion.Workbooks.Add
        End If
        lbExcel = True
        lbExisteHoja = False
        For Each xlHoja1 In xlLibro.Worksheets
            If xlHoja1.Name = lsNomHoja Then
                xlHoja1.Activate
                xlHoja1.Range("A1", "AZ10000") = ""
                lbExisteHoja = True
                Exit For
            End If
        Next
        If lbExisteHoja = False Then
            Set xlHoja1 = xlLibro.Worksheets.Add
            xlHoja1.Name = lsNomHoja
        End If
       RaiseEvent Progress(3, 10)
        lnFila = 2
        xlHoja1.PageSetup.Zoom = 75
        xlHoja1.PageSetup.Orientation = xlLandscape
        xlAplicacion.Range("A1:R1000").Font.Size = 9

        xlHoja1.Range("A1").ColumnWidth = 12
        xlHoja1.Range("B1").ColumnWidth = 12
        xlHoja1.Range("C1").ColumnWidth = 40
        xlHoja1.Range("D1").ColumnWidth = 30
        xlHoja1.Range("E1").ColumnWidth = 6
        xlHoja1.Range("F1").ColumnWidth = 12
        xlHoja1.Range("G1").ColumnWidth = 12
        xlHoja1.Range("H1").ColumnWidth = 12
        xlHoja1.Range("I1").ColumnWidth = 12
        xlHoja1.Range("J1").ColumnWidth = 30
        xlHoja1.Range("K1").ColumnWidth = 15

        xlHoja1.Cells(lnFila, 1) = gsNomCmac
        xlHoja1.Cells(lnFila, 7) = " -  Auditoria Interna"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
        lnFila = lnFila + 2
        xlHoja1.Cells(lnFila, 4) = "REPORTE DE REVISION DE CREDITOS EVALUADOS VIGENTES AL " & lsFecha
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Merge True
        lnFila = lnFila + 2
        Y1 = lnFila
        xlHoja1.Cells(lnFila, 1) = "Desembolso"
        xlHoja1.Cells(lnFila, 2) = "N° Credito"
        xlHoja1.Cells(lnFila, 3) = "Cliente"
        xlHoja1.Cells(lnFila, 4) = "Linea de Credito"
        xlHoja1.Cells(lnFila, 5) = "D/M"
        xlHoja1.Cells(lnFila, 6) = "Saldo Capital"
        xlHoja1.Cells(lnFila, 7) = "Ult Calific."
        xlHoja1.Cells(lnFila, 8) = "Ult Rev.Calif."
        xlHoja1.Cells(lnFila, 9) = "Documento"
        'xlHoja1.Cells(lnFila, 10) = "Direccion"
        'xlHoja1.Cells(lnFila, 11) = "Zona"
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 12)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 12)).HorizontalAlignment = xlCenter
        'CuadroExcel 1, Y1, 11, Y1
        xlHoja1.Range(xlHoja1.Cells(Y1, 1), xlHoja1.Cells(Y1, 11)).Cells.Borders.LineStyle = xlOutside
        xlHoja1.Range(xlHoja1.Cells(Y1, 1), xlHoja1.Cells(Y1, 11)).Cells.Borders.LineStyle = xlInside

        RaiseEvent Progress(3, 10)


        Sql = " select P.cPersNombre,"
        Sql = Sql & " ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=1) ,'')  cNudoci,"
        Sql = Sql & " ISNULL((Select B.cPersIdNro from PersID B where B.cPerscod=P.cPersCod and B.cPersIDTpo=2),'') AS cNudotr,"
        Sql = Sql & " AR.cCtaCod,  AR.cUltCalRev, Ar.dFecUltCal,"
        Sql = Sql & " C.nSaldoCap, C.nDiasAtraso, C.dFecVig, ar.cPersCod,"
        Sql = Sql & " (Select cDescripcion FROM ColocLineaCredito where cLineaCred = C.cLineaCred ) as  LineaCred"
        Sql = Sql & " From " & ServerCons & "Creditoconsol c join"
        Sql = Sql & " ("
        Sql = Sql & " Select ar.dEval dFecUltCal, ar.cPersCod, ar.cCtaCod, ar.cEvalCalifDet cUltCalRev"
        Sql = Sql & " from ColocEvalCalifDetalle ar where ar.dEval = (Select max (ar1.dEval)  from ColocEvalCalifDetalle ar1 where ar1.cCtaCod = ar.cCtaCod)"
        Sql = Sql & " ) ar on c.cCtaCod = ar.cCtaCod"
        Sql = Sql & " join  Persona P ON P.cPersCod =AR.cPersCod"

        Select Case psCodCmat
        Case "112": Sql = Sql & " Where C.nPrdEstado in(2201,2101,2104,2106,2107,2020,2021,2022,2030,2031,2032)"
        Case "102": Sql = Sql & " Where C.nPrdEstado in(2201,2802,2803,2804,2020,2021,2022,2030,2031,2032)"
        End Select



        Sql = Sql & lsCadena
        Sql = Sql & " ORDER BY Substring(AR.cCtaCod,9,1), P.cPersNombre"


        Conecta.AbreConexion
        Set rs = Conecta.CargaRecordSet(Sql)
        Conecta.CierraConexion
        '------------------------------------

        RaiseEvent Progress(7, 10)
        i = 0
        lsMoneda = ""
        lnTotal = rs.RecordCount
        m = 0

        If Not RSVacio(rs) Then
            Do While Not rs.EOF
                If lsMoneda <> Mid(rs!cCtaCod, 9, 1) And i <> 0 Then
                    lsTotalSaldoCap = lsTotalSaldoCap & ":" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
                    Y2 = lnFila

                    'CuadroExcel 1, Y1, 11, Y2
                    xlHoja1.Range(xlHoja1.Cells(Y1, 1), xlHoja1.Cells(Y2 + 2, 11)).Cells.Borders.LineStyle = xlOutside
                    xlHoja1.Range(xlHoja1.Cells(Y1, 1), xlHoja1.Cells(Y2 + 2, 11)).Cells.Borders.LineStyle = xlInside

                    lnFila = lnFila + 1
                    xlHoja1.Cells(lnFila, 3) = "TOTAL"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=sum(" & lsTotalSaldoCap & ")"
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
                    'CuadroExcel 1, lnFila, 11, lnFila
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).Cells.Borders.LineStyle = xlOutside
                    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).Cells.Borders.LineStyle = xlInside

                    lsTotalSaldoCap = ""
                    m = 0
                End If
                If lsMoneda <> Mid(rs!cCtaCod, 9, 1) Then
                    lnFila = lnFila + 2
                    xlHoja1.Cells(lnFila, 1) = IIf(Mid(rs!cCtaCod, 9, 1) = "1", "SOLES", "DOLARES") & " " & IIf(nChkMicro = 1, "-" & "Micro", "") _
                            & IIf(nChkComer = 1, "-" & "Comer", "") _
                            & IIf(nChkConsu = 1, "-" & "Consu", "")

                    xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True
                    Y1 = lnFila + 1
                    lsMoneda = Mid(rs!cCtaCod, 9, 1)
                End If

                i = i + 1
                m = m + 1
                lnFila = lnFila + 1
                xlHoja1.Cells(lnFila, 1) = Format(rs!dFecVig, "dd/mm/yyyy")
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 2) = rs!cCtaCod
                xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 3) = PstaNombre(rs!cpersnombre, False)
                xlHoja1.Cells(lnFila, 4) = Trim(rs!LineaCred)
                xlHoja1.Cells(lnFila, 5) = rs!nDiasAtraso
                xlHoja1.Range(xlHoja1.Cells(lnFila, 5), xlHoja1.Cells(lnFila, 5)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 6) = Format(rs!nSaldoCap, "#,#0.00")
                If m = 1 Then
                    lsTotalSaldoCap = xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
                End If
                xlHoja1.Cells(lnFila, 7) = rs!cUltCalRev
                xlHoja1.Range(xlHoja1.Cells(lnFila, 7), xlHoja1.Cells(lnFila, 7)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 8) = Format(rs!dFecUltCal, "dd/mm/yyyy")

                xlHoja1.Range(xlHoja1.Cells(lnFila, 8), xlHoja1.Cells(lnFila, 8)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 9) = IIf(Trim(rs!cNudoci) = "", Trim(rs!cNudotr), Trim(rs!cNudoci))
                'xlHoja1.Cells(lnFila, 10) = rs!cDirFI
                'xlHoja1.Cells(lnFila, 11) = rs!cDesZonFI
                rs.MoveNext
            Loop
            lsTotalSaldoCap = lsTotalSaldoCap & ":" & xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Address(False, False)
            Y2 = lnFila

            CuadroExcel 1, Y1, 11, Y2

            xlHoja1.Range(xlHoja1.Cells(Y1, 1), xlHoja1.Cells(Y2 + 2, 11)).Cells.Borders.LineStyle = xlOutside
            xlHoja1.Range(xlHoja1.Cells(Y1, 1), xlHoja1.Cells(Y2 + 2, 11)).Cells.Borders.LineStyle = xlInside

            lnFila = lnFila + 1
            xlHoja1.Cells(lnFila, 3) = "TOTAL"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 6), xlHoja1.Cells(lnFila, 6)).Formula = "=sum(" & lsTotalSaldoCap & ")"
            xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 10)).Font.Bold = True

            CuadroExcel 1, lnFila, 11, lnFila
            xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).Cells.Borders.LineStyle = xlOutside
            xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 11)).Cells.Borders.LineStyle = xlInside
            lsTotalSaldoCap = ""
            m = 0
        End If
        rs.Close
        Set rs = Nothing
        RaiseEvent Progress(9, 10)
xlHoja1.PageSetup.Zoom = 75
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
RaiseEvent Progress(10, 10)
MsgBox "Se genero el Reporte en " & App.path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
Set xlLibro = Nothing
Set xlHoja1 = Nothing
lbExcel = False
RaiseEvent CloseProgress
End Sub
Public Function Genera_Reporte178201(ByVal psServer As String, ByVal sMoneda As String, ByVal sSigno As String, ByVal nMonto As Currency, _
  ByVal bTipo As Boolean, ByVal bEstado As Boolean, ByVal sCodCMAC As String, ByVal psNomCmac As String, ByVal psNomAge As String) As String

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset

Dim RSAgencias As New ADODB.Recordset
Dim RSRecursos As New ADODB.Recordset

Dim Total As Currency
Dim Total_Prov As Currency

Dim Linea As Long
Dim Conta As Long
Dim lsCadena As String
Dim lnPagina As Integer
Dim lsCredito As String
Dim Matriz() As Variant
Dim MatrizTotal() As Variant

Dim J As Integer
Dim FechaFinMes As Date
Dim nTipo As String
Dim nEstado As String
FechaFinMes = GetUltimaFechaCierre
Set RSAgencias = Get_Agencias
Set RSRecursos = Get_Financiamiento


lnPagina = 1
If bTipo Then
    nTipo = "0" 'Comercial
    lsCredito = "Comercial-"
Else
    nTipo = "1" 'Micro
    lsCredito = "MicroEmpresa-"
End If

If bEstado Then
    nEstado = 0 'Normal
    lsCredito = lsCredito & "Normal"
Else
    nEstado = 1 'Refinanciado
    lsCredito = lsCredito & "Refinanciado"
End If


lsCadena = Cabecera_Pro(psNomCmac, psNomAge, gdFecSis, lnPagina, FechaFinMes, lsCredito, sMoneda, 178201, sSigno & " " & nMonto)
Linea = 8
Conta = 0

Dim i As Integer
 i = 1
RaiseEvent ShowProgress

While Not RSRecursos.EOF
    ReDim MatrizTotal(1 To 18)
    ReDim Matriz(1 To RSAgencias.RecordCount, 1 To 18)
    Linea = Linea + 1
    lsCadena = lsCadena & RSRecursos(1) & Chr(10)

    If Linea > 61 Then
        lsCadena = lsCadena & Chr(12) & Chr(10)
        lnPagina = lnPagina + 1
        Linea = 8
        lsCadena = lsCadena & Cabecera_Pro(psNomCmac, psNomAge, gdFecSis, lnPagina, FechaFinMes, lsCredito, sMoneda, 178201, sSigno & " " & nMonto)
    End If
    Set rs = ReporteGeneralProvisionesPyme_Comercial(psServer, nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), , , sCodCMAC)
    i = 1
    While Not rs.EOF
        Matriz(i, 1) = rs!cAgeDescripcion
        Matriz(i, 2) = rs!nTotal
        Matriz(i, 3) = rs!sk_Total
        MatrizTotal(2) = MatrizTotal(2) + CInt(rs!nTotal)
        MatrizTotal(3) = MatrizTotal(3) + CCur(rs!sk_Total)
        i = i + 1
        rs.MoveNext
    Wend

    'Normal
    Set rs = ReporteGeneralProvisionesPyme_Comercial(psServer, nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "0", sCodCMAC)
    i = 1
    While Not rs.EOF
        Matriz(i, 4) = rs!nTotal
        Matriz(i, 5) = rs!sk_Total
        Matriz(i, 6) = rs!Prov
        MatrizTotal(4) = MatrizTotal(4) + CInt(rs!nTotal)
        MatrizTotal(5) = MatrizTotal(5) + CCur(rs!sk_Total)
        MatrizTotal(6) = MatrizTotal(6) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Prob. Potencial
    Set rs = ReporteGeneralProvisionesPyme_Comercial(psServer, nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "1", sCodCMAC)
    i = 1
    While Not rs.EOF
        Matriz(i, 7) = rs!nTotal
        Matriz(i, 8) = rs!sk_Total
        Matriz(i, 9) = rs!Prov
        MatrizTotal(7) = MatrizTotal(7) + CInt(rs!nTotal)
        MatrizTotal(8) = MatrizTotal(8) + CCur(rs!sk_Total)
        MatrizTotal(9) = MatrizTotal(9) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Diferente
    Set rs = ReporteGeneralProvisionesPyme_Comercial(psServer, nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "2", sCodCMAC)
    i = 1
    While Not rs.EOF
        Matriz(i, 10) = rs!nTotal
        Matriz(i, 11) = rs!sk_Total
        Matriz(i, 12) = rs!Prov
        MatrizTotal(10) = MatrizTotal(10) + CInt(rs!nTotal)
        MatrizTotal(11) = MatrizTotal(11) + CCur(rs!sk_Total)
        MatrizTotal(12) = MatrizTotal(12) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Dudoso
    Set rs = ReporteGeneralProvisionesPyme_Comercial(psServer, nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "3", sCodCMAC)
    i = 1
    While Not rs.EOF
        Matriz(i, 13) = rs!nTotal
        Matriz(i, 14) = rs!sk_Total
        Matriz(i, 15) = rs!Prov
        MatrizTotal(13) = MatrizTotal(13) + CInt(rs!nTotal)
        MatrizTotal(14) = MatrizTotal(14) + CCur(rs!sk_Total)
        MatrizTotal(15) = MatrizTotal(15) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Perdida
    Set rs = ReporteGeneralProvisionesPyme_Comercial(psServer, nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "4", sCodCMAC)
    i = 1
    While Not rs.EOF
        Matriz(i, 16) = rs!nTotal
        Matriz(i, 17) = rs!sk_Total
        Matriz(i, 18) = rs!Prov
        MatrizTotal(16) = MatrizTotal(16) + CInt(rs!nTotal)
        MatrizTotal(17) = MatrizTotal(17) + CCur(rs!sk_Total)
        MatrizTotal(18) = MatrizTotal(18) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend
    Total_Prov = 0

    For i = 1 To 14
        Linea = Linea + 1
        Total_Prov = 0
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 1), 15) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 2), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 3), 10, 2) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 4), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 5), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 6), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 6))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 7), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 8), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 9), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 9))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 10), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 11), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 12), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 12))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 13), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 14), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 15), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 15))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 16), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 17), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 18), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 18))
        lsCadena = lsCadena & Space(2) & ImpreFormat(Total_Prov, 8, , True) & Chr(10)
    Next
    lsCadena = lsCadena & Chr(10)
    lsCadena = lsCadena & Space(20) & ImpreFormat(MatrizTotal(2), 4, 0) & " "
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(3), 10, 2) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(4), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(5), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(6), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(7), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(8), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(9), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(10), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(11), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(12), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(13), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(14), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(15), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(16), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(17), 10, 2, True) & Space(1)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(18), 10, 2, True) & Space(2) & Chr(10)
    lsCadena = lsCadena & Chr(10) & Chr(10) & Chr(10)
    Linea = Linea + 5
    RaiseEvent Progress(RSRecursos.Bookmark, RSRecursos.RecordCount)
    RSRecursos.MoveNext
Wend
RaiseEvent CloseProgress
Genera_Reporte178201 = lsCadena
End Function

Public Function Genera_Reporte178202(ByVal sMoneda As String, ByVal sSigno As String, ByVal nMonto As Currency, _
  ByVal bTipo As Boolean, ByVal bEstado As Boolean, ByVal psNomCmac As String, ByVal psNomAge As String) As String

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset

Dim RSAgencias As New ADODB.Recordset
Dim RSRecursos As New ADODB.Recordset

Dim nMoneda As String


Dim Total As Currency
Dim Total_Prov As Currency

Dim Linea As Long
Dim Conta As Long
Dim lsCadena As String
Dim lnPagina As Integer
Dim lsCredito As String
Dim Matriz() As Variant
Dim MatrizTotal() As Variant

Dim J As Integer
Dim nTipo As String
Dim nEstado As String
Dim FechaFinMes As Date
FechaFinMes = GetUltimaFechaCierre
Set RSAgencias = Get_Agencias
Set RSRecursos = Get_Financiamiento


lnPagina = 1

If bTipo Then
    nTipo = "0" 'Comercial
    lsCredito = "Comercial-"
Else
    nTipo = "1" 'Micro
    lsCredito = "MicroEmpresa-"
End If


If bEstado Then
    nEstado = 0 'Normal
    lsCredito = "Cred. Personales-Normal"
Else
    nEstado = 1 'Refinanciado
    lsCredito = "Cred. Personales-Refinanciado"
End If

'If OptMoneda(1).value = 0 Then
'    nMoneda = gMonedaNacional  ' Soles
'Else
'    nMoneda = gMonedaExtranjera ' Dolares
'End If


lsCadena = Cabecera_Pro(psNomCmac, psNomAge, gdFecSis, lnPagina, FechaFinMes, lsCredito, sMoneda, 178202, sSigno & " " & nMonto)
Linea = 8
Conta = 0
ReDim MatrizTotal(1 To 18)
ReDim Matriz(1 To RSAgencias.RecordCount, 1 To 18)
Dim i As Integer
i = 1
RaiseEvent ShowProgress
While Not RSRecursos.EOF
    ReDim MatrizTotal(1 To 18)
    ReDim Matriz(1 To RSAgencias.RecordCount, 1 To 18)

    Linea = Linea + 1
    lsCadena = lsCadena & RSRecursos(1) & Chr(10)

    If Linea > 61 Then
        lsCadena = lsCadena & Chr(12) & Chr(10)
        lnPagina = lnPagina + 1
        Linea = 8
        lsCadena = lsCadena & Cabecera_Pro(psNomCmac, psNomAge, gdFecSis, lnPagina, FechaFinMes, lsCredito, sMoneda, 178201, sSigno & " " & nMonto)
    End If
    Set rs = ReporteGeneralProvisionesPersonales(nEstado, sMoneda, sSigno, nMonto, RSRecursos(0))

    i = 1
    While Not rs.EOF
        Matriz(i, 1) = rs!cAgeDescripcion
        Matriz(i, 2) = rs!nTotal
        Matriz(i, 3) = rs!sk_Total
        MatrizTotal(2) = MatrizTotal(2) + CInt(rs!nTotal)
        MatrizTotal(3) = MatrizTotal(3) + CCur(rs!sk_Total)
        i = i + 1
        rs.MoveNext
    Wend

    'Normal
    Set rs = ReporteGeneralProvisionesPersonales(nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "0")
    i = 1
    While Not rs.EOF
        Matriz(i, 4) = rs!nTotal
        Matriz(i, 5) = rs!sk_Total
        Matriz(i, 6) = rs!Prov
        MatrizTotal(4) = MatrizTotal(4) + CInt(rs!nTotal)
        MatrizTotal(5) = MatrizTotal(5) + CCur(rs!sk_Total)
        MatrizTotal(6) = MatrizTotal(6) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Prob. Potencial
    Set rs = ReporteGeneralProvisionesPersonales(nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "1")
    i = 1
    While Not rs.EOF
        Matriz(i, 7) = rs!nTotal
        Matriz(i, 8) = rs!sk_Total
        Matriz(i, 9) = rs!Prov
        MatrizTotal(7) = MatrizTotal(7) + CInt(rs!nTotal)
        MatrizTotal(8) = MatrizTotal(8) + CCur(rs!sk_Total)
        MatrizTotal(9) = MatrizTotal(9) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Diferente
    Set rs = ReporteGeneralProvisionesPersonales(nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "2")
    i = 1
    While Not rs.EOF
        Matriz(i, 10) = rs!nTotal
        Matriz(i, 11) = rs!sk_Total
        Matriz(i, 12) = rs!Prov
        MatrizTotal(10) = MatrizTotal(10) + CInt(rs!nTotal)
        MatrizTotal(11) = MatrizTotal(11) + CCur(rs!sk_Total)
        MatrizTotal(12) = MatrizTotal(12) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Dudoso
    Set rs = ReporteGeneralProvisionesPersonales(nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "3")
    i = 1
    While Not rs.EOF
        Matriz(i, 13) = rs!nTotal
        Matriz(i, 14) = rs!sk_Total
        Matriz(i, 15) = rs!Prov
        MatrizTotal(13) = MatrizTotal(13) + CInt(rs!nTotal)
        MatrizTotal(14) = MatrizTotal(14) + CCur(rs!sk_Total)
        MatrizTotal(15) = MatrizTotal(15) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Perdida
    Set rs = ReporteGeneralProvisionesPersonales(nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "4")
    i = 1
    While Not rs.EOF
        Matriz(i, 16) = rs!nTotal
        Matriz(i, 17) = rs!sk_Total
        Matriz(i, 18) = rs!Prov
        MatrizTotal(16) = MatrizTotal(16) + CInt(rs!nTotal)
        MatrizTotal(17) = MatrizTotal(17) + CCur(rs!sk_Total)
        MatrizTotal(18) = MatrizTotal(18) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend
    Total_Prov = 0

    For i = 1 To 14
        Linea = Linea + 1
        Total_Prov = 0
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 1), 15) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 2), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 3), 10, 2) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 4), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 5), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 6), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 6))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 7), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 8), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 9), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 9))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 10), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 11), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 12), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 12))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 13), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 14), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 15), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 15))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 16), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 17), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 18), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 18))
        lsCadena = lsCadena & Space(2) & Total_Prov & Chr(10)
    Next
    lsCadena = lsCadena & Chr(10)
    lsCadena = lsCadena & Space(20) & ImpreFormat(MatrizTotal(2), 4, 0) & " "
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(3), 10, 2) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(4), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(5), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(6), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(7), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(8), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(9), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(10), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(11), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(12), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(13), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(14), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(15), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(16), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(17), 10, 2, True) & Space(1)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(18), 10, 2, True) & Space(2) & Chr(10)
    lsCadena = lsCadena & Chr(10) & Chr(10) & Chr(10)
    Linea = Linea + 5
    RaiseEvent Progress(RSRecursos.Bookmark, RSRecursos.RecordCount)
    RSRecursos.MoveNext
Wend

'Set loPrevio = New Previo.clsPrevio
'    loPrevio.Show lsCadena, "PROVISIONES", True
'Set loPrevio = Nothing
RaiseEvent CloseProgress
Genera_Reporte178202 = lsCadena
End Function

Public Function Genera_Reporte178203(ByVal sMoneda As String, ByVal sSigno As String, ByVal nMonto As Currency, _
  ByVal bTipo As Boolean, ByVal bEstado As Boolean, ByVal psCodCmac As String, ByVal psNomCmac As String, _
  ByVal psNomAge As String) As String

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset

Dim RSAgencias As New ADODB.Recordset
Dim RSRecursos As New ADODB.Recordset

Dim loPrevio As Previo.clsPrevio

Dim Total As Currency
Dim Total_Prov As Currency
Dim Linea As Long
Dim Conta As Long
Dim lsCadena As String
Dim lnPagina As Integer
Dim lsCredito As String
Dim Matriz() As Variant
Dim MatrizTotal() As Variant

Dim J As Integer
Dim nTipo  As String
Dim nEstado As String
Dim FechaFinMes As Date

FechaFinMes = GetUltimaFechaCierre

Set RSAgencias = Get_Agencias
Set RSRecursos = Get_Financiamiento
RaiseEvent ShowProgress
lnPagina = 1
If bTipo Then
    nTipo = "0" 'Comercial
    lsCredito = "Hipotecario-"
Else
    nTipo = "1 " 'Micro
    lsCredito = "Hipotecario-"
End If

If bEstado Then
    nEstado = 0 'Normal
    lsCredito = lsCredito & "Normal"
Else
    nEstado = 1 'Refinanciado
    lsCredito = lsCredito & "Refinanciado"
End If

lsCadena = Cabecera_Pro(psNomCmac, psNomAge, gdFecSis, lnPagina, FechaFinMes, lsCredito, sMoneda, 178203, sSigno & " " & nMonto)
Linea = 8
Conta = 0
ReDim MatrizTotal(1 To 15)
ReDim Matriz(1 To RSAgencias.RecordCount, 1 To 15)
Dim i As Integer
i = 1

While Not RSRecursos.EOF

    Linea = Linea + 1
    lsCadena = lsCadena & RSRecursos(1) & Chr(10)

    If Linea > 61 Then
        lsCadena = lsCadena & Chr(12) & Chr(10)
        lnPagina = lnPagina + 1
        Linea = 8
        lsCadena = lsCadena & Cabecera_Pro(psNomCmac, psNomAge, gdFecSis, lnPagina, FechaFinMes, lsCredito, sMoneda, 178203, sSigno & " " & nMonto)
    End If
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), , , psCodCmac)
    i = 1
    While Not rs.EOF
        Matriz(i, 1) = rs!cAgeDescripcion
        Matriz(i, 2) = rs!nTotal
        Matriz(i, 3) = rs!sk_Total
        MatrizTotal(2) = CInt(rs!nTotal)
        MatrizTotal(3) = CCur(rs!sk_Total)
        i = i + 1
        rs.MoveNext
    Wend

    'Normal
    'Set Rs = Co.ReporteGeneralProvisionesPyme_Comercial(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "0")
    'i = 1
    'While Not Rs.EOF
    '    Matriz(i, 4) = Rs!nTotal
    '    Matriz(i, 5) = Rs!sk_Total
    '    Matriz(i, 6) = Rs!Prov
    '    MatrizTotal(4) = MatrizTotal(4) + CInt(Rs!nTotal)
    '    MatrizTotal(5) = MatrizTotal(5) + CCur(Rs!sk_Total)
    '    MatrizTotal(6) = MatrizTotal(6) + CCur(Rs!Prov)
    '    i = i + 1
    '    Rs.MoveNext
    'Wend

    'Prob. Potencial
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "1", psCodCmac)
    i = 1
    While Not rs.EOF
        Matriz(i, 4) = rs!nTotal
        Matriz(i, 5) = rs!sk_Total
        Matriz(i, 6) = rs!Prov
        MatrizTotal(4) = MatrizTotal(4) + CInt(rs!nTotal)
        MatrizTotal(5) = MatrizTotal(5) + CCur(rs!sk_Total)
        MatrizTotal(6) = MatrizTotal(6) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Diferente
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "2", psCodCmac)
    i = 1
    While Not rs.EOF
        Matriz(i, 7) = rs!nTotal
        Matriz(i, 8) = rs!sk_Total
        Matriz(i, 9) = rs!Prov
        MatrizTotal(7) = MatrizTotal(7) + CInt(rs!nTotal)
        MatrizTotal(8) = MatrizTotal(8) + CCur(rs!sk_Total)
        MatrizTotal(9) = MatrizTotal(9) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Dudoso
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "3", psCodCmac)
    i = 1
    While Not rs.EOF
        Matriz(i, 10) = rs!nTotal
        Matriz(i, 11) = rs!sk_Total
        Matriz(i, 12) = rs!Prov
        MatrizTotal(10) = MatrizTotal(10) + CInt(rs!nTotal)
        MatrizTotal(11) = MatrizTotal(11) + CCur(rs!sk_Total)
        MatrizTotal(12) = MatrizTotal(12) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend

    'Perdida
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "4", psCodCmac)
    i = 1
    While Not rs.EOF
        Matriz(i, 13) = rs!nTotal
        Matriz(i, 14) = rs!sk_Total
        Matriz(i, 15) = rs!Prov
        MatrizTotal(13) = MatrizTotal(13) + CInt(rs!nTotal)
        MatrizTotal(14) = MatrizTotal(14) + CCur(rs!sk_Total)
        MatrizTotal(15) = MatrizTotal(15) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend
    Total_Prov = 0

    For i = 1 To 14
        Linea = Linea + 1
        Total_Prov = 0
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 1), 15) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 2), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 3), 10, 2) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 4), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 5), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 6), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 6))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 7), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 8), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 9), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 9))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 10), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 11), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 12), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 12))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 13), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 14), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 15), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 15))
        'lsCadena = lsCadena & ImpreFormat(Matriz(i, 16), 4, 0) & Space(2)
        'lsCadena = lsCadena & ImpreFormat(Matriz(i, 17), 10, 2, True) & Space(2)
        'lsCadena = lsCadena & ImpreFormat(Matriz(i, 18), 10, 2, True) & Space(2)
        'Total_Prov = Total_Prov + CCur(Matriz(i, 18))
        lsCadena = lsCadena & Space(2) & Total_Prov & Chr(10)
    Next
    lsCadena = lsCadena & Chr(10)
    lsCadena = lsCadena & Space(20) & ImpreFormat(MatrizTotal(2), 4, 0) & " "
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(3), 10, 2) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(4), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(5), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(6), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(7), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(8), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(9), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(10), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(11), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(12), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(13), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(14), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(15), 10, 2, True) & Space(2)
    'lsCadena = lsCadena & ImpreFormat(MatrizTotal(16), 4, 0) & Space(2)
    'lsCadena = lsCadena & ImpreFormat(MatrizTotal(17), 10, 2, True) & Space(1)
    'lsCadena = lsCadena & ImpreFormat(MatrizTotal(18), 10, 2, True) & Space(2) & Chr(10)
    lsCadena = lsCadena & Chr(10) & Chr(10) & Chr(10)
    Linea = Linea + 5
    RaiseEvent Progress(RSRecursos.Bookmark, RSRecursos.RecordCount)
    RSRecursos.MoveNext
Wend
RaiseEvent CloseProgress
Genera_Reporte178203 = lsCadena
End Function


'
Public Function Genera_Reporte178204(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal pdFecSis As Date) As String
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset

Dim RSAgencias As New ADODB.Recordset
'Dim RSRecursos As New ADODB.Recordset
'Dim loPrevio As Previo.clsPrevio

Dim nMoneda As String


Dim Total As Currency
Dim Total_Prov As Currency



Dim Linea As Long
Dim Conta As Long
Dim lsCadena As String
Dim lnPagina As Integer
Dim lsCredito As String
Dim Matriz() As Variant
Dim MatrizTotal() As Variant

Dim J As Integer
Dim FechaFinMes As Date

RaiseEvent ShowProgress
FechaFinMes = GetUltimaFechaCierre
Set RSAgencias = Get_Agencias
lnPagina = 1
lsCredito = "Pignoraticio"
lsCadena = Cabecera_Pro(psNomCmac, psNomAge, pdFecSis, lnPagina, FechaFinMes, lsCredito, "2", 178204, "")
Linea = 8
Conta = 0
ReDim MatrizTotal(1 To 18)
ReDim Matriz(1 To RSAgencias.RecordCount, 1 To 18)
Dim i As Integer
i = 1

'While Not RSRecursos.EOF
    Linea = Linea + 1
    lsCadena = lsCadena & "Recursos Propios" & Chr(10)
  RaiseEvent Progress(1, 11)
    If Linea > 61 Then
        lsCadena = lsCadena & Chr(12) & Chr(10)
        lnPagina = lnPagina + 1
        Linea = 8
        lsCadena = lsCadena & Cabecera_Pro(psNomCmac, psNomAge, pdFecSis, lnPagina, FechaFinMes, lsCredito, "2", 178204, "")
    End If
    Set rs = ReporteGeneralProvisionesPignoraticio()
       i = 1
    While Not rs.EOF
        Matriz(i, 1) = rs!cAgeDescripcion
        Matriz(i, 2) = rs!nTotal
        Matriz(i, 3) = rs!sk_Total
        MatrizTotal(2) = MatrizTotal(2) + CInt(rs!nTotal)
        MatrizTotal(3) = MatrizTotal(3) + CCur(rs!sk_Total)
        i = i + 1
        rs.MoveNext
    Wend
    RaiseEvent Progress(3, 11)
    'Normal
    Set rs = ReporteGeneralProvisionesPignoraticio(False, "0")
    i = 1
    While Not rs.EOF
        Matriz(i, 4) = rs!nTotal
        Matriz(i, 5) = rs!sk_Total
        Matriz(i, 6) = rs!Prov
        MatrizTotal(4) = MatrizTotal(4) + CInt(rs!nTotal)
        MatrizTotal(5) = MatrizTotal(5) + CCur(rs!sk_Total)
        MatrizTotal(6) = MatrizTotal(6) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend
    RaiseEvent Progress(4, 11)
    'Prob. Potencial
    Set rs = ReporteGeneralProvisionesPignoraticio(False, "1")
    i = 1
    While Not rs.EOF
        Matriz(i, 7) = rs!nTotal
        Matriz(i, 8) = rs!sk_Total
        Matriz(i, 9) = rs!Prov
        MatrizTotal(7) = MatrizTotal(7) + CInt(rs!nTotal)
        MatrizTotal(8) = MatrizTotal(8) + CCur(rs!sk_Total)
        MatrizTotal(9) = MatrizTotal(9) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend
    RaiseEvent Progress(5, 11)
    'Diferente
    Set rs = ReporteGeneralProvisionesPignoraticio(False, "2")
    i = 1
    While Not rs.EOF
        Matriz(i, 10) = rs!nTotal
        Matriz(i, 11) = rs!sk_Total
        Matriz(i, 12) = rs!Prov
        MatrizTotal(10) = MatrizTotal(10) + CInt(rs!nTotal)
        MatrizTotal(11) = MatrizTotal(11) + CCur(rs!sk_Total)
        MatrizTotal(12) = MatrizTotal(12) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend
    RaiseEvent Progress(7, 11)
    'Dudoso
    Set rs = ReporteGeneralProvisionesPignoraticio(False, "3")
    i = 1
    While Not rs.EOF
        Matriz(i, 13) = rs!nTotal
        Matriz(i, 14) = rs!sk_Total
        Matriz(i, 15) = rs!Prov
        MatrizTotal(13) = MatrizTotal(13) + CInt(rs!nTotal)
        MatrizTotal(14) = MatrizTotal(14) + CCur(rs!sk_Total)
        MatrizTotal(15) = MatrizTotal(15) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend
RaiseEvent Progress(8, 11)
    'Perdida
    Set rs = ReporteGeneralProvisionesPignoraticio(False, "4")
    i = 1
    While Not rs.EOF
        Matriz(i, 16) = rs!nTotal
        Matriz(i, 17) = rs!sk_Total
        Matriz(i, 18) = rs!Prov
        MatrizTotal(16) = MatrizTotal(16) + CInt(rs!nTotal)
        MatrizTotal(17) = MatrizTotal(17) + CCur(rs!sk_Total)
        MatrizTotal(18) = MatrizTotal(18) + CCur(rs!Prov)
        i = i + 1
        rs.MoveNext
    Wend
    Total_Prov = 0
    RaiseEvent Progress(9, 11)
    For i = 1 To 14
        Linea = Linea + 1
        Total_Prov = 0
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 1), 15) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 2), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 3), 10, 2) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 4), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 5), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 6), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 6))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 7), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 8), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 9), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 9))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 10), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 11), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 12), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 12))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 13), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 14), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 15), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 15))
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 16), 4, 0) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 17), 10, 2, True) & Space(2)
        lsCadena = lsCadena & ImpreFormat(Matriz(i, 18), 10, 2, True) & Space(2)
        Total_Prov = Total_Prov + CCur(Matriz(i, 18))
        lsCadena = lsCadena & Space(2) & Total_Prov & Chr(10)
    Next
    RaiseEvent Progress(10, 11)
    lsCadena = lsCadena & Chr(10)
    lsCadena = lsCadena & Space(20) & ImpreFormat(MatrizTotal(2), 4, 0) & " "
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(3), 10, 2) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(4), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(5), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(6), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(7), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(8), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(9), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(10), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(11), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(12), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(13), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(14), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(15), 10, 2, True) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(16), 4, 0) & Space(2)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(17), 10, 2, True) & Space(1)
    lsCadena = lsCadena & ImpreFormat(MatrizTotal(18), 10, 2, True) & Space(2) & Chr(10)
    lsCadena = lsCadena & Chr(10) & Chr(10) & Chr(10)
    Linea = Linea + 5
'    RSRecursos.MoveNext
'Wend
RaiseEvent Progress(11, 11)
RaiseEvent CloseProgress
Genera_Reporte178204 = lsCadena
End Function

Private Sub ImpHojaWProvisionCab(pnFila As Integer, psParteCab As String, ByVal dFecha As Date, Optional pbConsumo As Boolean = False)
Dim J As Integer
Dim lsNombre As String


    xlHoja1.Cells(pnFila, 1) = gsNomCmac
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 2) = "INFORME DE CLASIFICACION DE LA CARTERA DE CREDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 7) = "AL " & dFecha
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    If psParteCab = "1" Then
        xlHoja1.Cells(pnFila, 6) = "(  MONEDA NACIONAL  ) "
    Else
        xlHoja1.Cells(pnFila, 6) = "( MONEDA EXTRANJERA ) "
    End If
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 1) = " CREDITOS "
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Cells(pnFila, 2) = "TOTAL CARTERA"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 2), xlHoja1.Cells(pnFila, 3)).Merge True
    xlHoja1.Cells(pnFila + 1, 2) = "Nro"
    xlHoja1.Cells(pnFila + 1, 3) = "Monto"
    xlHoja1.Cells(pnFila, 4) = "NORMALES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 4), xlHoja1.Cells(pnFila, 5)).Merge True
    xlHoja1.Cells(pnFila + 1, 4) = "Nro"
    xlHoja1.Cells(pnFila + 1, 5) = "Monto"
    xlHoja1.Cells(pnFila, 6) = "POTENCIALES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 6), xlHoja1.Cells(pnFila, 7)).Merge True
    xlHoja1.Cells(pnFila + 1, 6) = "Nro"
    xlHoja1.Cells(pnFila + 1, 7) = "Monto"
    xlHoja1.Cells(pnFila, 8) = "DEFICIENTES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 8), xlHoja1.Cells(pnFila, 9)).Merge True
    xlHoja1.Cells(pnFila + 1, 8) = "Nro"
    xlHoja1.Cells(pnFila + 1, 9) = "Monto"
    xlHoja1.Cells(pnFila, 10) = "DUDOSOS"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 10), xlHoja1.Cells(pnFila, 11)).Merge True
    xlHoja1.Cells(pnFila + 1, 10) = "Nro"
    xlHoja1.Cells(pnFila + 1, 11) = "Monto"
    xlHoja1.Cells(pnFila, 12) = "PERDIDA"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 12), xlHoja1.Cells(pnFila, 13)).Merge True
    xlHoja1.Cells(pnFila + 1, 12) = "Nro"
    xlHoja1.Cells(pnFila + 1, 13) = "Monto"
    xlHoja1.Cells(pnFila, 14) = "PROVISION"
    xlHoja1.Cells(pnFila + 1, 14) = Format(dFecha, "dd/mm/yyyy")
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 14)).HorizontalAlignment = xlCenter
    CuadroExcel 1, pnFila, 14, pnFila + 1, True

    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 2), xlHoja1.Cells(350, 2)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 3), xlHoja1.Cells(350, 3)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 4), xlHoja1.Cells(350, 4)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 5), xlHoja1.Cells(350, 5)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 6), xlHoja1.Cells(350, 6)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 7), xlHoja1.Cells(350, 7)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 8), xlHoja1.Cells(350, 8)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 9), xlHoja1.Cells(350, 9)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 10), xlHoja1.Cells(350, 10)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 11), xlHoja1.Cells(350, 11)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 12), xlHoja1.Cells(350, 12)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 13), xlHoja1.Cells(350, 13)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 14), xlHoja1.Cells(350, 14)).NumberFormat = "#,#0.00;#,#0.00"


    xlHoja1.Range("A1").ColumnWidth = 18
    xlHoja1.Range("B1").ColumnWidth = 7
    xlHoja1.Range("C1").ColumnWidth = 15
    xlHoja1.Range("D1").ColumnWidth = 7
    xlHoja1.Range("E1").ColumnWidth = 15
    xlHoja1.Range("F1").ColumnWidth = 7
    xlHoja1.Range("G1").ColumnWidth = 15
    xlHoja1.Range("H1").ColumnWidth = 7
    xlHoja1.Range("I1").ColumnWidth = 15
    xlHoja1.Range("J1").ColumnWidth = 7
    xlHoja1.Range("K1").ColumnWidth = 15
    xlHoja1.Range("L1").ColumnWidth = 7
    xlHoja1.Range("M1").ColumnWidth = 15
    xlHoja1.Range("N1").ColumnWidth = 16

    ' Formato de celdas
    xlHoja1.Range("C1").NumberFormat = "#,#0.00"
    xlHoja1.Range("E1").NumberFormat = "#,#0.00"
    xlHoja1.Range("G1").NumberFormat = "#,#0.00"
    xlHoja1.Range("IC1").NumberFormat = "#,#0.00"
    xlHoja1.Range("K1").NumberFormat = "#,#0.00"
    xlHoja1.Range("M1").NumberFormat = "#,#0.00"
    xlHoja1.Range("N1").NumberFormat = "#,#0.00"

    pnFila = pnFila + 2


    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlAplicacion.Range("A1:Z100").Font.Size = 9

End Sub



'
Public Sub Genera_Reporte178205(ByVal TxtCambio As Double, ByVal txtMonto As Double, _
ByVal dFecha As Date)
'Dim fs As Scripting.FileSystemObject
Dim lbExisteHoja  As Boolean
Dim lsNomHoja As String
Dim pnFila As Integer, lnCol As Integer
Dim lsSQL As String
Dim rs As New ADODB.Recordset
Dim lnLinea As Integer
Dim lsMoneda As String, lsTipo As String, lsSubTipo As String, lsRango As String, lsFF As String
Dim lsTotFF(14) As String
Dim lsTotRango(14) As String
Dim lsTotSubTipo(14) As String
Dim lsTotTipo(14) As String
Dim lsTotCartera(14) As String
Dim lnIndice As Integer
Dim lbCambio As Boolean ' Indica el cambio condicion
Dim lnMoneda As Integer, lnTipo As Integer, lnSubTipo As Integer, lnRango As Integer, lnFF As Integer
Dim lnCuenta As Long
Dim FechaFinMes As Date
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim Conecta As DConecta
Dim ImpHojaWProvision As Boolean
Set Conecta = New DConecta
Dim Sql As String

ImpHojaWProvision = True
lnLinea = 1
FechaFinMes = GetUltimaFechaCierre
lbCambio = False
lsArchivo = "Anexo5_" & Format(FechaFinMes, "mmyyyy") & "HojaW_Provision.XLS"
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If
lbExcel = True
RaiseEvent ShowProgress
RaiseEvent Progress(0, 2)
For lnMoneda = 1 To 2 ' Moneda

    lbExisteHoja = False
    lsNomHoja = "HW_Moneda" & Str(lnMoneda)
    For Each xlHoja1 In xlLibro.Worksheets
        If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            lbExisteHoja = True
            Exit For
        End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlLibro.Worksheets.Add
        xlHoja1.Name = lsNomHoja
    End If
    'Me.EstatuBarICC.Panels(2).Text = "Generando Reporte en Excel Por favor espere..."
    pnFila = 2
    ' Cabecera del Reporte en Excel
    Call ImpHojaWProvisionCab(pnFila, Trim(Str(lnMoneda)), dFecha)
    ' Cuerpo del Reporte

    ' Limpiar Tot Cartera / Tipo / SubTipo
    For lnIndice = 2 To 14
        lsTotCartera(lnIndice) = ""
        lsTotTipo(lnIndice) = ""
        lsTotSubTipo(lnIndice) = ""
    Next lnIndice

    For lnTipo = 1 To 2

        ' Imprimir Encabezado
        xlHoja1.Cells(pnFila, 1) = IIf(Trim(Str(lnTipo)) = "1", "I.- CREDITOS COMERCIALES", "II.- CREDITOS MES")
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0C0FF

        pnFila = pnFila + 1
        ' Limpiar
        For lnIndice = 2 To 14
            lsTotSubTipo(lnIndice) = ""
        Next lnIndice

        For lnSubTipo = 1 To 2

            lsSQL = ""

            ' Imprimir Encabezado
            xlHoja1.Cells(pnFila, 1) = IIf(Trim(Str(lnSubTipo)) = "1", " 1.- EMPRESARIALES", " 2.- AGRICOLAS")
            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HFFC0C0
            pnFila = pnFila + 1
            ' Limpiar
            For lnIndice = 2 To 14
                lsTotRango(lnIndice) = ""
            Next lnIndice

            For lnRango = 1 To 2

                lsSQL = "SELECT Substring(CA.cCtaCod,6,3) as cSubtipo, " _
                      & "CASE WHEN (CA.cCtaCod like '________1%' AND CA.nMontoApr <=" & txtMonto & " * " & TxtCambio & " )" _
                      & "       OR (CA.cCtaCod like '________2%' AND CA.nMontoApr <=" & txtMonto & ") " _
                      & "     THEN  '1' ELSE '2' " _
                      & " END cRANGO, " _
                      & "Substring(CA.cLineaCred,1,2) as cFF, substring(CA.cCtaCod,4,2) as cAgencia, " _
                      & "Count (CA.cCtaCod) AS NroCred, " _
                      & "Sum (ISNULL(CA.nSaldoCap , 0)) AS SaldoCap, " _
                      & "Count( CASE WHEN CA.cCalGen = '0' THEN (CA.cCtaCod) END ) AS NroProv0, " _
                      & "Sum (ISNULL (CASE WHEN CA.cCalGen = '0' THEN (CA.nSaldoCap) END , 0 )) AS Prov0, " _
                      & "Count( CASE WHEN CA.cCalGen = '1' THEN (CA.cCtaCod) END ) AS NroProv1, " _
                      & "Sum (ISNULL (CASE WHEN CA.cCalGen = '1' THEN (CA.nSaldoCap) END , 0 )) AS Prov1, " _
                      & "Count(CASE WHEN CA.cCalGen = '2' THEN (CA.cCtaCod) END ) AS NroProv2, " _
                      & "Sum (ISNULL (CASE WHEN CA.cCalGen = '2' THEN (CA.nSaldoCap) END , 0 )) AS Prov2, " _
                      & "Count(CASE WHEN CA.cCalGen = '3' THEN (CA.cCtaCod) END ) AS NroProv3, " _
                      & "Sum (ISNULL (CASE WHEN CA.cCalGen = '3' THEN (CA.nSaldoCap) END , 0 )) AS Prov3, " _
                      & "Count(CASE WHEN CA.cCalGen = '4' THEN (CA.cCtaCod) END ) AS NroProv4, " _
                      & "Sum (ISNULL (CASE WHEN CA.cCalGen = '4' THEN (CA.nSaldoCap) END , 0 )) AS Prov4, " _
                      & "Sum(IsNull(ca.nProvision, 0)) As Provision "
                lsSQL = lsSQL & " " _
                      & "FROM ColocCalifProv CA " _
                      & "WHERE CA.nPrdEstado in ('2201','2020','2021','2022','2030','2031','2032') " _
                      & "AND Substring(CA.cCtaCod,6,2) in('10','20') " _
                      & "AND CA.nSaldoCap > 0 AND Substring(CA.cCtaCod,9,1) = '" & Trim(Str(lnMoneda)) & "' " _
                      & "AND Substring(CA.cCtaCod,6,1) ='" & Trim(Str(lnTipo)) & "' " _
                      & "AND Substring(CA.cCtaCod,8,1) ='" & Trim(Str(lnSubTipo)) & "' "

                If lnRango = 1 Then
                     lsSQL = lsSQL & " AND CA.nMontoapr <= (CASE WHEN CA.cCtaCod like '________1%' THEN (" & txtMonto & " * " & TxtCambio & " ) " _
                                & "                              WHEN CA.cCtaCod like '________2%' THEN (" & txtMonto & ")" _
                                & " END ) "
                Else
                     lsSQL = lsSQL & " AND CA.nMontoapr > (CASE WHEN CA.cCtaCod like '________1%' THEN (" & txtMonto & " * " & TxtCambio & " ) " _
                                & "                             WHEN CA.cCtaCod like '________2%' THEN (" & txtMonto & ")" _
                                & " END ) "
                End If
                Sql = " GROUP BY  Substring(CA.cCtaCod,6,3), " _
                      & "CASE WHEN (CA.cCtaCod like '________1%' AND CA.nMontoApr <=" & txtMonto & " * " & TxtCambio & " )" _
                      & "       OR (CA.cCtaCod like '________2%' AND CA.nMontoApr <=" & txtMonto & ") " _
                      & "     THEN  '1' ELSE '2' " _
                      & " END ,  " _
                      & "Substring(CA.cLineaCred,1,2), Substring(CA.cCtaCod,4,2) " _
                      & "ORDER By Substring(CA.cCtaCod,6,3), " _
                      & "CASE WHEN (CA.cCtaCod like '________1%' AND CA.nMontoApr <=" & txtMonto & " * " & TxtCambio & " )" _
                      & "       OR (CA.cCtaCod like '________2%' AND CA.nMontoApr <=" & txtMonto & ") " _
                      & "     THEN  '1' ELSE '2' " _
                      & " END ,  " _
                      & "Substring(CA.cLineaCred,1,2), Substring(CA.cCtaCod,4,2) "
                lsSQL = lsSQL & Sql

                Conecta.AbreConexion
                Set rs = Conecta.CargaRecordSet(lsSQL)
                Conecta.CierraConexion
                lsFF = ""

                If Not (rs.BOF And rs.EOF) Then
                    ' Imprimir Encabezado
                    xlHoja1.Cells(pnFila, 1) = "Rango(" & Trim(Str(lnRango)) & ") " & IIf(Trim(Str(lnRango)) = "1", "  MENORES A $ ", "  MAYORES A $ ") & Format(txtMonto, "#0.00")
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 1)).Font.Bold = True
                    pnFila = pnFila + 1
                    ' Limpiar
                    For lnIndice = 2 To 14
                        lsTotFF(lnIndice) = ""
                    Next lnIndice
                End If

                Do While Not rs.EOF

                    If lsFF <> rs!cFF Then ' Otro Fuente Financ
                        'Imprimir Total de anterior Fuente Fin
                        If lsFF <> "" Then
                            'Imprimir totales FF
                            xlHoja1.Cells(pnFila, 1) = "SubTotal F.F. " & Trim(lsFF)
                            For lnIndice = 2 To 14
                                xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotFF(lnIndice) & ")"
                            Next lnIndice
                            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                            'Asigno los valores al Rango
                            For lnIndice = 2 To 14
                                lsTotRango(lnIndice) = lsTotRango(lnIndice) & IIf(Len(lsTotRango(lnIndice)) > 0, "+", "") & lsTotFF(lnIndice)
                            Next lnIndice
                            pnFila = pnFila + 1

                        End If
                        ' Imprimir encabezado
                        xlHoja1.Cells(pnFila, 1) = "   F.F. " '---& ObtieneDescripcionFuenteFinanciamiento(Trim(rs!cFF))
                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 1)).Font.Bold = True
                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0FFC0
                        pnFila = pnFila + 1
                        For lnIndice = 2 To 14
                            lsTotFF(lnIndice) = ""
                        Next lnIndice

                    End If ' Fuente F

                    '*****

                    ' Imprimir el detalle
                    xlHoja1.Cells(pnFila, 1) = "      Agencia " & rs!cAgencia
                    xlHoja1.Cells(pnFila, 2) = Format(rs!NroCred, "#0")
                    xlHoja1.Cells(pnFila, 3) = Format(rs!SaldoCap, "#0.00")
                    xlHoja1.Cells(pnFila, 4) = Format(rs!NroProv0, "#0")
                    xlHoja1.Cells(pnFila, 5) = Format(rs!Prov0, "#0.00")
                    xlHoja1.Cells(pnFila, 6) = Format(rs!NroProv1, "#0")
                    xlHoja1.Cells(pnFila, 7) = Format(rs!Prov1, "#0.00")
                    xlHoja1.Cells(pnFila, 8) = Format(rs!NroProv2, "#0")
                    xlHoja1.Cells(pnFila, 9) = Format(rs!Prov2, "#0.00")
                    xlHoja1.Cells(pnFila, 10) = Format(rs!NroProv3, "#0")
                    xlHoja1.Cells(pnFila, 11) = Format(rs!Prov3, "#0.00")
                    xlHoja1.Cells(pnFila, 12) = Format(rs!NroProv4, "#0")
                    xlHoja1.Cells(pnFila, 13) = Format(rs!Prov4, "#0.00")
                    xlHoja1.Cells(pnFila, 14) = Format(rs!Provision, "#0.00")
                    'Acumulo totales
                    For lnIndice = 2 To 14
                        lsTotFF(lnIndice) = lsTotFF(lnIndice) & IIf(Len(lsTotFF(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)  'Nro Cred
                    Next lnIndice
                    pnFila = pnFila + 1
                    'lsTipo = rs!cTipo
                    lsSubTipo = rs!cSubTipo
                    lsRango = rs!cRango
                    lsFF = rs!cFF

                    rs.MoveNext
                Loop

                rs.Close
                Set rs = Nothing

                'Imprimir totales FF
                If Trim(lsTotFF(2)) <> "" Then
                    xlHoja1.Cells(pnFila, 1) = "SubTotal F.F. " & Trim(lsFF)
                    For lnIndice = 2 To 14
                        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotFF(lnIndice) & ")"
                    Next lnIndice
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                    'Asigno los valores al Rango
                    For lnIndice = 2 To 14
                        lsTotRango(lnIndice) = lsTotRango(lnIndice) & IIf(Len(lsTotRango(lnIndice)) > 0, "+", "") & lsTotFF(lnIndice)
                    Next lnIndice
                    pnFila = pnFila + 1

                    'Imprimir totales Rango
                    xlHoja1.Cells(pnFila, 1) = "SubTotal Rango  " & Trim(lsRango)
                    For lnIndice = 2 To 14
                        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotRango(lnIndice) & ")"
                    Next lnIndice
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                    'xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = RGB(10, 100, 10)
                    'Asigno los valores al SubTipo
                    For lnIndice = 2 To 14
                        lsTotSubTipo(lnIndice) = lsTotSubTipo(lnIndice) + IIf(Len(lsTotSubTipo(lnIndice)) > 0, "+", "") + lsTotRango(lnIndice)
                    Next lnIndice
                    pnFila = pnFila + 1
                End If

                ' Limpio la variable FF
                For lnIndice = 2 To 14
                    lsTotFF(lnIndice) = ""
                Next lnIndice
                ' Limpio la variable Rango
                For lnIndice = 2 To 14
                    lsTotRango(lnIndice) = ""
                Next lnIndice

            Next lnRango

            'Imprimir totales SubTipo
            If Trim(lsTotSubTipo(2)) <> "" Then
                xlHoja1.Cells(pnFila, 1) = "Total " & IIf(Trim(Str(lnSubTipo)) = "1", "EMPRESARIALES", "AGRICOLAS")
                For lnIndice = 2 To 14
                    xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotSubTipo(lnIndice) & ")"
                Next lnIndice
                xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                'Asigno los valores al tipo
                For lnIndice = 2 To 14
                    lsTotTipo(lnIndice) = lsTotTipo(lnIndice) & IIf(Len(lsTotTipo(lnIndice)) > 0, "+", "") & lsTotSubTipo(lnIndice)
                Next lnIndice
                pnFila = pnFila + 1
            End If

            ' Limpio la variable Subtipo
            For lnIndice = 2 To 14
                lsTotSubTipo(lnIndice) = ""
            Next lnIndice

         Next lnSubTipo

        'Imprimir totales Tipo
        If Trim(lsTotTipo(2)) <> "" Then
            xlHoja1.Cells(pnFila, 1) = "Total " & IIf(Trim(Str(lnTipo)) = "1", "COMERCIALES", "MES")
            For lnIndice = 2 To 14
                xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotTipo(lnIndice) & ")"
            Next lnIndice
            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
            'Asigno los valores a la Cartera
            For lnIndice = 2 To 14
                lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & lsTotTipo(lnIndice)
            Next lnIndice
            pnFila = pnFila + 2
        End If

        ' Limpio la variable tipo
        For lnIndice = 2 To 14
            lsTotTipo(lnIndice) = ""
        Next lnIndice

    Next lnTipo

    'Imprimir totales Tipo
    If Trim(lsTotTipo(2)) <> "" Then
        xlHoja1.Cells(pnFila, 1) = "Total " & IIf(Trim(Str(lnTipo)) = "1", "COMERCIALES", "MES")
        For lnIndice = 2 To 14
            xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotTipo(lnIndice) & ")"
        Next lnIndice
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
        'Asigno los valores a la Cartera
        For lnIndice = 2 To 14
            lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & lsTotTipo(lnIndice)
        Next lnIndice
        pnFila = pnFila + 2
    End If

    ' Limpio la variable tipo
    For lnIndice = 2 To 14
        lsTotTipo(lnIndice) = ""
    Next lnIndice


    '******************************************************************************************
    '**** PARA LO DE CONSUMO
    '******************************************************************************************

    For lnTipo = 3 To 4

        For lnIndice = 2 To 14
            lsTotTipo(lnIndice) = ""
        Next lnIndice

        ' Imprimir Encabezado
        xlHoja1.Cells(pnFila, 1) = IIf(Trim(Str(lnTipo)) = "3", "III.- CREDITOS CONSUMO", "IV.- CREDITOS HIPOTECARIOS")
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0C0FF

        pnFila = pnFila + 1
        ' Limpiar
        For lnIndice = 2 To 14
            lsTotSubTipo(lnIndice) = ""
        Next lnIndice

        ' Fuente Financiamiento
        For lnFF = 1 To 9

            lsSQL = "SELECT COUNT( ISNULL(CA.cCtaCod, 0)) as Cuenta " _
                      & "FROM ColocCalifProv CA " _
                      & "WHERE CA.nPrdEstado in ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') " _
                      & "AND Substring(CA.cCtaCod,9,1) = '" & Trim(Str(lnMoneda)) & "' " _
                      & "AND Substring(CA.cCtaCod,6,1) ='" & Trim(Str(lnTipo)) & "' " _
                      & "AND Substring(CA.cLineaCred,1,2) ='" & Format(Trim(Str(lnFF)), "00") & "' "

            Conecta.AbreConexion
            Set rs = Conecta.CargaRecordSet(lsSQL)
            Conecta.CierraConexion
            lnCuenta = rs!Cuenta
            rs.Close
            Set rs = Nothing
            If lnCuenta > 0 Then
                ' Imprimir Encabezado
                xlHoja1.Cells(pnFila, 1) = ObtieneDescripcionFuenteFinanciamiento(Format(Trim(Str(lnFF)), "00"))
                xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
                xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0FFC0
                pnFila = pnFila + 1
                ' Limpiar
                For lnIndice = 2 To 14
                    lsTotFF(lnIndice) = ""
                Next lnIndice

                ' Sub Tipo
                For lnSubTipo = 1 To 6
                    If lnTipo = 3 Then
                        lsSubTipo = IIf(Len(Trim(Str(lnSubTipo))) = 1, "30" & Trim(Str(lnSubTipo)), "3" & Trim(Str(lnSubTipo)))
                        lsSubTipo = IIf(lsSubTipo = "306", "320", lsSubTipo)
                    ElseIf lnTipo = 4 Then
                        lsSubTipo = IIf(Len(Trim(Str(lnSubTipo))) = 1, "40" & Trim(Str(lnSubTipo)), "3" & Trim(Str(lnSubTipo)))
                        'lsSubTipo = IIf(lsSubTipo = "402", "423", lsSubTipo) para trujillo habilitar
                    End If


                    lsSQL = "SELECT COUNT( ISNULL(CA.cCtaCod, 0)) as Cuenta " _
                              & "FROM ColocCalifProv CA " _
                              & "WHERE CA.nPrdEstado in ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') " _
                              & "AND Substring(CA.cCtaCod,9,1) = '" & Trim(Str(lnMoneda)) & "' " _
                              & "AND Substring(CA.cCtaCod,6,1) ='" & Trim(Str(lnTipo)) & "' " _
                              & "AND Substring(CA.cLineaCred,1,2) ='" & Format(Trim(Str(lnFF)), "00") & "' " _
                              & "AND Substring(CA.cCtaCod,6,3) ='" & Trim(lsSubTipo) & "' "

                    Conecta.AbreConexion
                    Set rs = Conecta.CargaRecordSet(lsSQL)
                    Conecta.CierraConexion
                    lnCuenta = rs!Cuenta
                    rs.Close
                    Set rs = Nothing
                    If lnCuenta > 0 Then

                        ' Imprimir Encabezado Subtipo
                        xlHoja1.Cells(pnFila, 1) = "55" 'ObtieneDescripcionSubTipoConsumo(Trim(Str(lsSubTipo)))
                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HFFC0C0
                        pnFila = pnFila + 1
                        ' Limpiar
                        For lnIndice = 2 To 14
                            lsTotSubTipo(lnIndice) = ""
                        Next lnIndice

                        lsSQL = "SELECT Substring(CA.cLineaCred,1,2) as cFF, substring(CA.cCtaCod,4,2) as cAgencia, " _
                                & "Count (CA.cCtaCod) AS NroCred, " _
                                & "Sum (ISNULL(CA.nSaldoCap , 0)) AS SaldoCap, " _
                                & "Count( CASE WHEN CA.cCalGen = '0' THEN (CA.cCtaCod) END ) AS NroProv0, " _
                                & "Sum (ISNULL (CASE WHEN CA.cCalGen = '0' THEN (CA.nSaldoCap) END , 0 )) AS Prov0, " _
                                & "Count( CASE WHEN CA.cCalGen = '1' THEN (CA.cCtaCod) END ) AS NroProv1, " _
                                & "Sum (ISNULL (CASE WHEN CA.cCalGen = '1' THEN (CA.nSaldoCap) END , 0 )) AS Prov1, " _
                                & "Count(CASE WHEN CA.cCalGen = '2' THEN (CA.cCtaCod) END ) AS NroProv2, " _
                                & "Sum (ISNULL (CASE WHEN CA.cCalGen = '2' THEN (CA.nSaldoCap) END , 0 )) AS Prov2, " _
                                & "Count(CASE WHEN CA.cCalGen = '3' THEN (CA.cCtaCod) END ) AS NroProv3, " _
                                & "Sum (ISNULL (CASE WHEN CA.cCalGen = '3' THEN (CA.nSaldoCap) END , 0 )) AS Prov3, " _
                                & "Count(CASE WHEN CA.cCalGen = '4' THEN (CA.cCtaCod) END ) AS NroProv4, " _
                                & "Sum (ISNULL (CASE WHEN CA.cCalGen = '4' THEN (CA.nSaldoCap) END , 0 )) AS Prov4, " _
                                & "Sum(IsNull(ca.nProvision, 0)) As Provision "

                        lsSQL = lsSQL & " " _
                                & "FROM ColocCalifProv CA " _
                                & "WHERE CA.nPrdEstado in ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807')  " _
                                & "AND Substring(CA.cCtaCod,9,1) = '" & Trim(Str(lnMoneda)) & "' " _
                                & "AND Substring(CA.cCtaCod,6,1) ='" & Trim(Str(lnTipo)) & "' " _
                                & "AND Substring(CA.cLineaCred,1,2) ='" & Format(Trim(Str(lnFF)), "00") & "' " _
                                & "AND Substring(CA.cCtaCod,6,3) ='" & Trim(lsSubTipo) & "' "
                        lsSQL = lsSQL & "GROUP BY  Substring(CA.cLineaCred,1,2), Substring(CA.cCtaCod,4,2) " _
                                & "ORDER By Substring(CA.cLineacred,1,2), Substring(CA.cCtaCod,4,2) "
                        Conecta.AbreConexion
                        Set rs = Conecta.CargaRecordSet(lsSQL)
                        Conecta.CierraConexion

                        Do While Not rs.EOF
                            ' Imprimir el detalle
                            xlHoja1.Cells(pnFila, 1) = "      Agencia " & rs!cAgencia
                            xlHoja1.Cells(pnFila, 2) = Format(rs!NroCred, "#0")
                            xlHoja1.Cells(pnFila, 3) = Format(rs!SaldoCap, "#0.00")
                            xlHoja1.Cells(pnFila, 4) = Format(rs!NroProv0, "#0")
                            xlHoja1.Cells(pnFila, 5) = Format(rs!Prov0, "#0.00")
                            xlHoja1.Cells(pnFila, 6) = Format(rs!NroProv1, "#0")
                            xlHoja1.Cells(pnFila, 7) = Format(rs!Prov1, "#0.00")
                            xlHoja1.Cells(pnFila, 8) = Format(rs!NroProv2, "#0")
                            xlHoja1.Cells(pnFila, 9) = Format(rs!Prov2, "#0.00")
                            xlHoja1.Cells(pnFila, 10) = Format(rs!NroProv3, "#0")
                            xlHoja1.Cells(pnFila, 11) = Format(rs!Prov3, "#0.00")
                            xlHoja1.Cells(pnFila, 12) = Format(rs!NroProv4, "#0")
                            xlHoja1.Cells(pnFila, 13) = Format(rs!Prov4, "#0.00")
                            xlHoja1.Cells(pnFila, 14) = Format(rs!Provision, "#0.00")
                            'Acumulo totales
                            For lnIndice = 2 To 14
                                lsTotSubTipo(lnIndice) = lsTotSubTipo(lnIndice) & IIf(Len(lsTotSubTipo(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)  'Nro Cred
                            Next lnIndice
                            pnFila = pnFila + 1

                            rs.MoveNext

                        Loop

                        rs.Close
                        Set rs = Nothing

                        'Imprimir totales Subtipo
                        If Trim(lsTotSubTipo(2)) <> "" Then
                            xlHoja1.Cells(pnFila, 1) = "Sub Total "
                            For lnIndice = 2 To 14
                                xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotSubTipo(lnIndice) & ")"
                            Next lnIndice
                            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                            'Asigno los valores a FF
                            For lnIndice = 2 To 14
                                lsTotFF(lnIndice) = lsTotFF(lnIndice) & IIf(Len(lsTotFF(lnIndice)) > 0, "+", "") & lsTotSubTipo(lnIndice)
                            Next lnIndice
                            pnFila = pnFila + 1

                        End If

                        ' Limpio la variable SubTipo
                        For lnIndice = 2 To 14
                            lsTotSubTipo(lnIndice) = ""
                        Next lnIndice

                    End If

                Next lnSubTipo


                'Imprimir totales FF
                If Trim(lsTotFF(2)) <> "" Then
                    xlHoja1.Cells(pnFila, 1) = "Total " & ObtieneDescripcionFuenteFinanciamiento(Format(Trim(Str(lnFF)), "00"))
                    For lnIndice = 2 To 14
                        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotFF(lnIndice) & ")"
                    Next lnIndice
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                    'Asigno los valores a Total
                    For lnIndice = 2 To 14
                        lsTotTipo(lnIndice) = lsTotTipo(lnIndice) & IIf(Len(lsTotTipo(lnIndice)) > 0, "+", "") & lsTotFF(lnIndice)
                    Next lnIndice

                    pnFila = pnFila + 2
                End If

                ' Limpio la variable FF
                For lnIndice = 2 To 14
                    lsTotFF(lnIndice) = ""
                Next lnIndice

            End If

       Next lnFF

    'Imprimir totales Tipo
    If Trim(lsTotTipo(2)) <> "" Then
        xlHoja1.Cells(pnFila, 1) = "Total " & IIf(Trim(Str(lnTipo)) = "3", "CONSUMO", "HIPOTECARIO")
        For lnIndice = 2 To 14
            xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotTipo(lnIndice) & ")"
        Next lnIndice
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
        'Asigno los valores a la Cartera
        For lnIndice = 2 To 14
            lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & lsTotTipo(lnIndice)
        Next lnIndice
        pnFila = pnFila + 2
    End If

    Next lnTipo
    ' ********************

    'Imprime Total Cartera
    xlHoja1.Cells(pnFila, 1) = "Total CARTERA"
    For lnIndice = 2 To 14
        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotCartera(lnIndice) & ")"
    Next lnIndice
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0FFFF

    'CuadroExcel 1, 6, 14, pnFila, True
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(pnFila, 14)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(pnFila, 14)).Cells.Borders.LineStyle = xlInside
    RaiseEvent Progress(CInt(lnMoneda), 2)
Next lnMoneda
RaiseEvent CloseProgress
MsgBox "Se Genero el archivo " & App.path & "\SPOOLER\" & lsArchivo

'*******
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
lbExcel = False
End Sub


Public Sub Genera_Reporte178207(ByVal psServerCons As String, ByVal TxtCambio As Double, ByVal psNomCmact As String, _
ByVal pdFecDataFin As Date)
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim RCD As nRcdReportes
Dim lsArchivo As String
Dim nFil As Integer
Dim nCol As Integer
Dim lsCodigo As String
Dim Titulo As String
Dim i As Integer
Dim J As Integer
Dim k As Integer
Dim L As Integer


Dim lsEstadoProd(1 To 4) As String
lsEstadoProd(1) = " nPrdEstado in (2020,2022) "
lsEstadoProd(2) = " nPrdEstado in (2021) "
lsEstadoProd(3) = " nPrdEstado in (2030,2031,2032) "
lsEstadoProd(4) = " nPrdEstado in (2201) "

Set RCD = New nRcdReportes

On Error Resume Next
Sql = "select * from ConstSistema where nConsSisCod=3"
Set rs = RCD.GetQuery(Sql)
lsCodigo = IIf(IsNull(rs!NCONSSISVALOR), "", rs!NCONSSISVALOR)

lsArchivo = "ReporteAnexo5C_" & Format(pdFecDataFin, "YYYYMMDD") & ".xls"
Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

Set xlHoja1 = xlLibro.Worksheets.Add
'RaiseEvent Progress(2, 20)

xlHoja1.Range("A2").Cells.Font.Bold = True
xlHoja1.Range("A2").Cells.Font.Size = 10
xlHoja1.Cells(2, 1) = "SUPERINTENDENCIA DE BANCA Y SEGUROS" 'psNomCmac
xlHoja1.Range("A3").Cells.Font.Size = 8
xlHoja1.Cells(3, 1) = "                 ANEXO 5-C"


xlHoja1.Cells(4, 4) = "INFORME SOBRE CREDITOS HIPOTECARIOS - FONDO MI VIVIENDA"
'xlHoja1.Range("A6").HorizontalAlignment = xlCenter
xlHoja1.Range(4, 4).Cells.Font.Bold = True
xlHoja1.Range(4, 4).Cells.Font.Size = 15
xlHoja1.Cells(5, 5) = "             AL " & Format(pdFecDataFin, "dd mmmm yyyy")
xlHoja1.Cells(6, 5) = "               (en miles de soles)"

xlHoja1.Cells(8, 1) = " EMPRESA :" & psNomCmact
xlHoja1.Range("A10:M32").Cells.Font.Size = 8
xlHoja1.Range("A10:M11").HorizontalAlignment = xlCenter
xlHoja1.Cells(8, 12) = "CODIGO :"
xlHoja1.Cells(8, 13) = lsCodigo

xlHoja1.Range("A1").ColumnWidth = 25
xlHoja1.Range("B1").ColumnWidth = 10
xlHoja1.Range("C1").ColumnWidth = 10
xlHoja1.Range("D1").ColumnWidth = 10
xlHoja1.Range("E1").ColumnWidth = 10
xlHoja1.Range("F1").ColumnWidth = 10
xlHoja1.Range("G1").ColumnWidth = 12
xlHoja1.Range("H1").ColumnWidth = 12
xlHoja1.Range("I1").ColumnWidth = 12
xlHoja1.Range("J1").ColumnWidth = 12
xlHoja1.Range("K1").ColumnWidth = 12
xlHoja1.Range("L1").ColumnWidth = 12
xlHoja1.Range("M1").ColumnWidth = 12


xlHoja1.Cells(10, 1) = "CREDITOS HIPOTECARIOS FONDO"
xlHoja1.Cells(11, 1) = "MI VIVIENDA 1"

xlHoja1.Cells(10, 2) = "NORMAL"
xlHoja1.Cells(10, 3) = "CPP"
xlHoja1.Cells(10, 4) = "DEFICIENTE"
xlHoja1.Cells(10, 5) = "DUDOSO"
xlHoja1.Cells(10, 6) = "PERDIDA"
xlHoja1.Cells(10, 7) = "TOTAL"

xlHoja1.Cells(10, 8) = "SUB - TOTAL"
xlHoja1.Cells(11, 8) = "COBERTURADO 3"

xlHoja1.Cells(10, 9) = "SUB - TOTAL NO"
xlHoja1.Cells(11, 9) = "COBERTURADO"

xlHoja1.Cells(10, 10) = "PROVISION"
xlHoja1.Cells(11, 10) = "SUB - TOTAL"

xlHoja1.Cells(10, 10) = "PROVISION"
xlHoja1.Cells(11, 10) = "REQUERIDA 4"

xlHoja1.Cells(10, 11) = "PROVISION"
xlHoja1.Cells(11, 11) = "CONSTITUIDA"

xlHoja1.Cells(10, 12) = "DEFICIT DE"
xlHoja1.Cells(11, 12) = "PROVISION"

xlHoja1.Cells(10, 13) = "GARANTIAS 1"

xlHoja1.Cells(12, 1) = "VIGENTES"
xlHoja1.Cells(13, 1) = "VENCIDOS"
xlHoja1.Cells(14, 1) = "REFINANCIADOS"
xlHoja1.Cells(15, 1) = "COBRANZA JUDICIAL"
xlHoja1.Cells(16, 1) = "TOTAL"
xlHoja1.Cells(17, 1) = "SUB TOTAL COBERTURADO 2"
xlHoja1.Cells(18, 1) = "SUB TOTAL COBERTURADO 3"
xlHoja1.Cells(19, 1) = "PROVISION REQUERIDA 4"
xlHoja1.Cells(20, 1) = "PROVISION COSNTITUIDA"
xlHoja1.Cells(21, 1) = "DEFICIT DE PROVISION"
xlHoja1.Cells(22, 1) = "GARANTIAS 5"

k = 12
For i = 1 To 4  ' Estados de Productos
    J = 2
    Sql = " Select T.nConsValor, isnull(sum(C.SaldoCap),0) Total"
    Sql = Sql & " From"
    Sql = Sql & " ("
    Sql = Sql & " Select case when substring(cCtaCod,9,1)='1' then nSaldoCap "
    Sql = Sql & "             else nSaldoCap * " & TxtCambio & " end SaldoCap, "
    Sql = Sql & "        cCalGen from ColocCalifProv "
    Sql = Sql & " where  " & lsEstadoProd(i) & "  and substring(cCtaCod,6,3) = '402'"
    Sql = Sql & " ) C"
    Sql = Sql & " Right Outer Join"
    Sql = Sql & " ( Select nConsValor, cConsDescripcion"
    Sql = Sql & "   From Constante"
    Sql = Sql & "   Where nConsCod = 3010 And nConsValor <> 3010"
    Sql = Sql & " ) T on T.nConsValor = C.cCalGen"
    Sql = Sql & " group by T.nConsValor"
    Set rs = RCD.GetQuery(Sql)
    While Not rs.EOF
        xlHoja1.Cells(k, J) = Format(rs!Total, "0.00")
        rs.MoveNext
        J = J + 1
    Wend
    k = k + 1
Next i

' Calculando Provisiones
Sql = "  Select T.nConsValor, isnull(sum(C.nProvision),0) Total"
Sql = Sql & " From"
Sql = Sql & " ("
Sql = Sql & "  Select case when substring(cCtaCod,9,1)='1' then nProvision"
Sql = Sql & "              else nProvision * " & TxtCambio & "  end nProvision,"
Sql = Sql & "         cCalGen from ColocCalifProv "
Sql = Sql & "  where  nPrdEstado in (2020,2021,2022,2030,2031,2032,2201) and substring(cCtaCod,6,3) = '402'"
Sql = Sql & " ) C"
Sql = Sql & " Right Outer Join"
Sql = Sql & " ( Select nConsValor, cConsDescripcion"
Sql = Sql & "   From Constante"
Sql = Sql & "   Where nConsCod = 3010 And nConsValor <> 3010"
Sql = Sql & " ) T on T.nConsValor = C.cCalGen"
Sql = Sql & " group by T.nConsValor"
Sql = Sql & " order by T.nConsValor"
Set rs = RCD.GetQuery(Sql)
J = 2
k = 12
While Not rs.EOF
    xlHoja1.Cells(19, J) = Format(rs!Total, "0.00")
    xlHoja1.Cells(20, J) = Format(rs!Total, "0.00")
    rs.MoveNext
    J = 1 + J
Wend




' Calculando las Garantias
Sql = " Select T.nConsValor, isnull(sum(C.SumGar),0) Total"
Sql = Sql & " From"
Sql = Sql & " ("
Sql = Sql & " Select  c.cCalGen,"
Sql = Sql & "     isnull(case when substring(CG.cCtaCod,9,1)='1' then  G.nRealizacion"
Sql = Sql & "         else round( G.nRealizacion  * " & TxtCambio & " ,2) end,0) SumGar"
Sql = Sql & " from ColocCalifProv  C"
Sql = Sql & " Inner Join ColocGarantia CG on C.cCtaCod = CG.cCtaCod"
Sql = Sql & " Inner Join Garantias G on G.cNumgarant = CG.cNumGarant"
Sql = Sql & " where  C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201) and substring(C.cCtaCod,6,3) = '402'"
Sql = Sql & " ) C"
Sql = Sql & " Right Outer Join"
Sql = Sql & " ( Select nConsValor, cConsDescripcion"
Sql = Sql & "   From Constante"
Sql = Sql & "   Where nConsCod = 3010 And nConsValor <> 3010"
Sql = Sql & " ) T on T.nConsValor = C.cCalGen"
Sql = Sql & " group by T.nConsValor"
Set rs = RCD.GetQuery(Sql)
J = 2
While Not rs.EOF
    xlHoja1.Cells(22, J) = Format(rs!Total, "0.00")
    J = 1 + J
    rs.MoveNext
Wend
J = 12
For i = 1 To 4
    Sql = " Select  isnull(sum(case when substring(cCtaCod,9,1)='1' then nProvision"
    Sql = Sql & "          else nProvision * " & TxtCambio & "  end),0) Provision"
    Sql = Sql & "      From ColocCalifProv"
    Sql = Sql & " where  " & lsEstadoProd(i) & " and substring(cCtaCod,6,3) = '402'"
    Set rs = RCD.GetQuery(Sql)
    While Not rs.EOF
        xlHoja1.Range("J" & J) = Format(rs!Provision, "0.00")
        xlHoja1.Range("K" & J) = Format(rs!Provision, "0.00")
        rs.MoveNext
    Wend
    J = 1 + J
Next i


xlHoja1.Cells(25, 2) = "FUNCIONARIO RESPONSABLE"
xlHoja1.Range("B25:C25").MergeCells = True
xlHoja1.Range("B25").HorizontalAlignment = xlCenter

xlHoja1.Range("E25:F25").MergeCells = True
xlHoja1.Cells(25, 5) = "GERENTE GENERAL"
xlHoja1.Range("E25").HorizontalAlignment = xlCenter

xlHoja1.Cells(25, 8) = "CONTADOR GENERAL"
xlHoja1.Range("H25:I25").MergeCells = True
xlHoja1.Range("H25").HorizontalAlignment = xlCenter

xlHoja1.Cells(27, 1) = "1. Se considerará los montos d los créditos hipotecarios concedidos con recursos de Fondo, clasificados por cada rubro contables, sin prejuicio de que"
xlHoja1.Cells(28, 1) = "   también se incluya en el anexo n° 5."
xlHoja1.Cells(29, 1) = "2. Corresponde a la porción del crédito que cuenta con cobertura de riesgo del Fondo."
xlHoja1.Cells(30, 1) = "3. Corresponde a la porción del crédito que no cuenta con cobertura de riesgo del Fondo"
xlHoja1.Cells(31, 1) = "4. Se anotará la provisión conforme a lo señalado en el Art.1 de la Resolucion SBS. 0865-99."
xlHoja1.Cells(32, 1) = "5. Corresponde al valor de realización de las garantías hipotecarias conforme a las normas vigentes."

' Bordes
xlHoja1.Range("A12:F22").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("A12:F22").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("A12:F22").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("A12:F22").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("A12:F22").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("A12:F22").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("A12:F22").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A12:F22").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

xlHoja1.Range("G12:M16").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
xlHoja1.Range("G12:M16").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
xlHoja1.Range("G12:M16").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
xlHoja1.Range("G12:M16").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("G12:M16").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("G12:M16").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("G12:M16").Cells.Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("G12:M16").Cells.Borders(xlInsideHorizontal).LineStyle = xlContinuous

For i = Asc("A") To Asc("M")
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlDiagonalUp).LineStyle = xlNone
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlEdgeLeft).LineStyle = xlContinuous
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlEdgeRight).LineStyle = xlContinuous
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlEdgeLeft).Weight = xlMedium
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlEdgeTop).Weight = xlMedium
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlEdgeBottom).Weight = xlMedium
    xlHoja1.Range(Chr(i) & "10:" & Chr(i) & "11").Cells.Borders(xlEdgeRight).Weight = xlMedium
Next i

For i = Asc("B") To Asc("F")
    xlHoja1.Range(Chr(i) & "16").Formula = "=SUM(" & Chr(i) & "12.." & Chr(i) & "15)"
    xlHoja1.Range(Chr(i) & "17").Formula = "= (1/3) * " & Chr(i) & 16
    xlHoja1.Range(Chr(i) & "18").Formula = "= (2/3) * " & Chr(i) & 16
    xlHoja1.Range(Chr(i) & "21").Formula = "=(" & Chr(i) & "20" & "-" & Chr(i) & "19)"
Next i
For i = 12 To 16
    xlHoja1.Range("G" & i).Formula = "=SUM(B" & i & "..F" & i & ")"
    xlHoja1.Range("M" & i).Formula = "=SUM(H" & i & "..L" & i & ")"
        xlHoja1.Range("H" & i).Formula = "= (1/3) * G" & i
    xlHoja1.Range("I" & i).Formula = "= (2/3) * G" & i
Next i
For i = Asc("H") To Asc("L")
    xlHoja1.Range(Chr(i) & "16").Formula = "=SUM(" & Chr(i) & "12.." & Chr(i) & "15)"
Next i

For i = 12 To 15
    xlHoja1.Range("L" & i).Formula = "=(K" & i & " - J" & i & ")"
Next i




xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
MsgBox "Se genero el archivo " & App.path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'RaiseEvent Progress(20, 20)
'---------------->
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
Set rs = Nothing
Set RCD = Nothing
'Set oTipCambio = Nothing
'-------------->

'RaiseEvent CloseProgress


Exit Sub
ErrorExcel:
    MsgBox "Error Nº [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
    xlLibro.Close
    ' Cierra Microsoft Excel con el método Quit.
    xlAplicacion.Quit
    'Libera los objetos.
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Public Sub Genera_Reporte178206(ByVal TxtCambio As Double, ByVal txtMonto As Double, ByVal sFuente As String, _
                ByVal dFecha As Date, Optional ByVal psNomCmac As String, Optional ByVal psCodCmac As String)

Dim Sql As String
Dim rs As New ADODB.Recordset
Dim lsProvision() As String
Dim lsRecurso As String, lsRecursos As String
Dim lsAgencias As String
Dim lnNumRec As Integer
Dim lnNumAge  As Integer
Dim lnNumProv As Integer
Dim lsMoneda  As String
Dim lsRefinan  As String
Dim lsTipo  As String
Dim lsTipoCredito As String
Dim lsSQLRecurso As String, lsSQLRecursoRef As String
Dim lsMonto As String
Dim rsP As New ADODB.Recordset
Dim lnMoneda As Integer
Dim lnTotal As Integer
Dim lnNumRec1 As Integer
Dim lsNomHoja As String

Dim liProd As Integer
Dim liCatCal As Integer

Dim lsArchivo As String
Dim lbExcel As Boolean
Dim FechaFinMes As Date

Dim pnFila As Integer, lnCol As Integer
Dim lbExisteHoja  As Boolean
Dim lsTotRub(6) As String
Dim lsTotxCatNroDeu(6) As String
Dim lsTotxCatMonto(6) As String
Dim lsTotxCatProviR(6) As String
Dim lsTotxCatProviC(6) As String
Dim lsTotxCatDefic(6) As String
Dim lsTotxCatGarant(6) As String

Dim Conecta As DConecta

Dim lsTotGeneral(6) As String
Dim lsFFinan As String
Dim nTotGarPrefReal As Double

Set Conecta = New DConecta
FechaFinMes = GetUltimaFechaCierre
'If Len(Trim(nFuente)) > 0 Then
    lsFFinan = " And Substring (ca.cLineaCred,1,4)  in " & Trim(sFuente) & ""
'End If
RaiseEvent ShowProgress
lsArchivo = "Anexo5_" & Format(FechaFinMes, "mmyyyy") & ".XLS"
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If
lbExcel = True
lbExisteHoja = False
lsNomHoja = "Anexo 5"
For Each xlHoja1 In xlLibro.Worksheets
    If xlHoja1.Name = lsNomHoja Then
        xlHoja1.Activate
        lbExisteHoja = True
        Exit For
    End If
Next

If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

pnFila = 2
' Cabecera del Reporte en Excel
ImpAnexo5Cab pnFila, "1", sFuente, dFecha, , , psNomCmac, psCodCmac, "5"
RaiseEvent Progress(1, 9)

' Cuerpo del Reporte
For liProd = 1 To 4


    If liProd = 1 Or liProd = 2 Or liProd = 4 Then
        Select Case liProd
            Case 1
                xlHoja1.Cells(pnFila, 1) = "COMERCIALES"
            Case 2
                xlHoja1.Cells(pnFila, 1) = "MICROEMPRESAS"
            Case 4
                xlHoja1.Cells(pnFila, 1) = "HIPOTECARIO VIVIENDA"
        End Select
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 15)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 15)).HorizontalAlignment = xlCenter
        xlHoja1.Cells(pnFila + 1, 1) = "Monto"
        xlHoja1.Cells(pnFila + 2, 1) = "Provision Requerida"
        xlHoja1.Cells(pnFila + 3, 1) = "Provision Constituida"
        xlHoja1.Cells(pnFila + 4, 1) = "Deficit"
        xlHoja1.Cells(pnFila + 5, 1) = "Garantias Preferidas y Preferidas de Muy rapida Realizacion"
        xlHoja1.Cells(pnFila + 6, 1) = "Numero de Deudores"

        lnCol = 2
        For liCatCal = 0 To 4
            If liCatCal = 0 Then   ' Para Normales

                Sql = "SELECT Count ( distinct (cPersCod) ) as Deudores, "
                Sql = Sql & "           Count (CA.cCtaCod ) AS NroCred, "
                Sql = Sql & "           ISNULL(Sum (CASE WHEN CA.cCtaCod like  '________1%' THEN (CA.nSaldoCap) "
                Sql = Sql & "                            WHEN CA.cCtaCod like  '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
                Sql = Sql & "                   END ), 0 ) AS SaldoCap, "
                Sql = Sql & "           ISNULL (Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
                Sql = Sql & "                             WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
                Sql = Sql & "                   END ), 0 ) AS Provision, "
                Sql = Sql & "           ISNULL(SUM (ISNULL(G.nMontoGar,0)),0) AS GarantHip   "
                Sql = Sql & " FROM  ColocCalifProv CA LEFT JOIN " & ServerCons & "CreditoConsol C ON CA.cCtaCod = C.cCtaCod"
                Sql = Sql & "  LEFT JOIN ( "
                Sql = Sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                Sql = Sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                Sql = Sql & "       Where G.nGarClase=1 AND G.nGarTpoRealiz=1  "
                Sql = Sql & "        Group by GC.cCtaCod  "
                Sql = Sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
                Sql = Sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "' "
                Sql = Sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan


                '-------------------------------------------

                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(Sql)
                Conecta.CierraConexion

                xlHoja1.Cells(pnFila + 1, lnCol + liCatCal) = Format(rsP!SaldoCap, "#0.00")
                xlHoja1.Cells(pnFila + 2, lnCol + liCatCal) = Format(rsP!Provision, "#0.00")
                xlHoja1.Cells(pnFila + 3, lnCol + liCatCal) = Format(rsP!Provision, "#0.00")
                xlHoja1.Cells(pnFila + 4, lnCol + liCatCal) = "0.00"
                xlHoja1.Cells(pnFila + 5, lnCol + liCatCal) = Format(rsP!GarantHip, "#0.00")
                xlHoja1.Cells(pnFila + 6, lnCol + liCatCal) = Format(rsP!Deudores, "#")
                xlHoja1.Range(xlHoja1.Cells(pnFila + 6, lnCol + liCatCal), xlHoja1.Cells(pnFila + 6, lnCol + liCatCal)).NumberFormat = "#,#0"

                rsP.Close
                Set rsP = Nothing

                'Total por Rubro

                lsTotRub(1) = xlHoja1.Range(xlHoja1.Cells(pnFila + 1, lnCol + liCatCal), xlHoja1.Cells(pnFila + 1, lnCol + liCatCal)).Address(False, False)
                lsTotRub(2) = xlHoja1.Range(xlHoja1.Cells(pnFila + 2, lnCol + liCatCal), xlHoja1.Cells(pnFila + 2, lnCol + liCatCal)).Address(False, False)
                lsTotRub(3) = xlHoja1.Range(xlHoja1.Cells(pnFila + 3, lnCol + liCatCal), xlHoja1.Cells(pnFila + 3, lnCol + liCatCal)).Address(False, False)
                lsTotRub(4) = xlHoja1.Range(xlHoja1.Cells(pnFila + 4, lnCol + liCatCal), xlHoja1.Cells(pnFila + 4, lnCol + liCatCal)).Address(False, False)
                lsTotRub(5) = xlHoja1.Range(xlHoja1.Cells(pnFila + 5, lnCol + liCatCal), xlHoja1.Cells(pnFila + 5, lnCol + liCatCal)).Address(False, False)
                lsTotRub(6) = xlHoja1.Range(xlHoja1.Cells(pnFila + 6, lnCol + liCatCal), xlHoja1.Cells(pnFila + 6, lnCol + liCatCal)).Address(False, False)

                'Para Tot x Calificacion
                lsTotxCatMonto(liCatCal) = lsTotxCatMonto(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 1, lnCol + liCatCal), xlHoja1.Cells(pnFila + 1, lnCol + liCatCal)).Address(False, False)
                lsTotxCatProviR(liCatCal) = lsTotxCatProviR(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 2, lnCol + liCatCal), xlHoja1.Cells(pnFila + 2, lnCol + liCatCal)).Address(False, False)
                lsTotxCatProviC(liCatCal) = lsTotxCatProviC(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 3, lnCol + liCatCal), xlHoja1.Cells(pnFila + 3, lnCol + liCatCal)).Address(False, False)
                lsTotxCatDefic(liCatCal) = lsTotxCatDefic(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 4, lnCol + liCatCal), xlHoja1.Cells(pnFila + 4, lnCol + liCatCal)).Address(False, False)
                lsTotxCatGarant(liCatCal) = lsTotxCatGarant(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 5, lnCol + liCatCal), xlHoja1.Cells(pnFila + 5, lnCol + liCatCal)).Address(False, False)
                lsTotxCatNroDeu(liCatCal) = lsTotxCatNroDeu(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 6, lnCol + liCatCal), xlHoja1.Cells(pnFila + 6, lnCol + liCatCal)).Address(False, False)


            Else  ' Todas menos Normales

               Sql = "  SELECT Count ( distinct (CA.cPersCod) ) as Deudores, "
                Sql = Sql & " Count (CA.cCtaCod ) AS NroCred, "
                Sql = Sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nSaldoCap) "
                Sql = Sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
                Sql = Sql & "               END ), 0 ) AS SaldoCap, "
                Sql = Sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
                Sql = Sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
                Sql = Sql & "               END ), 0 ) AS Provision,"
                Sql = Sql & "  GarantPrefRapRea = SUM(ISNULL(G.nMontoGar,0))  "
                Sql = Sql & " FROM ColocCalifProv CA"
                Sql = Sql & "  LEFT JOIN ( "
                Sql = Sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                Sql = Sql & "        "
                Sql = Sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                Sql = Sql & "        Where G.nGarClase=1 AND G.nGarTpoRealiz=1 "
                Sql = Sql & "        Group by GC.cCtaCod  "
                Sql = Sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
                Sql = Sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "' "
                Sql = Sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan


                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(Sql)
                Conecta.CierraConexion

                xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 0) = Format(rsP!SaldoCap, "#0.00")
                xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 0) = Format(rsP!Provision, "#0.00")
                xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 0) = Format(rsP!Provision, "#0.00")
                xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 0) = "0.00"
                xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 0) = Format(rsP!GarantPrefRapRea, "#0.00")
                xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 0) = Format(rsP!Deudores, "#")
                xlHoja1.Range(xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 0)).NumberFormat = "#,#0"

                rsP.Close
                Set rsP = Nothing

                'Totales x Rubro
                lsTotRub(1) = lsTotRub(1) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 0)).Address(False, False)
                lsTotRub(2) = lsTotRub(2) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 0)).Address(False, False)
                lsTotRub(3) = lsTotRub(3) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 0)).Address(False, False)
                lsTotRub(4) = lsTotRub(4) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 0)).Address(False, False)
                lsTotRub(5) = lsTotRub(5) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 0)).Address(False, False)
                lsTotRub(6) = lsTotRub(6) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 0)).Address(False, False)

                'Para Tot x Calificacion
                lsTotxCatMonto(liCatCal) = lsTotxCatMonto(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 0)).Address(False, False)
                lsTotxCatProviR(liCatCal) = lsTotxCatProviR(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 0)).Address(False, False)
                lsTotxCatProviC(liCatCal) = lsTotxCatProviC(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 0)).Address(False, False)
                lsTotxCatDefic(liCatCal) = lsTotxCatDefic(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 0)).Address(False, False)
                lsTotxCatGarant(liCatCal) = lsTotxCatGarant(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 0)).Address(False, False)
                lsTotxCatNroDeu(liCatCal) = lsTotxCatNroDeu(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 0), xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 0)).Address(False, False)

                'Para Creditos MAYORES ( Tabla 2 )
                ' Cred Vigentes mayores (sin incluir agricolas)

                Sql = "  SELECT Count ( distinct (CA.cPersCod) ) as Deudores, "
                Sql = Sql & " Count (CA.cCtaCod ) AS NroCred, "
                Sql = Sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nSaldoCap) "
                Sql = Sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
                Sql = Sql & "               END ), 0 ) AS SaldoCap, "
                Sql = Sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
                Sql = Sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
                Sql = Sql & "               END ), 0 ) AS Provision,"
                Sql = Sql & "  GarantPrefRapRea = ISNULL(SUM(ISNULL(GR.nMontoGar,0)),0)  "
                Sql = Sql & " FROM ColocCalifProv CA"
                Sql = Sql & "  INNER JOIN ( "
                Sql = Sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                Sql = Sql & "        "
                Sql = Sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                Sql = Sql & "        Where G.nGarClase=1 "
                Sql = Sql & "        Group by GC.cCtaCod  "
                Sql = Sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
                Sql = Sql & "  LEFT JOIN ( "
                Sql = Sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                Sql = Sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                Sql = Sql & "        Where G.nGarClase=1 AND  G.nGarTpoRealiz=1 "
                Sql = Sql & "        Group by GC.cCtaCod  "
                Sql = Sql & "       ) GR ON GR.cCtaCod = CA.cCtaCod "
                Sql = Sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "' "
                Sql = Sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan

            '------------------------------------
                'rsP.CursorLocation = adUseClient
                'rsP.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                'Set rsP.ActiveConnection = Nothing
                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(Sql)
                Conecta.CierraConexion
                xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 1) = Format(rsP!SaldoCap, "#0.00")
                xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 1) = Format(rsP!Provision, "#0.00")
                xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 1) = Format(rsP!Provision, "#0.00")
                xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 1) = "0.00"
                xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 1) = Format(rsP!GarantPrefRapRea, "#0.00")
                'xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 1) = Format(rsP!Deudores, "#")
                xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 1) = ""
                xlHoja1.Range(xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 1), xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 1)).NumberFormat = "#,#0"

                rsP.Close
                Set rsP = Nothing

                'Totales x Rubro
                lsTotRub(1) = lsTotRub(1) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 1), xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 1)).Address(False, False)
                lsTotRub(2) = lsTotRub(2) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 1), xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 1)).Address(False, False)
                lsTotRub(3) = lsTotRub(3) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 1), xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 1)).Address(False, False)
                lsTotRub(4) = lsTotRub(4) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 1), xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 1)).Address(False, False)
                lsTotRub(5) = lsTotRub(5) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 1), xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 1)).Address(False, False)
                lsTotRub(6) = lsTotRub(6) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 1), xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 1)).Address(False, False)


                '**************************************************
                'Para ( Tabla 2 B )
                '**************************************************

                Sql = "  SELECT Count ( distinct (CA.cPersCod) ) as Deudores, "
                Sql = Sql & " Count (CA.cCtaCod ) AS NroCred, "
                Sql = Sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nSaldoCap) "
                Sql = Sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
                Sql = Sql & "               END ), 0 ) AS SaldoCap, "
                Sql = Sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
                Sql = Sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
                Sql = Sql & "               END ), 0 ) AS Provision,"
                Sql = Sql & "  GarantPrefRapRea = ISNULL(SUM(ISNULL(G.nMontoGar,0)),0)  "
                Sql = Sql & " FROM ColocCalifProv CA"
                Sql = Sql & "  INNER JOIN ( "
                Sql = Sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                Sql = Sql & "       "
                Sql = Sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                Sql = Sql & "        Where G.nGarClase = 1 AND G.nGarTpoRealiz = 1 "
                Sql = Sql & "        Group by GC.cCtaCod  "
                Sql = Sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
                Sql = Sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "' "
                Sql = Sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan

            '------------------------------------
                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(Sql)
                Conecta.CierraConexion
                xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 2) = Format(rsP!SaldoCap, "#0.00")
                xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 2) = Format(rsP!Provision, "#0.00")
                xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 2) = Format(rsP!Provision, "#0.00")
                xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 2) = "0.00"
                xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 2) = Format(rsP!GarantPrefRapRea, "#0.00")
                'xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 2) = Format(rsP!Deudores, "#")
                xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 2) = ""
                xlHoja1.Range(xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 2), xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 2)).NumberFormat = "#,#0"

                rsP.Close
                Set rsP = Nothing

                'Totales x Rubro
                lsTotRub(1) = lsTotRub(1) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 2), xlHoja1.Cells(pnFila + 1, liCatCal * 3 + 2)).Address(False, False)
                lsTotRub(2) = lsTotRub(2) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 2), xlHoja1.Cells(pnFila + 2, liCatCal * 3 + 2)).Address(False, False)
                lsTotRub(3) = lsTotRub(3) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 2), xlHoja1.Cells(pnFila + 3, liCatCal * 3 + 2)).Address(False, False)
                lsTotRub(4) = lsTotRub(4) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 2), xlHoja1.Cells(pnFila + 4, liCatCal * 3 + 2)).Address(False, False)
                lsTotRub(5) = lsTotRub(5) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 2), xlHoja1.Cells(pnFila + 5, liCatCal * 3 + 2)).Address(False, False)
                lsTotRub(6) = lsTotRub(6) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 2), xlHoja1.Cells(pnFila + 6, liCatCal * 3 + 2)).Address(False, False)

            End If
        Next liCatCal
        'Imprime totales
        xlHoja1.Range(xlHoja1.Cells(pnFila + 1, 15), xlHoja1.Cells(pnFila + 1, 15)).Formula = "=SUM(" & lsTotRub(1) & ")"
        xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 15), xlHoja1.Cells(pnFila + 2, 15)).Formula = "=SUM(" & lsTotRub(2) & ")"
        xlHoja1.Range(xlHoja1.Cells(pnFila + 3, 15), xlHoja1.Cells(pnFila + 3, 15)).Formula = "=SUM(" & lsTotRub(3) & ")"
        xlHoja1.Range(xlHoja1.Cells(pnFila + 4, 15), xlHoja1.Cells(pnFila + 4, 15)).Formula = "=SUM(" & lsTotRub(4) & ")"
        xlHoja1.Range(xlHoja1.Cells(pnFila + 5, 15), xlHoja1.Cells(pnFila + 5, 15)).Formula = "=SUM(" & lsTotRub(5) & ")"
        xlHoja1.Range(xlHoja1.Cells(pnFila + 6, 15), xlHoja1.Cells(pnFila + 6, 15)).Formula = "=SUM(" & lsTotRub(6) & ")"

        pnFila = pnFila + 8

    Else   ' CONSUMO
            ' se reporta al final
    End If
    RaiseEvent Progress(CInt(liProd) + 1, 9)
Next liProd

' *** CONSUMO

liProd = 3
' Cabecera del Reporte en Excel
ImpAnexo5Cab pnFila, "2", sFuente, dFecha, , , psNomCmac, psCodCmac, "5"

    xlHoja1.Cells(pnFila, 1) = "CONSUMO"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 15)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 15)).HorizontalAlignment = xlCenter
    xlHoja1.Cells(pnFila + 1, 1) = "Monto"
    xlHoja1.Cells(pnFila + 2, 1) = "Provision Requerida"
    xlHoja1.Cells(pnFila + 3, 1) = "Provision Constituida"
    xlHoja1.Cells(pnFila + 4, 1) = "Deficit"
    xlHoja1.Cells(pnFila + 5, 1) = "Garantias Preferidas y Preferidas de Muy rapida Realizacion"
    xlHoja1.Cells(pnFila + 6, 1) = "Numero de Deudores"

    For liCatCal = 0 To 4

            Sql = "SELECT Count (Distinct (CA.cPersCod) ) as Deudores, "
            Sql = Sql & "           Count (CA.cCtaCod ) AS NroCred, "
            Sql = Sql & "           ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nSaldoCap) "
            Sql = Sql & "                            WHEN CA.cCtaCod like '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
            Sql = Sql & "                   END ), 0 ) AS SaldoCap, "
            Sql = Sql & "           ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
            Sql = Sql & "                            WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
            Sql = Sql & "                   END ), 0 ) AS Provision, "
            Sql = Sql & "  GarantPrefRapRea = ISNULL(SUM(ISNULL(G.nMontoGar,0)),0)  "
            Sql = Sql & " FROM ColocCalifProv CA "
             Sql = Sql & "  LEFT JOIN ( "
                Sql = Sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                Sql = Sql & "       "
                Sql = Sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                Sql = Sql & "        Where G.nGarClase = 1 AND G.nGarTpoRealiz = 1 "
                Sql = Sql & "        Group by GC.cCtaCod  "
                Sql = Sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
            Sql = Sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "'"
            Sql = Sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan

            'rsP.CursorLocation = adUseClient
            'rsP.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
            'Set rsP.ActiveConnection = Nothing
                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(Sql)
                Conecta.CierraConexion
            xlHoja1.Cells(pnFila + 1, lnCol + liCatCal) = Format(rsP!SaldoCap, "#0.00")
            xlHoja1.Cells(pnFila + 2, lnCol + liCatCal) = Format(rsP!Provision, "#0.00")
            xlHoja1.Cells(pnFila + 3, lnCol + liCatCal) = Format(rsP!Provision, "#0.00")
            xlHoja1.Cells(pnFila + 4, lnCol + liCatCal) = "0.00"
            xlHoja1.Cells(pnFila + 5, lnCol + liCatCal) = Format(rsP!GarantPrefRapRea, "#0.00")
            xlHoja1.Cells(pnFila + 6, lnCol + liCatCal) = Format(rsP!Deudores, "#")
            xlHoja1.Range(xlHoja1.Cells(pnFila + 6, lnCol + liCatCal), xlHoja1.Cells(pnFila + 6, lnCol + liCatCal)).NumberFormat = "#,#0"

            rsP.Close
            Set rsP = Nothing

            'Total por Rubro
            If liCatCal = 0 Then
                lsTotRub(1) = xlHoja1.Range(xlHoja1.Cells(pnFila + 1, lnCol + liCatCal), xlHoja1.Cells(pnFila + 1, lnCol + liCatCal)).Address(False, False)
                lsTotRub(2) = xlHoja1.Range(xlHoja1.Cells(pnFila + 2, lnCol + liCatCal), xlHoja1.Cells(pnFila + 2, lnCol + liCatCal)).Address(False, False)
                lsTotRub(3) = xlHoja1.Range(xlHoja1.Cells(pnFila + 3, lnCol + liCatCal), xlHoja1.Cells(pnFila + 3, lnCol + liCatCal)).Address(False, False)
                lsTotRub(4) = xlHoja1.Range(xlHoja1.Cells(pnFila + 4, lnCol + liCatCal), xlHoja1.Cells(pnFila + 4, lnCol + liCatCal)).Address(False, False)
                lsTotRub(5) = xlHoja1.Range(xlHoja1.Cells(pnFila + 5, lnCol + liCatCal), xlHoja1.Cells(pnFila + 5, lnCol + liCatCal)).Address(False, False)
                lsTotRub(6) = xlHoja1.Range(xlHoja1.Cells(pnFila + 6, lnCol + liCatCal), xlHoja1.Cells(pnFila + 6, lnCol + liCatCal)).Address(False, False)
            Else
                lsTotRub(1) = lsTotRub(1) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 1, lnCol + liCatCal), xlHoja1.Cells(pnFila + 1, lnCol + liCatCal)).Address(False, False)
                lsTotRub(2) = lsTotRub(2) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 2, lnCol + liCatCal), xlHoja1.Cells(pnFila + 2, lnCol + liCatCal)).Address(False, False)
                lsTotRub(3) = lsTotRub(3) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 3, lnCol + liCatCal), xlHoja1.Cells(pnFila + 3, lnCol + liCatCal)).Address(False, False)
                lsTotRub(4) = lsTotRub(4) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 4, lnCol + liCatCal), xlHoja1.Cells(pnFila + 4, lnCol + liCatCal)).Address(False, False)
                lsTotRub(5) = lsTotRub(5) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 5, lnCol + liCatCal), xlHoja1.Cells(pnFila + 5, lnCol + liCatCal)).Address(False, False)
                lsTotRub(6) = lsTotRub(6) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 6, lnCol + liCatCal), xlHoja1.Cells(pnFila + 6, lnCol + liCatCal)).Address(False, False)
            End If


            'Para Tot x Calificacion
            lsTotxCatMonto(liCatCal) = lsTotxCatMonto(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 1, lnCol + liCatCal), xlHoja1.Cells(pnFila + 1, lnCol + liCatCal)).Address(False, False)
            lsTotxCatProviR(liCatCal) = lsTotxCatProviR(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 2, lnCol + liCatCal), xlHoja1.Cells(pnFila + 2, lnCol + liCatCal)).Address(False, False)
            lsTotxCatProviC(liCatCal) = lsTotxCatProviC(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 3, lnCol + liCatCal), xlHoja1.Cells(pnFila + 3, lnCol + liCatCal)).Address(False, False)
            lsTotxCatDefic(liCatCal) = lsTotxCatDefic(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 4, lnCol + liCatCal), xlHoja1.Cells(pnFila + 4, lnCol + liCatCal)).Address(False, False)
            lsTotxCatGarant(liCatCal) = lsTotxCatGarant(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 5, lnCol + liCatCal), xlHoja1.Cells(pnFila + 5, lnCol + liCatCal)).Address(False, False)
            lsTotxCatNroDeu(liCatCal) = lsTotxCatNroDeu(liCatCal) & "+" & xlHoja1.Range(xlHoja1.Cells(pnFila + 6, lnCol + liCatCal), xlHoja1.Cells(pnFila + 6, lnCol + liCatCal)).Address(False, False)

    RaiseEvent Progress(CInt(liCatCal) + 4, 9)
    Next liCatCal

    'Imprime totales
    xlHoja1.Range(xlHoja1.Cells(pnFila + 1, 7), xlHoja1.Cells(pnFila + 1, 7)).Formula = "=SUM(" & lsTotRub(1) & ")"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 7), xlHoja1.Cells(pnFila + 2, 7)).Formula = "=SUM(" & lsTotRub(2) & ")"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 3, 7), xlHoja1.Cells(pnFila + 3, 7)).Formula = "=SUM(" & lsTotRub(3) & ")"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 4, 7), xlHoja1.Cells(pnFila + 4, 7)).Formula = "=SUM(" & lsTotRub(4) & ")"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 5, 7), xlHoja1.Cells(pnFila + 5, 7)).Formula = "=SUM(" & lsTotRub(5) & ")"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 6, 7), xlHoja1.Cells(pnFila + 6, 7)).Formula = "=SUM(" & lsTotRub(6) & ")"


    pnFila = pnFila + 8


' *** TOTAL

' Cabecera del Reporte en Excel
ImpAnexo5Cab pnFila, "3", sFuente, dFecha, , , psNomCmac, psCodCmac, "5"
pnFila = pnFila + 2

lsTotGeneral(1) = "": lsTotGeneral(2) = "": lsTotGeneral(3) = "": lsTotGeneral(4) = "": lsTotGeneral(5) = "": lsTotGeneral(6) = ""


nTotGarPrefReal = 0

For liCatCal = 0 To 4

        Sql = " Select ISNULL(SUM(CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END),0) as nGarMonto "
        Sql = Sql & " from ColocCalifProv CA Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cCtaCod = CA.cCtaCod "
        Sql = Sql & " Inner Join " & ServerCons & "GarantiasConsol G ON G.cNumGarant = GC.cNumGarant"
        Sql = Sql & " Where G.nGarClase = 1 AND CA.cCalgen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan
        Conecta.AbreConexion
        Set rsP = Conecta.CargaRecordSet(Sql)
        Conecta.CierraConexion
        nTotGarPrefReal = nTotGarPrefReal + rsP!nGarMonto

        'Imprime totales x Calific
        xlHoja1.Range(xlHoja1.Cells(pnFila + liCatCal, 2), xlHoja1.Cells(pnFila + liCatCal, 2)).Formula = "=SUM(" & Mid(lsTotxCatNroDeu(liCatCal), 2) & ")"
        xlHoja1.Range(xlHoja1.Cells(pnFila + liCatCal, 2), xlHoja1.Cells(pnFila + liCatCal, 2)).NumberFormat = "#,#0"
        xlHoja1.Range(xlHoja1.Cells(pnFila + liCatCal, 3), xlHoja1.Cells(pnFila + liCatCal, 3)).Formula = "=SUM(" & Mid(lsTotxCatMonto(liCatCal), 2) & ")"
        xlHoja1.Range(xlHoja1.Cells(pnFila + liCatCal, 4), xlHoja1.Cells(pnFila + liCatCal, 4)).Formula = "=SUM(" & Mid(lsTotxCatProviR(liCatCal), 2) & ")"
        xlHoja1.Range(xlHoja1.Cells(pnFila + liCatCal, 5), xlHoja1.Cells(pnFila + liCatCal, 5)).Formula = "=SUM(" & Mid(lsTotxCatProviC(liCatCal), 2) & ")"
        xlHoja1.Range(xlHoja1.Cells(pnFila + liCatCal, 6), xlHoja1.Cells(pnFila + liCatCal, 6)).Formula = "=SUM(" & Mid(lsTotxCatDefic(liCatCal), 2) & ")"

        xlHoja1.Range(xlHoja1.Cells(pnFila + liCatCal, 8), xlHoja1.Cells(pnFila + liCatCal, 8)).value = Format(rsP!nGarMonto, "#0.00")
        xlHoja1.Range(xlHoja1.Cells(pnFila + liCatCal, 9), xlHoja1.Cells(pnFila + liCatCal, 9)).Formula = "=SUM(" & Mid(lsTotxCatGarant(liCatCal), 2) & ")"

        'Asigno los valores al total general
        lsTotGeneral(1) = lsTotGeneral(1) & IIf(Trim(lsTotGeneral(1)) = "", "", "+") & Mid(lsTotxCatNroDeu(liCatCal), 2)
        lsTotGeneral(2) = lsTotGeneral(2) & IIf(Trim(lsTotGeneral(2)) = "", "", "+") & Mid(lsTotxCatMonto(liCatCal), 2)
        lsTotGeneral(3) = lsTotGeneral(3) & IIf(Trim(lsTotGeneral(3)) = "", "", "+") & Mid(lsTotxCatProviR(liCatCal), 2)
        lsTotGeneral(4) = lsTotGeneral(4) & IIf(Trim(lsTotGeneral(4)) = "", "", "+") & Mid(lsTotxCatProviC(liCatCal), 2)
        lsTotGeneral(5) = lsTotGeneral(5) & IIf(Trim(lsTotGeneral(5)) = "", "", "+") & Mid(lsTotxCatDefic(liCatCal), 2)
        lsTotGeneral(6) = lsTotGeneral(6) & IIf(Trim(lsTotGeneral(6)) = "", "", "+") & Mid(lsTotxCatGarant(liCatCal), 2)

        rsP.Close
Next
pnFila = pnFila + 5
'Imprime totales
xlHoja1.Range(xlHoja1.Cells(pnFila, 2), xlHoja1.Cells(pnFila, 2)).Formula = "=SUM(" & lsTotGeneral(1) & ")"
xlHoja1.Range(xlHoja1.Cells(pnFila, 2), xlHoja1.Cells(pnFila, 2)).NumberFormat = "#,#0"
xlHoja1.Range(xlHoja1.Cells(pnFila, 3), xlHoja1.Cells(pnFila, 3)).Formula = "=SUM(" & lsTotGeneral(2) & ")"
xlHoja1.Range(xlHoja1.Cells(pnFila, 4), xlHoja1.Cells(pnFila, 4)).Formula = "=SUM(" & lsTotGeneral(3) & ")"
xlHoja1.Range(xlHoja1.Cells(pnFila, 5), xlHoja1.Cells(pnFila, 5)).Formula = "=SUM(" & lsTotGeneral(4) & ")"
xlHoja1.Range(xlHoja1.Cells(pnFila, 6), xlHoja1.Cells(pnFila, 6)).Formula = "=SUM(" & lsTotGeneral(5) & ")"
xlHoja1.Range(xlHoja1.Cells(pnFila, 8), xlHoja1.Cells(pnFila, 8)).value = Format(nTotGarPrefReal, "#0.00")
xlHoja1.Range(xlHoja1.Cells(pnFila, 9), xlHoja1.Cells(pnFila, 9)).Formula = "=SUM(" & lsTotGeneral(6) & ")"

xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 10)).Font.Bold = True
xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 10)).Interior.Color = &HC0C0FF

'*******
RaiseEvent CloseProgress

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
MsgBox "Se genero el archivo " & App.path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"

'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing

lbExcel = False
End Sub

Public Sub Genera_Reporte178208(ByVal TxtCambio As Double, ByVal dFecha As Date, _
                        ByVal nCF As Integer, ByVal psServer As String, _
                        ByVal dFecData As String, _
                        Optional ByVal psNomCmac As String, _
                        Optional ByVal psCodCmac As String)


Dim Sql As String
Dim rs As New ADODB.Recordset
Dim fs As New Scripting.FileSystemObject
Dim tc As Double
Dim i As Integer
Dim J As Integer
Dim k As Integer
Dim lsNomHoja As String
Dim lbExisteHoja As Boolean
Dim SqlM As String
Dim RsM As New ADODB.Recordset
Dim lbMesActual As Boolean
Dim lbIncluyeCF As Boolean
Dim lsArchivo As String
Dim lsServConsol As String
Dim Co As DConecta
Dim RCC As DRCC
Set RCC = New DRCC
lsServConsol = RCC.ServConsol(psServer)
Set RCC = Nothing

'Tipo de Cambio de Pantalla
tc = TxtCambio
Set Co = New DConecta
'--------
If dFecha = dFecData Then
    lbMesActual = True
Else
    lbMesActual = False
End If
'
If nCF = 1 Then
    lbIncluyeCF = True
Else
    lbIncluyeCF = False
End If

lsArchivo = "Rep_Anexo05_" & Format(dFecha, "YYYYMM") & IIf(lbIncluyeCF, "", "SinCF") & ".XLS"

Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

 lsNomHoja = "Anexo5"
 For Each xlHoja1 In xlLibro.Worksheets
 If xlHoja1.Name = lsNomHoja Then
    xlHoja1.Activate
    xlHoja1.Range("A1", "AZ10000") = ""
    xlHoja1.Range("A1", "AZ10000").Rows.Delete
    lbExisteHoja = True
    Exit For
 End If
 Next

If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

 xlHoja1.Range("A1:AZ1000") = ""
 xlHoja1.Range("A2") = "ANEXO N° 5"
 xlHoja1.Range("A3") = "INFORME DE CLASIFICACION DE DEUDORES Y PROVISIONES"
 xlHoja1.Range("A6") = "EMPRESA : " & psNomCmac
 xlHoja1.Range("G6") = "CODIGO : " & psCodCmac
 xlHoja1.Range("A8") = "AL   " & Format(dFecha, "DD") & "  DE " & UCase(Format(dFecha, "MMMM")) & "  DEL " & Format(dFecha, "YYYY")
 xlHoja1.Range("A9") = "(En miles de nuevos soles)"
 xlHoja1.Range("A11") = "I.- INFORME DE CLASIFICACION DE LOS DEUDORES DE LA CARTERA DE CRÉDITOS, CONTINGENCIASY ARRENDAMIENTOS FINANCIEROS"

 xlHoja1.Range("A14") = "A. MONTO DE LOS CREDITOS"
 xlHoja1.Range("A15") = "CONTINGENTES Y ARRENDAMIENTOS"
 xlHoja1.Range("A16") = "FINANCIEROS 2/"
 xlHoja1.Range("B16") = "NORMAL"
 xlHoja1.Range("C16") = "CPP"
 xlHoja1.Range("D16") = "DEFICIENTE"
 xlHoja1.Range("E16") = "DUDOSO"
 xlHoja1.Range("F16") = "PERDIDA"
 xlHoja1.Range("G16") = "TOTAL"
 xlHoja1.Range("A16:G16").Font.Bold = True


 xlHoja1.Range("A17") = " Comercial"
 xlHoja1.Range("A18") = " Mes"
 xlHoja1.Range("A19") = " Hipotecario para vivienda"
 xlHoja1.Range("A20") = " Consumo"

 xlHoja1.Range("A21") = "B. NUMERO DE DEUDORES"
 xlHoja1.Range("B21") = "NORMAL"
 xlHoja1.Range("C21") = "CPP"
 xlHoja1.Range("D21") = "DEFICIENTE"
 xlHoja1.Range("E21") = "DUDOSO"
 xlHoja1.Range("F21") = "PERDIDA"
 xlHoja1.Range("G21") = "TOTAL"
 xlHoja1.Range("A21:G21").Font.Bold = True

 xlHoja1.Range("A22") = " Comercial"
 xlHoja1.Range("A23") = " Mes"
 xlHoja1.Range("A24") = " Hipotecario para vivienda"
 xlHoja1.Range("A25") = " Consumo"
 xlHoja1.Range("A26") = "TOTAL 4/"

 xlHoja1.Range("A27") = "C.  MONTO DE LOS CREDITOS,"
 xlHoja1.Range("A28") = "CONTINGENTES Y ARRENDAMIENTOS"
 xlHoja1.Range("A29") = "FINANCIERO EFECTO A SUSTITUCION"
 xlHoja1.Range("A30") = "DE CONTRAPARTE CREDITICIA 5/"
 xlHoja1.Range("A31") = " Comercial"
 xlHoja1.Range("A32") = " Mes"
 xlHoja1.Range("A33") = " Hipotecario para vivienda"
 xlHoja1.Range("A34") = " Consumo"
 xlHoja1.Range("A30").Font.Bold = True
 xlHoja1.Range("B30") = "NORMAL"
 xlHoja1.Range("C30") = "CPP"
 xlHoja1.Range("D30") = "DEFICIENTE"
 xlHoja1.Range("E30") = "DUDOSO"
 xlHoja1.Range("F30") = "PERDIDA"
 xlHoja1.Range("G30") = "TOTAL"
 xlHoja1.Range("B30:G30").Font.Bold = True


 xlHoja1.Range("A35") = "D. MONTO DE LOS CREDITOS,"
 xlHoja1.Range("A36") = "CONTINGENTES Y ARRENDAMIENTOS"
 xlHoja1.Range("A37") = "FINANCIEROS QUE CUENTAN CON"
 xlHoja1.Range("A38") = "GARANTIAS PREFERIDAS"
 xlHoja1.Range("A39") = "AUTOLIQUIDABLES 6/"
 xlHoja1.Range("A40") = " Comercial"
 xlHoja1.Range("A41") = " Mes"
 xlHoja1.Range("A42") = " Hipotecario para vivienda"
 xlHoja1.Range("A43") = " Consumo"

 xlHoja1.Range("B39") = "NORMAL"
 xlHoja1.Range("C39") = "CPP"
 xlHoja1.Range("D39") = "DEFICIENTE"
 xlHoja1.Range("E39") = "DUDOSO"
 xlHoja1.Range("F39") = "PERDIDA"
 xlHoja1.Range("G39") = "TOTAL"
 xlHoja1.Range("B39:G39").Font.Bold = True

 xlHoja1.Range("A44") = "E. MONTO DE LOS CREDITOS,"
 xlHoja1.Range("A45") = "CONTINGENTES Y ARRENDAMIENTOS"
 xlHoja1.Range("A46") = "FINANCIEROS QUE CUENTAN CON"
 xlHoja1.Range("A47") = "GARANTIAS PREFERIDAS DE MUY"
 xlHoja1.Range("A48") = "RAPIDA REALIZACION 7/"
 xlHoja1.Range("A49") = " Comercial"
 xlHoja1.Range("A50") = " Mes"
 xlHoja1.Range("A51") = " Hipotecario para vivienda"


 xlHoja1.Range("B48") = "NORMAL"
 xlHoja1.Range("C48") = "CPP"
 xlHoja1.Range("D48") = "DEFICIENTE"
 xlHoja1.Range("E48") = "DUDOSO"
 xlHoja1.Range("F48") = "PERDIDA"
 xlHoja1.Range("G48") = "TOTAL"
 xlHoja1.Range("B48:G48").Font.Bold = True

 xlHoja1.Range("A52") = "F. MONTO DE LOS CREDITOS,"
 xlHoja1.Range("A53") = "CONTINGENTES Y ARRENDAMIENTOS"
 xlHoja1.Range("A54") = "FINANCIEROS QUE CUENTAN CON"
 xlHoja1.Range("A55") = "GARANTIAS PREFERIDAS 8/"
 xlHoja1.Range("A56") = " Comercial"
 xlHoja1.Range("A57") = " Mes"
 xlHoja1.Range("A58") = " Hipotecario para vivienda"

 xlHoja1.Range("B55") = "NORMAL"
 xlHoja1.Range("C55") = "CPP"
 xlHoja1.Range("D55") = "DEFICIENTE"
 xlHoja1.Range("E55") = "DUDOSO"
 xlHoja1.Range("F55") = "PERDIDA"
 xlHoja1.Range("G55") = "TOTAL"
 xlHoja1.Range("B55:G55").Font.Bold = True

 xlHoja1.Range("A59") = "G. MONTO DE LOS FINANCIAMIENTOS,"
 xlHoja1.Range("A60") = "DE CORTO PLAZO 9/"
 xlHoja1.Range("A61") = " Comercial"
 xlHoja1.Range("A62") = " Mes"


 xlHoja1.Range("B60") = "NORMAL"
 xlHoja1.Range("C60") = "CPP"
 xlHoja1.Range("D60") = "DEFICIENTE"
 xlHoja1.Range("E60") = "DUDOSO"
 xlHoja1.Range("F60") = "PERDIDA"
 xlHoja1.Range("G60") = "TOTAL"
 xlHoja1.Range("B60:G60").Font.Bold = True


 xlHoja1.Range("A63") = "H. CREDITOS, CONTINGENTES Y"
 xlHoja1.Range("A64") = "ARRENDAMIENTOS FINANCIEROS "
 xlHoja1.Range("A65") = "COMERCIALES SUJETOS AL REGIMEN"
 xlHoja1.Range("A66") = "DE PROVISIONES GENERICAS"
 xlHoja1.Range("A67") = "BASADO EN METODOLOGIAS"
 xlHoja1.Range("A68") = "INTERNAS 10/"

 xlHoja1.Range("B68") = "NORMAL"
 xlHoja1.Range("B68").Font.Bold = True

 xlHoja1.Range("A69") = " Monto"
 xlHoja1.Range("A70") = " Tasa de Provision Impírica"

 xlHoja1.Range("A71") = "I. CREDITOS HIPOTECARIOS PARA"
 xlHoja1.Range("A72") = "VIVIENDA - MIVIVIENDA 11/"

 xlHoja1.Range("B72") = "NORMAL"
 xlHoja1.Range("C72") = "CPP"
 xlHoja1.Range("D72") = "DEFICIENTE"
 xlHoja1.Range("E72") = "DUDOSO"
 xlHoja1.Range("F72") = "PERDIDA"
 xlHoja1.Range("G72") = "TOTAL"
 xlHoja1.Range("B72:G72").Font.Bold = True

 xlHoja1.Range("A73") = " Cobertura del Fondo Mivivienda 12/"
 xlHoja1.Range("A74") = " Cobertura Garantía autoliquidable 13/"
 xlHoja1.Range("A75") = " Monto afecto de sustitución de contraparte"
 xlHoja1.Range("A76") = " Crediticia 14/"
 xlHoja1.Range("A77") = " Cobertura GPMRR 15/"
 xlHoja1.Range("A78") = " Cobertura GP 16/"
 xlHoja1.Range("A79") = " Sin Cobertura 17/"
 xlHoja1.Range("A80") = " Cobertura GPMRR 18/"
 xlHoja1.Range("A81") = " Cobertura GP 18/"
 xlHoja1.Range("A82") = " Sin Cobertura 18/"

 xlHoja1.Range("A83") = "J. MONTO DE LOS CREDITOS"
 xlHoja1.Range("A84") = "CONTINGENTES Y ARRENDAMIENTOS"
 xlHoja1.Range("A85") = "FINANCIEROS QUE NO CUENTEN CON"
 xlHoja1.Range("A86") = "COBERTURA 19/"
 xlHoja1.Range("A87") = " Comerciales"
 xlHoja1.Range("A88") = " MES"
 xlHoja1.Range("A89") = " Hipotecario para Vivienda"
 xlHoja1.Range("A90") = " Consumo"

 xlHoja1.Range("B86") = "NORMAL"
 xlHoja1.Range("C86") = "CPP"
 xlHoja1.Range("D86") = "DEFICIENTE"
 xlHoja1.Range("E86") = "DUDOSO"
 xlHoja1.Range("F86") = "PERDIDA"
 xlHoja1.Range("G86") = "TOTAL"
 xlHoja1.Range("B86:G86").Font.Bold = True

 xlHoja1.Range("A91") = "K. PROVISIONES CONSTITUIDAS 20/"
 xlHoja1.Range("A92") = " Comerciales"
 xlHoja1.Range("A93") = " MES"
 xlHoja1.Range("A94") = " Hipotecario para Vivienda"
 xlHoja1.Range("A95") = " Consumo"

 xlHoja1.Range("B91") = "NORMAL"
 xlHoja1.Range("C91") = "CPP"
 xlHoja1.Range("D91") = "DEFICIENTE"
 xlHoja1.Range("E91") = "DUDOSO"
 xlHoja1.Range("F91") = "PERDIDA"
 xlHoja1.Range("G91") = "TOTAL"
 xlHoja1.Range("B91:G91").Font.Bold = True

 xlHoja1.Range("A96") = "L. PROVISIONES REQUERIDAS 21/"
 xlHoja1.Range("A97") = " Comerciales"
 xlHoja1.Range("A98") = " MES"
 xlHoja1.Range("A99") = " Hipotecario para Vivienda"
 xlHoja1.Range("A100") = " Consumo"

 xlHoja1.Range("B96") = "NORMAL"
 xlHoja1.Range("C96") = "CPP"
 xlHoja1.Range("D96") = "DEFICIENTE"
 xlHoja1.Range("E96") = "DUDOSO"
 xlHoja1.Range("F96") = "PERDIDA"
 xlHoja1.Range("G96") = "TOTAL"
 xlHoja1.Range("B96:G96").Font.Bold = True

 xlHoja1.Range("A101") = "M. SUPERAVIT (DEFICIT) DE"
 xlHoja1.Range("A102") = "PROVISIONES 22/"
 xlHoja1.Range("A103") = " Comerciales"
 xlHoja1.Range("A104") = " MES"
 xlHoja1.Range("A105") = " Hipotecario para Vivienda"
 xlHoja1.Range("A106") = " Consumo"

 xlHoja1.Range("B102") = "NORMAL"
 xlHoja1.Range("C102") = "CPP"
 xlHoja1.Range("D102") = "DEFICIENTE"
 xlHoja1.Range("E102") = "DUDOSO"
 xlHoja1.Range("F102") = "PERDIDA"
 xlHoja1.Range("G102") = "TOTAL"
 xlHoja1.Range("B102:G102").Font.Bold = True

 xlHoja1.Range("A107") = "TOTAL"
 xlHoja1.Range("A1").ColumnWidth = 45
 xlHoja1.Range("B1").ColumnWidth = 16
 xlHoja1.Range("C1").ColumnWidth = 16
 xlHoja1.Range("D1").ColumnWidth = 16
 xlHoja1.Range("E1").ColumnWidth = 16
 xlHoja1.Range("F1").ColumnWidth = 16
 xlHoja1.Range("G1").ColumnWidth = 16

 xlHoja1.Application.ActiveWindow.Zoom = 75
 xlHoja1.Range("A2:G2").MergeCells = True
 xlHoja1.Range("A3:G3").MergeCells = True
 xlHoja1.Range("A2:G11").Font.Bold = True
 xlHoja1.Range("A2:G3").Font.Size = 12

 xlHoja1.Range("A2").Font.Bold = True
 xlHoja1.Range("A2").HorizontalAlignment = xlCenter

 xlHoja1.Range("A6").Font.Bold = True
 xlHoja1.Range("G6").Font.Bold = True

 xlHoja1.Range("A8:G8").MergeCells = True
 xlHoja1.Range("A8").Font.Bold = True
 xlHoja1.Range("A8").HorizontalAlignment = xlCenter

 xlHoja1.Range("A9:G9").MergeCells = True
 xlHoja1.Range("A9").Font.Bold = True
 xlHoja1.Range("A9").HorizontalAlignment = xlCenter

 xlHoja1.Range("A11:G11").MergeCells = True
 xlHoja1.Range("A11").Font.Bold = True
 xlHoja1.Range("A11").HorizontalAlignment = xlCenter

xlHoja1.Range("A14:A16").Font.Bold = True
xlHoja1.Range("A121").Font.Bold = True
xlHoja1.Range("A22:A29").Font.Bold = True
xlHoja1.Range("A35:A39").Font.Bold = True
xlHoja1.Range("A44:A48").Font.Bold = True
xlHoja1.Range("A52:A55").Font.Bold = True
xlHoja1.Range("A59:A60").Font.Bold = True
xlHoja1.Range("A63:A68").Font.Bold = True
xlHoja1.Range("A71:A72").Font.Bold = True
xlHoja1.Range("A83:A86").Font.Bold = True
xlHoja1.Range("A91").Font.Bold = True
xlHoja1.Range("A96").Font.Bold = True
xlHoja1.Range("A101:A102").Font.Bold = True
xlHoja1.Range("A107").Font.Bold = True

'--------
xlHoja1.Range("A109") = "II.- INFORME DE CLASIFICACION DE LOS DEUDORES DE LA CARTERA TRANSFERIDA SEGUN D.S.N° 114-98-EF 23/"
xlHoja1.Range("A109:G109").MergeCells = True
xlHoja1.Range("A109").Font.Bold = True
xlHoja1.Range("A109").HorizontalAlignment = xlCenter
xlHoja1.Range("A110") = "N.- MONTO PENDIENTE DE"
xlHoja1.Range("A111") = "RECOMPRA 24/"
xlHoja1.Range("B111") = "NORMAL"
xlHoja1.Range("C111") = "CPP"
xlHoja1.Range("D111") = "DEFICIENTE"
xlHoja1.Range("E111") = "DUDOSO"
xlHoja1.Range("F111") = "PERDIDA"
xlHoja1.Range("G111") = "TOTAL"
xlHoja1.Range("A111:G111").Font.Bold = True

xlHoja1.Range("A112") = " Comerciales"
xlHoja1.Range("A113") = " MES"
xlHoja1.Range("A114") = " Hipotecario para Vivienda"
xlHoja1.Range("A115") = " Consumo"
xlHoja1.Range("A116") = "O.- PROVISIONES CONSTITUIDAS 25/"
xlHoja1.Range("B116") = "NORMAL"
xlHoja1.Range("C116") = "CPP"
xlHoja1.Range("D116") = "DEFICIENTE"
xlHoja1.Range("E116") = "DUDOSO"
xlHoja1.Range("F116") = "PERDIDA"
xlHoja1.Range("G116") = "TOTAL"
xlHoja1.Range("A116:G116").Font.Bold = True

xlHoja1.Range("A117") = " Comerciales"
xlHoja1.Range("A118") = " MES"
xlHoja1.Range("A119") = " Hipotecario para Vivienda"
xlHoja1.Range("A120") = " Consumo"
xlHoja1.Range("A121") = "P.- PROVISIONES REQUERIDAS 26/"
xlHoja1.Range("B121") = "NORMAL"
xlHoja1.Range("C121") = "CPP"
xlHoja1.Range("D121") = "DEFICIENTE"
xlHoja1.Range("E121") = "DUDOSO"
xlHoja1.Range("F121") = "PERDIDA"
xlHoja1.Range("G121") = "TOTAL"
xlHoja1.Range("A121:G121").Font.Bold = True

xlHoja1.Range("A122") = " Comerciales"
xlHoja1.Range("A123") = " MES"
xlHoja1.Range("A124") = " Hipotecario para Vivienda"
xlHoja1.Range("A125") = " Consumo"
xlHoja1.Range("A126") = "Q.- SUPERAVIT (DEFICIT DE)"
xlHoja1.Range("A127") = "PROVISIONES 27/"
xlHoja1.Range("B127") = "NORMAL"
xlHoja1.Range("C127") = "CPP"
xlHoja1.Range("D127") = "DEFICIENTE"
xlHoja1.Range("E127") = "DUDOSO"
xlHoja1.Range("F127") = "PERDIDA"
xlHoja1.Range("G127") = "TOTAL"
xlHoja1.Range("B127:G127").Font.Bold = True
xlHoja1.Range("A127").Font.Bold = True
xlHoja1.Range("A128") = " Comerciales"
xlHoja1.Range("A129") = " MES"
xlHoja1.Range("A130") = " Hipotecario para Vivienda"
xlHoja1.Range("A131") = " Consumo"
xlHoja1.Range("A131") = "TOTAL"
xlHoja1.Range("A131").Font.Bold = True


xlHoja1.Range("A133") = "II.- INFORME DE CLASIFICACION DE LOS DEUDORES DE LA CARTERA TRANSFERIDA SEGUN D.S.N° 114-98-EF 23/"
xlHoja1.Range("A133:G133").MergeCells = True
xlHoja1.Range("A133").Font.Bold = True
xlHoja1.Range("A133").HorizontalAlignment = xlCenter
xlHoja1.Range("A134") = "N.- MONTO PENDIENTE DE"
xlHoja1.Range("A135") = "RECOMPRA 24/"
xlHoja1.Range("B135") = "NORMAL"
xlHoja1.Range("C135") = "CPP"
xlHoja1.Range("D135") = "DEFICIENTE"
xlHoja1.Range("E135") = "DUDOSO"
xlHoja1.Range("F135") = "PERDIDA"
xlHoja1.Range("G135") = "TOTAL"
xlHoja1.Range("B135:G135").Font.Bold = True
xlHoja1.Range("A134:A135").Font.Bold = True
xlHoja1.Range("A136") = " Comerciales"
xlHoja1.Range("A137") = " MES"
xlHoja1.Range("A138") = " Hipotecario para Vivienda"
xlHoja1.Range("A139") = " Consumo"
xlHoja1.Range("A140") = "O.- PROVISIONES CONSTITUIDAS 25/"
xlHoja1.Range("B140") = "NORMAL"
xlHoja1.Range("C140") = "CPP"
xlHoja1.Range("D140") = "DEFICIENTE"
xlHoja1.Range("E140") = "DUDOSO"
xlHoja1.Range("F140") = "PERDIDA"
xlHoja1.Range("G140") = "TOTAL"
xlHoja1.Range("B140:G140").Font.Bold = True
xlHoja1.Range("A140").Font.Bold = True
xlHoja1.Range("A141") = " Comerciales"
xlHoja1.Range("A142") = " MES"
xlHoja1.Range("A143") = " Hipotecario para Vivienda"
xlHoja1.Range("A144") = " Consumo"
xlHoja1.Range("A145") = "P.- PROVISIONES REQUERIDAS 26/"
xlHoja1.Range("B145") = "NORMAL"
xlHoja1.Range("C145") = "CPP"
xlHoja1.Range("D145") = "DEFICIENTE"
xlHoja1.Range("E145") = "DUDOSO"
xlHoja1.Range("F145") = "PERDIDA"
xlHoja1.Range("G145") = "TOTAL"
xlHoja1.Range("B145:G145").Font.Bold = True
xlHoja1.Range("A145").Font.Bold = True
xlHoja1.Range("A146") = " Comerciales"
xlHoja1.Range("A147") = " MES"
xlHoja1.Range("A148") = " Hipotecario para Vivienda"
xlHoja1.Range("A149") = " Consumo"
xlHoja1.Range("A150") = "Q.- SUPERAVIT (DEFICIT DE)"
xlHoja1.Range("A151") = "PROVISIONES 27/"
xlHoja1.Range("B151") = "NORMAL"
xlHoja1.Range("C151") = "CPP"
xlHoja1.Range("D151") = "DEFICIENTE"
xlHoja1.Range("E151") = "DUDOSO"
xlHoja1.Range("F151") = "PERDIDA"
xlHoja1.Range("G151") = "TOTAL"
xlHoja1.Range("B151:G151").Font.Bold = True
xlHoja1.Range("A150:A151").Font.Bold = True
xlHoja1.Range("A152") = " Comerciales"
xlHoja1.Range("A153") = " MES"
xlHoja1.Range("A154") = " Hipotecario para Vivienda"
xlHoja1.Range("A155") = " Consumo"
xlHoja1.Range("A156") = "TOTAL"
xlHoja1.Range("A156").Font.Bold = True


xlHoja1.Range("A159") = "IV.- CUADRE DEL ANEXO N°5 CON LAS CIFRAS DEL BALANCE 33/"
xlHoja1.Range("A159:G159").MergeCells = True
xlHoja1.Range("A159").HorizontalAlignment = xlCenter
xlHoja1.Range("A159").Font.Bold = True
xlHoja1.Range("A161") = "V.- CIFRAS DE BALANCE"
xlHoja1.Range("A161:B161").MergeCells = True
xlHoja1.Range("A161").Font.Bold = True
xlHoja1.Range("A161").HorizontalAlignment = xlCenter

xlHoja1.Range("D161") = "W.- ANEXO N°5"
xlHoja1.Range("C161:E161").MergeCells = True
xlHoja1.Range("D161").Font.Bold = True
xlHoja1.Range("D161").HorizontalAlignment = xlCenter

xlHoja1.Range("A163") = "MONTO CREDITOS, CONTINGENTE Y ARRENDAMIENTOS"
xlHoja1.Range("A164") = "FINANCIEROS"

xlHoja1.Range("B163") = "PROVISIONES"
xlHoja1.Range("B164") = "GENERICAS"

xlHoja1.Range("C163") = "PROVISIONES"
xlHoja1.Range("C164") = "ESPECIFICAS"

xlHoja1.Range("D162") = "MONTO CREDITOS"
xlHoja1.Range("D163") = "CONTINGENTES Y"
xlHoja1.Range("D164") = "ARRENDAMIENTO"
xlHoja1.Range("D165") = "FINANCIEROS"

xlHoja1.Range("E163") = "PROVISIONES "
xlHoja1.Range("E164") = "GENERICAS"
xlHoja1.Range("E165") = "CONSTITUIDAS"

xlHoja1.Range("E163") = "PROVISIONES "
xlHoja1.Range("E164") = "ESPECIFICAS"
xlHoja1.Range("E165") = "CONSTITUIDAS"

xlHoja1.Range("A162:F165").HorizontalAlignment = xlCenter
xlHoja1.Range("A162:F165").Font.Bold = True
xlHoja1.Range("A161:F165").Font.Size = 8

 xlHoja1.Range("B170").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("B170").Borders(xlEdgeTop).Weight = xlThin
xlHoja1.Range("B170:B171").HorizontalAlignment = xlCenter
xlHoja1.Range("B170") = "GERENTE"
xlHoja1.Range("B171") = "GENERAL"

xlHoja1.Range("D170").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("D170").Borders(xlEdgeTop).Weight = xlThin
xlHoja1.Range("D170:D171").HorizontalAlignment = xlCenter
xlHoja1.Range("D170") = "CONTADOR"
xlHoja1.Range("D171") = "GENERAL"

xlHoja1.Range("F170").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("F170").Borders(xlEdgeTop).Weight = xlThin
xlHoja1.Range("F170:F171").HorizontalAlignment = xlCenter
xlHoja1.Range("F170") = "FUNCIONARIO"
xlHoja1.Range("F171") = "RESPONSABLE"

'--------


Sql = " select substring(cCtaCod,6,1) cta, cCalGen ,"
Sql = Sql & " Isnull(sum(case when substring(cCtaCod,9,1) = '1' then nSaldoCap"
Sql = Sql & "          when substring(cCtaCod,9,1) = '2' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCap,"
Sql = Sql & " Count( distinct(cPersCod) ) Deudores,"
Sql = Sql & " Isnull(sum(case when nGarant = 1 and substring(cCtaCod,9,1) = '1' and substring(cCtaCod,6,3) <> '423' then nSaldoCap"
Sql = Sql & "          when nGarant = 1 and substring(cCtaCod,9,1) = '2' and substring(cCtaCod,6,3) <> '423' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCapGarAutoLiq,"
Sql = Sql & " Isnull(sum(case when nGarant = 2 and substring(cCtaCod,9,1) = '1' and substring(cCtaCod,6,3) <> '423' then nSaldoCap"
Sql = Sql & "          when nGarant = 2 and substring(cCtaCod,9,1) = '2' and substring(cCtaCod,6,3) <> '423' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCapGarMuyRR,"
Sql = Sql & " IsNull(sum(case when nGarant = 3 and substring(cCtaCod,9,1) = '1' and substring(cCtaCod,6,3) <> '423' then nSaldoCap"
Sql = Sql & "          when nGarant = 3 and substring(cCtaCod,9,1) = '2' and substring(cCtaCod,6,3) <> '423' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCapGarPref,"
Sql = Sql & " Isnull(sum(case when substring(cCtaCod,6,3) = '423' And substring(cCtaCod,9,1) = '1' then nSaldoCap"
Sql = Sql & "          when substring(cCtaCod,6,3) = '423' And substring(cCtaCod,9,1) = '2' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCapMiVivienda,"
Sql = Sql & " Isnull(sum(case when nGarant = 4 and substring(cCtaCod,9,1) = '1' and substring(cCtaCod,6,3) <> '423' then nSaldoCap"
Sql = Sql & "          when nGarant = 4 and substring(cCtaCod,9,1) = '2' and substring(cCtaCod,6,3) <> '423' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCapSinG,"
Sql = Sql & " Isnull(sum(case when substring(cCtaCod,9,1) = '1'  then nProvision"
Sql = Sql & "          when substring(cCtaCod,9,1) = '2'  then nProvision * " & tc
Sql = Sql & "  end ),0) ProvConstituida"
If lbMesActual = True Then
    Sql = Sql & " FROM " & lsServConsol & "ColocCalifProv ca Where"
Else
    Sql = Sql & " FROM " & lsServConsol & "ColocCalifProvTotal ca"
    Sql = Sql & " Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
End If
Sql = Sql & " nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2101,2104,2106,2107)"

If lbIncluyeCF = False Then
    Sql = Sql & " And substring(ca.cCtaCod,7,2) <> '21' "
End If
Sql = Sql & " Group by substring(cCtaCod,6,1), cCalGen"
Sql = Sql & " Order by substring(cCtaCod,6,1), cCalGen"
Co.AbreConexion
Set rs = Co.CargaRecordSet(Sql)
While Not rs.EOF
    Select Case rs!Cta
        Case "1"
            Select Case rs!cCalGen
                Case "0": xlHoja1.Range("B17") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("B22") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("B31") = 0
                          xlHoja1.Range("B40") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("B56") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("B87") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("B92") = Format(rs!ProvConstituida, "#.000")

                Case "1": xlHoja1.Range("C17") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("C22") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("C31") = 0
                          xlHoja1.Range("C40") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("C56") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("C87") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("C92") = Format(rs!ProvConstituida, "#.000")

                Case "2": xlHoja1.Range("D17") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("D22") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("D31") = 0
                          xlHoja1.Range("D40") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("D56") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("D87") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("D92") = Format(rs!ProvConstituida, "#.000")

                Case "3": xlHoja1.Range("E17") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("E22") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("E31") = 0
                          xlHoja1.Range("E40") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("E56") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("E87") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("E92") = Format(rs!ProvConstituida, "#.000")


                Case "4": xlHoja1.Range("F17") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("F22") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("F31") = 0
                          xlHoja1.Range("F40") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("F56") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("F87") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("F92") = Format(rs!ProvConstituida, "#.000")
            End Select
        Case "2"
            Select Case rs!cCalGen
                Case "0": xlHoja1.Range("B18") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("B23") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("B41") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("B57") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("B88") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("B93") = Format(rs!ProvConstituida, "#.000")

                Case "1": xlHoja1.Range("C18") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("C23") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("C41") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("C57") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("C88") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("C93") = Format(rs!ProvConstituida, "#.000")

                Case "2": xlHoja1.Range("D18") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("D23") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("D41") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("D57") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("D88") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("D93") = Format(rs!ProvConstituida, "#.000")

                Case "3": xlHoja1.Range("E18") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("E23") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("E41") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("E57") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("E88") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("E93") = Format(rs!ProvConstituida, "#.000")

                Case "4": xlHoja1.Range("F18") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("F23") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("F41") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("F57") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("F88") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("F93") = Format(rs!ProvConstituida, "#.000")
            End Select
        Case "3"
            Select Case rs!cCalGen
                Case "0": xlHoja1.Range("B20") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("B25") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("B43") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("B90") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("B95") = Format(rs!ProvConstituida, "#.000")

                Case "1": xlHoja1.Range("C20") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("C25") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("C43") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("C90") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("C95") = Format(rs!ProvConstituida, "#.000")

                Case "2": xlHoja1.Range("D20") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("D25") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("D43") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("D90") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("D95") = Format(rs!ProvConstituida, "#.000")

                Case "3": xlHoja1.Range("E20") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("E25") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("E43") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("E90") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("E95") = Format(rs!ProvConstituida, "#.000")

                Case "4": xlHoja1.Range("F20") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("F25") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("F43") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("F90") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("F95") = Format(rs!ProvConstituida, "#.000")
            End Select

        Case "4"

            Select Case rs!cCalGen
                Case "0": xlHoja1.Range("B19") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("B24") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("B42") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("B58") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("B89") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("B94") = Format(rs!ProvConstituida, "#.000")

                Case "1": xlHoja1.Range("C19") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("C24") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("C42") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("C58") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("C89") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("C94") = Format(rs!ProvConstituida, "#.000")

                Case "2": xlHoja1.Range("D19") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("D24") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("D42") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("D58") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("D89") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("D94") = Format(rs!ProvConstituida, "#.000")

                Case "3": xlHoja1.Range("E19") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("E24") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("E42") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("E58") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("E89") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("E94") = Format(rs!ProvConstituida, "#.000")

                Case "4": xlHoja1.Range("F19") = Format(rs!SaldoCap, "#.00")
                          xlHoja1.Range("F24") = Format(rs!Deudores, "#0")
                          xlHoja1.Range("F42") = Format(rs!SaldoCapGarAutoLiq, "#.00")
                          xlHoja1.Range("F58") = Format(rs!SaldoCapGarPref, "#.00")
                          xlHoja1.Range("F89") = Format(rs!SaldoCapSinG, "#.00")
                          xlHoja1.Range("F94") = Format(rs!ProvConstituida, "#.000")
            End Select

    End Select
    rs.MoveNext
 Wend

Sql = " select substring(cCtaCod,6,1), cCalGen ,"
Sql = Sql & " Isnull(sum(case when substring(cCtaCod,9,1) = '1' then nSaldoCap"
Sql = Sql & "   when substring(cCtaCod,9,1) = '2' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCap,"
Sql = Sql & " Count( distinct(cPersCod) ) Deudores,"
Sql = Sql & " Isnull(sum(case when nGarant = 1 and substring(cCtaCod,9,1) = '1'  then nSaldoCap"
Sql = Sql & "   when nGarant = 1 and substring(cCtaCod,9,1) = '2'  then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCapGarAutoLiq,"
Sql = Sql & " Isnull(sum(case when nGarant = 2 and substring(cCtaCod,9,1) = '1'  then nSaldoCap"
Sql = Sql & "   when nGarant = 2 and substring(cCtaCod,9,1) = '2'  then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCapGarMuyRR,"
Sql = Sql & " Isnull(sum(case when nGarant = 3 and substring(cCtaCod,9,1) = '1'  then nSaldoCap"
Sql = Sql & "   when nGarant = 3 and substring(cCtaCod,9,1) = '2' then nSaldoCap * " & tc
Sql = Sql & " end ),0) SaldoCapGarPref,"
Sql = Sql & " Isnull(sum(case when  substring(cCtaCod,9,1) = '1' then nSaldoCap"
Sql = Sql & "   when  substring(cCtaCod,9,1) = '2' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCapMiVivienda,"
Sql = Sql & " Isnull(sum(case when nGarant = 4 and substring(cCtaCod,9,1) = '1'  then nSaldoCap"
Sql = Sql & "   when nGarant = 4 and substring(cCtaCod,9,1) = '2' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCapSinG,"
Sql = Sql & " Isnull(sum(case when substring(cCtaCod,9,1) = '1'  then nProvision"
Sql = Sql & "   when substring(cCtaCod,9,1) = '2'  then nProvision * " & tc
Sql = Sql & "  end ),0) ProvConstituida "

If lbMesActual = True Then
    Sql = Sql & " FROM " & lsServConsol & "ColocCalifProv ca Where "
Else
    Sql = Sql & " FROM " & lsServConsol & "ColocCalifProvTotal ca"
    Sql = Sql & " Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
End If
Sql = Sql & " nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2101,2104,2106,2107) And "
Sql = Sql & " substring(cCtaCod,6,3) = '423' "
Sql = Sql & " Group by substring(cCtaCod,6,1), cCalGen "
Sql = Sql & " Order by substring(cCtaCod,6,1), cCalGen "

Set rs = Co.CargaRecordSet(Sql)

While Not rs.EOF
    Select Case rs!cCalGen
        Case "0":
             xlHoja1.Range("B73") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap / 3, 2))
             xlHoja1.Range("B78") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap * 2 / 3, 2))
        Case "1"
             xlHoja1.Range("C73") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap / 3, 2))
             xlHoja1.Range("C78") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap * 2 / 3, 2))
        Case "2"
             xlHoja1.Range("D73") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap / 3, 2))
             xlHoja1.Range("D78") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap * 2 / 3, 2))
        Case "3"
             xlHoja1.Range("E73") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap / 3, 2))
             xlHoja1.Range("E78") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap * 2 / 3, 2))
        Case "4"
             xlHoja1.Range("F73") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap / 3, 2))
             xlHoja1.Range("F78") = IIf(IsNull(rs!SaldoCap), 0, Round(rs!SaldoCap * 2 / 3, 2))
    End Select
    rs.MoveNext
Wend
'
' '-----------------------------------------------
' ' Calculos
' '-----------------------------------------------
'
 xlHoja1.Range("B97").Formula = "= (B31+B40+B49+B56+B61)*1% + B70+B69+B87* 1%"
 xlHoja1.Range("B98").Formula = "= (B32+B41+B50+B57+B62)*1% + B70+B69+B88* 1%"
 xlHoja1.Range("B99").Formula = "= (B33+B42+B51+B58)    *1% + B70+B69+B89* 1%"
 xlHoja1.Range("B100").Formula = "= (B34+B43+B52)        *1% + B70+B69+B90* 1%"

 xlHoja1.Range("C97").Formula = "= (C31*5%)+ (C40*1%) + (C49*1.25%) + (C56*2.5%) + (C61*1%)  + B70+B71+B87* 1%"

 For i = 17 To 20
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 22 To 25
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 31 To 34
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 40 To 43
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 49 To 51
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 56 To 58
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 61 To 62
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 73 To 74
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 76 To 79
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 87 To 90
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 92 To 95
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 97 To 100
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 103 To 106
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i

 xlHoja1.Range("B26").Formula = "= B22+B23+B24+B25"
 xlHoja1.Range("C26").Formula = "= C22+C23+C24+C25"
 xlHoja1.Range("D26").Formula = "= D22+D23+D24+D25"
 xlHoja1.Range("E26").Formula = "= E22+E23+E24+E25"
 xlHoja1.Range("F26").Formula = "= F22+F23+F24+F25"

 xlHoja1.Range("B97").Formula = "= (B31+B40+B49+B56+B61)*1% + B70+B69+B87* 1%"
 xlHoja1.Range("B98").Formula = "= (B32+B41+B50+B57+B62)*1% +B88* 1%"
 xlHoja1.Range("B99").Formula = "= (B33+B42+B51+B58)    *1% + B73 * 0% + (B74+B76+B77+B78+B79)*1% + B89 * 1%"
 xlHoja1.Range("B100").Formula = "= (B34+B43)*1% + B90 *1%"

 xlHoja1.Range("C97").Formula = "= (C31*5%)+ (C40*1%) + (C49*1.25%) + (C56*2.5%) + (C61*1%)  +C87* 5%"
 xlHoja1.Range("C98").Formula = "= (C32*5%)+ (C41*1%) + (C50*1.25%) + (C57*2.5%) + (C62*1%)  +C88* 5%"
 xlHoja1.Range("C99").Formula = "= (C33*5%)+ (C42*1%) + (C51*1.25%) + (C58*2.5%) + (C73*0%)+(C74*1%)+(C76*5%)+(C77*1.25%)+(C78*2.5%)+(C79*5%)+(C89*5%)"
 xlHoja1.Range("C100").Formula = "=(C34*5%)+(C43*1%)+C90*5%"

 xlHoja1.Range("D97").Formula = "=(D31*25%)+(D40*1%)+(D49*6.25%)+(D56*12.5%)+(D61*1%)+(D87*25%)"
 xlHoja1.Range("D98").Formula = "=(D32*25%)+(D41*1%)+(D50*6.25%)+(D57*12.5%)+(D62*1%)+(D88*25%)"
 xlHoja1.Range("D99").Formula = "=(D33*25%)+(D42*1%)+(D51*6.25%)+(D58*12.5%)+(D73*0%)+(D74*1%)+(D76*25%)+(D77*6.25%)+(D78*12.5%)+(D79*25%)+(D89*25%)"
 xlHoja1.Range("D100").Formula = "=(D34*25%)+(D43*1%)+(D90*25%)"

 xlHoja1.Range("E97").Formula = "=(E31*60%)+(E40*1%)+(E49*15%)+(E56*30%)+(E61*1%)+(E87*60%)"
 xlHoja1.Range("E98").Formula = "=(E32*60%)+(E41*1%)+(E50*15%)+(E57*30%)+(E62*1%)+(E88*60%)"
 xlHoja1.Range("E99").Formula = "=(E33*60%)+(E42*1%)+(E51*15%)+(E58*30%)+(E73*0%)+(E74*1%)+(E76*60%)+(E77*15%)+(E78*30%)+(E79*60%)+(E89*25%)"
 xlHoja1.Range("E100").Formula = "=(E34*60%)+(E43*1%)+(E90*60%)"

 xlHoja1.Range("F97").Formula = "=(F31*100%)+(F40*1%)+(F56*60%)+(F61*1%)+(F87*100%)"
 xlHoja1.Range("F98").Formula = "=(F32*100%)+(F41*1%)+(F57*60%)+(F62*1%)+(F88*100%)"
 xlHoja1.Range("F99").Formula = "=(F33*100%)+(F42*19%)+(F51*30%)+(F58*60%)+(F73*0%)+(F74*1%)+(F76*100%)+(F77*30%)+(F78*60%)+(F79*100%)+(F89*100%)"
 xlHoja1.Range("F100").Formula = "=(F34*100%)+(F43*1%)+(F90*100%)"

 xlHoja1.Range("C166").Formula = "= G17+G18+G19+G20"
 xlHoja1.Range("D166").Formula = "= B92+B93+B94+B95"
 xlHoja1.Range("E166").Formula = "= G92+G93+G94+G95+G117+G118+G119+G120+G146+G147+G148+G149-D166"

 '--- formato
 '----------

 xlHoja1.Range("B17:G20").Style = "Comma"
 xlHoja1.Range("B40:G43").Style = "Comma"
 xlHoja1.Range("B56:G58").Style = "Comma"
 xlHoja1.Range("B73:G78").Style = "Comma"
 xlHoja1.Range("B87:G90").Style = "Comma"
 xlHoja1.Range("B92:G95").Style = "Comma"
 xlHoja1.Range("B97:G100").Style = "Comma"

 xlHoja1.Range("A14:G107").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A14:G107").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A14:G107").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A14:G107").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A14:G107").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("A14:G107").Borders(xlEdgeLeft).Weight = xlThin
 xlHoja1.Range("A14:G107").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("A14:G107").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("A16:G16").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A16:G16").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A21:G21").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A21:G21").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A21:G21").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A26:G26").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A26:G26").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A26:G26").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A30:G30").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A30:G30").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A35:G39").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A35:G39").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A35:G39").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A44:G48").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A44:G48").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A44:G48").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A52:G55").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A52:G55").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A52:G55").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A59:G60").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A59:G60").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A59:G60").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A63:G68").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A63:G68").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A63:G68").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A71:G72").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A71:G72").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A71:G72").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A79:G79").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A79:G79").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A83:G86").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A83:G86").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A83:G86").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A91:G91").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A91:G91").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A91:G91").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A96:G96").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A96:G96").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A96:G96").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A101:G102").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A101:G102").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A101:G102").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A101:G102").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A101:G102").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A107:G107").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A107:G107").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("B14:F62").Borders(xlInsideVertical).LineStyle = xlContinuous
 xlHoja1.Range("B14:F62").Borders(xlInsideVertical).Weight = xlThin

 xlHoja1.Range("B63:F68").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("B63:F68").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("B63:F68").Borders(xlInsideVertical).Weight = xlThin

 xlHoja1.Range("B14:F62").Borders(xlInsideVertical).LineStyle = xlContinuous
 xlHoja1.Range("B14:F62").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("B14:F62").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("B14:F62").Borders(xlInsideVertical).Weight = xlThin

 xlHoja1.Range("B73:F107").Borders(xlInsideVertical).LineStyle = xlContinuous
 xlHoja1.Range("B73:F107").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("B73:F107").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("B73:F107").Borders(xlInsideVertical).Weight = xlThin

xlHoja1.Range("C63:G70").MergeCells = True

'---
 xlHoja1.Range("A110:G131").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A110:G131").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A110:G131").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A110:G131").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A110:G131").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("A110:G131").Borders(xlEdgeLeft).Weight = xlThin
 xlHoja1.Range("A110:G131").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("A110:G131").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("A111:G111").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A111:G111").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A116:G116").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A116:G116").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A116:G116").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A116:G116").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A121:G121").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A121:G121").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A121:G121").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A121:G121").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A126:G127").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A126:G127").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A126:G127").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A126:G127").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A131:G131").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A131:G131").Borders(xlEdgeTop).Weight = xlThin

 xlHoja1.Range("B110:F131").Borders(xlInsideVertical).LineStyle = xlContinuous
 xlHoja1.Range("B110:F131").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("B110:F131").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("B110:F131").Borders(xlInsideVertical).Weight = xlThin


 xlHoja1.Range("A134:G156").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A134:G156").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A134:G156").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A134:G156").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A134:G156").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("A134:G156").Borders(xlEdgeLeft).Weight = xlThin
 xlHoja1.Range("A134:G156").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("A134:G156").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("A135:G135").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A135:G135").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A140:G140").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A140:G140").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A140:G140").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A140:G140").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A145:G145").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A145:G145").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A145:G145").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A145:G145").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A150:G151").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A150:G151").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A150:G151").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A150:G151").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A156:G156").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A156:G156").Borders(xlEdgeTop).Weight = xlThin

 xlHoja1.Range("B134:F156").Borders(xlInsideVertical).LineStyle = xlContinuous
 xlHoja1.Range("B134:F156").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("B134:F156").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("B134:F156").Borders(xlInsideVertical).Weight = xlThin

 xlHoja1.Range("A161:E166").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("A161:E166").Borders(xlEdgeLeft).Weight = xlThin
 xlHoja1.Range("A161:E166").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("A161:E166").Borders(xlEdgeRight).Weight = xlThin
 xlHoja1.Range("A161:E166").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A161:E166").Borders(xlEdgeBottom).Weight = xlThin
 xlHoja1.Range("A161:E166").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A161:E166").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A161:E166").Borders(xlInsideVertical).LineStyle = xlContinuous
 xlHoja1.Range("A161:E166").Borders(xlInsideVertical).Weight = xlThin
'---
Co.CierraConexion
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set Co = Nothing
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
MsgBox "Se genero el Anexo. Verificar en la carpeta Spooler", vbInformation, "AVISO"
'CargaArchivo App.path & "\SPOOLER\" & lsArchivo, App.path & "\SPOOLER\"
'lbExcel = False
End Sub

Public Sub Genera_Reporte178210(ByVal TxtCambio As Double, ByVal dFecha As Date, _
                        ByVal nCF As Integer, ByVal psServer As String, _
                        ByVal dFecData As String, _
                        Optional ByVal psNomCmac As String, _
                        Optional ByVal psCodCmac As String)

'Dim fs As Scripting.FileSystemObject
Dim lbExisteHoja  As Boolean
Dim lsNomHoja As String
Dim pnFila As Integer, lnCol As Integer
Dim lsSQL As String
Dim rs As New ADODB.Recordset
Dim rsSubT As New ADODB.Recordset
Dim rsDesc As New ADODB.Recordset
Dim lnLinea As Integer
Dim lsMoneda As String, lsTipo As String, lsSubTipo As String, lsRango As String, lsFF As String
Dim lsTotFF(14) As String
Dim lsTotRango(14) As String
Dim lsTotSubTipo(14) As String
Dim lsTotTipo(14) As String
Dim lsTotCartera(14) As String
Dim lnIndice As Integer
Dim lbCambio As Boolean ' Indica el cambio condicion
Dim lnMoneda As Integer, lnTipo As Integer, lnSubTipo As Integer, lnRango As Integer, lnFF As Integer
Dim lnCuenta As Long
Dim lsCFianza As String
Dim lsTipoDesc As String
Dim lsSubTDesc As String
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim J As Integer
Dim lsNombre As String

'Rep_HojaWProvision808 = True
lnLinea = 1
Dim lbMesActual As Boolean

Dim lbIncluyeCF As Boolean
If nCF = 1 Then
    lbIncluyeCF = True
Else
    lbIncluyeCF = False
End If

Dim lsServConsol As String
Dim Co As DConecta
Dim RCC As DRCC
Set RCC = New DRCC
lsServConsol = RCC.ServConsol(psServer)
Set RCC = Nothing

Set Co = New DConecta

lbCambio = False
lsArchivo = "Anexo5_" & Format(dFecha, "yyyymm") & "_HojaW_808" & IIf(lbIncluyeCF = True, "", "_SinCF") & ".XLS"
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If
lbExcel = True

If dFecha = dFecData Then
    lbMesActual = True
Else
    lbMesActual = False
End If
If lbIncluyeCF = False Then
    lsCFianza = " And Substring(CA.cCtaCod,7,2) <> '21' "
End If

For lnMoneda = 1 To 2 ' Moneda

    lbExisteHoja = False
    lsNomHoja = "HW_Moneda" & Str(lnMoneda)
    For Each xlHoja1 In xlLibro.Worksheets
        If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
            xlHoja1.Range("A1", "AZ10000") = ""
            xlHoja1.Range("A1", "AZ10000").Rows.Delete
            lbExisteHoja = True
            Exit For
        End If
    Next
    If lbExisteHoja = False Then
        Set xlHoja1 = xlLibro.Worksheets.Add
        xlHoja1.Name = lsNomHoja
    End If
    'Me.EstatuBarICC.Panels(2).Text = "Generando Reporte en Excel Por favor espere..."
    pnFila = 2
    ' Cabecera del Reporte en Excel
    '*********  falta procedimiento
    Call ImpHojaWProvisionCab(pnFila, Trim(Str(lnMoneda)), dFecha)
    '*****************************


    xlHoja1.Cells(pnFila, 1) = psNomCmac
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 2) = "INFORME DE CLASIFICACION DE LA CARTERA DE CREDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 7) = "AL " & Format(dFecha, "DD/MM/YYYY")
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    If Trim(Str(lnMoneda)) = "1" Then
        xlHoja1.Cells(pnFila, 6) = "(  MONEDA NACIONAL  ) "
    Else
        xlHoja1.Cells(pnFila, 6) = "( MONEDA EXTRANJERA ) "
    End If
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 1) = " CREDITOS "
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Cells(pnFila, 2) = "TOTAL CARTERA"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 2), xlHoja1.Cells(pnFila, 3)).Merge True
    xlHoja1.Cells(pnFila + 1, 2) = "Nro"
    xlHoja1.Cells(pnFila + 1, 3) = "Monto"
    xlHoja1.Cells(pnFila, 4) = "NORMALES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 4), xlHoja1.Cells(pnFila, 5)).Merge True
    xlHoja1.Cells(pnFila + 1, 4) = "Nro"
    xlHoja1.Cells(pnFila + 1, 5) = "Monto"
    xlHoja1.Cells(pnFila, 6) = "POTENCIALES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 6), xlHoja1.Cells(pnFila, 7)).Merge True
    xlHoja1.Cells(pnFila + 1, 6) = "Nro"
    xlHoja1.Cells(pnFila + 1, 7) = "Monto"
    xlHoja1.Cells(pnFila, 8) = "DEFICIENTES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 8), xlHoja1.Cells(pnFila, 9)).Merge True
    xlHoja1.Cells(pnFila + 1, 8) = "Nro"
    xlHoja1.Cells(pnFila + 1, 9) = "Monto"
    xlHoja1.Cells(pnFila, 10) = "DUDOSOS"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 10), xlHoja1.Cells(pnFila, 11)).Merge True
    xlHoja1.Cells(pnFila + 1, 10) = "Nro"
    xlHoja1.Cells(pnFila + 1, 11) = "Monto"
    xlHoja1.Cells(pnFila, 12) = "PERDIDA"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 12), xlHoja1.Cells(pnFila, 13)).Merge True
    xlHoja1.Cells(pnFila + 1, 12) = "Nro"
    xlHoja1.Cells(pnFila + 1, 13) = "Monto"
    xlHoja1.Cells(pnFila, 14) = "PROVISION"
    xlHoja1.Cells(pnFila + 1, 14) = Format(dFecha, "dd/mm/yyyy")
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 14)).HorizontalAlignment = xlCenter
    CuadroExcel 1, pnFila, 14, pnFila + 1, True

    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 2), xlHoja1.Cells(500, 2)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 3), xlHoja1.Cells(500, 3)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 4), xlHoja1.Cells(500, 4)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 5), xlHoja1.Cells(500, 5)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 6), xlHoja1.Cells(500, 6)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 7), xlHoja1.Cells(500, 7)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 8), xlHoja1.Cells(500, 8)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 9), xlHoja1.Cells(500, 9)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 10), xlHoja1.Cells(500, 10)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 11), xlHoja1.Cells(500, 11)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 12), xlHoja1.Cells(500, 12)).NumberFormat = "#,#0;#,#0"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 13), xlHoja1.Cells(500, 13)).NumberFormat = "#,#0.00;#,#0.00"
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 14), xlHoja1.Cells(500, 14)).NumberFormat = "#,#0.00;#,#0.00"


    xlHoja1.Range("A1").ColumnWidth = 18
    xlHoja1.Range("B1").ColumnWidth = 7
    xlHoja1.Range("C1").ColumnWidth = 15
    xlHoja1.Range("D1").ColumnWidth = 7
    xlHoja1.Range("E1").ColumnWidth = 15
    xlHoja1.Range("F1").ColumnWidth = 7
    xlHoja1.Range("G1").ColumnWidth = 15
    xlHoja1.Range("H1").ColumnWidth = 7
    xlHoja1.Range("I1").ColumnWidth = 15
    xlHoja1.Range("J1").ColumnWidth = 7
    xlHoja1.Range("K1").ColumnWidth = 15
    xlHoja1.Range("L1").ColumnWidth = 7
    xlHoja1.Range("M1").ColumnWidth = 15
    xlHoja1.Range("N1").ColumnWidth = 16

    ' Formato de celdas
    xlHoja1.Range("C1").NumberFormat = "#,#0.00"
    xlHoja1.Range("E1").NumberFormat = "#,#0.00"
    xlHoja1.Range("G1").NumberFormat = "#,#0.00"
    xlHoja1.Range("IC1").NumberFormat = "#,#0.00"
    xlHoja1.Range("K1").NumberFormat = "#,#0.00"
    xlHoja1.Range("M1").NumberFormat = "#,#0.00"
    xlHoja1.Range("N1").NumberFormat = "#,#0.00"

    pnFila = pnFila + 2


    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlAplicacion.Range("A1:Z100").Font.Size = 9
    '*****************************
    pnFila = 8
    ' Cuerpo del Reporte
    ' Limpiar Tot Cartera / Tipo / SubTipo
    For lnIndice = 2 To 14
        lsTotCartera(lnIndice) = ""
        lsTotTipo(lnIndice) = ""
        lsTotSubTipo(lnIndice) = ""
    Next lnIndice

    For lnTipo = 1 To 4

        ' Imprimir Encabezado
        If Trim(Str(lnTipo)) = "1" Then
            lsTipoDesc = "I.- CREDITOS COMERCIALES"
        ElseIf Trim(Str(lnTipo)) = "2" Then
            lsTipoDesc = "II.- CREDITOS MICROEMPRESA"
        ElseIf Trim(Str(lnTipo)) = "3" Then
            lsTipoDesc = "III.- CREDITOS CONSUMO"
        ElseIf Trim(Str(lnTipo)) = "4" Then
            lsTipoDesc = "IV.- CREDITOS HIPOTECARIOS"
        End If
        xlHoja1.Cells(pnFila, 1) = lsTipoDesc
'        If pCartaFianza = True Then
'            xlHoja1.Cells(pnFila, 1) = xlHoja1.Cells(pnFila, 1) & " (Incluye Cartas Fianza ) "
'        Else
'            xlHoja1.Cells(pnFila, 1) = xlHoja1.Cells(pnFila, 1) & " (No Incluye Cartas Fianza )"
'        End If
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0C0FF

        pnFila = pnFila + 1
        ' Limpiar
        For lnIndice = 2 To 14
            lsTotSubTipo(lnIndice) = ""
        Next lnIndice

        lsSQL = " Select distinct(Substring(CA.cCtaCod,6,3)) cSubT FROM " & lsServConsol & "ColocCalifProv CA " _
              & " Where substring(Ca.cCtaCod,6,1) = '" & lnTipo & "' AND Substring(CA.cLineaCred,5,1) = '" & Trim(Str(lnMoneda)) & "' " & lsCFianza & " ORDER BY Substring(CA.cCtaCod,6,3) "

        Set rsSubT = Co.CargaRecordSet(lsSQL)

        Do While Not rsSubT.EOF
            lnSubTipo = Trim(Str(rsSubT!cSubT))
        'For lnSubTipo = 1 To 2
            If Mid(rsSubT!cSubT, 2, 2) = "21" Then
                lsSubTDesc = "CARTA FIANZA"
            Else
'                lsSQL = " Select * from Tablacod " _
'                      & " Where cCodtab like '6" & lnTipo & "__' And substring(cCodTab,2,3) = '" & Trim(Str(rsSubT!cSubT)) & "' "
'                rsDesc.Open lsSQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
'                lsSubTDesc = Trim(IIf(IsNull(rsDesc!cNomTab), "", rsDesc!cNomTab))
'
            End If

            ' Imprimir Encabezado
            xlHoja1.Cells(pnFila, 1) = lsSubTDesc
            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HFFC0C0
            pnFila = pnFila + 1
            ' Limpiar
            For lnIndice = 2 To 14
                lsTotRango(lnIndice) = ""
            Next lnIndice

            For lnRango = 1 To 4   ' Por tipo de Garantia

'                lsSQL = "SELECT Substring(CA.cLineaCred,7,3) as cSubtipo, " _
'                      & "Substring(CA.cCodLinCred,6,1) as cFF, substring(CA.cCtaCod,1,2) as cAgencia, " _
'                      & "Count (CA.cCtaCod) AS NroCred, " _
'                      & "Isnull(Sum (CA.nSaldoCap), 0) AS SaldoCap, " _
'                      & "Count( CASE WHEN CA.cCalGen = '0' THEN (CA.cCtaCod) END ) AS NroProv0, " _
'                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '0' THEN (CA.nSaldoCap) END ), 0) AS Prov0, " _
'                      & "Count( CASE WHEN CA.cCalGen = '1' THEN (CA.cCtaCod) END ) AS NroProv1, " _
'                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '1' THEN (CA.nSaldoCap) END ), 0) AS Prov1, " _
'                      & "Count(CASE WHEN CA.cCalGen = '2' THEN (CA.cCtaCod) END ) AS NroProv2, " _
'                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '2' THEN (CA.nSaldoCap) END ), 0) AS Prov2, " _
'                      & "Count(CASE WHEN CA.cCalGen = '3' THEN (CA.cCtaCod) END ) AS NroProv3, " _
'                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '3' THEN (CA.nSaldoCap) END ), 0) AS Prov3, " _
'                      & "Count(CASE WHEN CA.cCalGen = '4' THEN (CA.cCtaCod) END ) AS NroProv4, " _
'                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '4' THEN (CA.nSaldoCap) END ), 0) AS Prov4, " _
'                      & "Sum(IsNull(ca.nProvision, 0)) As Provision "
'                If lbMesActual = True Then
'                    lsSQL = lsSQL & " FROM ColocCalifProv ca Where "
'                Else
'                    lsSQL = lsSQL & " FROM ColocCalifProvTotal ca"
'                    lsSQL = lsSQL & " Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
'                End If
'                lsSQL = lsSQL & " " _
'                      & " CA.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2101,2104,2106,2107)   " _
'                      '& "AND Substring(CA.cLineaCred,5,1) = '" & Trim(Str(lnMoneda)) & "' " _
'                      & "AND CA.nSaldoCap > 0  " _
'                      & "AND Substring(CA.cLineaCred,1,1) ='" & Trim(Str(lnTipo)) & "' " _
'                      & "AND Substring(CA.cLineaCred,1,3) ='" & Trim(Str(lnSubTipo)) & "' " _
'                      & "AND CA.nGarant = " & lnRango & " " _
'                      & lsCFianza
'                lsSQL = lsSQL & "GROUP BY  Substring(CA.cLineaCred,1,3), " _
'                      & "Substring(CA.cLineaCred,6,1), Substring(CA.cCtaCod,1,2) " _
'                      & "ORDER By Substring(CA.cLineaCred,1,3), " _
'                      & "Substring(CA.cCodLinCred,6,1), Substring(CA.cCtaCod,1,2) "


                lsFF = ""

                If Not (rs.BOF And rs.EOF) Then
                    ' Imprimir Encabezado
                    If lnRango = 1 Then
                        xlHoja1.Cells(pnFila, 1) = "Garantias AutoLiquidable"
                    ElseIf lnRango = 2 Then
                        xlHoja1.Cells(pnFila, 1) = "Garantias Muy Rapida Realizacion"
                    ElseIf lnRango = 3 Then
                        xlHoja1.Cells(pnFila, 1) = "Garantias Preferida"
                    ElseIf lnRango = 4 Then
                        xlHoja1.Cells(pnFila, 1) = "Garantias No Preferida"
                    End If
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 1)).Font.Bold = True
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HFF00C
                    pnFila = pnFila + 1
                    ' Limpiar
                    For lnIndice = 2 To 14
                        lsTotFF(lnIndice) = ""
                    Next lnIndice
                End If

                Do While Not rs.EOF

                    If lsFF <> rs!cFF Then ' Otro Fuente Financ
                        'Imprimir Total de anterior Fuente Fin
                        If lsFF <> "" Then
                            'Imprimir totales FF
                            xlHoja1.Cells(pnFila, 1) = "SubTotal F.F. " & Trim(lsFF)
                            For lnIndice = 2 To 14
                                xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotFF(lnIndice) & ")"
                            Next lnIndice
                            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                            'Asigno los valores al Rango
                            For lnIndice = 2 To 14
                                lsTotRango(lnIndice) = lsTotRango(lnIndice) & IIf(Len(lsTotRango(lnIndice)) > 0, "+", "") & lsTotFF(lnIndice)
                            Next lnIndice
                            pnFila = pnFila + 1

                        End If
                        ' Imprimir encabezado
                        xlHoja1.Cells(pnFila, 1) = "   F.F. " & ObtieneDescripcionFuenteFinanciamiento(Trim(rs!cFF))
                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 1)).Font.Bold = True
                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0FFC0
                        pnFila = pnFila + 1
                        For lnIndice = 2 To 14
                            lsTotFF(lnIndice) = ""
                        Next lnIndice

                    End If ' Fuente F

                    '*****

                    ' Imprimir el detalle
                    xlHoja1.Cells(pnFila, 1) = "     Agencia " & rs!cAgencia
                    xlHoja1.Cells(pnFila, 2) = Format(rs!NroCred, "#0")
                    xlHoja1.Cells(pnFila, 3) = Format(rs!SaldoCap, "#0.00")
                    xlHoja1.Cells(pnFila, 4) = Format(rs!NroProv0, "#0")
                    xlHoja1.Cells(pnFila, 5) = Format(rs!Prov0, "#0.00")
                    xlHoja1.Cells(pnFila, 6) = Format(rs!NroProv1, "#0")
                    xlHoja1.Cells(pnFila, 7) = Format(rs!Prov1, "#0.00")
                    xlHoja1.Cells(pnFila, 8) = Format(rs!NroProv2, "#0")
                    xlHoja1.Cells(pnFila, 9) = Format(rs!Prov2, "#0.00")
                    xlHoja1.Cells(pnFila, 10) = Format(rs!NroProv3, "#0")
                    xlHoja1.Cells(pnFila, 11) = Format(rs!Prov3, "#0.00")
                    xlHoja1.Cells(pnFila, 12) = Format(rs!NroProv4, "#0")
                    xlHoja1.Cells(pnFila, 13) = Format(rs!Prov4, "#0.00")
                    xlHoja1.Cells(pnFila, 14) = Format(rs!Provision, "#0.00")
                    'Acumulo totales
                    For lnIndice = 2 To 14
                        lsTotFF(lnIndice) = lsTotFF(lnIndice) & IIf(Len(lsTotFF(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)  'Nro Cred
                    Next lnIndice
                    pnFila = pnFila + 1
                    'lsTipo = rs!cTipo
                    lsSubTipo = rs!cSubTipo
                    'lnRango = rs!cRango
                    lsFF = rs!cFF

                    rs.MoveNext
                Loop

                rs.Close
                Set rs = Nothing

                'Imprimir totales FF
                If Trim(lsTotFF(2)) <> "" Then
                    xlHoja1.Cells(pnFila, 1) = "SubTotal F.F. " & Trim(lsFF)
                    For lnIndice = 2 To 14
                        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotFF(lnIndice) & ")"
                    Next lnIndice
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                    'Asigno los valores al Rango
                    For lnIndice = 2 To 14
                        lsTotRango(lnIndice) = lsTotRango(lnIndice) & IIf(Len(lsTotRango(lnIndice)) > 0, "+", "") & lsTotFF(lnIndice)
                    Next lnIndice
                    pnFila = pnFila + 1

                    'Imprimir totales Rango
                    xlHoja1.Cells(pnFila, 1) = "SubTotal Garantia  " & Trim(lsRango)
                    For lnIndice = 2 To 14
                        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotRango(lnIndice) & ")"
                    Next lnIndice
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                    'xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = RGB(10, 100, 10)
                    'Asigno los valores al SubTipo
                    For lnIndice = 2 To 14
                        lsTotSubTipo(lnIndice) = lsTotSubTipo(lnIndice) + IIf(Len(lsTotSubTipo(lnIndice)) > 0, "+", "") + lsTotRango(lnIndice)
                    Next lnIndice
                    pnFila = pnFila + 1
                End If

                ' Limpio la variable FF
                For lnIndice = 2 To 14
                    lsTotFF(lnIndice) = ""
                Next lnIndice
                ' Limpio la variable Rango
                For lnIndice = 2 To 14
                    lsTotRango(lnIndice) = ""
                Next lnIndice

            Next lnRango

            'Imprimir totales SubTipo
            If Trim(lsTotSubTipo(2)) <> "" Then
                xlHoja1.Cells(pnFila, 1) = "Total " & lsSubTDesc
                For lnIndice = 2 To 14
                    xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotSubTipo(lnIndice) & ")"
                Next lnIndice
                xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                'Asigno los valores al tipo
                For lnIndice = 2 To 14
                    'lsTotTipo(lnIndice) = lsTotTipo(lnIndice) & IIf(Len(lsTotTipo(lnIndice)) > 0, "+", "") & lsTotSubTipo(lnIndice)
                    lsTotTipo(lnIndice) = lsTotTipo(lnIndice) & IIf(Len(lsTotTipo(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
                Next lnIndice
                pnFila = pnFila + 1
            End If

            ' Limpio la variable Subtipo
            For lnIndice = 2 To 14
                lsTotSubTipo(lnIndice) = ""
            Next lnIndice

            rsSubT.MoveNext
         Loop

        rsSubT.Close
        'Imprimir totales Tipo
        If Trim(lsTotTipo(2)) <> "" Then
            xlHoja1.Cells(pnFila, 1) = lsTipoDesc
            For lnIndice = 2 To 14
                xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotTipo(lnIndice) & ")"
            Next lnIndice
            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
            'Asigno los valores a la Cartera
            For lnIndice = 2 To 14
                'lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & lsTotTipo(lnIndice)
                lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
            Next lnIndice
            pnFila = pnFila + 2
        End If

        ' Limpio la variable tipo
        For lnIndice = 2 To 14
            lsTotTipo(lnIndice) = ""
        Next lnIndice

    Next lnTipo

    'Imprimir totales Tipo
    If Trim(lsTotTipo(2)) <> "" Then
        xlHoja1.Cells(pnFila, 1) = "Total " & lsTipoDesc
        For lnIndice = 2 To 14
            xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotTipo(lnIndice) & ")"
        Next lnIndice
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
        'Asigno los valores a la Cartera
        For lnIndice = 2 To 14
            'lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & lsTotTipo(lnIndice)
            lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
        Next lnIndice
        pnFila = pnFila + 2
    End If

    ' Limpio la variable tipo
    For lnIndice = 2 To 14
        lsTotTipo(lnIndice) = ""
    Next lnIndice


    ' ********************

    'Imprime Total Cartera
    xlHoja1.Cells(pnFila, 1) = "Total CARTERA"
    For lnIndice = 2 To 14
        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotCartera(lnIndice) & ")"
    Next lnIndice
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0FFFF

    'CuadroExcel 1, 6, 14, pnFila, True
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(pnFila, 14)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(pnFila, 14)).Cells.Borders.LineStyle = xlInside

Next lnMoneda

'*******
xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
End Sub
Public Sub Genera_Reporte178209(ByVal TxtCambio As Double, ByVal dFecha As Date, _
                        ByVal nCF As Integer, ByVal psServer As String, _
                        ByVal dFecData As String, _
                        ByVal psFFinan As String, _
                        Optional ByVal psNomCmac As String, _
                        Optional ByVal psCodCmac As String)


Dim Sql As String
Dim rs As New ADODB.Recordset
Dim fs As New Scripting.FileSystemObject
Dim tc As Double
Dim i As Integer
Dim J As Integer
Dim k As Integer
Dim lsNomHoja As String
Dim lbExisteHoja As Boolean

Dim lbMesActual As Boolean
Dim lbIncluyeCF As Boolean
Dim lsArchivo As String
Dim lsServConsol As String

Dim RCC As DRCC
Set RCC = New DRCC
lsServConsol = RCC.ServConsol(psServer)
Set RCC = Nothing

tc = TxtCambio
If dFecha = dFecData Then
    lbMesActual = True
Else
    lbMesActual = False
End If

If nCF = 1 Then
    lbIncluyeCF = True
Else
    lbIncluyeCF = False
End If

lsArchivo = "Rep_Anexo5D_" & Format(dFecha, "YYYYMM") & ".XLS"
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

 lsNomHoja = "Anexo5D"
 For Each xlHoja1 In xlLibro.Worksheets
 If xlHoja1.Name = lsNomHoja Then
    xlHoja1.Activate
    xlHoja1.Range("A1", "AZ10000") = ""
    lbExisteHoja = True
    Exit For
 End If
 Next

If lbExisteHoja = False Then
    Set xlHoja1 = xlLibro.Worksheets.Add
    xlHoja1.Name = lsNomHoja
End If

 'xlHoja1.Range("A1") = gsNomCmac
 'xlHoja1.Range("A2") = "Al: " & txtFecha
 'xlHoja1.Range("A1:A2").Font.Bold = True
 xlHoja1.Range("A1:AZ1000") = ""
 xlHoja1.Range("A2") = "ANEXO N° 5 D"
 xlHoja1.Range("A3") = "INFORME DE CLASIFICACION DE DEUDORES DE LA CARTERA DE CREDITOS,"
 xlHoja1.Range("A4") = "CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS QUE RESPALDAN FINANCIAMIENTOS O LINEAS DE CREDITO"
 xlHoja1.Range("A6") = "EMPRESA : " & psNomCmac
 xlHoja1.Range("G6") = "CODIGO : " & psCodCmac
 xlHoja1.Range("A8") = "AL " & Format(dFecha, "DD") & " DE  " & UCase(Format(dFecha, "MMMM")) & " DEL  " & Format(dFecha, "YYYY")
 xlHoja1.Range("A9") = "(En miles de nuevos soles)"

 xlHoja1.Range("A1").ColumnWidth = 42
 xlHoja1.Range("B1").ColumnWidth = 15
 xlHoja1.Range("C1").ColumnWidth = 15
 xlHoja1.Range("D1").ColumnWidth = 15
 xlHoja1.Range("E1").ColumnWidth = 15
 xlHoja1.Range("F1").ColumnWidth = 15
 xlHoja1.Range("G1").ColumnWidth = 16

 xlHoja1.Application.ActiveWindow.Zoom = 75
 xlHoja1.Range("A2:G2").MergeCells = True
 xlHoja1.Range("A2:A2").HorizontalAlignment = xlCenter
 xlHoja1.Range("A3:A3").HorizontalAlignment = xlCenter
 xlHoja1.Range("A3:G3").MergeCells = True
 xlHoja1.Range("A4:G4").MergeCells = True
 xlHoja1.Range("A8:G8").MergeCells = True
 xlHoja1.Range("A9:G9").MergeCells = True
 xlHoja1.Range("A8:A8").HorizontalAlignment = xlCenter
 xlHoja1.Range("A9:A9").HorizontalAlignment = xlCenter
 xlHoja1.Range("A2:G11").Font.Bold = True

 xlHoja1.Range("A11") = "A. MONTO DE LOS CREDITOS"
 xlHoja1.Range("A12") = "CONTINGENTES Y ARRENDAMIENTOS"
 xlHoja1.Range("A13") = "FINANCIEROS 2/"
 xlHoja1.Range("B13") = "NORMAL"
 xlHoja1.Range("C13") = "CPP"
 xlHoja1.Range("D13") = "DEFICIENTE"
 xlHoja1.Range("E13") = "DUDOSO"
 xlHoja1.Range("F13") = "PERDIDA"
 xlHoja1.Range("G13") = "TOTAL"
 xlHoja1.Range("A11:G13").Font.Bold = True

 xlHoja1.Range("A14") = " Comercial"
 xlHoja1.Range("A15") = " Mes"
 xlHoja1.Range("A16") = " Hipotecario para vivienda"
 xlHoja1.Range("A17") = " Consumo"

 xlHoja1.Range("A18") = "B. PROVISION CONSTITUIDA 3/"
 xlHoja1.Range("B18") = "NORMAL"
 xlHoja1.Range("C18") = "CPP"
 xlHoja1.Range("D18") = "DEFICIENTE"
 xlHoja1.Range("E18") = "DUDOSO"
 xlHoja1.Range("F18") = "PERDIDA"
 xlHoja1.Range("G18") = "TOTAL"
 xlHoja1.Range("A18:G18").Font.Bold = True

 xlHoja1.Range("A19") = " Comercial"
 xlHoja1.Range("A20") = " Mes"
 xlHoja1.Range("A21") = " Hipotecario para vivienda"
 xlHoja1.Range("A22") = " Consumo"

 xlHoja1.Range("A23") = "C. PROVISIONES REQUERIDAS 4/"
 xlHoja1.Range("B23") = "NORMAL"
 xlHoja1.Range("C23") = "CPP"
 xlHoja1.Range("D23") = "DEFICIENTE"
 xlHoja1.Range("E23") = "DUDOSO"
 xlHoja1.Range("F23") = "PERDIDA"
 xlHoja1.Range("G23") = "TOTAL"
 xlHoja1.Range("A23:G23").Font.Bold = True

 xlHoja1.Range("A24") = " Comercial"
 xlHoja1.Range("A25") = " Mes"
 xlHoja1.Range("A26") = " Hipotecario para vivienda"
 xlHoja1.Range("A27") = " Consumo"

 xlHoja1.Range("A28") = "D. SUPERAVIT (DEFICIT) DE"
 xlHoja1.Range("A29") = "PROVISIONES 5/ "
 xlHoja1.Range("B29") = "NORMAL"
 xlHoja1.Range("C29") = "CPP"
 xlHoja1.Range("D29") = "DEFICIENTE"
 xlHoja1.Range("E29") = "DUDOSO"
 xlHoja1.Range("F29") = "PERDIDA"
 xlHoja1.Range("G29") = "TOTAL"
 xlHoja1.Range("A28:G29").Font.Bold = True

 xlHoja1.Range("A30") = " Comercial"
 xlHoja1.Range("A31") = " Mes"
 xlHoja1.Range("A32") = " Hipotecario para vivienda"
 xlHoja1.Range("A33") = " Consumo"

 xlHoja1.Range("A34") = "TOTAL"

 xlHoja1.Range("A36") = "E. MONTO DE LOS CREDITOS"
 xlHoja1.Range("A37") = "CONTINGENTES Y ARRENDAMIENTOS"
 xlHoja1.Range("A38") = "FINANCIEROS CEDIDOS EN"
 xlHoja1.Range("A39") = "EJECUCION"

 xlHoja1.Range("B37") = "NORMAL"
 xlHoja1.Range("C37") = "CPP"
 xlHoja1.Range("D37") = "DEFICIENTE"
 xlHoja1.Range("E37") = "DUDOSO"
 xlHoja1.Range("F37") = "PERDIDA"
 xlHoja1.Range("G37") = "TOTAL"
 xlHoja1.Range("A36:G39").Font.Bold = True

 xlHoja1.Range("A40") = " Comercial"
 xlHoja1.Range("A41") = " Mes"
 xlHoja1.Range("A42") = " Hipotecario para vivienda"
 xlHoja1.Range("A43") = " Consumo"


 xlHoja1.Range("B47").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("B47").Borders(xlEdgeTop).Weight = xlThin
xlHoja1.Range("B47:B48").HorizontalAlignment = xlCenter
xlHoja1.Range("B47") = "GERENTE"
xlHoja1.Range("B48") = "GENERAL"

xlHoja1.Range("D47").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("D47").Borders(xlEdgeTop).Weight = xlThin
xlHoja1.Range("D47:D48").HorizontalAlignment = xlCenter
xlHoja1.Range("D47") = "CONTADOR"
xlHoja1.Range("D48") = "GENERAL"

xlHoja1.Range("F47").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("F47").Borders(xlEdgeTop).Weight = xlThin
xlHoja1.Range("F47:F48").HorizontalAlignment = xlCenter
xlHoja1.Range("F47") = "FUNCIONARIO"
xlHoja1.Range("F48") = "RESPONSABLE"

'--------
Sql = " select substring(cCtaCod,6,1) cta, cCalGen ,"
Sql = Sql & " Isnull(sum(case when substring(cCtaCod,9,1) = '1' then nSaldoCap"
Sql = Sql & "                 when substring(cCtaCod,9,1) = '2' then nSaldoCap * " & tc
Sql = Sql & "  end ),0) SaldoCap,"
Sql = Sql & " Isnull(sum(case when substring(cCtaCod,9,1) = '1'  then nProvision"
Sql = Sql & "                 when substring(cCtaCod,9,1) = '2'  then nProvision * " & tc
Sql = Sql & "  end ),0) ProvConstituida"

If lbMesActual = True Then
    Sql = Sql & " FROM " & lsServConsol & "ColocCalifProv ca Where "
Else
    Sql = Sql & " FROM " & lsServConsol & "ColocCalifProvTotal ca"
    Sql = Sql & " Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
End If
Sql = Sql & " nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2101,2104,2106,2107) "
Sql = Sql & " And Substring(cLineaCred,1,4) in " & psFFinan & " "
If lbIncluyeCF = False Then
    Sql = Sql & " And substring(ca.cCtaCod,7,2) <> '21' "
End If
Sql = Sql & " Group by substring(cCtaCod,6,1), cCalGen"
Sql = Sql & " Order by substring(cCtaCod,6,1), cCalGen"

While Not rs.EOF
    Select Case rs!Cta
        Case "1"
            Select Case rs!cCalGen
                Case "0": xlHoja1.Range("B14") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("B19") = Format(rs!ProvConstituida, "#,###.00")

                Case "1": xlHoja1.Range("C14") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("C19") = Format(rs!ProvConstituida, "#,###.00")

                Case "2": xlHoja1.Range("D14") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("D19") = Format(rs!ProvConstituida, "#,###.00")

                Case "3": xlHoja1.Range("E14") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("E19") = Format(rs!ProvConstituida, "#,###.00")

                Case "4": xlHoja1.Range("F14") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("F19") = Format(rs!ProvConstituida, "#,###.00")
            End Select
        Case "2"
            Select Case rs!cCalGen
                Case "0": xlHoja1.Range("B15") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("B20") = Format(rs!ProvConstituida, "#,###.00")

                Case "1": xlHoja1.Range("C15") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("C20") = Format(rs!ProvConstituida, "#,###.00")

                Case "2": xlHoja1.Range("D15") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("D20") = Format(rs!ProvConstituida, "#,###.00")

                Case "3": xlHoja1.Range("E15") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("E20") = Format(rs!ProvConstituida, "#,###.00")

                Case "4": xlHoja1.Range("F15") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("F20") = Format(rs!ProvConstituida, "#,###.00")
            End Select
        Case "4"
            Select Case rs!cCalGen
                Case "0": xlHoja1.Range("B16") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("B21") = Format(rs!ProvConstituida, "#,###.00")

                Case "1": xlHoja1.Range("C16") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("C21") = Format(rs!ProvConstituida, "#,###.00")

                Case "2": xlHoja1.Range("D16") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("D21") = Format(rs!ProvConstituida, "#,###.00")

                Case "3": xlHoja1.Range("E16") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("E21") = Format(rs!ProvConstituida, "#,###.00")

                Case "4": xlHoja1.Range("F16") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("F21") = Format(rs!ProvConstituida, "#,###.00")

            End Select

        Case "3"

            Select Case rs!cCalGen
                Case "0": xlHoja1.Range("B17") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("B22") = Format(rs!ProvConstituida, "#,###.00")

                Case "1": xlHoja1.Range("C17") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("C22") = Format(rs!ProvConstituida, "#,###.00")

                Case "2": xlHoja1.Range("D17") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("D22") = Format(rs!ProvConstituida, "#,###.00")

                Case "3": xlHoja1.Range("E17") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("E22") = Format(rs!ProvConstituida, "#,###.00")

                Case "4": xlHoja1.Range("F17") = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("F22") = Format(rs!ProvConstituida, "#,###.00")
            End Select

    End Select
    rs.MoveNext
 Wend


 '-----------------------------------------------
 ' Calculos
 '-----------------------------------------------

 For i = 14 To 17
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 19 To 22
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 24 To 27
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 30 To 33
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i
 For i = 40 To 43
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i

 xlHoja1.Range("B34").Formula = "= B30+B31+B32+B33"
 xlHoja1.Range("C34").Formula = "= C30+C31+C32+C33"
 xlHoja1.Range("D34").Formula = "= D30+D31+D32+D33"
 xlHoja1.Range("E34").Formula = "= E30+E31+E32+E33"
 xlHoja1.Range("F34").Formula = "= F30+F31+F32+F33"

 xlHoja1.Range("B24").Formula = "= (B14)*1% "
 xlHoja1.Range("B25").Formula = "= (B15)*1% "
 xlHoja1.Range("B26").Formula = "= (B16)*1% "
 xlHoja1.Range("B27").Formula = "= (B17)*1% "
' xlHoja1.Range("C26").Formula = "= (B33+B42+B51+B58)    *1% + B70+B69+B89* 1%"
' xlHoja1.Range("B27").Formula = "= (B34+B43+B52)        *1% + B70+B69+B90* 1%"

' xlHoja1.Range("C97").Formula = "= (C31*5%)+ (C40*1%) + (C49*1.25%) + (C56*2.5%) + (C61*1%)  + B70+B71+B87* 1%"
 For i = 19 To 22
    xlHoja1.Range("B" & 11 + i).Formula = "= B" & i & "-B" & i + 5
    xlHoja1.Range("C" & 11 + i).Formula = "= C" & i & "-C" & i + 5
    xlHoja1.Range("D" & 11 + i).Formula = "= D" & i & "-D" & i + 5
    xlHoja1.Range("E" & 11 + i).Formula = "= E" & i & "-E" & i + 5
    xlHoja1.Range("F" & 11 + i).Formula = "= F" & i & "-F" & i + 5
 Next
 '--- formato
 '----------

 xlHoja1.Range("B17:G20").Style = "Comma"
 xlHoja1.Range("B40:G43").Style = "Comma"
 xlHoja1.Range("B56:G58").Style = "Comma"

 xlHoja1.Range("A11:G43").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A11:G43").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A11:G43").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A14:G43").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A11:G43").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("A11:G43").Borders(xlEdgeLeft).Weight = xlThin
 xlHoja1.Range("A11:G43").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("A11:G43").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("A13:G13").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A13:G13").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A18:G18").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A18:G18").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A18:G18").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A23:G23").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A23:G23").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A23:G23").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A27:G27").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A27:G27").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A29:G29").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A29:G29").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A33:G33").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A33:G33").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A35:G39").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A35:G39").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A35:G39").Borders(xlEdgeBottom).Weight = xlThin

xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
'CargaArchivo App.path & "\SPOOLER\" & lsArchivo, App.path & "\SPOOLER\"
MsgBox "Se genero el Archivo. Verifique en la Carpeta Spooler", vbInformation, "AVISO"
End Sub



Private Sub ImpAnexo5Cab(pnFila As Integer, psParteCab As String, ByVal sFuente As String, _
                        ByVal dFecha As Date, Optional pbConsumo As Boolean = False, _
                        Optional ByVal psTitulo As String = "", Optional ByVal psNomCmac As String = "", _
                        Optional ByVal psCodCmac As String = "", Optional ByVal psCodAnexo As String = "")

Dim J As Integer
Dim lsNombre As String

If psParteCab = "1" Then
    Select Case psCodAnexo
    Case "5"
        xlHoja1.Cells(pnFila, 6) = "A N E X O     0 5 " & IIf(Len(Trim(sFuente)) > 0, " - Fuente Financ. " & Val(sFuente), "")
    Case "5I"
        xlHoja1.Cells(pnFila, 6) = "A N E X O     0 5 - I " & IIf(Len(Trim(sFuente)) > 0, " - Fuente Financ. " & sFuente, "")
    Case "5G"
        xlHoja1.Cells(pnFila, 6) = "A N E X O     0 5 - G "
    End Select

    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).Font.Bold = True
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 1) = "EMPRESA :  " & psNomCmac: xlHoja1.Cells(pnFila, 11) = "CODIGO : " & psCodCmac
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).Font.Bold = True
    pnFila = pnFila + 1

    Select Case psCodAnexo
    Case "5"
        xlHoja1.Cells(pnFila, 2) = "INFORME DE CLASIFICACION DE LOS DEUDORES DE LA CARTERA DE CREDITOS, CONTINGENTES Y ARRENDAMIENTOS FINANCIEROS"
    Case "5I"
        xlHoja1.Cells(pnFila, 2) = "INFORME DE CLASIFICACION DE DEUDORES DE LA CARTERA CREDITICIA QUE RESPALDA FINANCIAMIENTOS O LINEAS DE CREDITO"
    Case "5G"
        xlHoja1.Cells(pnFila, 2) = "INFORME SOBRE CREDITOS REFINANCIADOS, REESTRUCTURADOS Y PROGRAMAS RFA - FOPE"
    End Select

    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 6) = "AL  " & dFecha
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 6) = "(  E N    N U E V O S    S O L E S   )"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).Merge True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).HorizontalAlignment = xlCenter
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 1) = "1. INFORME POR TIPOS DE CREDITOS "
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 11)).Font.Bold = True
    pnFila = pnFila + 2

    xlHoja1.Cells(pnFila + 1, 1) = "TIPO DE CREDITO "

    xlHoja1.Cells(pnFila, 2) = "NORMAL"
    xlHoja1.Cells(pnFila + 1, 2) = "Tabla 1-2"

    xlHoja1.Cells(pnFila, 3) = "POTENCIALES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 3), xlHoja1.Cells(pnFila, 5)).Merge True
    xlHoja1.Cells(pnFila + 1, 3) = "Tabla 1"
    xlHoja1.Cells(pnFila + 1, 4) = "Tabla 2(a)"
    xlHoja1.Cells(pnFila + 1, 5) = "Tabla 2(b)"

    xlHoja1.Cells(pnFila, 6) = "DEFICIENTES"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 6), xlHoja1.Cells(pnFila, 8)).Merge True
    xlHoja1.Cells(pnFila + 1, 6) = "Tabla 1"
    xlHoja1.Cells(pnFila + 1, 7) = "Tabla 2(a)"
    xlHoja1.Cells(pnFila + 1, 8) = "Tabla 2(b)"

    xlHoja1.Cells(pnFila, 9) = "DUDOSOS"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 9), xlHoja1.Cells(pnFila, 11)).Merge True
    xlHoja1.Cells(pnFila + 1, 9) = "Tabla 1"
    xlHoja1.Cells(pnFila + 1, 10) = "Tabla 2(a)"
    xlHoja1.Cells(pnFila + 1, 11) = "Tabla 2(b)"

    xlHoja1.Cells(pnFila, 12) = "PERDIDA"
    xlHoja1.Range(xlHoja1.Cells(pnFila, 12), xlHoja1.Cells(pnFila, 14)).Merge True
    xlHoja1.Cells(pnFila + 1, 12) = "Tabla 1"
    xlHoja1.Cells(pnFila + 1, 13) = "Tabla 2(a)"
    xlHoja1.Cells(pnFila + 1, 14) = "Tabla 2(b)"

    xlHoja1.Cells(pnFila, 15) = "TOTAL"
    xlHoja1.Range("A1..A36").ColumnWidth = 25
    xlHoja1.Range("B1").ColumnWidth = 15
    xlHoja1.Range("C1").ColumnWidth = 15
    xlHoja1.Range("D1").ColumnWidth = 15
    xlHoja1.Range("E1").ColumnWidth = 15
    xlHoja1.Range("F1").ColumnWidth = 15
    xlHoja1.Range("G1").ColumnWidth = 15
    xlHoja1.Range("H1").ColumnWidth = 15
    xlHoja1.Range("I1").ColumnWidth = 15
    xlHoja1.Range("J1").ColumnWidth = 15
    xlHoja1.Range("K1").ColumnWidth = 15
    xlHoja1.Range("L1").ColumnWidth = 15
    xlHoja1.Range("M1").ColumnWidth = 15
    xlHoja1.Range("N1").ColumnWidth = 15
    xlHoja1.Range("O1").ColumnWidth = 20

    'CuadroExcel 1, pnFila, 11, pnFila + 2, True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 2, 15)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 2, 15)).Cells.Borders.LineStyle = xlInside

    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 2, 15)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 2, 15)).Interior.Color = &HC0C0C0
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 2, 15)).Font.Bold = True

    'CuadroExcel 1, pnFila + 2, 11, pnFila + 26, True
    If psCodAnexo = "5G" Then
        xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 1), xlHoja1.Cells(pnFila + 27, 15)).Cells.Borders.LineStyle = xlOutside
        xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 1), xlHoja1.Cells(pnFila + 27, 15)).Cells.Borders.LineStyle = xlInside
    Else
        xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 1), xlHoja1.Cells(pnFila + 25, 15)).Cells.Borders.LineStyle = xlOutside
        xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 1), xlHoja1.Cells(pnFila + 25, 15)).Cells.Borders.LineStyle = xlInside
    End If

    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 2), xlHoja1.Cells(pnFila + 25, 15)).NumberFormat = "#,#0.00"
    If psCodAnexo = "5G" Then
        CuadroExcel 1, pnFila, 15, pnFila + 27, True
    Else
        CuadroExcel 1, pnFila, 15, pnFila + 25, True
    End If

   pnFila = pnFila + 3


ElseIf psParteCab = "2" Then

    xlHoja1.Cells(pnFila + 1, 1) = "TIPO DE CREDITO "

    xlHoja1.Cells(pnFila, 2) = "NORMAL"
    xlHoja1.Cells(pnFila + 1, 2) = "Tabla 3"

    xlHoja1.Cells(pnFila, 3) = "POTENCIALES"
    xlHoja1.Cells(pnFila + 1, 3) = "Tabla 3"

    xlHoja1.Cells(pnFila, 4) = "DEFICIENTE"
    xlHoja1.Cells(pnFila + 1, 4) = "Tabla 3"

    xlHoja1.Cells(pnFila, 5) = "DUDOSOS"
    xlHoja1.Cells(pnFila + 1, 5) = "Tabla 3"

    xlHoja1.Cells(pnFila, 6) = "PERDIDA"
    xlHoja1.Cells(pnFila + 1, 6) = "Tabla 3"

    xlHoja1.Cells(pnFila, 7) = "TOTAL"

    'CuadroExcel 1, pnFila, 7, pnFila + 1, True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 7)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 7)).Cells.Borders.LineStyle = xlInside
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 2, 7)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 2, 7)).Font.Bold = True

    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 2, 7)).Interior.Color = &HC0C0C0

    'CuadroExcel 1, pnFila + 2, 7, pnFila + 9, True
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 1), xlHoja1.Cells(pnFila + 8, 7)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 1), xlHoja1.Cells(pnFila + 8, 7)).Cells.Borders.LineStyle = xlInside

    CuadroExcel 1, pnFila, 7, pnFila + 8, True
    pnFila = pnFila + 2

Else

    pnFila = pnFila + 3
    xlHoja1.Cells(pnFila, 1) = "2. INFORME POR CATEGORIA DE DEUDORES "
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 15)).Font.Bold = True
    pnFila = pnFila + 2


    xlHoja1.Cells(pnFila, 1) = "CATEGORIA DE "
    xlHoja1.Cells(pnFila + 1, 1) = "DEUDORES"

    xlHoja1.Cells(pnFila, 2) = "Nro DE "
    xlHoja1.Cells(pnFila + 1, 2) = "DEUDORES"

    xlHoja1.Cells(pnFila, 3) = "MONTO"

    xlHoja1.Cells(pnFila, 4) = "PROVISION "
    xlHoja1.Cells(pnFila + 1, 4) = "REQUERIDA"

    xlHoja1.Cells(pnFila, 5) = "PROVISION"
    xlHoja1.Cells(pnFila + 1, 5) = "CONSTITUIDA"

    xlHoja1.Cells(pnFila, 6) = "DEFIC/EXCES"

    xlHoja1.Cells(pnFila, 7) = "DEFIC/EXCES"
    xlHoja1.Cells(pnFila + 1, 7) = "%"

    xlHoja1.Cells(pnFila, 8) = "GARANTIAS "
    xlHoja1.Cells(pnFila + 1, 8) = "PREFERIDAS"

    xlHoja1.Cells(pnFila, 9) = "GARANTIAS PREFERIDAS DE "
    xlHoja1.Cells(pnFila + 1, 9) = "MUY RAPIDA REALIZACION"

    xlHoja1.Cells(pnFila + 2, 1) = "Normal"
    xlHoja1.Cells(pnFila + 3, 1) = "Potencial"
    xlHoja1.Cells(pnFila + 4, 1) = "Deficiente"
    xlHoja1.Cells(pnFila + 5, 1) = "Dudoso"
    xlHoja1.Cells(pnFila + 6, 1) = "Pérdida"
    xlHoja1.Cells(pnFila + 7, 1) = " "

    'CuadroExcel 1, pnFila, 10, pnFila + 1, True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 10)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 10)).Cells.Borders.LineStyle = xlInside

    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 8)).HorizontalAlignment = xlCenter
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 8)).Font.Bold = True

    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila + 1, 10)).Interior.Color = &HC0C0C0
    'CuadroExcel 1, pnFila + 2, 10, pnFila + 6, True
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 1), xlHoja1.Cells(pnFila + 6, 10)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(pnFila + 2, 1), xlHoja1.Cells(pnFila + 6, 8)).Cells.Borders.LineStyle = xlInside
    CuadroExcel 1, pnFila, 10, pnFila + 6, True
    CuadroExcel 1, pnFila + 7, 10, pnFila + 7, True
End If
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 70
xlHoja1.PageSetup.Orientation = xlLandscape
xlAplicacion.Range("A1:Z100").Font.Size = 9
End Sub

Private Sub CuadroExcel(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, Optional lbLineasVert As Boolean = False)
Dim i, J As Integer

For i = X1 To X2
    xlHoja1.Range(xlHoja1.Cells(Y1, i), xlHoja1.Cells(Y1, i)).Borders(xlEdgeTop).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(Y2, i), xlHoja1.Cells(Y2, i)).Borders(xlEdgeBottom).LineStyle = xlContinuous
Next i
If lbLineasVert = False Then
    For i = X1 To X2
        For J = Y1 To Y2
            xlHoja1.Range(xlHoja1.Cells(J, i), xlHoja1.Cells(J, i)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        Next J
    Next i
End If
If lbLineasVert Then
    For J = Y1 To Y2
        xlHoja1.Range(xlHoja1.Cells(J, X1), xlHoja1.Cells(J, X1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    Next J
End If

For J = Y1 To Y2
    xlHoja1.Range(xlHoja1.Cells(J, X2), xlHoja1.Cells(J, X2)).Borders(xlEdgeRight).LineStyle = xlContinuous
Next J
End Sub




Public Function GetUltimaFechaCierre() As Date
Dim loConecta As DConecta
Dim rs As ADODB.Recordset
Dim Sql As String
Sql = "select nConsSisValor from constsistema where nConsSisCod=14"
Set loConecta = New DConecta
Set rs = New ADODB.Recordset
loConecta.AbreConexion
Set rs = loConecta.CargaRecordSet(Sql)
loConecta.CierraConexion
GetUltimaFechaCierre = rs!NCONSSISVALOR
End Function
Public Function ObtieneDescripcionFuenteFinanciamiento(ByVal psFuente As String) As String
Dim Sql As String
Dim ServerCons As String
Dim rs As New ADODB.Recordset

Dim RCD As nRcdReportes
Set RCD = New nRcdReportes
ServerCons = RCD.GetServerConsol
Sql = "select cDescripcion from " & ServerCons & "ColocLineaCredito where clineaCred='" & psFuente & "'"
Set rs = RCD.GetQuery(Sql)
ObtieneDescripcionFuenteFinanciamiento = IIf(IsNull(rs!cDescripcion), " ", rs!cDescripcion)
Set RCD = Nothing
Set rs = Nothing
End Function

Public Function ReporteCreditosEvaluados(ByVal pdFecha As Date, ByVal pnMicro As Byte, ByVal pnComer As Byte, _
ByVal pnConsu As Byte, ByVal pnHipot As Byte, ByVal pnMonedaN As Byte, pnMonedaE As Byte) As Recordset


Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsMoneda As String
Dim cadena As String
Dim lsSQL As String


    cadena = " and CED.cCtaCod in ("
    If pnMicro = 1 Then
        cadena = cadena & "201,202,203,"
    End If

    If pnComer = 1 Then
        cadena = cadena & "101,102,103,"
    End If

    If pnConsu = 1 Then
        cadena = cadena & "301,302,303,304,"
    End If

    If pnHipot = 1 Then
        cadena = cadena & "401,423,"
    End If

    cadena = Left(cadena, Len(cadena) - 1)

    cadena = cadena & ")"

    lsMoneda = " "
    If pnMonedaN = 1 And pnMonedaE = 0 Then
        lsMoneda = " and substring(CED.cCtaCod,9,1)=1"
    End If

    If pnMonedaN = 0 And pnMonedaE = 1 Then
        lsMoneda = " and substring(CED.cCtaCod,9,1)=2"
    End If


lsSQL = " Select Desembolso=(Select dVigencia from Colocaciones Where cCtaCod=CED.cCtaCod), "
lsSQL = lsSQL & " CED.cCtaCod, Per.cPersNombre Cliente,"
lsSQL = lsSQL & " Col.cLineaCred, nDiasAtraso,"
lsSQL = lsSQL & " Sal.nSaldoCap, cEvalCalif,"
lsSQL = lsSQL & " cEvalCalifDet,"
lsSQL = lsSQL & " Dni = (Select cPersIdNro from PersId where cPersCod=Per.cPersCod and cPersIdTpo=1),"
lsSQL = lsSQL & " Ruc = (Select cPersIdNro from PersId where cPersCod=Per.cPersCod and cPersIdTpo=2),"
lsSQL = lsSQL & " Per.cPersDireccDomicilio,"
lsSQL = lsSQL & " Zona=(Select cUbigeoDescripcion from  ubicacionGeografica where cubiGeoCod=cPersDireccubiGeo)"
lsSQL = lsSQL & " from ColocEvalCalif CEC"
lsSQL = lsSQL & " Inner Join ColocEvalCalifDetalle CED on CED.cPersCod=CEC.cPersCod"
lsSQL = lsSQL & " Inner Join ColocacSaldo Sal on Sal.cCtaCod=CED.cCtaCod"
lsSQL = lsSQL & " Inner Join Persona Per on Per.cPersCod=CEC.cPersCod"
lsSQL = lsSQL & " Inner Join Colocaciones Col on Col.cCtaCod=CED.cCtaCod"
lsSQL = lsSQL & " Where   CEC.nEvalTipo=0 and sal.dFecha='" & Format(pdFecha, "mm/dd/yyyy") & "' "
lsSQL = lsSQL & cadena & lsMoneda
lsSQL = lsSQL & " Order by  substring(CED.cCtaCod,9,1),Per.cPersNombre"




    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrRs = loReg.CargaRecordSet(lsSQL)
    If lrRs Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrRs.BOF And lrRs.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If

    Set ReporteCreditosEvaluados = lrRs
    Set lrRs = Nothing
End Function


Public Function ReporteCreditosRevisados(ByVal pdFecha As Date, ByVal pnMicro As Byte, ByVal pnComer As Byte, _
ByVal pnConsu As Byte, ByVal pnHipot As Byte, ByVal pnMonedaN As String, ByVal pnMonedaE As String) As ADODB.Recordset


Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsMoneda As String
Dim cadena As String
Dim lsSQL As String


    cadena = " and DEC.cCtaCod in ("
    If pnMicro = 1 Then
        cadena = cadena & "201,202,203,"
    End If

    If pnComer = 1 Then
        cadena = cadena & "101,102,103,"
    End If

    If pnConsu = 1 Then
        cadena = cadena & "301,302,303,304,"
    End If

    If pnHipot = 1 Then
        cadena = cadena & "401,423,"
    End If

    cadena = Left(cadena, Len(cadena) - 1)
    cadena = cadena & cadena & ")"

    lsMoneda = " "
    If pnMonedaN = 1 And pnMonedaE = 0 Then
        lsMoneda = " and substring(DEC.cCtaCod,9,1)=1"
    End If

    If pnMonedaN = 0 And pnMonedaE = 1 Then
        lsMoneda = " and substring(DEC.cCtaCod,9,1)=2"
    End If


lsSQL = " Select Desembolso=(Select dVigencia from Colocaciones Where cCtaCod=CED.cCtaCod), "
lsSQL = lsSQL & " CED.cCtaCod, Per.cPersNombre Cliente,"
lsSQL = lsSQL & " Col.cLineaCred, nDiasAtraso,"
lsSQL = lsSQL & " Sal.nSaldoCap, cEvalCalif,"
lsSQL = lsSQL & " cEvalCalifDet,"
lsSQL = lsSQL & " Dni = (Select cPersIdNro from PersId where cPersCod=Per.cPersCod and cPersIdTpo=1),"
lsSQL = lsSQL & " Ruc = (Select cPersIdNro from PersId where cPersCod=Per.cPersCod and cPersIdTpo=2),"
lsSQL = lsSQL & " Per.cPersDireccDomicilio,"
lsSQL = lsSQL & " Zona=(Select cUbigeoDescripcion from  ubicacionGeografica where cubiGeoCod=cPersDireccubiGeo)"
lsSQL = lsSQL & " from ColocEvalCalif CEC"
lsSQL = lsSQL & " Inner Join ColocEvalCalifDetalle CED on CED.cPersCod=CEC.cPersCod"
lsSQL = lsSQL & " Inner Join ColocacSaldo Sal on Sal.cCtaCod=CED.cCtaCod"
lsSQL = lsSQL & " Inner Join Persona Per on Per.cPersCod=CEC.cPersCod"
lsSQL = lsSQL & " Inner Join Colocaciones Col on Col.cCtaCod=CED.cCtaCod"
lsSQL = lsSQL & " Where   CEC.nEvalTipo=1 and sal.dFecha='" & Format(pdFecha, "mm/dd/yyyy") & "' "
lsSQL = lsSQL & cadena & lsMoneda
lsSQL = lsSQL & " Order by  substring(CED.cCtaCod,9,1),Per.cPersNombre"




    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrRs = loReg.CargaRecordSet(lsSQL)
    If lrRs Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrRs.BOF And lrRs.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If

    Set ReporteCreditosRevisados = lrRs
    Set lrRs = Nothing
End Function

Public Function ReporteGeneralProvisionesPyme_Comercial(ByVal psServer As String, ByVal pnTipo As Byte, ByVal pnEstado As Byte, _
psMoneda As String, ByVal psSigno As String, ByVal pnMonto As Currency, ByVal psLineaCred As String, _
Optional Flag As Boolean = True, Optional pnCalGen As String = "", Optional psCodCmac As String) As ADODB.Recordset
Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsCadena As String
Dim lsSQL As String


lsCadena = IIf(psMoneda = "1", " and substring(cCtaCod,9,1)=1 ", " and substring(cCtaCod,9,1)=2 ")
Select Case psCodCmac
    Case "112"
        If pnTipo = 0 Then
            lsCadena = lsCadena & " and substring(cCtaCod,6,3)=101"
        Else
            lsCadena = lsCadena & " and substring(cCtaCod,6,3)=201"
        End If
    Case "102"
        If pnTipo = 0 Then
            lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('101','102','103','121')"
        Else
            lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('201','221')"
        End If
End Select

If pnEstado = 0 Then
    lsCadena = lsCadena & " and cRefinan='N' "
  Else
    lsCadena = lsCadena & " and cRefinan='R' "
End If
lsCadena = lsCadena & " and  substring(cLineaCred,1,4)='" & psLineaCred & "' "


If Flag Then
lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total"
lsSQL = lsSQL & " from Agencias Ag Left outer join"
lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  From ColocCalifProv"
lsSQL = lsSQL & " where nMontoApr" & psSigno & pnMonto
lsSQL = lsSQL & lsCadena
lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
lsSQL = lsSQL & " ORDER BY cAgeCod"

Else
lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
lsSQL = lsSQL & " from Agencias Ag Left outer join"
lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan From ColocCalifProv"
lsSQL = lsSQL & " where nMontoApr" & psSigno & pnMonto
lsSQL = lsSQL & " and cCalGen=" & pnCalGen
lsSQL = lsSQL & lsCadena
lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
lsSQL = lsSQL & " ORDER BY cAgeCod"

End If

    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrRs = loReg.CargaRecordSet(lsSQL)
    If lrRs Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrRs.BOF And lrRs.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If

    Set ReporteGeneralProvisionesPyme_Comercial = lrRs
    Set lrRs = Nothing
End Function

Public Function ReporteGeneralProvisionesPignoraticio(Optional Flag As Boolean = True, Optional pnCalGen As String = "") As ADODB.Recordset
Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsCadena As String
Dim lsSQL As String


'lscadena = IIf(psMoneda = "1", " and substring(cCtaCod,9,1)=1 ", " and substring(cCtaCod,9,1)=2 ")
lsCadena = "  substring(cCtaCod,6,3)='305' "

If Flag Then
lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total"
lsSQL = lsSQL & " from Agencias Ag Left outer join"
lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  from ColocCalifProv"
lsSQL = lsSQL & " where " 'nMontoApr" & psSigno & pnMonto
lsSQL = lsSQL & lsCadena
lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
lsSQL = lsSQL & " ORDER BY cAgeCod"

Else
lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
lsSQL = lsSQL & " from Agencias Ag Left outer join"
lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  from ColocCalifProv"
lsSQL = lsSQL & " where " ' nMontoApr" & psSigno & pnMonto
lsSQL = lsSQL & " cCalGen=" & pnCalGen
lsSQL = lsSQL & " and " & lsCadena
lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
lsSQL = lsSQL & " ORDER BY cAgeCod"

End If

    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrRs = loReg.CargaRecordSet(lsSQL)
    If lrRs Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrRs.BOF And lrRs.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If

    Set ReporteGeneralProvisionesPignoraticio = lrRs
    Set lrRs = Nothing
End Function

Public Function ReporteGeneralProvisionesAgricola(ByVal pnTipo As Byte, ByVal pnEstado As Byte, _
psMoneda As String, ByVal psSigno As String, ByVal pnMonto As Currency, ByVal psLineaCred As String, _
Optional Flag As Boolean = True, Optional pnCalGen As String = "") As ADODB.Recordset
Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsCadena As String
Dim lsSQL As String


lsCadena = IIf(psMoneda = "1", " and substring(cCtaCod,9,1)=1 ", " and substring(cCtaCod,9,1)=2 ")
If pnTipo And gsCodCMAC = "112" Then ' Trujillo
    lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('102','103')"
  Else
    lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('202','203')"
End If
If pnTipo And gsCodCMAC = "102" Then
    lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('104')"
  Else
    lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('204')"
End If
If pnEstado = 0 Then
    lsCadena = lsCadena & " and cRefinan='N' "
  Else
    lsCadena = lsCadena & " and cRefinan='R' "
End If
lsCadena = lsCadena & " and  substring(cLineaCred,1,4)=" & psLineaCred


If Flag Then
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  from ColocCalifProv"
    lsSQL = lsSQL & " where nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & lsCadena
    lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
    lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
    lsSQL = lsSQL & " ORDER BY cAgeCod"
Else
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  from ColocCalifProv"
    lsSQL = lsSQL & " where nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & " and cCalGen=" & pnCalGen
    lsSQL = lsSQL & lsCadena
    lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
    lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
    lsSQL = lsSQL & " ORDER BY cAgeCod"
End If

    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrRs = loReg.CargaRecordSet(lsSQL)
    If lrRs Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrRs.BOF And lrRs.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If

    Set ReporteGeneralProvisionesAgricola = lrRs
    Set lrRs = Nothing
End Function

Public Function ReporteGeneralProvisionesHipotecario(ByVal pnTipo As Byte, ByVal pnEstado As Byte, _
psMoneda As String, ByVal psSigno As String, ByVal pnMonto As Currency, ByVal psLineaCred As String, _
Optional Flag As Boolean = True, Optional pnCalGen As String = "", Optional ByVal psCodCmac As String = "") As ADODB.Recordset
Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsCadena As String
Dim lsSQL As String


lsCadena = IIf(psMoneda = "1", " and substring(cCtaCod,9,1)=1 ", " and substring(cCtaCod,9,1)=2 ")
If psCodCmac = "112" Then ' Trujillo
    lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('401','402')"
End If
If psCodCmac = "102" Then
    lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('401','402')"

End If
If pnEstado = 0 Then
    lsCadena = lsCadena & " and cRefinan='N' "
  Else
    lsCadena = lsCadena & " and cRefinan='R' "
End If
lsCadena = lsCadena & " and  substring(cLineaCred,1,4)=" & psLineaCred


If Flag Then
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  from ColocCalifProv"
    lsSQL = lsSQL & " where nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & lsCadena
    lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
    lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
    lsSQL = lsSQL & " ORDER BY cAgeCod"
Else
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  from ColocCalifProv"
    lsSQL = lsSQL & " where nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & " and cCalGen=" & pnCalGen
    lsSQL = lsSQL & lsCadena
    lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
    lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
    lsSQL = lsSQL & " ORDER BY cAgeCod"
End If

    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrRs = loReg.CargaRecordSet(lsSQL)
    If lrRs Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrRs.BOF And lrRs.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If

    Set ReporteGeneralProvisionesHipotecario = lrRs
    Set lrRs = Nothing
End Function


Public Function Get_Agencias() As ADODB.Recordset
Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsSQL As String

lsSQL = " select cAgeCod, cAgeDescripcion Nombre from Agencias "
    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrRs = loReg.CargaRecordSet(lsSQL)


    If lrRs.BOF And lrRs.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If

    Set Get_Agencias = lrRs
    Set lrRs = Nothing
End Function
Public Function Get_Financiamiento() As ADODB.Recordset
Dim lrRs As New ADODB.Recordset
Dim lsSQL As String
Dim nCred As nCredRepoFinMes
Set nCred = New nCredRepoFinMes
Set lrRs = nCred.ObtieneLineaCredito
    If lrRs.BOF And lrRs.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If

    Set Get_Financiamiento = lrRs
    Set lrRs = Nothing
    Set nCred = Nothing

End Function

Public Function ReporteGeneralProvisionesPersonales(ByVal pnEstado As Byte, _
psMoneda As String, ByVal psSigno As String, ByVal pnMonto As Currency, ByVal psLineaCred As String, _
Optional Flag As Boolean = True, Optional pnCalGen As String = "") As ADODB.Recordset
Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsCadena As String
Dim lsSQL As String

 'cambio lima
lsCadena = IIf(psMoneda = "1", " and substring(cCtaCod,9,1)='1' ", " and substring(cCtaCod,9,1)='2' ")
lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('301','302','303','304','320')"

If pnEstado = 0 Then
    lsCadena = lsCadena & " and cRefinan='N' "
  Else
    lsCadena = lsCadena & " and cRefinan='R' "
End If
lsCadena = lsCadena & " and  substring(cLineaCred,1,4)=" & psLineaCred


If Flag Then
lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total"
lsSQL = lsSQL & " from Agencias Ag Left outer join"
lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan From ColocCalifProv"
lsSQL = lsSQL & " where nMontoApr" & psSigno & pnMonto
lsSQL = lsSQL & lsCadena
lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
lsSQL = lsSQL & " ORDER BY cAgeCod"

Else
lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
lsSQL = lsSQL & " from Agencias Ag Left outer join"
lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan From ColocCalifProv "
lsSQL = lsSQL & " where nMontoApr" & psSigno & pnMonto
lsSQL = lsSQL & " and cCalGen=" & pnCalGen
lsSQL = lsSQL & lsCadena
lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
lsSQL = lsSQL & " ORDER BY cAgeCod"

End If

    Set loReg = New DConecta
    loReg.AbreConexion
    Set lrRs = loReg.CargaRecordSet(lsSQL)
    If lrRs Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Validación ", vbInformation, "Aviso"
        Exit Function
    End If
    If lrRs.BOF And lrRs.EOF Then
        MsgBox "ERROR: No existen datos ", vbInformation, "Aviso"
        Exit Function
    End If

    Set ReporteGeneralProvisionesPersonales = lrRs
    Set lrRs = Nothing
End Function

Public Function Cabecera_Pro(ByVal NomCmac As String, ByVal NomAge As String, FecSis As Date, NroPag As Integer, FecCierre As Date, psTitTipoCred As String, nMoneda As String, nrpRepo As Long, _
ByVal pnSigMonto As String) As String
Dim lsCadenaPrint As String



lsCadenaPrint = ImpreFormat(NomCmac, 30, 0, False) & Space(180) & "Pagina: " & NroPag & Chr(10)
lsCadenaPrint = lsCadenaPrint & ImpreFormat(NomAge, 30, 0, False) & Space(180) & "Fecha : " & FecSis & Chr(10)
lsCadenaPrint = lsCadenaPrint & Space(80) & "<<Provi>> : SALDOS Y PROVISIONES PARA CALIFICACION DE LA CARTERA DE CREDITOS AL : " & FecCierre & Chr(10)
If nrpRepo = 178204 Then
    lsCadenaPrint = lsCadenaPrint & "Creditos : " & ImpreFormat(psTitTipoCred, 30, 0, False) & Chr(10)
Else

    lsCadenaPrint = lsCadenaPrint & "Creditos : " & ImpreFormat(psTitTipoCred, 30, 0, False) & "  Moneda : " & IIf(nMoneda = "1", "Soles", "Dolares") & Space(10) & "Monto Aprobado " & pnSigMonto & " " & IIf(nMoneda = "1", "S/.", "U$") & Chr(10)
End If
lsCadenaPrint = lsCadenaPrint & String(230, "-") & Chr(10)

Select Case nrpRepo
 Case 178201, 178202, 178204:
      lsCadenaPrint = lsCadenaPrint & Space(18) & "N° Total" & Space(3) & "SK TOTAL" & Space(2) & "#NORMAL" & Space(3) & "SK NORMAL" & Space(2) & "PRO NORMAL" & Space(3) & " & # P.P" & Space(3) & "SK PROB.POT" & Space(3) & "PRO.PROB.POT" & Space(3) & "# DEF" & Space(3) & "SK DEFIC." & Space(3) & "PRO.DEFIC" & Space(4) & "# DUDOS" & Space(3) & "SK DUDOS." & Space(3) & "PRO.DUDOSO" & Space(3) & "# PERDIDA" & Space(3) & "SK PERDID." & Space(3) & "PRO.PERDIDA" & Space(4) & "PRO TOTAL" & Chr(10)

Case 178203:
      lsCadenaPrint = lsCadenaPrint & Space(18) & "N° Total" & Space(3) & "SK TOTAL" & Space(3) & " & # P.P" & Space(3) & "SK PROB.POT" & Space(3) & "PRO.PROB.POT" & Space(3) & "# DEF" & Space(3) & "SK DEFIC." & Space(3) & "PRO.DEFIC" & Space(4) & "# DUDOS" & Space(3) & "SK DUDOS." & Space(3) & "PRO.DUDOSO" & Space(3) & "# PERDIDA" & Space(3) & "SK PERDID." & Space(3) & "PRO.PERDIDA" & Space(4) & "PRO TOTAL" & Chr(10)

End Select
lsCadenaPrint = lsCadenaPrint & String(230, "-") & Chr(10)

Cabecera_Pro = lsCadenaPrint
End Function


Private Sub Class_Initialize()
Set RCD = New nRcdReportes
ServerCons = RCD.GetServerConsol

End Sub

Private Sub Class_Terminate()
Set RCD = Nothing
End Sub


