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

Public Sub Genera_Reporte178101(ByVal psServer As String, ByVal lsFecha As String, ByVal nChkCor As Byte, _
ByVal nChkGraEmp As Byte, ByVal nChkMedEmp As Byte, ByVal nChkPeqEmp As Byte, ByVal nChkMicEmp As Byte, _
ByVal nChkConRev As Byte, ByVal nChkConNoRev As Byte, ByVal nChkHip As Byte)
Dim sql As String


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
    'lsCadena = "  AND SUBSTRING(AD.cCtaCod,6,1) IN ("
    lsCadena = "  AND SUBSTRING(C.cTpoCredCod,1,1) IN (" 'BRGO 20100608 BASILEA II
    
    'BRGO 20100608 BASILEA II
    If nChkCor = 1 Then
         lsCadena = lsCadena & "'1',"
    End If
    If nChkGraEmp = 1 Then
        lsCadena = lsCadena & "'2',"
    End If
    If nChkMedEmp = 1 Then
        lsCadena = lsCadena & "'3',"
    End If
    If nChkPeqEmp = 1 Then
        lsCadena = lsCadena & "'4',"
    End If
    If nChkMicEmp = 1 Then
        lsCadena = lsCadena & "'5',"
    End If
    If nChkConRev = 1 Then
        lsCadena = lsCadena & "'6',"
    End If
    If nChkConNoRev = 1 Then
        lsCadena = lsCadena & "'7',"
    End If
    If nChkHip = 1 Then
        lsCadena = lsCadena & "'8',"
    End If
    ' Fin BRGO
    
    RaiseEvent Progress(1, 10)
    If InStr(Len(lsCadena) - 2, lsCadena, "',", vbTextCompare) <> 0 Then
        lsCadena = Mid(lsCadena, 1, Len(lsCadena) - 1)
    End If
    lsCadena = lsCadena & ")"
    'If nChkMicro = 0 And nChkComer = 0 And nChkConsu = 0 And nChkHipo = 0 Then
    If nChkCor = 0 And nChkGraEmp = 0 And nChkMedEmp = 0 And nChkPeqEmp = 0 And nChkMicEmp = 0 And nChkConRev = 0 And nChkConNoRev = 0 And nChkHip = 0 Then ' BRGO 20100608 BASILEA II
        lsCadena = ""
    End If
    'BRGO 20100806
    lsNomHoja = Format(lsFecha, "ddmmyyyy") & "-" _
                & IIf(nChkCor = 1, "Micro" & "-", "") _
                & IIf(nChkGraEmp = 1, "Gran.Emp" & "-", "") _
                & IIf(nChkMedEmp = 1, "Med.Emp" & "-", "") _
                & IIf(nChkPeqEmp = 1, "Peq.Emp" & "-", "") _
                & IIf(nChkMicEmp = 1, "Mic.Emp" & "-", "") _
                & IIf(nChkConRev = 1, "Cons.Rev" & "-", "") _
                & IIf(nChkConNoRev = 1, "Cons.No Rev" & "-", "") _
                & IIf(nChkHip = 1, "Hipo", "")

    lsNomHoja = IIf(Right(lsNomHoja, 1) = "-", Mid(lsNomHoja, 1, Len(lsNomHoja) - 1), lsNomHoja)
    RaiseEvent Progress(3, 10)
        lsArchivo = "CredEval" & Format(ldFechaFinMes, "ddmmyyyy") & ".XLS"
        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
            Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & lsArchivo)
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
            xlHoja1.Name = Mid(lsNomHoja, 1, 31) 'brgo 20100806 Se corta la longitud del nombre del archivo
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

        sql = " SELECT DISTINCT  P.cPersNombre,"
        sql = sql & " ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=1) ,'') AS cNudoci,"
        sql = sql & " ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=2) ,'') AS cNudotr,"
        sql = sql & " AP.cEvalCalif cCalGen, AD.cCtaCod, AD.dEval DFECCAL, CS.NSALDOCAP,"
        sql = sql & " CS.NDIASATRASO, SUBSTRING(AD.cCtaCod,9,1) cValor, cNomTab  = (case when SUBSTRING(AD.cCtaCod,9,1)='1' then 'Soles' else 'Dolares' end),"
        sql = sql & " AD.cEvalCalifDet cCalAud"
        sql = sql & " FROM ColocEvalCalif AP"
        sql = sql & " JOIN Persona P  ON P.cPersCod =AP.cPersCod"
        sql = sql & " JOIN ColocEvalCalifDetalle AD  ON AD.cPersCod=AP.cPersCod"
        sql = sql & " JOIN ColocacSaldo CS ON CS.cCtaCod=AD.cCtaCOd"
        sql = sql & " JOIN Colocaciones C ON C.cCtaCod = CS.cCtaCod" '20100608 BRGO
        sql = sql & " WHERE "
        sql = sql & " datediff(day,AD.dEval,'" & Format(lsFecha, "mm/dd/yyyy") & "')  = 0 and "
        sql = sql & " datediff(day,CS.dFecha,'" & Format(lsFecha, "mm/dd/yyyy") & "')  = 0  "
        sql = sql & lsCadena
        'SQL = SQL & " --AD.cEvalESTADO IN ('1'" & IIf(chkCancelados.value = 1, ",'0'", "") & ")"
        sql = sql & " ORDER BY SUBSTRING(AD.cCtaCod,9,1), P.cPersNombre, AD.dEval"
        

        Conecta.AbreConexion
            Set rs = Conecta.CargaRecordSet(sql)
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
                    xlHoja1.Cells(lnFila, 1) = Trim(rs!cNomTab) & " " _
                            & IIf(nChkCor = 1, "-" & "Corporat", "") _
                            & IIf(nChkGraEmp = 1, "-" & "Gran.Emp", "") _
                            & IIf(nChkMedEmp = 1, "-" & "Med.Emp", "") _
                            & IIf(nChkPeqEmp = 1, "-" & "Peq.Emp", "") _
                            & IIf(nChkMicEmp = 1, "-" & "Mic.Emp", "") _
                            & IIf(nChkConRev = 1, "-" & "Cons.Rev", "") _
                            & IIf(nChkConNoRev = 1, "-" & "Cons.No Rev", "") _
                            & IIf(nChkHip = 1, "-" & "Hipo", "")
                            
                            
                            

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

                sql = " Select C.dFecApr, C.nSaldoCap , C.nDiasAtraso, C.dFecVig,"
                sql = sql & " LineaCred = (Select cDescripcion FROM ColocLineaCredito where cLineaCred = C.cLineaCred ),"
                sql = sql & " FI.cDireccion as cDirFI, FI.cCodZon as cZonFI, ZFI.cUbiGeoDescripcion as cDesZonFI"
                sql = sql & " FROM " & psServer & "CREDITOCONSOL C"
                sql = sql & " LEFT JOIN " & psServer & "FUENTEINGRESOConsol FI ON FI.cNumFuente = C.cNumFuente"
                sql = sql & " LEFT JOIN UBICACIONgEOGRAFICA ZFI ON FI.cCodZon = ZFI.cUbiGeoCod"
                sql = sql & " WHERE C.cCtaCod ='" & rs!cCtaCod & "' "
                sql = sql & " UNION "
                sql = sql & " Select C.dFecApr, C.nMontoApr as nSaldoCap , 0 as nDiasAtraso, C.dFecVig,"
                sql = sql & " LineaCred = 'Carta Fianza' ,"
                sql = sql & " FI.cDireccion as cDirFI, FI.cCodZon as cZonFI, ZFI.cUbiGeoDescripcion as cDesZonFI"
                sql = sql & " FROM " & psServer & "CartaFianzaCONSOL C"
                sql = sql & " LEFT JOIN " & psServer & "FUENTEINGRESOConsol FI ON FI.cNumFuente = C.cNumFuente"
                sql = sql & " LEFT JOIN UBICACIONgEOGRAFICA ZFI ON FI.cCodZon = ZFI.cUbiGeoCod"
                sql = sql & " WHERE C.cCtaCod ='" & rs!cCtaCod & "' "
                

                'rsFuente.CursorLocation = adUseClient
                'rsFuente.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                'rsFuente.ActiveConnection = Nothing
                Conecta.AbreConexion
                    Set rsFuente = Conecta.CargaRecordSet(sql)
                Conecta.CierraConexion

                xlHoja1.Cells(lnFila, 1) = "'" & Format(rsFuente!dFecVig, "dd/mm/yyyy")
                xlHoja1.Range(xlHoja1.Cells(lnFila, 1), xlHoja1.Cells(lnFila, 1)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 2) = rs!cCtaCod
                xlHoja1.Range(xlHoja1.Cells(lnFila, 2), xlHoja1.Cells(lnFila, 2)).HorizontalAlignment = xlCenter
                xlHoja1.Cells(lnFila, 3) = PstaNombre(rs!cPersNombre, False)
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
                xlHoja1.Cells(lnFila, 9) = IIf(Trim(rs!cNudoci) = "", Trim(rs!cNudoTr), Trim(rs!cNudoci))
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
xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
RaiseEvent Progress(10, 10)
MsgBox "Se genero el archivo en " & App.Path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
lbExcel = False
RaiseEvent CloseProgress
End Sub

Public Sub Genera_Reporte178102(ByVal lsFecha As String, ByVal nChkMicro As Byte, _
ByVal nChkComer As Byte, ByVal nChkConsu As Byte, ByVal nChkHipo As Byte, ByVal pbCartaFianza As Boolean)
Dim sql As String
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
Dim vExcelObj  As Excel.Application

Set RCD = New nRcdReportes
Server = RCD.GetServerConsol
Set RCD = Nothing
ldFechaFinMes = GetUltimaFechaCierre

RaiseEvent ShowProgress
    lsCadena = "  AND SUBSTRING(CC.cCtaCod,6,1) IN ("
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
    lsCadena = lsCadena & ")"
    If nChkMicro = 0 And nChkComer = 0 And nChkConsu = 0 And nChkHipo = 0 Then
        lsCadena = ""
    End If
    
    If pbCartaFianza = False Then
            sql = "     Select      SUBSTRING(CC.cCtaCod,4,2) as cCodAge,C.cConsDescripcion AS TIPO, CC.cCtaCod, P.cPersNombre,"
            sql = sql & "           ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=1) ,'') AS cNudoci,"
            sql = sql & "           ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=2) ,'') AS cNudotr,"
            sql = sql & "           convert(varchar(10), CC.DFECVIG, 103) as dFecVig, CC.nMontoDesemb, CC.cRefinan, CC.nSaldoCap, CC.nCapVencido, CC.nDiasAtraso, CC.cCodAnalista,"
            sql = sql & "           CC.nHipoSoles, CC.nHipoDol, CC.nGarantSoles, CC.nGarantDol, L.cDescripcion,"
            sql = sql & "           ISNULL((SELECT TOP 1 CT.cCalGen FROM " & Server & "ColocCalifProvTotal CT WHERE CT.CCTACOD = CC.CCTACOD AND convert(char(10),CT.dFecha,112)<'" & Format(lsFecha, "yyyymmdd") & "' ORDER BY CT.dFecha DESC) ,'')  as cCalAnt, CCred.CRFA "
            sql = sql & " From      " & Server & "CreditoConsol CC"
            sql = sql & "           JOIN " & Server & "ProductoPersonaConsol RC ON RC.CCTACOD = CC.CCTACOD"
            sql = sql & "           JOIN PERSONA P ON P.CPERSCOD = RC.cPersCod  and nPrdPersRelac= 20"
            sql = sql & "           JOIN ColocLineaCredito L on L.cLineaCred = CC.cLineaCred"
            sql = sql & "           JOIN Constante C on c.nConsValor = SUBSTRING(CC.cCtaCod,6,3) AND nConsCod= 1001 "
            sql = sql & "           JOIN ColocacCred CCred on cCred.cCtaCod = CC.cCtaCod "
            sql = sql & " WHERE     CC.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107) " & lsCadena
            sql = sql & "           ORDER BY SUBSTRING(CC.cCtaCod,4,2), P.cPersNombre, CC.cCtaCod"
    Else
        
        sql = "  Select             SUBSTRING(CC.cCtaCod,4,2) as cCodAge,C.cConsDescripcion AS TIPO, CC.cCtaCod, P.cPersNombre,"
            sql = sql & "           ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=1) ,'') AS cNudoci,"
            sql = sql & "           ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=2) ,'') AS cNudotr,"
            sql = sql & "           convert(varchar(10), CC.DFECVIG, 103) as dFecVig, CC.nMontoApr as nMontoDesemb, 'N' as cRefinan, CC.nMontoApr as nSaldoCap, 0 as nCapVencido, 0 as nDiasAtraso, CC.cCodAnalista,"
            sql = sql & "           0 AS nHipoSoles, 0 AS nHipoDol, 0 AS nGarantSoles, 0 AS nGarantDol, '' AS cDescripcion,"
            sql = sql & "           ISNULL((SELECT TOP 1 CT.cCalGen FROM " & Server & "ColocCalifProvTotal CT WHERE CT.CCTACOD = CC.CCTACOD AND convert(char(10),CT.dFecha,112)<'" & Format(lsFecha, "yyyymmdd") & "' ORDER BY CT.dFecha DESC) ,'')  as cCalAnt, '' as CRFA "
            sql = sql & " From      " & Server & "CartaFianzaConsol CC"
            sql = sql & "           JOIN " & Server & "ProductoPersonaConsol RC ON RC.CCTACOD = CC.CCTACOD"
            sql = sql & "           JOIN PERSONA P ON P.CPERSCOD = RC.cPersCod  and nPrdPersRelac= 20"
            sql = sql & "           JOIN Constante C on c.nConsValor = SUBSTRING(CC.cCtaCod,6,3) AND nConsCod= 1001 "
            sql = sql & " WHERE     CC.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107) "
            sql = sql & "           ORDER BY SUBSTRING(CC.cCtaCod,4,2), P.cPersNombre, CC.cCtaCod"
    End If
        Conecta.AbreConexion
        Set rs = Conecta.CargaRecordSet(sql)
        Conecta.CierraConexion
        
        If pbCartaFianza Then
            lsArchivo = "ListaCredEvalAnt" & Format(lsFecha, "mmyyyy") & "CF.XLS"
        Else
            lsArchivo = "ListaCredEvalAnt" & Format(lsFecha, "mmyyyy") & ".XLS"
        End If
    
        lsNomHoja = Format(lsFecha, "ddmmyyyy") & "-" _
                    & IIf(nChkMicro = 1, "Micro" & "-", "") _
                    & IIf(nChkComer = 1, "Comer" & "-", "") _
                    & IIf(nChkConsu = 1, "Cons" & "-", "") & IIf(nChkHipo = 1, "Hipo", "")

        Dim lsTipoCred As String
        lsTipoCred = IIf(nChkMicro = 1, "MicroEmpresa" & "-", "") _
                    & IIf(nChkComer = 1, "Comercial" & "-", "") _
                    & IIf(nChkConsu = 1, "Consumo" & "-", "") & IIf(nChkHipo = 1, "Hipotecario", "")
        If pbCartaFianza Then
            lsTipoCred = "CARTAFIANZA"
            lsNomHoja = "CARTAFIANZA"
        Else
            lsTipoCred = IIf(Right(lsTipoCred, 1) = "-", Mid(lsTipoCred, 1, Len(lsTipoCred) - 1), lsTipoCred)
        
            lsNomHoja = IIf(Right(lsNomHoja, 1) = "-", Mid(lsNomHoja, 1, Len(lsNomHoja) - 1), lsNomHoja)
        End If
    
    If Not rs.BOF And Not rs.EOF Then
    
        Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
        vExcelObj.DisplayAlerts = False
        
        vExcelObj.Workbooks.Add
        vExcelObj.Sheets("Hoja1").Select
        vExcelObj.Sheets("Hoja1").Name = lsNomHoja
        
        vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
        vExcelObj.Range("A1:IV65536").Font.Size = 8
        vExcelObj.Columns("A:IV").Select
        vExcelObj.Selection.VerticalAlignment = 3
        
        vExcelObj.Columns("G").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
        vExcelObj.Columns("I:J").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    
        vExcelObj.Columns("M:P").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
        vExcelObj.Range("A6:T6").Select
        vExcelObj.Range("A6:T6").Font.Bold = True
        vExcelObj.Range("A6:T6").HorizontalAlignment = 3
        vExcelObj.Selection.AutoFilter

        
        vExcelObj.Columns("A").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        vExcelObj.Columns("B:H").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        
        vExcelObj.Range("A1").Select
        vExcelObj.Range("A1").Font.Bold = True
        vExcelObj.Range("A1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "CAJA MUNICIPAL DE ICA"
        
        vExcelObj.Range("P1").Select
        vExcelObj.Range("P1").Font.Bold = True
        vExcelObj.Range("P1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = Format(ldFechaFinMes, "dd/mm/yyyy")
        
        vExcelObj.Range("F4").Select
        vExcelObj.Range("F4").Font.Bold = True
        vExcelObj.Range("F4").Font.Size = 12
        vExcelObj.Range("F4").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "LISTADO DE CREDITOS " & lsTipoCred & " CON CALIFICACION ANTERIOR"
        
        vExcelObj.Range("A6").Select
        vExcelObj.Range("A6").Font.Bold = True
        vExcelObj.Range("A6").ColumnWidth = 25
        vExcelObj.ActiveCell.value = "Tipo Credito"
        
        vExcelObj.Range("B6").Select
        vExcelObj.Range("B6").Font.Bold = True
        vExcelObj.Range("B6").ColumnWidth = 20
        vExcelObj.ActiveCell.value = "N° de Credito"
        
        vExcelObj.Range("C6").Select
        vExcelObj.Range("C6").Font.Bold = True
        vExcelObj.Range("C6").ColumnWidth = 40
        vExcelObj.ActiveCell.value = "Titular"
        
        vExcelObj.Range("D6").Select
        vExcelObj.Range("D6").Font.Bold = True
        vExcelObj.Range("D6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "N° Doc.Natural"
        
        vExcelObj.Range("E6").Select
        vExcelObj.Range("E6").Font.Bold = True
        vExcelObj.Range("E6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "N° Doc.Juridico"
        
        vExcelObj.Range("F6").Select
        vExcelObj.Range("F6").Font.Bold = True
        vExcelObj.Range("F6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Vigencia"
        
        vExcelObj.Range("G6").Select
        vExcelObj.Range("G6").Font.Bold = True
        vExcelObj.Range("G6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Desembolso"
        
        vExcelObj.Range("H6").Select
        vExcelObj.Range("H6").Font.Bold = True
        vExcelObj.Range("H6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Refinan."
        
        vExcelObj.Range("I6").Select
        vExcelObj.Range("I6").Font.Bold = True
        vExcelObj.Range("I6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Saldo.Cap."
        
        vExcelObj.Range("J6").Select
        vExcelObj.Range("J6").Font.Bold = True
        vExcelObj.Range("J6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Cap.Vencido"
        
        vExcelObj.Range("K6").Select
        vExcelObj.Range("K6").Font.Bold = True
        vExcelObj.Range("K6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Atraso"
        
        vExcelObj.Range("L6").Select
        vExcelObj.Range("L6").Font.Bold = True
        vExcelObj.Range("L6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Analista"
        
        vExcelObj.Range("M6").Select
        vExcelObj.Range("M6").Font.Bold = True
        vExcelObj.Range("M6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Monto.Gar.Hipo.Soles"
        
        vExcelObj.Range("N6").Select
        vExcelObj.Range("N6").Font.Bold = True
        vExcelObj.Range("N6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Monto.Gar.Hipo.Dolares"
        
        vExcelObj.Range("O6").Select
        vExcelObj.Range("O6").Font.Bold = True
        vExcelObj.Range("O6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Monto.Gar.Soles"
        
        vExcelObj.Range("P6").Select
        vExcelObj.Range("P6").Font.Bold = True
        vExcelObj.Range("P6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Monto.Gar.Dolares"
        
        vExcelObj.Range("Q6").Select
        vExcelObj.Range("Q6").Font.Bold = True
        vExcelObj.Range("Q6").ColumnWidth = 30
        vExcelObj.ActiveCell.value = "Linea Credito"
        
        vExcelObj.Range("R6").Select
        vExcelObj.Range("R6").Font.Bold = True
        vExcelObj.Range("R6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.Ant."
        
        vExcelObj.Range("S6").Select
        vExcelObj.Range("S6").Font.Bold = True
        vExcelObj.Range("S6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "RFA"
        
        
        vExcelObj.Range("T6").Select
        vExcelObj.Range("T6").Font.Bold = True
        vExcelObj.Range("T6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "AGENCIA"
        
        Dim lsCelda As String
        lnFila = 6
        Do While Not rs.EOF
            lnFila = lnFila + 1
            
            lsCelda = "A" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!Tipo
            
            lsCelda = "B" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cCtaCod
            
            lsCelda = "C" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cPersNombre
            
            lsCelda = "D" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cNudoci
            
            lsCelda = "E" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cNudoTr
            
            lsCelda = "F" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!dFecVig
            
            lsCelda = "G" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nMontoDesemb, "#,#0.00")
            
            lsCelda = "H" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cRefinan
            
            lsCelda = "I" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nSaldoCap, "#,#0.00")
            
            lsCelda = "J" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nCapVencido, "#,#0.00")
            
            lsCelda = "K" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!nDiasAtraso
            
            lsCelda = "L" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cCodAnalista
            
            lsCelda = "M" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nHipoSoles, "#,#0.00")
            
            lsCelda = "N" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nhipodol, "#,#0.00")
            
            lsCelda = "O" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarantSoles, "#,#0.00")
            
            lsCelda = "P" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarantDol, "#,#0.00")
            
            lsCelda = "Q" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!CDescripcion
            
            lsCelda = "R" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cCalAnt
            
            lsCelda = "S" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cRFA
            
            lsCelda = "T" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!CCodage
            
            RaiseEvent Progress(rs.Bookmark, rs.RecordCount)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    RaiseEvent CloseProgress
    
    vExcelObj.Range("A1").Select
    vExcelObj.ActiveWorkbook.SaveAs (App.Path & "\SPOOLER\" & lsArchivo)
    vExcelObj.ActiveWorkbook.Close
    MsgBox App.Path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
    vExcelObj.Workbooks.Open (App.Path & "\SPOOLER\" & lsArchivo)
    vExcelObj.Visible = True
    
    Set vExcelObj = Nothing

Set Conecta = Nothing

End Sub
Public Sub Genera_Reporte178104(ByVal lsFecha As String, ByVal lbRefinan As Boolean)
Dim sql As String
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
Dim vExcelObj  As Excel.Application

Set RCD = New nRcdReportes
Server = RCD.GetServerConsol
Set RCD = Nothing
ldFechaFinMes = GetUltimaFechaCierre

RaiseEvent ShowProgress
    If lbRefinan = True Then
        lsCadena = "  AND CC.cRefinan ='R' "
    Else
        lsCadena = "  AND CC.cRefinan ='N' "
    End If

    sql = "     Select          C.cConsDescripcion AS TIPO, CC.cCtaCod, P.cPersNombre,"
        sql = sql & "           ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=1) ,'') AS cNudoci,"
        sql = sql & "           ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=2) ,'') AS cNudotr,"
        sql = sql & "           convert(varchar(10), CC.DFECVIG, 103) as dFecVig , CC.nMontoDesemb, CC.cRefinan, CC.nSaldoCap, CC.nCapVencido, CC.nDiasAtraso, CC.cCodAnalista,"
        sql = sql & "           CC.nHipoSoles, CC.nHipoDol, CC.nGarantSoles, CC.nGarantDol, L.cDescripcion,"
        sql = sql & "           ISNULL((SELECT TOP 1 CT.cCalGen FROM " & Server & "ColocCalifProvTotal CT WHERE CT.CCTACOD = CC.CCTACOD AND convert(char(10),CT.dFecha,112)<'" & Format(lsFecha, "yyyymmdd") & "' ORDER BY CT.dFecha DESC) ,'')  as cCalAnt, "
        sql = sql & "           CCred.CRFA, "
        sql = sql & "           case WHEN   ISNULL((SELECT  count(*) "
        sql = sql & "                               FROM    " & Server & "PlandesPagConsol PD"
        sql = sql & "                               WHERE   datediff(day,PD.dFecVenc,'" & Format(lsFecha, "mm/dd/yyyy") & "')<=180 and PD.dFecPag <>'1900-01-01'"
        sql = sql & "                                       and PD.nEstado = 1 and datediff(day,pd.dfecvenc,PD.dfecpag )<8 and PD.nTipo = 1"
        sql = sql & "                                       and NOT EXISTS (SELECT  PD1.cCtaCod"
        sql = sql & "                                                       FROM    " & Server & "PlandesPagConsol PD1"
        sql = sql & "                                                       WHERE   datediff(day,pd1.dFecVenc,'" & Format(lsFecha, "mm/dd/yyyy") & "')<=180 and pd1.dFecPag <>'1900-01-01'"
        sql = sql & "                                                       and pd1.nEstado = 1 and datediff(day,pd1.dfecvenc,pd1.dfecpag )>=8 and PD1.nTipo = 1"
        sql = sql & "                                                       and PD1.cCtaCod = PD.cCtaCod )"
        sql = sql & "                                       and PD.cCtaCod = CC.cCtaCod and PD.cCtaCod not in (SELECT CCTACOD FROM COLOCACCRED WHERE CRFA IN ('RFA','DIF','RFC') )),-1)>0 THEN 'PUNTUAL' ELSE '' END AS CREFPUNTUAL "
        sql = sql & " From      " & Server & "CreditoConsol CC"
        sql = sql & "           JOIN " & Server & "ProductoPersonaConsol RC ON RC.CCTACOD = CC.CCTACOD "
        sql = sql & "           JOIN PERSONA P ON P.CPERSCOD = RC.cPersCod  and nPrdPersRelac= 20"
        sql = sql & "           JOIN ColocLineaCredito L on L.cLineaCred = CC.cLineaCred"
        sql = sql & "           JOIN Constante C on c.nConsValor = SUBSTRING(CC.cCtaCod,6,3) AND nConsCod= 1001 "
        sql = sql & "           JOIN ColocacCred CCred on cCred.cCtaCod = CC.cCtaCod "
        sql = sql & " WHERE     CC.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107) " & lsCadena
        sql = sql & "           ORDER BY P.cPersNombre, CC.cCtaCod"

        Conecta.AbreConexion
        Set rs = Conecta.CargaRecordSet(sql)
        Conecta.CierraConexion
        
        lsArchivo = "ListaCredXTIPO" & Format(lsFecha, "mmyyyy") & ".XLS"
    
    
        lsNomHoja = IIf(lbRefinan = True, "REFINANCIADOS", "NORMALES")
        
        Dim lsTipoCred As String
        
        lsTipoCred = IIf(lbRefinan = True, "REFINANCIADOS", "NORMALES")
        
        
        
    If Not rs.BOF And Not rs.EOF Then
    
        Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
        vExcelObj.DisplayAlerts = False
        
        vExcelObj.Workbooks.Add
        vExcelObj.Sheets("Hoja1").Select
        vExcelObj.Sheets("Hoja1").Name = lsNomHoja
        
        vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
        vExcelObj.Range("A1:IV65536").Font.Size = 8
        vExcelObj.Columns("A:IV").Select
        vExcelObj.Selection.VerticalAlignment = 3
        
        vExcelObj.Columns("G").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
        vExcelObj.Columns("I:J").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    
        vExcelObj.Columns("M:P").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
        vExcelObj.Range("A6:S6").Select
        vExcelObj.Range("A6:S6").Font.Bold = True
        vExcelObj.Range("A6:S6").HorizontalAlignment = 3
        vExcelObj.Selection.AutoFilter

        
        vExcelObj.Columns("A").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        vExcelObj.Columns("B:H").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        
        vExcelObj.Range("A1").Select
        vExcelObj.Range("A1").Font.Bold = True
        vExcelObj.Range("A1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "CAJA MUNICIPAL DE ICA"
        
        vExcelObj.Range("P1").Select
        vExcelObj.Range("P1").Font.Bold = True
        vExcelObj.Range("P1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = Format(ldFechaFinMes, "dd/mm/yyyy")
        
        vExcelObj.Range("F4").Select
        vExcelObj.Range("F4").Font.Bold = True
        vExcelObj.Range("F4").Font.Size = 12
        vExcelObj.Range("F4").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "LISTADO DE CREDITOS " & lsTipoCred & " CON CALIFICACION ANTERIOR"
        
        vExcelObj.Range("A6").Select
        vExcelObj.Range("A6").Font.Bold = True
        vExcelObj.Range("A6").ColumnWidth = 25
        vExcelObj.ActiveCell.value = "Tipo Credito"
        
        vExcelObj.Range("B6").Select
        vExcelObj.Range("B6").Font.Bold = True
        vExcelObj.Range("B6").ColumnWidth = 20
        vExcelObj.ActiveCell.value = "N° de Credito"
        
        vExcelObj.Range("C6").Select
        vExcelObj.Range("C6").Font.Bold = True
        vExcelObj.Range("C6").ColumnWidth = 40
        vExcelObj.ActiveCell.value = "Titular"
        
        vExcelObj.Range("D6").Select
        vExcelObj.Range("D6").Font.Bold = True
        vExcelObj.Range("D6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "N° Doc.Natural"
        
        vExcelObj.Range("E6").Select
        vExcelObj.Range("E6").Font.Bold = True
        vExcelObj.Range("E6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "N° Doc.Juridico"
        
        vExcelObj.Range("F6").Select
        vExcelObj.Range("F6").Font.Bold = True
        vExcelObj.Range("F6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Vigencia"
        
        vExcelObj.Range("G6").Select
        vExcelObj.Range("G6").Font.Bold = True
        vExcelObj.Range("G6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Desembolso"
        
        vExcelObj.Range("H6").Select
        vExcelObj.Range("H6").Font.Bold = True
        vExcelObj.Range("H6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Refinan."
        
        vExcelObj.Range("I6").Select
        vExcelObj.Range("I6").Font.Bold = True
        vExcelObj.Range("I6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Saldo.Cap."
        
        vExcelObj.Range("J6").Select
        vExcelObj.Range("J6").Font.Bold = True
        vExcelObj.Range("J6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Cap.Vencido"
        
        vExcelObj.Range("K6").Select
        vExcelObj.Range("K6").Font.Bold = True
        vExcelObj.Range("K6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Atraso"
        
        vExcelObj.Range("L6").Select
        vExcelObj.Range("L6").Font.Bold = True
        vExcelObj.Range("L6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Analista"
        
        vExcelObj.Range("M6").Select
        vExcelObj.Range("M6").Font.Bold = True
        vExcelObj.Range("M6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Monto.Gar.Hipo.Soles"
        
        vExcelObj.Range("N6").Select
        vExcelObj.Range("N6").Font.Bold = True
        vExcelObj.Range("N6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Monto.Gar.Hipo.Dolares"
        
        vExcelObj.Range("O6").Select
        vExcelObj.Range("O6").Font.Bold = True
        vExcelObj.Range("O6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Monto.Gar.Soles"
        
        vExcelObj.Range("P6").Select
        vExcelObj.Range("P6").Font.Bold = True
        vExcelObj.Range("P6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Monto.Gar.Dolares"
        
        vExcelObj.Range("Q6").Select
        vExcelObj.Range("Q6").Font.Bold = True
        vExcelObj.Range("Q6").ColumnWidth = 30
        vExcelObj.ActiveCell.value = "Linea Credito"
        
        vExcelObj.Range("R6").Select
        vExcelObj.Range("R6").Font.Bold = True
        vExcelObj.Range("R6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.Ant."
        
        vExcelObj.Range("S6").Select
        vExcelObj.Range("S6").Font.Bold = True
        vExcelObj.Range("S6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "RFA"
        
        vExcelObj.Range("T6").Select
        vExcelObj.Range("T6").Font.Bold = True
        vExcelObj.Range("T6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "TIPOPAGO"
        
        Dim lsCelda As String
        lnFila = 6
        Do While Not rs.EOF
            lnFila = lnFila + 1
            
            lsCelda = "A" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!Tipo
            
            lsCelda = "B" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cCtaCod
            
            lsCelda = "C" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cPersNombre
            
            lsCelda = "D" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cNudoci
            
            lsCelda = "E" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cNudoTr
            
            lsCelda = "F" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!dFecVig
            
            lsCelda = "G" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nMontoDesemb, "#,#0.00")
            
            lsCelda = "H" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cRefinan
            
            lsCelda = "I" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nSaldoCap, "#,#0.00")
            
            lsCelda = "J" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nCapVencido, "#,#0.00")
            
            lsCelda = "K" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!nDiasAtraso
            
            lsCelda = "L" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cCodAnalista
            
            lsCelda = "M" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nHipoSoles, "#,#0.00")
            
            lsCelda = "N" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nhipodol, "#,#0.00")
            
            lsCelda = "O" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarantSoles, "#,#0.00")
            
            lsCelda = "P" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarantDol, "#,#0.00")
            
            lsCelda = "Q" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!CDescripcion
            
            lsCelda = "R" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cCalAnt
            
            lsCelda = "S" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cRFA
            
            lsCelda = "T" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!CREFPUNTUAL
            
            RaiseEvent Progress(rs.Bookmark, rs.RecordCount)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
    End If
    RaiseEvent CloseProgress
    
    vExcelObj.Range("A1").Select
    vExcelObj.ActiveWorkbook.SaveAs (App.Path & "\SPOOLER\" & lsArchivo)
    vExcelObj.ActiveWorkbook.Close
    MsgBox App.Path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
    vExcelObj.Workbooks.Open (App.Path & "\SPOOLER\" & lsArchivo)
    vExcelObj.Visible = True
    
    Set vExcelObj = Nothing
    
    Set Conecta = Nothing
    
    
End Sub
Public Sub Genera_Reporte178106(ByVal lsFecha As String, ByVal lbRefinan As Boolean, Optional pbTotalCart As Boolean = False, Optional pbCartaFianza As Boolean)
Dim sql As String
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
Dim lsTipoCred As String
Dim vExcelObj  As Excel.Application

Set RCD = New nRcdReportes
Server = RCD.GetServerConsol
Set RCD = Nothing
ldFechaFinMes = GetUltimaFechaCierre

    If pbTotalCart Then
        lsCadena = ""
    Else
        If lbRefinan = True Then
            lsCadena = "  AND CC.cRefinan ='R' "
        Else
            lsCadena = "  AND CC.cRefinan ='N' "
        End If
    End If
    
        sql = "       Select  C.cConsDescripcion AS TIPO, CC.cCtaCod, P.cPersNombre,  "
        sql = sql & "       ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=1) ,'') AS cNudoci,"
        sql = sql & "       ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=2) ,'') AS cNudotr,"
        sql = sql & "       convert(varchar(10), CC.DFECVIG, 103) as dFecVig , CC.nMontoApr as nMontoDesemb, CC.cRefinan, CC.nSaldoCap, CC.nDiasAtraso, CC.cCodAnalista,"
        sql = sql & "       L.cDescripcion,"
        sql = sql & "       CC.nGarPref, CC.nGarMuyRR, CC.nGarAutoL, "
        sql = sql & "       ISNULL((SELECT TOP 1 CT.cCalGen FROM " & Server & "ColocCalifProvTotal CT WHERE CT.CCTACOD = CC.CCTACOD AND convert(char(10),CT.dFecha,112)<'" & Format(lsFecha, "yyyymmdd") & "' ORDER BY CT.dFecha DESC) ,'')  as cCalAnt, "
        sql = sql & "       ISNULL((SELECT TOP 1 CT1.nProvision FROM " & Server & "ColocCalifProvTotal CT1 WHERE CT1.CCTACOD = CC.CCTACOD AND convert(char(10),CT1.dFecha,112)<'" & Format(lsFecha, "yyyymmdd") & "' ORDER BY CT1.dFecha DESC) ,0)  as nProvAnt, "
        sql = sql & "       isnull(CC.cCalGen ,'') as cCalAct, CC.nProvision AS nProvAct, ISNULL(CC.cCalRFA,'') AS cCalRFA,"
        sql = sql & "       ISNULL(CC.cCalCMAC,'') AS cCalCMAC, ISNULL(CC.cCalSistF,'') AS cCalSistF"
        sql = sql & " From  ColocCalifProv CC"
        sql = sql & "       JOIN DBConsolidada..ProductoPersonaConsol RC ON RC.CCTACOD = CC.CCTACOD"
        sql = sql & "       JOIN PERSONA P ON P.CPERSCOD = RC.cPersCod  and nPrdPersRelac= 20"
        sql = sql & "       LEFT JOIN ColocLineaCredito L on L.cLineaCred = CC.cLineaCred"
        sql = sql & "       JOIN Constante c on c.nConsValor = SUBSTRING(CC.cCtaCod,6,3) AND nConsCod= 1001"
        sql = sql & " WHERE CC.nPrdEstado in (2020,2022,2021,2030,2032,2031,2100,2104,2106,2107,2201,2050,2101,2104,2106,2107) " & IIf(pbCartaFianza, " and Substring(CC.cCtaCod,7,2)=21 ", lsCadena)
        sql = sql & "       ORDER BY P.cPersNombre, CC.cCtaCod"

        Conecta.AbreConexion
        Set rs = Conecta.CargaRecordSet(sql)
        Conecta.CierraConexion
        If pbTotalCart Then
            lsArchivo = "TotalCartera_" & Format(lsFecha, "mmyyyy") & ".XLS"
        Else
            lsArchivo = "CreditosxTipoDespCal" & Format(lsFecha, "mmyyyy") & ".XLS"
        End If
        If pbCartaFianza Then
            lsArchivo = "CreditosCartaFianza" & Format(lsFecha, "mmyyyy") & ".XLS"
        End If
        
        If pbTotalCart Then
            lsNomHoja = "TOTAL_CARTERA"
            lsTipoCred = ""
        Else
            lsNomHoja = IIf(lbRefinan = True, "REFINANCIADOS", "NORMALES")
            lsTipoCred = IIf(lbRefinan = True, "REFINANCIADOS", "NORMALES")
        End If
        If pbCartaFianza Then
            lsNomHoja = "CARTAFIANZA"
            lsTipoCred = "CARTAFIANZA"
        End If
        
    If Not rs.BOF And Not rs.EOF Then
        RaiseEvent ShowProgress
        
        Set vExcelObj = New Excel.Application
        vExcelObj.DisplayAlerts = False
        
        vExcelObj.Workbooks.Add
        vExcelObj.Sheets("Hoja1").Select
        vExcelObj.Sheets("Hoja1").Name = lsNomHoja
        
        vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
        vExcelObj.Range("A1:IV65536").Font.Size = 8
        vExcelObj.Columns("A:IV").Select
        vExcelObj.Selection.VerticalAlignment = 3
        
        vExcelObj.Columns("G").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
        vExcelObj.Columns("I").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    
        vExcelObj.Columns("M:P").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
        vExcelObj.Range("A6:W6").Select
        vExcelObj.Range("A6:W6").Font.Bold = True
        vExcelObj.Range("A6:W6").HorizontalAlignment = 3
        vExcelObj.Selection.AutoFilter

        
        vExcelObj.Columns("A").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        vExcelObj.Columns("B:H").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        
        vExcelObj.Range("A1").Select
        vExcelObj.Range("A1").Font.Bold = True
        vExcelObj.Range("A1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "CAJA MUNICIPAL DE ICA"
        
        vExcelObj.Range("P1").Select
        vExcelObj.Range("P1").Font.Bold = True
        vExcelObj.Range("P1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = Format(ldFechaFinMes, "dd/mm/yyyy")
        
        vExcelObj.Range("F4").Select
        vExcelObj.Range("F4").Font.Bold = True
        vExcelObj.Range("F4").Font.Size = 12
        vExcelObj.Range("F4").HorizontalAlignment = 1
        If pbCartaFianza = False Then
            vExcelObj.ActiveCell.value = IIf(pbTotalCart = False, "LISTADO DE CREDITOS " & lsTipoCred & " DESPUES DE CALIFICACION", "LISTADO DE CREDITOS EVALUADOS AL :" & ldFechaFinMes)
        Else
            vExcelObj.ActiveCell.value = "LISTADO DE CARTAS FIANZAS EVALUADOS AL :" & ldFechaFinMes
        End If
        
        vExcelObj.Range("A6").Select
        vExcelObj.Range("A6").Font.Bold = True
        vExcelObj.Range("A6").ColumnWidth = 25
        vExcelObj.ActiveCell.value = "Tipo Credito"
        
        vExcelObj.Range("B6").Select
        vExcelObj.Range("B6").Font.Bold = True
        vExcelObj.Range("B6").ColumnWidth = 20
        vExcelObj.ActiveCell.value = "N° de Credito"
        
        vExcelObj.Range("C6").Select
        vExcelObj.Range("C6").Font.Bold = True
        vExcelObj.Range("C6").ColumnWidth = 40
        vExcelObj.ActiveCell.value = "Titular"
        
        vExcelObj.Range("D6").Select
        vExcelObj.Range("D6").Font.Bold = True
        vExcelObj.Range("D6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "N° Doc.Natural"
        
        vExcelObj.Range("E6").Select
        vExcelObj.Range("E6").Font.Bold = True
        vExcelObj.Range("E6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "N° Doc.Juridico"
        
        vExcelObj.Range("F6").Select
        vExcelObj.Range("F6").Font.Bold = True
        vExcelObj.Range("F6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Vigencia"
        
        vExcelObj.Range("G6").Select
        vExcelObj.Range("G6").Font.Bold = True
        vExcelObj.Range("G6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Desembolso"
        
        vExcelObj.Range("H6").Select
        vExcelObj.Range("H6").Font.Bold = True
        vExcelObj.Range("H6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Refinan."
        
        vExcelObj.Range("I6").Select
        vExcelObj.Range("I6").Font.Bold = True
        vExcelObj.Range("I6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Saldo.Cap."
        
        vExcelObj.Range("J6").Select
        vExcelObj.Range("J6").Font.Bold = True
        vExcelObj.Range("J6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Moneda"
        
        vExcelObj.Range("K6").Select
        vExcelObj.Range("K6").Font.Bold = True
        vExcelObj.Range("K6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Atraso"
        
        vExcelObj.Range("L6").Select
        vExcelObj.Range("L6").Font.Bold = True
        vExcelObj.Range("L6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Analista"
        
        vExcelObj.Range("M6").Select
        vExcelObj.Range("M6").Font.Bold = True
        vExcelObj.Range("M6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Linea Credito"
        
        vExcelObj.Range("N6").Select
        vExcelObj.Range("N6").Font.Bold = True
        vExcelObj.Range("N6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "GarPref"
        
        vExcelObj.Range("O6").Select
        vExcelObj.Range("O6").Font.Bold = True
        vExcelObj.Range("O6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "GarMuyRR"
        
        vExcelObj.Range("P6").Select
        vExcelObj.Range("P6").Font.Bold = True
        vExcelObj.Range("P6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "nGarAutoL"
        
        vExcelObj.Range("Q6").Select
        vExcelObj.Range("Q6").Font.Bold = True
        vExcelObj.Range("Q6").ColumnWidth = 30
        vExcelObj.ActiveCell.value = "Cal.Ant."
        
        vExcelObj.Range("R6").Select
        vExcelObj.Range("R6").Font.Bold = True
        vExcelObj.Range("R6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Prov.Ant."
        
        vExcelObj.Range("S6").Select
        vExcelObj.Range("S6").Font.Bold = True
        vExcelObj.Range("S6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.Act."
        
        vExcelObj.Range("T6").Select
        vExcelObj.Range("T6").Font.Bold = True
        vExcelObj.Range("T6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Prov.Act."
        
        vExcelObj.Range("U6").Select
        vExcelObj.Range("U6").Font.Bold = True
        vExcelObj.Range("U6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.RFA."
        
        vExcelObj.Range("V6").Select
        vExcelObj.Range("V6").Font.Bold = True
        vExcelObj.Range("V6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.CMAC."
        
        vExcelObj.Range("W6").Select
        vExcelObj.Range("W6").Font.Bold = True
        vExcelObj.Range("W6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.SIST.FINAN"
        
        Dim lsCelda As String
        lnFila = 6
        Do While Not rs.EOF
            lnFila = lnFila + 1
            
            lsCelda = "A" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!Tipo
            
            lsCelda = "B" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cCtaCod
            
            lsCelda = "C" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cPersNombre
            
            lsCelda = "D" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cNudoci
            
            lsCelda = "E" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cNudoTr
            
            lsCelda = "F" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!dFecVig
            
            lsCelda = "G" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nMontoDesemb, "#,#0.00")
            
            lsCelda = "H" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cRefinan
            
            lsCelda = "I" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nSaldoCap, "#,#0.00")
            
            lsCelda = "J" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Mid(rs!cCtaCod, 9, 1)
            
            lsCelda = "K" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!nDiasAtraso
            
            lsCelda = "L" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cCodAnalista
            
            lsCelda = "M" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!CDescripcion
            
            lsCelda = "N" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarPref, "#,#0.00")
            
            lsCelda = "O" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarMuyRR, "#,#0.00")
            
            lsCelda = "P" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarAutoL, "#,#0.00")
            
            lsCelda = "Q" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalAnt
            
            lsCelda = "R" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nProvAnt, "#,#0.00")
            
            lsCelda = "S" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalAct
            
            lsCelda = "T" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nProvAct, "#,#0.00")
            
            lsCelda = "U" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalRFA
            
            lsCelda = "V" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalCMAC
            
            lsCelda = "W" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalSistF
            
            RaiseEvent Progress(rs.Bookmark, rs.RecordCount)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        RaiseEvent CloseProgress
        
        vExcelObj.Range("A1").Select
        vExcelObj.ActiveWorkbook.SaveAs (App.Path & "\SPOOLER\" & lsArchivo)
        vExcelObj.ActiveWorkbook.Close
        MsgBox App.Path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
        vExcelObj.Workbooks.Open (App.Path & "\SPOOLER\" & lsArchivo)
        vExcelObj.Visible = True
        
    End If
    
    Set vExcelObj = Nothing
    Set Conecta = Nothing
    
    
End Sub

Public Sub Genera_Reporte178105(ByVal lsFecha As String, ByVal nChkMicro As Byte, _
                            ByVal nChkComer As Byte, ByVal nChkConsu As Byte, ByVal nChkHipo As Byte)
Dim sql As String
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
Dim lsTipoCred As String
Dim vExcelObj  As Excel.Application

Set RCD = New nRcdReportes
Server = RCD.GetServerConsol
Set RCD = Nothing
ldFechaFinMes = GetUltimaFechaCierre

        lsCadena = "  AND SUBSTRING(CC.cCtaCod,6,1) IN ("
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
    lsCadena = lsCadena & ")"
    If nChkMicro = 0 And nChkComer = 0 And nChkConsu = 0 And nChkHipo = 0 Then
        lsCadena = ""
    End If
        
        sql = "       Select  C.cConsDescripcion AS TIPO, CC.cCtaCod, P.cPersNombre,  "
        sql = sql & "       ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=1) ,'') AS cNudoci,"
        sql = sql & "       ISNULL((Select A.cPersIdNro from PersID A where A.cPerscod=P.cPersCod and A.cPersIDTpo=2) ,'') AS cNudotr,"
        sql = sql & "       convert(varchar(10), CC.DFECVIG, 103) as dFecVig , CC.nMontoApr as nMontoDesemb, CC.cRefinan, CC.nSaldoCap, CC.nDiasAtraso, CC.cCodAnalista,"
        sql = sql & "       L.cDescripcion,"
        sql = sql & "       CC.nGarPref, CC.nGarMuyRR, CC.nGarAutoL, isnull(CT.cCalGen ,'') as cCalAnt, ISNULL(CT.nProvision,0) AS nProvAnt,"
        sql = sql & "       isnull(CC.cCalGen ,'') as cCalAct, CC.nProvision AS nProvAct, ISNULL(CC.cCalRFA,'') AS cCalRFA,"
        sql = sql & "       ISNULL(CC.cCalCMAC,'') AS cCalCMAC, ISNULL(CC.cCalSistF,'') AS cCalSistF"
        sql = sql & " From  ColocCalifProv CC"
        sql = sql & "       LEFT JOIN DBConsolidada..ColocCalifProvTotal CT ON CT.CCTACOD = CC.CCTACOD AND CT.dFecha = (SELECT MAX(dFecha) FROM DBConsolidada..ColocCalifProvTotal CT1 WHERE  CT1.dFecha<'" & Format(lsFecha, "mm/dd/yyyy") & "')"
        sql = sql & "       JOIN DBConsolidada..ProductoPersonaConsol RC ON RC.CCTACOD = CC.CCTACOD"
        sql = sql & "       JOIN PERSONA P ON P.CPERSCOD = RC.cPersCod  and nPrdPersRelac= 20"
        sql = sql & "       LEFT JOIN ColocLineaCredito L on L.cLineaCred = CC.cLineaCred"
        sql = sql & "       JOIN Constante c on c.nConsValor = SUBSTRING(CC.cCtaCod,6,3) AND nConsCod= 1001"
        sql = sql & " WHERE CC.nPrdEstado in (2020,2022,2021,2030,2032,2031,2100,2104,2106,2107,2201,2050,2101,2104,2106,2107) " & lsCadena
        sql = sql & "       ORDER BY P.cPersNombre, CC.cCtaCod"

        Conecta.AbreConexion
        Set rs = Conecta.CargaRecordSet(sql)
        Conecta.CierraConexion
        
        lsArchivo = "ListaCredEvalDespCal_" & Format(lsFecha, "mmyyyy") & ".XLS"
        
        lsNomHoja = Format(lsFecha, "ddmmyyyy") & "-" _
                    & IIf(nChkMicro = 1, "Micro" & "-", "") _
                    & IIf(nChkComer = 1, "Comer" & "-", "") _
                    & IIf(nChkConsu = 1, "Cons" & "-", "") & IIf(nChkHipo = 1, "Hipo", "")

        lsTipoCred = IIf(nChkMicro = 1, "MicroEmpresa" & "-", "") _
                    & IIf(nChkComer = 1, "Comercial" & "-", "") _
                    & IIf(nChkConsu = 1, "Consumo" & "-", "") & IIf(nChkHipo = 1, "Hipotecario", "")
        
    If Not rs.BOF And Not rs.EOF Then
        RaiseEvent ShowProgress
        
        Set vExcelObj = New Excel.Application
        vExcelObj.DisplayAlerts = False
        
        vExcelObj.Workbooks.Add
        vExcelObj.Sheets("Hoja1").Select
        vExcelObj.Sheets("Hoja1").Name = lsNomHoja
        
        vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
        vExcelObj.Range("A1:IV65536").Font.Size = 8
        vExcelObj.Columns("A:IV").Select
        vExcelObj.Selection.VerticalAlignment = 3
        
        vExcelObj.Columns("G").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
        vExcelObj.Columns("I").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
    
        vExcelObj.Columns("M:P").Select
        vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
        vExcelObj.Range("A6:W6").Select
        vExcelObj.Range("A6:W6").Font.Bold = True
        vExcelObj.Range("A6:W6").HorizontalAlignment = 3
        vExcelObj.Selection.AutoFilter
        
        vExcelObj.Columns("A").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        vExcelObj.Columns("B:H").Select
        vExcelObj.Selection.HorizontalAlignment = 1
        
        vExcelObj.Range("A1").Select
        vExcelObj.Range("A1").Font.Bold = True
        vExcelObj.Range("A1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "CAJA MUNICIPAL DE ICA"
        
        vExcelObj.Range("P1").Select
        vExcelObj.Range("P1").Font.Bold = True
        vExcelObj.Range("P1").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = Format(ldFechaFinMes, "dd/mm/yyyy")
        
        vExcelObj.Range("F4").Select
        vExcelObj.Range("F4").Font.Bold = True
        vExcelObj.Range("F4").Font.Size = 12
        vExcelObj.Range("F4").HorizontalAlignment = 1
        vExcelObj.ActiveCell.value = "LISTADO DE CREDITOS " & lsTipoCred & " DESPUES DE CALIFICACION"
        
        vExcelObj.Range("A6").Select
        vExcelObj.Range("A6").Font.Bold = True
        vExcelObj.Range("A6").ColumnWidth = 25
        vExcelObj.ActiveCell.value = "Tipo Credito"
        
        vExcelObj.Range("B6").Select
        vExcelObj.Range("B6").Font.Bold = True
        vExcelObj.Range("B6").ColumnWidth = 20
        vExcelObj.ActiveCell.value = "N° de Credito"
        
        vExcelObj.Range("C6").Select
        vExcelObj.Range("C6").Font.Bold = True
        vExcelObj.Range("C6").ColumnWidth = 40
        vExcelObj.ActiveCell.value = "Titular"
        
        vExcelObj.Range("D6").Select
        vExcelObj.Range("D6").Font.Bold = True
        vExcelObj.Range("D6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "N° Doc.Natural"
        
        vExcelObj.Range("E6").Select
        vExcelObj.Range("E6").Font.Bold = True
        vExcelObj.Range("E6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "N° Doc.Juridico"
        
        vExcelObj.Range("F6").Select
        vExcelObj.Range("F6").Font.Bold = True
        vExcelObj.Range("F6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Vigencia"
        
        vExcelObj.Range("G6").Select
        vExcelObj.Range("G6").Font.Bold = True
        vExcelObj.Range("G6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Desembolso"
        
        vExcelObj.Range("H6").Select
        vExcelObj.Range("H6").Font.Bold = True
        vExcelObj.Range("H6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Refinan."
        
        vExcelObj.Range("I6").Select
        vExcelObj.Range("I6").Font.Bold = True
        vExcelObj.Range("I6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Saldo.Cap."
        
        vExcelObj.Range("J6").Select
        vExcelObj.Range("J6").Font.Bold = True
        vExcelObj.Range("J6").ColumnWidth = 12
        vExcelObj.ActiveCell.value = "Moneda"
        
        vExcelObj.Range("K6").Select
        vExcelObj.Range("K6").Font.Bold = True
        vExcelObj.Range("K6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Atraso"
        
        vExcelObj.Range("L6").Select
        vExcelObj.Range("L6").Font.Bold = True
        vExcelObj.Range("L6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Analista"
        
        vExcelObj.Range("M6").Select
        vExcelObj.Range("M6").Font.Bold = True
        vExcelObj.Range("M6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "Linea Credito"
        
        vExcelObj.Range("N6").Select
        vExcelObj.Range("N6").Font.Bold = True
        vExcelObj.Range("N6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "GarPref"
        
        vExcelObj.Range("O6").Select
        vExcelObj.Range("O6").Font.Bold = True
        vExcelObj.Range("O6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "GarMuyRR"
        
        vExcelObj.Range("P6").Select
        vExcelObj.Range("P6").Font.Bold = True
        vExcelObj.Range("P6").ColumnWidth = 15
        vExcelObj.ActiveCell.value = "nGarAutoL"
        
        vExcelObj.Range("Q6").Select
        vExcelObj.Range("Q6").Font.Bold = True
        vExcelObj.Range("Q6").ColumnWidth = 30
        vExcelObj.ActiveCell.value = "Cal.Ant."
        
        vExcelObj.Range("R6").Select
        vExcelObj.Range("R6").Font.Bold = True
        vExcelObj.Range("R6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Prov.Ant."
        
        vExcelObj.Range("S6").Select
        vExcelObj.Range("S6").Font.Bold = True
        vExcelObj.Range("S6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.Act."
        
        vExcelObj.Range("T6").Select
        vExcelObj.Range("T6").Font.Bold = True
        vExcelObj.Range("T6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Prov.Act."
        
        vExcelObj.Range("U6").Select
        vExcelObj.Range("U6").Font.Bold = True
        vExcelObj.Range("U6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.RFA."
        
        vExcelObj.Range("V6").Select
        vExcelObj.Range("V6").Font.Bold = True
        vExcelObj.Range("V6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.CMAC."
        
        vExcelObj.Range("W6").Select
        vExcelObj.Range("W6").Font.Bold = True
        vExcelObj.Range("W6").ColumnWidth = 10
        vExcelObj.ActiveCell.value = "Cal.SIST.FINAN"
        
        Dim lsCelda As String
        lnFila = 6
        Do While Not rs.EOF
            lnFila = lnFila + 1
            
            lsCelda = "A" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!Tipo
            
            lsCelda = "B" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cCtaCod
            
            lsCelda = "C" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cPersNombre
            
            lsCelda = "D" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cNudoci
            
            lsCelda = "E" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = "'" & rs!cNudoTr
            
            lsCelda = "F" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!dFecVig
            
            lsCelda = "G" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nMontoDesemb, "#,#0.00")
            
            lsCelda = "H" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cRefinan
            
            lsCelda = "I" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nSaldoCap, "#,#0.00")
            
            lsCelda = "J" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Mid(rs!cCtaCod, 9, 1)
            
            lsCelda = "K" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!nDiasAtraso
            
            lsCelda = "L" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!cCodAnalista
            
            lsCelda = "M" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = rs!CDescripcion
            
            lsCelda = "N" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarPref, "#,#0.00")
            
            lsCelda = "O" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarMuyRR, "#,#0.00")
            
            lsCelda = "P" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nGarAutoL, "#,#0.00")
            
            lsCelda = "Q" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalAnt
            
            lsCelda = "R" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nProvAnt, "#,#0.00")
            
            lsCelda = "S" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalAct
            
            lsCelda = "T" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.ActiveCell.value = Format(rs!nProvAct, "#,#0.00")
            
            lsCelda = "U" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalRFA
            
            lsCelda = "V" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalCMAC
            
            lsCelda = "W" & Trim(Str(lnFila))
            vExcelObj.Range(lsCelda).Select
            vExcelObj.Range(lsCelda).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = rs!cCalSistF
            
            RaiseEvent Progress(rs.Bookmark, rs.RecordCount)
            rs.MoveNext
        Loop
        rs.Close
        Set rs = Nothing
        
        RaiseEvent CloseProgress
        
        vExcelObj.Range("A1").Select
        vExcelObj.ActiveWorkbook.SaveAs (App.Path & "\SPOOLER\" & lsArchivo)
        vExcelObj.ActiveWorkbook.Close
        MsgBox App.Path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
        vExcelObj.Workbooks.Open (App.Path & "\SPOOLER\" & lsArchivo)
        vExcelObj.Visible = True
        
        
    End If
    
    Set vExcelObj = Nothing
    Set Conecta = Nothing
    
    
End Sub

Public Sub Genera_Reporte178103(ByVal lsFecha As String, ByVal nChkMicro As Byte, _
ByVal nChkComer As Byte, ByVal nChkConsu As Byte, ByVal psCodCmat As String, ByVal nChkHipo As Byte)
End Sub
Public Function Genera_Reporte178201(ByVal psServer As String, ByVal sMoneda As String, ByVal sSigno As String, ByVal nMonto As Currency, _
  ByVal bTipo As Boolean, ByVal bEstado As Boolean, ByVal sCodCMAC As String, ByVal psNomCmac As String, ByVal psNomAge As String) As String

Dim sql As String
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

    'Deficiente
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

    For i = 1 To RSAgencias.RecordCount
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

Dim sql As String
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
    'lsCredito = "Comercial-"
Else
    nTipo = "1" 'Micro
    'lsCredito = "MicroEmpresa-"
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

    For i = 1 To RSAgencias.RecordCount
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
  ByVal bTipo As Boolean, ByVal bEstado As Boolean, ByVal psCodCMAC As String, ByVal psNomCmac As String, _
  ByVal psNomAge As String) As String

Dim sql As String
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim Rs3 As New ADODB.Recordset
Dim Rs4 As New ADODB.Recordset

Dim RSAgencias As New ADODB.Recordset
Dim RSRecursos As New ADODB.Recordset

Dim loPrevio As previo.clsPrevio

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
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), , , psCodCMAC)
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
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "1", psCodCMAC)
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
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "2", psCodCMAC)
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
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "3", psCodCMAC)
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
    Set rs = ReporteGeneralProvisionesHipotecario(nTipo, nEstado, sMoneda, sSigno, nMonto, RSRecursos(0), False, "4", psCodCMAC)
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

    For i = 1 To RSAgencias.RecordCount
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
Dim sql As String
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
    For i = 1 To RSAgencias.RecordCount
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
        lsCadena = lsCadena & Space(2) & ImpreFormat(Total_Prov, 10, 2, True) & Chr(10)
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


    xlHoja1.Range("A1").ColumnWidth = 22
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
    xlHoja1.PageSetup.Zoom = 70
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlAplicacion.Range("A1:Z300").Font.Size = 9

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
Dim sql As String

ImpHojaWProvision = True
lnLinea = 1
FechaFinMes = GetUltimaFechaCierre
lbCambio = False
lsArchivo = "Anexo5_" & Format(FechaFinMes, "mmyyyy") & "HojaW_Provision.XLS"
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & lsArchivo)
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
                sql = " GROUP BY  Substring(CA.cCtaCod,6,3), " _
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
                lsSQL = lsSQL & sql

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
                    lsSubTipo = rs!cSubtipo
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
MsgBox "Se Genero el archivo " & App.Path & "\SPOOLER\" & lsArchivo

'*******
xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo
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
Dim sql As String
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

'On Error Resume Next
sql = "select * from ConstSistema where nConsSisCod=3"
Set rs = RCD.GetQuery(sql)
lsCodigo = IIf(IsNull(rs!nConsSisValor), "", rs!nConsSisValor)

lsArchivo = "ReporteAnexo5C_" & Format(pdFecDataFin, "YYYYMMDD") & ".xls"
Set fs = New Scripting.FileSystemObject

Set xlAplicacion = New Excel.Application
If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & lsArchivo)
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
    sql = " Select T.nConsValor, isnull(sum(C.SaldoCap),0) Total"
    sql = sql & " From"
    sql = sql & " ("
    sql = sql & " Select case when substring(cCtaCod,9,1)='1' then nSaldoCap "
    sql = sql & "             else nSaldoCap * " & TxtCambio & " end SaldoCap, "
    sql = sql & "        cCalGen from ColocCalifProv "
    sql = sql & " where  " & lsEstadoProd(i) & "  and substring(cCtaCod,6,3) = '402'"
    sql = sql & " ) C"
    sql = sql & " Right Outer Join"
    sql = sql & " ( Select nConsValor, cConsDescripcion"
    sql = sql & "   From Constante"
    sql = sql & "   Where nConsCod = 3010 And nConsValor <> 3010"
    sql = sql & " ) T on T.nConsValor = C.cCalGen"
    sql = sql & " group by T.nConsValor"
    Set rs = RCD.GetQuery(sql)
    While Not rs.EOF
        xlHoja1.Cells(k, J) = Format(rs!Total, "0.00")
        rs.MoveNext
        J = J + 1
    Wend
    k = k + 1
Next i

' Calculando Provisiones
sql = "  Select T.nConsValor, isnull(sum(C.nProvision),0) Total"
sql = sql & " From"
sql = sql & " ("
sql = sql & "  Select case when substring(cCtaCod,9,1)='1' then nProvision"
sql = sql & "              else nProvision * " & TxtCambio & "  end nProvision,"
sql = sql & "         cCalGen from ColocCalifProv "
sql = sql & "  where  nPrdEstado in (2020,2021,2022,2030,2031,2032,2201) and substring(cCtaCod,6,3) = '402'"
sql = sql & " ) C"
sql = sql & " Right Outer Join"
sql = sql & " ( Select nConsValor, cConsDescripcion"
sql = sql & "   From Constante"
sql = sql & "   Where nConsCod = 3010 And nConsValor <> 3010"
sql = sql & " ) T on T.nConsValor = C.cCalGen"
sql = sql & " group by T.nConsValor"
sql = sql & " order by T.nConsValor"
Set rs = RCD.GetQuery(sql)
J = 2
k = 12
While Not rs.EOF
    xlHoja1.Cells(19, J) = Format(rs!Total, "0.00")
    xlHoja1.Cells(20, J) = Format(rs!Total, "0.00")
    rs.MoveNext
    J = 1 + J
Wend




' Calculando las Garantias
sql = " Select T.nConsValor, isnull(sum(C.SumGar),0) Total"
sql = sql & " From"
sql = sql & " ("
sql = sql & " Select  c.cCalGen,"
sql = sql & "     isnull(case when substring(CG.cCtaCod,9,1)='1' then  G.nRealizacion"
sql = sql & "         else round( G.nRealizacion  * " & TxtCambio & " ,2) end,0) SumGar"
sql = sql & " from ColocCalifProv  C"
sql = sql & " Inner Join ColocGarantia CG on C.cCtaCod = CG.cCtaCod"
sql = sql & " Inner Join Garantias G on G.cNumgarant = CG.cNumGarant"
sql = sql & " where  C.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201) and substring(C.cCtaCod,6,3) = '402'"
sql = sql & " ) C"
sql = sql & " Right Outer Join"
sql = sql & " ( Select nConsValor, cConsDescripcion"
sql = sql & "   From Constante"
sql = sql & "   Where nConsCod = 3010 And nConsValor <> 3010"
sql = sql & " ) T on T.nConsValor = C.cCalGen"
sql = sql & " group by T.nConsValor"
Set rs = RCD.GetQuery(sql)
J = 2
While Not rs.EOF
    xlHoja1.Cells(22, J) = Format(rs!Total, "0.00")
    J = 1 + J
    rs.MoveNext
Wend
J = 12
For i = 1 To 4
    sql = " Select  isnull(sum(case when substring(cCtaCod,9,1)='1' then nProvision"
    sql = sql & "          else nProvision * " & TxtCambio & "  end),0) Provision"
    sql = sql & "      From ColocCalifProv"
    sql = sql & " where  " & lsEstadoProd(i) & " and substring(cCtaCod,6,3) = '402'"
    Set rs = RCD.GetQuery(sql)
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




xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo
MsgBox "Se genero el archivo " & App.Path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
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
                ByVal dFecha As Date, Optional ByVal psNomCmac As String, Optional ByVal psCodCMAC As String)

Dim sql As String
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
If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & lsArchivo)
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
ImpAnexo5Cab pnFila, "1", sFuente, dFecha, , , psNomCmac, psCodCMAC, "5"
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

                sql = "SELECT Count ( distinct (cPersCod) ) as Deudores, "
                sql = sql & "           Count (CA.cCtaCod ) AS NroCred, "
                sql = sql & "           ISNULL(Sum (CASE WHEN CA.cCtaCod like  '________1%' THEN (CA.nSaldoCap) "
                sql = sql & "                            WHEN CA.cCtaCod like  '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
                sql = sql & "                   END ), 0 ) AS SaldoCap, "
                sql = sql & "           ISNULL (Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
                sql = sql & "                             WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
                sql = sql & "                   END ), 0 ) AS Provision, "
                sql = sql & "           ISNULL(SUM (ISNULL(G.nMontoGar,0)),0) AS GarantHip   "
                sql = sql & " FROM  ColocCalifProv CA LEFT JOIN " & ServerCons & "CreditoConsol C ON CA.cCtaCod = C.cCtaCod"
                sql = sql & "  LEFT JOIN ( "
                sql = sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                sql = sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                sql = sql & "       Where G.nGarClase=1 AND G.nGarTpoRealiz=1  "
                sql = sql & "        Group by GC.cCtaCod  "
                sql = sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
                sql = sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "' "
                sql = sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan


                '-------------------------------------------

                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(sql)
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

               sql = "  SELECT Count ( distinct (CA.cPersCod) ) as Deudores, "
                sql = sql & " Count (CA.cCtaCod ) AS NroCred, "
                sql = sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nSaldoCap) "
                sql = sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
                sql = sql & "               END ), 0 ) AS SaldoCap, "
                sql = sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
                sql = sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
                sql = sql & "               END ), 0 ) AS Provision,"
                sql = sql & "  GarantPrefRapRea = SUM(ISNULL(G.nMontoGar,0))  "
                sql = sql & " FROM ColocCalifProv CA"
                sql = sql & "  LEFT JOIN ( "
                sql = sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                sql = sql & "        "
                sql = sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                sql = sql & "        Where G.nGarClase=1 AND G.nGarTpoRealiz=1 "
                sql = sql & "        Group by GC.cCtaCod  "
                sql = sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
                sql = sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "' "
                sql = sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan


                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(sql)
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

                sql = "  SELECT Count ( distinct (CA.cPersCod) ) as Deudores, "
                sql = sql & " Count (CA.cCtaCod ) AS NroCred, "
                sql = sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nSaldoCap) "
                sql = sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
                sql = sql & "               END ), 0 ) AS SaldoCap, "
                sql = sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
                sql = sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
                sql = sql & "               END ), 0 ) AS Provision,"
                sql = sql & "  GarantPrefRapRea = ISNULL(SUM(ISNULL(GR.nMontoGar,0)),0)  "
                sql = sql & " FROM ColocCalifProv CA"
                sql = sql & "  INNER JOIN ( "
                sql = sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                sql = sql & "        "
                sql = sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                sql = sql & "        Where G.nGarClase=1 "
                sql = sql & "        Group by GC.cCtaCod  "
                sql = sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
                sql = sql & "  LEFT JOIN ( "
                sql = sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                sql = sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                sql = sql & "        Where G.nGarClase=1 AND  G.nGarTpoRealiz=1 "
                sql = sql & "        Group by GC.cCtaCod  "
                sql = sql & "       ) GR ON GR.cCtaCod = CA.cCtaCod "
                sql = sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "' "
                sql = sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan

            '------------------------------------
                'rsP.CursorLocation = adUseClient
                'rsP.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                'Set rsP.ActiveConnection = Nothing
                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(sql)
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

                sql = "  SELECT Count ( distinct (CA.cPersCod) ) as Deudores, "
                sql = sql & " Count (CA.cCtaCod ) AS NroCred, "
                sql = sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nSaldoCap) "
                sql = sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
                sql = sql & "               END ), 0 ) AS SaldoCap, "
                sql = sql & "  ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
                sql = sql & "                   WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
                sql = sql & "               END ), 0 ) AS Provision,"
                sql = sql & "  GarantPrefRapRea = ISNULL(SUM(ISNULL(G.nMontoGar,0)),0)  "
                sql = sql & " FROM ColocCalifProv CA"
                sql = sql & "  INNER JOIN ( "
                sql = sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                sql = sql & "       "
                sql = sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                sql = sql & "        Where G.nGarClase = 1 AND G.nGarTpoRealiz = 1 "
                sql = sql & "        Group by GC.cCtaCod  "
                sql = sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
                sql = sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "' "
                sql = sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan

            '------------------------------------
                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(sql)
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
ImpAnexo5Cab pnFila, "2", sFuente, dFecha, , , psNomCmac, psCodCMAC, "5"

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

            sql = "SELECT Count (Distinct (CA.cPersCod) ) as Deudores, "
            sql = sql & "           Count (CA.cCtaCod ) AS NroCred, "
            sql = sql & "           ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nSaldoCap) "
            sql = sql & "                            WHEN CA.cCtaCod like '________2%' THEN (CA.nSaldoCap * " & TxtCambio & ") "
            sql = sql & "                   END ), 0 ) AS SaldoCap, "
            sql = sql & "           ISNULL(Sum (CASE WHEN CA.cCtaCod like '________1%' THEN (CA.nProvision) "
            sql = sql & "                            WHEN CA.cCtaCod like '________2%' THEN (CA.nProvision * " & TxtCambio & ") "
            sql = sql & "                   END ), 0 ) AS Provision, "
            sql = sql & "  GarantPrefRapRea = ISNULL(SUM(ISNULL(G.nMontoGar,0)),0)  "
            sql = sql & " FROM ColocCalifProv CA "
             sql = sql & "  LEFT JOIN ( "
                sql = sql & "       Select GC.cCtaCod, SUM((CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END) ) as nMontoGar From " & ServerCons & "GarantiasConsol G "
                sql = sql & "       "
                sql = sql & "        Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cNumGarant = G.cNumGarant  "
                sql = sql & "        Where G.nGarClase = 1 AND G.nGarTpoRealiz = 1 "
                sql = sql & "        Group by GC.cCtaCod  "
                sql = sql & "       ) G ON G.cCtaCod = CA.cCtaCod "
            sql = sql & " WHERE CA.nPrdEstado IN ('2201','2020','2021','2022','2030','2031','2032','2802','2803','2804','2807') and Substring(CA.cCtaCod,6,1) = '" & Trim(Str(liProd)) & "'"
            sql = sql & " AND CA.cCalGen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan

            'rsP.CursorLocation = adUseClient
            'rsP.Open SQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
            'Set rsP.ActiveConnection = Nothing
                Conecta.AbreConexion
                Set rsP = Conecta.CargaRecordSet(sql)
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
ImpAnexo5Cab pnFila, "3", sFuente, dFecha, , , psNomCmac, psCodCMAC, "5"
pnFila = pnFila + 2

lsTotGeneral(1) = "": lsTotGeneral(2) = "": lsTotGeneral(3) = "": lsTotGeneral(4) = "": lsTotGeneral(5) = "": lsTotGeneral(6) = ""


nTotGarPrefReal = 0

For liCatCal = 0 To 4

        sql = " Select ISNULL(SUM(CASE WHEN G.nMoneda = 2 THEN G.nMontoRealiz * " & TxtCambio & " ELSE G.nMontoRealiz END),0) as nGarMonto "
        sql = sql & " from ColocCalifProv CA Inner Join " & ServerCons & "GarantCredConsol GC ON GC.cCtaCod = CA.cCtaCod "
        sql = sql & " Inner Join " & ServerCons & "GarantiasConsol G ON G.cNumGarant = GC.cNumGarant"
        sql = sql & " Where G.nGarClase = 1 AND CA.cCalgen = '" & Trim(Str(liCatCal)) & "' " & lsFFinan
        Conecta.AbreConexion
        Set rsP = Conecta.CargaRecordSet(sql)
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

xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo
MsgBox "Se genero el archivo " & App.Path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"

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

Public Sub DefineFormato178208()

     xlHoja1.Range("A14") = "A. MONTO DE LOS CREDITOS DIRECTOS E"
     xlHoja1.Range("A15") = "INDIRECTOS 1/"
     xlHoja1.Range("B15") = "NORMAL"
     xlHoja1.Range("C15") = "CPP"
     xlHoja1.Range("D15") = "DEFICIENTE"
     xlHoja1.Range("E15") = "DUDOSO"
     xlHoja1.Range("F15") = "PERDIDA"
     xlHoja1.Range("G15") = "TOTAL"
     xlHoja1.Range("A14:G15").Font.Bold = True
    
     xlHoja1.Range("A16") = " Corporativos"
     xlHoja1.Range("A17") = " Grandes empresas"
     xlHoja1.Range("A18") = " Medianas empresas"
     xlHoja1.Range("A19") = " Pequeñas empresas"
     xlHoja1.Range("A20") = " Microempresas"
     xlHoja1.Range("A21") = " Consumo revolvente"
     xlHoja1.Range("A22") = " Consumo no revolvente"
     xlHoja1.Range("A23") = " Hipotecario para vivienda"
     xlHoja1.Range("A24") = " Total"
     
     xlHoja1.Range("A25") = "A'.- MONTO DE LOS CREDITOS DIRECTOS"
     xlHoja1.Range("A26") = "Y EL EQUIVALENTE A RIESGO CREDITICIO DE LOS"
     xlHoja1.Range("A27") = "CREDITOS INDIRECTOS 2/"
     xlHoja1.Range("B27") = "NORMAL"
     xlHoja1.Range("C27") = "CPP"
     xlHoja1.Range("D27") = "DEFICIENTE"
     xlHoja1.Range("E27") = "DUDOSO"
     xlHoja1.Range("F27") = "PERDIDA"
     xlHoja1.Range("G27") = "TOTAL"
     xlHoja1.Range("A25:G27").Font.Bold = True
    
     xlHoja1.Range("A28") = " Corporativos"
     xlHoja1.Range("A29") = " Grandes empresas"
     xlHoja1.Range("A30") = " Medianas empresas"
     xlHoja1.Range("A31") = " Pequeñas empresas"
     xlHoja1.Range("A32") = " Microempresas"
     xlHoja1.Range("A33") = " Consumo revolvente"
     xlHoja1.Range("A34") = " Consumo no revolvente"
     xlHoja1.Range("A35") = " Hipotecario para vivienda"
     xlHoja1.Range("A36") = " Total"
    
     xlHoja1.Range("A37") = "B.  NUMERO DE DEUDORES 3/,"
     xlHoja1.Range("B37") = "NORMAL"
     xlHoja1.Range("C37") = "CPP"
     xlHoja1.Range("D37") = "DEFICIENTE"
     xlHoja1.Range("E37") = "DUDOSO"
     xlHoja1.Range("F37") = "PERDIDA"
     xlHoja1.Range("G37") = "TOTAL"
     xlHoja1.Range("A37:G37").Font.Bold = True
     
     xlHoja1.Range("A38") = " Corporativos"
     xlHoja1.Range("A39") = " Grandes empresas"
     xlHoja1.Range("A40") = " Medianas empresas"
     xlHoja1.Range("A41") = " Pequeñas empresas"
     xlHoja1.Range("A42") = " Microempresas"
     xlHoja1.Range("A43") = " Consumo revolvente"
     xlHoja1.Range("A44") = " Consumo no revolvente"
     xlHoja1.Range("A45") = " Hipotecario para vivienda"
     xlHoja1.Range("A46") = " Total"
     
     xlHoja1.Range("A47") = "C.- MONTO DE LOS CRÉDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO"
     xlHoja1.Range("A48") = "CREDITICIO DE LOS CRÉDITOS INDIRECTOS CON SUSTITUCIÓN DE"
     xlHoja1.Range("A49") = "CONTRAPARTE CREDITICIA - ANTES DE LA SUSTITUCIÓN 5/"
     xlHoja1.Range("B49") = "NORMAL"
     xlHoja1.Range("C49") = "CPP"
     xlHoja1.Range("D49") = "DEFICIENTE"
     xlHoja1.Range("E49") = "DUDOSO"
     xlHoja1.Range("F49") = "PERDIDA"
     xlHoja1.Range("G49") = "TOTAL"
     xlHoja1.Range("A47:G49").Font.Bold = True
     
     xlHoja1.Range("A50") = " Corporativos"
     xlHoja1.Range("A51") = " Grandes empresas"
     xlHoja1.Range("A52") = " Medianas empresas"
     xlHoja1.Range("A53") = " Pequeñas empresas"
     xlHoja1.Range("A54") = " Microempresas"
     xlHoja1.Range("A55") = " Consumo revolvente"
     xlHoja1.Range("A56") = " Consumo no revolvente"
     xlHoja1.Range("A57") = " Hipotecario para vivienda"
     xlHoja1.Range("A58") = " Total"
    
     xlHoja1.Range("A59") = "C'.- MONTO DE LOS CRÉDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO"
     xlHoja1.Range("A60") = "CREDITICIO DE LOS CRÉDITOS INDIRECTOS CON SUSTITUCIÓN DE"
     xlHoja1.Range("A61") = "CONTRAPARTE CREDITICIA - DESPUES DE LA SUSTITUCIÓN 5b/"
     xlHoja1.Range("B61") = "NORMAL"
     xlHoja1.Range("C61") = "CPP"
     xlHoja1.Range("D61") = "DEFICIENTE"
     xlHoja1.Range("E61") = "DUDOSO"
     xlHoja1.Range("F61") = "PERDIDA"
     xlHoja1.Range("G61") = "TOTAL"
     xlHoja1.Range("A59:G61").Font.Bold = True
     
     xlHoja1.Range("A62") = " Corporativos"
     xlHoja1.Range("A63") = " Grandes empresas"
     xlHoja1.Range("A64") = " Medianas empresas"
     xlHoja1.Range("A65") = " Pequeñas empresas"
     xlHoja1.Range("A66") = " Microempresas"
     xlHoja1.Range("A67") = " Consumo revolvente"
     xlHoja1.Range("A68") = " Consumo no revolvente"
     xlHoja1.Range("A69") = " Hipotecario para vivienda"
     xlHoja1.Range("A70") = " Total"
     
     xlHoja1.Range("A71") = "D.- MONTO DE LOS CRÉDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO"
     xlHoja1.Range("A72") = "CREDITICIO DE LOS CRÉDITOS INDIRECTOS"
     xlHoja1.Range("A73") = "QUE CUENTAN CON GARANTIAS PREFERIDAS"
     xlHoja1.Range("A74") = "AUTOLIQUIDABLES 6/"
     xlHoja1.Range("B74") = "NORMAL"
     xlHoja1.Range("C74") = "CPP"
     xlHoja1.Range("D74") = "DEFICIENTE"
     xlHoja1.Range("E74") = "DUDOSO"
     xlHoja1.Range("F74") = "PERDIDA"
     xlHoja1.Range("G74") = "TOTAL"
     xlHoja1.Range("A71:G74").Font.Bold = True
     
     xlHoja1.Range("A75") = " Corporativos"
     xlHoja1.Range("A76") = " Grandes empresas"
     xlHoja1.Range("A77") = " Medianas empresas"
     xlHoja1.Range("A78") = " Pequeñas empresas"
     xlHoja1.Range("A79") = " Microempresas"
     xlHoja1.Range("A80") = " Consumo revolvente"
     xlHoja1.Range("A81") = " Consumo no revolvente"
     xlHoja1.Range("A82") = " Hipotecario para vivienda"
     xlHoja1.Range("A83") = " Total"
    
     xlHoja1.Range("A84") = "D'.- MONTO DE LOS CRÉDITOS QUE"
     xlHoja1.Range("A85") = "CUENTAN CON CONVENIOS ELEGIBLES 7/"
     xlHoja1.Range("B85") = "NORMAL"
     xlHoja1.Range("C85") = "CPP"
     xlHoja1.Range("D85") = "DEFICIENTE"
     xlHoja1.Range("E85") = "DUDOSO"
     xlHoja1.Range("F85") = "PERDIDA"
     xlHoja1.Range("G85") = "TOTAL"
     xlHoja1.Range("A84:G85").Font.Bold = True
     
     xlHoja1.Range("A86") = " Consumo no revolvente"
    
     xlHoja1.Range("A87") = "E.- MONTO DE LOS CRÉDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO"
     xlHoja1.Range("A88") = "CREDITICIO DE LOS CRÉDITOS INDIRECTOS"
     xlHoja1.Range("A89") = "QUE CUENTAN CON GARANTÍAS PREFERIDAS"
     xlHoja1.Range("A90") = "DE MUY RÁPIDA REALIZACIÓN 8/"
     xlHoja1.Range("B90") = "NORMAL"
     xlHoja1.Range("C90") = "CPP"
     xlHoja1.Range("D90") = "DEFICIENTE"
     xlHoja1.Range("E90") = "DUDOSO"
     xlHoja1.Range("F90") = "PERDIDA"
     xlHoja1.Range("G90") = "TOTAL"
     xlHoja1.Range("A87:G90").Font.Bold = True
     
     xlHoja1.Range("A91") = " Corporativos"
     xlHoja1.Range("A92") = " Grandes empresas"
     xlHoja1.Range("A93") = " Medianas empresas"
     xlHoja1.Range("A94") = " Pequeñas empresas"
     xlHoja1.Range("A95") = " Microempresas"
     xlHoja1.Range("A96") = " Hipotecario para vivienda"
     xlHoja1.Range("A97") = " Total"
     
     xlHoja1.Range("A98") = "F.- MONTO DE LOS CRÉDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO"
     xlHoja1.Range("A99") = "CREDITICIO DE LOS CRÉDITOS INDIRECTOS"
     xlHoja1.Range("A100") = "QUE CUENTAN CON GARANTIAS PREFERIDAS 9/"
     xlHoja1.Range("B100") = "NORMAL"
     xlHoja1.Range("C100") = "CPP"
     xlHoja1.Range("D100") = "DEFICIENTE"
     xlHoja1.Range("E100") = "DUDOSO"
     xlHoja1.Range("F100") = "PERDIDA"
     xlHoja1.Range("G100") = "TOTAL"
     xlHoja1.Range("A98:G100").Font.Bold = True
     
     xlHoja1.Range("A101") = " Corporativos"
     xlHoja1.Range("A102") = " Grandes empresas"
     xlHoja1.Range("A103") = " Medianas empresas"
     xlHoja1.Range("A104") = " Pequeñas empresas"
     xlHoja1.Range("A105") = " Microempresas"
     xlHoja1.Range("A106") = " Consumo revolvente 10/"
     xlHoja1.Range("A107") = " Consumo no revolvente 10/"
     xlHoja1.Range("A108") = " Hipotecario para vivienda"
     xlHoja1.Range("A109") = " Total"
    
     xlHoja1.Range("A110") = "G.- MONTO DE LOS CRÉDITOS HIPOTECARIOS"
     xlHoja1.Range("A111") = "QUE CUENTAN CON COBERTURA DEL FONDO MIVIVIENDA"
     xlHoja1.Range("B111") = "NORMAL"
     xlHoja1.Range("C111") = "CPP"
     xlHoja1.Range("D111") = "DEFICIENTE"
     xlHoja1.Range("E111") = "DUDOSO"
     xlHoja1.Range("F111") = "PERDIDA"
     xlHoja1.Range("G111") = "TOTAL"
     xlHoja1.Range("A110:G111").Font.Bold = True
     
     xlHoja1.Range("A112") = " Con cobertura del Fondo Mivivienda 11/"
    
     xlHoja1.Range("A113") = "H.- MONTO DE LOS CRÉDITOS DIRECTOS Y EL EQUIVALENTE A RIESGO"
     xlHoja1.Range("A114") = "CREDITICIO DE LOS CRÉDITOS INDIRECTOS"
     xlHoja1.Range("A115") = "QUE CUENTAN NO CUENTAN CON COBERTURA 12/"
     xlHoja1.Range("B115") = "NORMAL"
     xlHoja1.Range("C115") = "CPP"
     xlHoja1.Range("D115") = "DEFICIENTE"
     xlHoja1.Range("E115") = "DUDOSO"
     xlHoja1.Range("F115") = "PERDIDA"
     xlHoja1.Range("G115") = "TOTAL"
     xlHoja1.Range("A113:G115").Font.Bold = True
     
     xlHoja1.Range("A116") = " Corporativos"
     xlHoja1.Range("A117") = " Grandes empresas"
     xlHoja1.Range("A118") = " Medianas empresas"
     xlHoja1.Range("A119") = " Pequeñas empresas"
     xlHoja1.Range("A120") = " Microempresas"
     xlHoja1.Range("A121") = " Consumo revolvente"
     xlHoja1.Range("A122") = " Consumo no revolvente"
     xlHoja1.Range("A123") = " Hipotecario para vivienda"
     xlHoja1.Range("A124") = " Total"
     
     xlHoja1.Range("A125") = "I.- PROVISIONES CONSTITUIDAS 13/"
     xlHoja1.Range("B125") = "NORMAL"
     xlHoja1.Range("C125") = "CPP"
     xlHoja1.Range("D125") = "DEFICIENTE"
     xlHoja1.Range("E125") = "DUDOSO"
     xlHoja1.Range("F125") = "PERDIDA"
     xlHoja1.Range("G125") = "TOTAL"
     xlHoja1.Range("A125:G125").Font.Bold = True
     
     xlHoja1.Range("A126") = " Corporativos"
     xlHoja1.Range("A127") = " Grandes empresas"
     xlHoja1.Range("A128") = " Medianas empresas"
     xlHoja1.Range("A129") = " Pequeñas empresas"
     xlHoja1.Range("A130") = " Microempresas"
     xlHoja1.Range("A131") = " Consumo revolvente"
     xlHoja1.Range("A132") = " Consumo no revolvente"
     xlHoja1.Range("A133") = " Hipotecario para vivienda"
     xlHoja1.Range("A134") = " Total"
     
     xlHoja1.Range("A135") = "J.- PROVISIONES REQUERIDAS 14/"
     xlHoja1.Range("B135") = "NORMAL"
     xlHoja1.Range("C135") = "CPP"
     xlHoja1.Range("D135") = "DEFICIENTE"
     xlHoja1.Range("E135") = "DUDOSO"
     xlHoja1.Range("F135") = "PERDIDA"
     xlHoja1.Range("G135") = "TOTAL"
     xlHoja1.Range("A135:G135").Font.Bold = True
     
     xlHoja1.Range("A136") = " Corporativos"
     xlHoja1.Range("A137") = " Grandes empresas"
     xlHoja1.Range("A138") = " Medianas empresas"
     xlHoja1.Range("A139") = " Pequeñas empresas"
     xlHoja1.Range("A140") = " Microempresas"
     xlHoja1.Range("A141") = " Consumo revolvente"
     xlHoja1.Range("A142") = " Consumo no revolvente"
     xlHoja1.Range("A143") = " Hipotecario para vivienda"
     xlHoja1.Range("A144") = " Total"
    
     '*****
     xlHoja1.Range("A145") = "K.- SUPERÁVIT (DÉFICIT) DE"
     xlHoja1.Range("A146") = "PROVISIONES 15/"
     xlHoja1.Range("B146") = "NORMAL"
     xlHoja1.Range("C146") = "CPP"
     xlHoja1.Range("D146") = "DEFICIENTE"
     xlHoja1.Range("E146") = "DUDOSO"
     xlHoja1.Range("F146") = "PERDIDA"
     xlHoja1.Range("G146") = "TOTAL"
     xlHoja1.Range("A145:G146").Font.Bold = True
     
     xlHoja1.Range("A147") = " Corporativos"
     xlHoja1.Range("A148") = " Grandes empresas"
     xlHoja1.Range("A149") = " Medianas empresas"
     xlHoja1.Range("A150") = " Pequeñas empresas"
     xlHoja1.Range("A151") = " Microempresas"
     xlHoja1.Range("A152") = " Consumo revolvente"
     xlHoja1.Range("A153") = " Consumo no revolvente"
     xlHoja1.Range("A154") = " Hipotecario para vivienda"
     xlHoja1.Range("A155") = " Total"
    
     xlHoja1.Range("A1").ColumnWidth = 72
     xlHoja1.Range("B1").ColumnWidth = 27
     xlHoja1.Range("C1").ColumnWidth = 27
     xlHoja1.Range("D1").ColumnWidth = 27
     xlHoja1.Range("E1").ColumnWidth = 27
     xlHoja1.Range("F1").ColumnWidth = 27
     xlHoja1.Range("G1").ColumnWidth = 27
    
     xlHoja1.Application.ActiveWindow.Zoom = 80
     xlHoja1.Range("A2:G2").MergeCells = True
     xlHoja1.Range("A3:G3").MergeCells = True
     xlHoja1.Range("A2:G3").Font.Size = 12
    
     xlHoja1.Range("A2").Font.Bold = True
     xlHoja1.Range("A2").HorizontalAlignment = xlCenter
    
     xlHoja1.Range("A8:G8").MergeCells = True
     xlHoja1.Range("A8").HorizontalAlignment = xlCenter
    
     xlHoja1.Range("A9:G9").MergeCells = True
     xlHoja1.Range("A9").HorizontalAlignment = xlCenter
    
     xlHoja1.Range("A11:G11").MergeCells = True
     xlHoja1.Range("A11").HorizontalAlignment = xlCenter
    
    xlHoja1.Range("A157") = "CUADRE DEL ANEXO N°5 CON LAS CIFRAS DEL BALANCE 16/"
    xlHoja1.Range("A157:G157").MergeCells = True
    xlHoja1.Range("A157").HorizontalAlignment = xlCenter
    xlHoja1.Range("A157").Font.Bold = True
    
    xlHoja1.Range("B159") = "V.- CIFRAS DE BALANCE"
    xlHoja1.Range("A159:E159").MergeCells = True
    xlHoja1.Range("A159").Font.Bold = True
    xlHoja1.Range("A159").HorizontalAlignment = xlCenter
    
    xlHoja1.Range("B160") = "Saldo"
    xlHoja1.Range("B160").WrapText = True
    xlHoja1.Range("C160") = "Exposicion equivalente a riesgo crediticio"
    xlHoja1.Range("C160").WrapText = True
    xlHoja1.Range("D160") = "Provisiones genéricas"
    xlHoja1.Range("D160").WrapText = True
    xlHoja1.Range("E160") = "Provisiones específicas"
    xlHoja1.Range("E160").WrapText = True
    xlHoja1.Range("B160:E160").Font.Bold = True
    xlHoja1.Range("B160:E160").HorizontalAlignment = xlCenter
    
    xlHoja1.Range("A161") = "CREDITOS DIRECTOS"
    xlHoja1.Range("A161").Font.Bold = True
    
    xlHoja1.Range("A162") = "Créditos directos: 1401+1403+1404+1405+1406-2901.01-2901.02-2901.04"
    
    xlHoja1.Range("A163") = "CREDITOS INDIRECTOS"
    xlHoja1.Range("A163").Font.Bold = True

    xlHoja1.Range("A164") = "a) Confirmaciones de cartas de crédito irrevocables, de hasta un año, cuando el banco emisor sea una empresa del sistema financiero del exterior de primer nivel"
    xlHoja1.Range("A164").WrapText = True
    xlHoja1.Range("A165") = "b) Emisiones de cartas fianzas que respalden obligaciones de hacer y no hacer"
    xlHoja1.Range("A165").WrapText = True
    xlHoja1.Range("A166") = "c) Emisiones de avales, cartas de crédito de importación y cartas fianzas no incluidas en el literal “b)”, y las confirmaciones de cartas de crédito no incluidas en el literal “a)”, así como  las aceptaciones bancarias "
    xlHoja1.Range("A166").WrapText = True
    xlHoja1.Range("A167") = "Total"
    
    xlHoja1.Range("B168") = "W.- ANEXO 5"
    xlHoja1.Range("A168:E168").MergeCells = True
    xlHoja1.Range("A168").Font.Bold = True
    xlHoja1.Range("A168").HorizontalAlignment = xlCenter
    
    xlHoja1.Range("C169") = "Créditos directos e indirectos afectos a provisiones"
    xlHoja1.Range("C169").Font.Bold = True
    xlHoja1.Range("C169").HorizontalAlignment = xlCenter
    xlHoja1.Range("C169").WrapText = True
    
    xlHoja1.Range("D169") = "Provisiones genéricas constituidas"
    xlHoja1.Range("D169").Font.Bold = True
    xlHoja1.Range("D169").HorizontalAlignment = xlCenter
    xlHoja1.Range("D169").WrapText = True
    
    xlHoja1.Range("E169") = "Provisiones específicas constituidas"
    xlHoja1.Range("E169").Font.Bold = True
    xlHoja1.Range("E169").HorizontalAlignment = xlCenter
    xlHoja1.Range("E169").WrapText = True
    xlHoja1.Range("E170") = "Total"
    
    xlHoja1.Range("C174:C175").HorizontalAlignment = xlCenter
    xlHoja1.Range("C174") = "GERENTE"
    xlHoja1.Range("C175") = "GENERAL"
    
    xlHoja1.Range("D174:D175").HorizontalAlignment = xlCenter
    xlHoja1.Range("D174") = "CONTADOR"
    xlHoja1.Range("D175") = "GENERAL"
    
    xlHoja1.Range("E174:E175").HorizontalAlignment = xlCenter
    xlHoja1.Range("E174") = "FUNCIONARIO"
    xlHoja1.Range("E175") = "RESPONSABLE"
    
    xlHoja1.Range("A2:G174").Font.Name = "Arial Narrow"
    xlHoja1.Range("A2:G174").Font.Size = 10
    
    xlHoja1.Range("B16:G24").Style = "Comma"
    xlHoja1.Range("B28:G36").Style = "Comma"
    xlHoja1.Range("B38:G46").Style = "Comma"
    xlHoja1.Range("B75:G82").Style = "Comma"
    xlHoja1.Range("B101:G108").Style = "Comma"
    xlHoja1.Range("B112:G112").Style = "Comma"
    xlHoja1.Range("B116:G123").Style = "Comma"
    xlHoja1.Range("B126:G133").Style = "Comma"
    xlHoja1.Range("B136:G143").Style = "Comma"

 xlHoja1.Range("A14:G155").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A14:G155").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A14:G155").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A14:G155").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A14:A155").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("A14:A155").Borders(xlEdgeLeft).Weight = xlThin
 xlHoja1.Range("A14:A155").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("A14:A155").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("B14:B155").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("B14:B155").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("C14:C155").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("C14:C155").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("D14:D155").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("D14:D155").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("E14:E155").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("E14:E155").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("F14:F155").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("F14:F155").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("G14:G155").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("G14:G155").Borders(xlEdgeRight).Weight = xlThin

 xlHoja1.Range("A15:G15").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A15:G15").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A24:G24").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A24:G24").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A24:G24").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A27:G27").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A27:G27").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A36:G36").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A36:G36").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A36:G36").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A37:G37").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A37:G37").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A46:G46").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A46:G46").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A46:G46").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A49:G49").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A49:G49").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A58:G58").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A58:G58").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A58:G58").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A61:G61").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A61:G61").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A70:G70").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A70:G70").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A70:G70").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A74:G74").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A74:G74").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A83:G83").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A83:G83").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A83:G83").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A85:G85").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A85:G85").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A86:G86").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A86:G86").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A90:G90").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A90:G90").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A97:G97").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A97:G97").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A97:G97").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A100:G100").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A100:G100").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A109:G109").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A109:G109").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A109:G109").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A111:G111").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A111:G111").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A112:G112").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A112:G112").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A115:G115").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A115:G115").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A124:G124").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A124:G124").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A124:G124").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A125:G125").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A125:G125").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A134:G134").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A134:G134").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A134:G134").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A135:G135").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A135:G135").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A144:G144").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A144:G144").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A144:G144").Borders(xlEdgeBottom).Weight = xlThin

 xlHoja1.Range("A146:G146").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A146:G146").Borders(xlEdgeBottom).Weight = xlThin
 
 xlHoja1.Range("A155:G155").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A155:G155").Borders(xlEdgeBottom).Weight = xlThin
 xlHoja1.Range("A155:G155").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A155:G155").Borders(xlEdgeTop).Weight = xlThin
 
 xlHoja1.Range("A159:E167").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A159:E167").Borders(xlEdgeBottom).Weight = xlThin
 xlHoja1.Range("A159:E167").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A159:E167").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A159:E167").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("A159:E167").Borders(xlEdgeLeft).Weight = xlThin
 xlHoja1.Range("A159:E167").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("A159:E167").Borders(xlEdgeRight).Weight = xlThin
 xlHoja1.Range("A159:E167").Borders(xlInsideVertical).LineStyle = xlContinuous
 xlHoja1.Range("A159:E167").Borders(xlInsideVertical).Weight = xlThin
 xlHoja1.Range("A159:E167").Borders(xlInsideHorizontal).LineStyle = xlContinuous
 xlHoja1.Range("A159:E167").Borders(xlInsideHorizontal).Weight = xlThin
 
 xlHoja1.Range("A169:E170").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A169:E170").Borders(xlEdgeBottom).Weight = xlThin
 xlHoja1.Range("A169:E170").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A169:E170").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A169:E170").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("A169:E170").Borders(xlEdgeLeft).Weight = xlThin
 xlHoja1.Range("A169:E170").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("A169:E170").Borders(xlEdgeRight).Weight = xlThin
 xlHoja1.Range("A169:E170").Borders(xlInsideVertical).LineStyle = xlContinuous
 xlHoja1.Range("A169:E170").Borders(xlInsideVertical).Weight = xlThin
 xlHoja1.Range("A169:E170").Borders(xlInsideHorizontal).LineStyle = xlContinuous
 xlHoja1.Range("A169:E170").Borders(xlInsideHorizontal).Weight = xlThin
End Sub

Public Sub DefineFormato178208A(ByVal psNomCmac As String, ByVal dFecha As Date)
     xlHoja1.Range("A1:AZ1000") = ""
     xlHoja1.Range("A1").ColumnWidth = 16
     xlHoja1.Range("B1").ColumnWidth = 45
     xlHoja1.Range("C1").ColumnWidth = 28
     xlHoja1.Range("D1").ColumnWidth = 29
     xlHoja1.Range("E1").ColumnWidth = 28
     
     xlHoja1.Range("B9:E29").Borders(xlEdgeLeft).LineStyle = xlContinuous
     xlHoja1.Range("B9:E29").Borders(xlEdgeLeft).Weight = xlThin
     xlHoja1.Range("B9:E29").Borders(xlEdgeRight).LineStyle = xlContinuous
     xlHoja1.Range("B9:E29").Borders(xlEdgeRight).Weight = xlThin
     xlHoja1.Range("B9:E29").Borders(xlEdgeBottom).LineStyle = xlContinuous
     xlHoja1.Range("B9:E29").Borders(xlEdgeBottom).Weight = xlThin
     xlHoja1.Range("B9:E29").Borders(xlEdgeTop).LineStyle = xlContinuous
     xlHoja1.Range("B9:E29").Borders(xlEdgeTop).Weight = xlThin
     xlHoja1.Range("B9:E29").Borders(xlInsideVertical).LineStyle = xlContinuous
     xlHoja1.Range("B9:E29").Borders(xlInsideVertical).Weight = xlThin
     ''2
     xlHoja1.Range("B10:E28").Borders(xlEdgeLeft).LineStyle = xlContinuous
     xlHoja1.Range("B10:E28").Borders(xlEdgeLeft).Weight = xlThin
     xlHoja1.Range("B10:E28").Borders(xlEdgeRight).LineStyle = xlContinuous
     xlHoja1.Range("B10:E28").Borders(xlEdgeRight).Weight = xlThin
     xlHoja1.Range("B10:E28").Borders(xlEdgeBottom).LineStyle = xlContinuous
     xlHoja1.Range("B10:E28").Borders(xlEdgeBottom).Weight = xlThin
     xlHoja1.Range("B10:E28").Borders(xlEdgeTop).LineStyle = xlContinuous
     xlHoja1.Range("B10:E28").Borders(xlEdgeTop).Weight = xlThin
     xlHoja1.Range("B10:E28").Borders(xlInsideVertical).LineStyle = xlContinuous
     xlHoja1.Range("B10:E28").Borders(xlInsideVertical).Weight = xlThin
     ''3
     xlHoja1.Range("C31:C32").Borders(xlEdgeLeft).LineStyle = xlContinuous
     xlHoja1.Range("C31:C32").Borders(xlEdgeLeft).Weight = xlThin
     xlHoja1.Range("C31:C32").Borders(xlEdgeRight).LineStyle = xlContinuous
     xlHoja1.Range("C31:C32").Borders(xlEdgeRight).Weight = xlThin
     xlHoja1.Range("C31:C31").Borders(xlEdgeTop).LineStyle = xlContinuous
     xlHoja1.Range("C31:C31").Borders(xlEdgeTop).Weight = xlThin
     xlHoja1.Range("C31:C31").Borders(xlEdgeBottom).LineStyle = xlContinuous
     xlHoja1.Range("C31:C31").Borders(xlEdgeBottom).Weight = xlThin
     xlHoja1.Range("C32:C32").Borders(xlEdgeBottom).LineStyle = xlContinuous
     xlHoja1.Range("C32:C32").Borders(xlEdgeBottom).Weight = xlThin
     
     'xlHoja1.Range("C25:C26").Borders(xlInsideVertical).LineStyle = xlContinuous
     'xlHoja1.Range("C25:C26").Borders(xlInsideVertical).Weight = xlThin
     'Terminar*********************************************************
     '*****************************************************************
     
     xlHoja1.Range("C1") = "ANEXO N° 5-A"
     
     xlHoja1.Range("A3") = "EMPRESA : " & psNomCmac
     xlHoja1.Range("C4") = "AL   " & Format(dFecha, "DD") & "  DE " & UCase(Format(dFecha, "MMMM")) & "  DEL " & Format(dFecha, "YYYY")
     xlHoja1.Range("C5") = "(En miles de nuevos soles)"
     xlHoja1.Range("C7") = "RESUMEN DE PROVISIONES PROCICLICAS"
     xlHoja1.Range("F3").Font.Bold = True
     xlHoja1.Range("A9:A9").MergeCells = True
     
     xlHoja1.Range("B10") = "Corporativos"
     xlHoja1.Range("B11") = "Corporativos con garantia autoliquidables"
     xlHoja1.Range("B12") = "Grandes empresas"
     xlHoja1.Range("B13") = "Grandes empresas con garantía autoliquidable"
     xlHoja1.Range("B14") = "Medianas empresas"
     xlHoja1.Range("B15") = "Pequeñas empresas"
     xlHoja1.Range("B16") = "Microempresas"
     xlHoja1.Range("B17") = "Consumo revolvente"
     xlHoja1.Range("B18") = "Consumo no-revolvente"
     xlHoja1.Range("B19") = "Consumo no-revolvente bajo convenios elegibles"
     xlHoja1.Range("B20") = "Hipotecario para vivienda"
     xlHoja1.Range("B21") = "Hipotecario para vivienda con garantía autoliquidables"
     xlHoja1.Range("B22") = "Grandes empresas - sustitución de contraparte"
     xlHoja1.Range("B23") = "Medianas empresas - sustitución de contraparte"
     xlHoja1.Range("B24") = "Pequeñas empresas - sustitución de contraparte"
     xlHoja1.Range("B25") = "Microempresas - sustitución de contraparte"
     xlHoja1.Range("B26") = "Consumo revolvente - sustitución de contraparte"
     xlHoja1.Range("B27") = "Consumo no-revolvente - sustitución de contraparte"
     xlHoja1.Range("B28") = "Hipotecario para Vivienda - sustitución de contraparte"
     xlHoja1.Range("B29") = "Total"
     
     xlHoja1.Range("C9") = "Endeudamiento en categoria normal 1/"
     xlHoja1.Range("C9").HorizontalAlignment = xlJustify
     xlHoja1.Range("A9").RowHeight = 39
     xlHoja1.Range("D9") = "Provisión Prociclica constituida"
     xlHoja1.Range("D9").HorizontalAlignment = xlJustify
     xlHoja1.Range("E9") = "Provisión Prociclica requerida 2/"
     xlHoja1.Range("E9").HorizontalAlignment = xlJustify
     
    xlHoja1.Range("A1:E9").Font.Bold = True
    xlHoja1.Range("A12:A20").Font.Bold = True
    xlHoja1.Range("B31") = "Indicar si se encuentra en periodo de adecuación"
    xlHoja1.Range("C31") = "0"
    xlHoja1.Range("B32") = "Si está en período de adecuación, indicar mes"
    xlHoja1.Range("C32") = "0"
    xlHoja1.Range("D31") = "1=si, 0=no"
    xlHoja1.Range("D32") = "0=primer mes, 1=segundo y tercer mes, 2=cuarto y quinto mes"
    
    xlHoja1.Range("B36") = "______________"
    xlHoja1.Range("B36").HorizontalAlignment = xlCenter
    xlHoja1.Range("C36") = "______________"
    xlHoja1.Range("C36").HorizontalAlignment = xlCenter
    xlHoja1.Range("E36") = "______________"
    xlHoja1.Range("E36").HorizontalAlignment = xlCenter
    
    xlHoja1.Range("C37").HorizontalAlignment = xlCenter
    xlHoja1.Range("C38").HorizontalAlignment = xlCenter
     
    xlHoja1.Range("B37") = "GERENTE"
    xlHoja1.Range("B38") = "GENERAL"
    xlHoja1.Range("B37").HorizontalAlignment = xlCenter
    xlHoja1.Range("B38").HorizontalAlignment = xlCenter
    
    xlHoja1.Range("C37") = "CONTADOR"
    xlHoja1.Range("C38") = "GENERAL"
    
    xlHoja1.Range("E37") = "FUNCIONARIO"
    xlHoja1.Range("E38") = "RESPONSABLE"
    xlHoja1.Range("E37").HorizontalAlignment = xlCenter
    xlHoja1.Range("E38").HorizontalAlignment = xlCenter
End Sub
Public Sub Genera_Reporte178208(ByVal TxtCambio As Double, ByVal dFecha As Date, _
                        ByVal nCF As Integer, ByVal psServer As String, _
                        ByVal dFecData As String, _
                        Optional ByVal psNomCmac As String, _
                        Optional ByVal psCodCMAC As String)


Dim sql As String
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
Dim MatrixAnexo5A() As String
ReDim Preserve MatrixAnexo5A(1 To 3, 1 To 13)
Dim pbAdecuacion As Boolean
Dim lsColumna As String
Dim nValProvProcCorporativo As Double
Dim nValProvProcGrandeEmp As Double
Dim nValProvProcMedianaEmp As Double
Dim nValProvProcPequenaEmp As Double
Dim nValProvProcMicroEmp As Double
Dim nValProvProcConsumoRev As Double
Dim nValProvProcConsumoNoRev As Double
Dim nValProvProcHipotecario As Double
'*******************************

Dim nTempo As Double 'ARCV 13-06-2007

Set RCC = New DRCC
lsServConsol = RCC.ServConsol(psServer)
Set RCC = Nothing

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

'MAVM 20120702 Observ Cert ***
'lsArchivo = "Rep_Anexo05_" & Format(dFecha, "YYYYMM") & IIf(lbIncluyeCF, "", "SinCF") & ".XLS"
lsArchivo = "Rep_Anexo05_" & Format(dFecha, "YYYYMM") & Format(Time, "hhmmss") & IIf(lbIncluyeCF, "", "SinCF") & ".XLS"
'***

Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & lsArchivo)
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

''ALPA 20150722***********************************************************************************
'Dim lnCapitalRN1 As Currency
'Dim lnInteresRN1 As Currency
'Dim lnCapitalRN2 As Currency
'Dim lnInteresRN2 As Currency
'Dim lnCapitalRN3 As Currency
'Dim lnInteresRN3 As Currency
'Dim lnCapitalRN4 As Currency
'Dim lnInteresRN4 As Currency
'Dim lnCapitalRN5 As Currency
'Dim lnInteresRN5 As Currency
'Dim lnCapitalRN6 As Currency
'Dim lnInteresRN6 As Currency
'Dim lnCapitalRN7 As Currency
'Dim lnInteresRN7 As Currency
'Dim nTipoGarantia As Integer
'Dim lrs As ADODB.Recordset
'Set lrs = New ADODB.Recordset
'Set lrs = ObtieneSaldosDeRefinanciadoANormal(dFecha, TxtCambio)
'If Not (lrs.BOF Or lrs.EOF) Then
'Do While Not lrs.EOF
'nTipoGarantia1 = lrs!nTipoGarantia
'lrs.MoveNext
'Loop
''************************************************************************************************

'Modificado por BRGO 20100609 (Encabezado de la Hoja)********************************************
 xlHoja1.Range("A1:G200") = ""
 xlHoja1.Range("A2") = "ANEXO N° 5"
 xlHoja1.Range("A3") = "INFORME DE CLASIFICACION DE DEUDORES Y PROVISIONES"
 xlHoja1.Range("A6") = "EMPRESA : " & psNomCmac
 xlHoja1.Range("G6") = "CODIGO : " & psCodCMAC
 xlHoja1.Range("A8") = "AL   " & Format(dFecha, "DD") & "  DE " & UCase(Format(dFecha, "MMMM")) & "  DEL " & Format(dFecha, "YYYY")
 xlHoja1.Range("A9") = "(En nuevos soles)"
 xlHoja1.Range("A11") = "I.- INFORME DE CLASIFICACION DE LOS DEUDORES DE LA CARTERA DE CRÉDITOS, CONTINGENCIASY ARRENDAMIENTOS FINANCIEROS"
 Call DefineFormato178208

sql = " select substring(x.cTpoCredCod,1,1) cta, cCalGen ,"
sql = sql & " Isnull(sum(case when substring(cCtaCod,9,1) = '1' AND nGarant>0 then nSaldoCap"
sql = sql & "          when substring(cCtaCod,9,1) = '2' AND nGarant>0 then nSaldoCap * " & tc
sql = sql & "  end ),0) SaldoCap,"
sql = sql & " Count( distinct(cPersCod) ) Deudores,"
sql = sql & " Isnull(sum(case when nGarant > 0 and substring(cCtaCod,9,1) = '1' /*and x.cTpoCredCod <> '853'*/ AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then nSaldoCap" 'JUEZ 20121212 Se comentaron la no inclusion de los tipos de creditos 853 (MiVivienda)
sql = sql & "                 when nGarant > 0 and substring(cCtaCod,9,1) = '2' /*and x.cTpoCredCod <> '853'*/ AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then nSaldoCap * " & tc

If lbIncluyeCF = True Then
  sql = sql & "                 when nGarant > 0 and substring(cCtaCod,9,1) = '1' /*and x.cTpoCredCod <> '853'*/ AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then nSaldoCap * 0.5"
  sql = sql & "                 when nGarant > 0 and substring(cCtaCod,9,1) = '2' /*and x.cTpoCredCod <> '853'*/ AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then nSaldoCap * 0.5 * " & tc
End If

sql = sql & "  end ),0) SaldoCapEquivRiesCred,"
'JUEZ 20160210 *************************************
'sql = sql & " Isnull(sum(case when nGarant = 1 and substring(cCtaCod,9,1) = '1' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then nSaldoCap"
'sql = sql & "                 when nGarant = 1 and substring(cCtaCod,9,1) = '2' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then nSaldoCap * " & tc
  sql = sql & " Isnull(sum(case when nGarant = 1 and substring(cCtaCod,9,1) = '1' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan)"
  sql = sql & "                 when nGarant = 1 and substring(cCtaCod,9,1) = '2' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan) * " & tc
'END JUEZ ******************************************
If lbIncluyeCF = True Then
'JUEZ 20160210 *************************************
'  sql = sql & "                 when nGarant = 1 and substring(cCtaCod,9,1) = '1' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then nSaldoCap * 0.5"
'  sql = sql & "                 when nGarant = 1 and substring(cCtaCod,9,1) = '2' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then nSaldoCap * 0.5 * " & tc
  sql = sql & "                 when nGarant = 1 and substring(cCtaCod,9,1) = '1' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan) * 0.5"
  sql = sql & "                 when nGarant = 1 and substring(cCtaCod,9,1) = '2' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan) * 0.5 * " & tc
'END JUEZ ******************************************
End If

sql = sql & "  end ),0) SaldoCapGarAutoLiq,"
sql = sql & " Isnull(sum(case when nGarant = 2 and substring(cCtaCod,9,1) = '1' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then nSaldoCap"
sql = sql & "                 when nGarant = 2 and substring(cCtaCod,9,1) = '2' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then nSaldoCap * " & tc

If lbIncluyeCF = True Then
  sql = sql & "                 when nGarant = 2 and substring(cCtaCod,9,1) = '1' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then nSaldoCap * 0.5"
  sql = sql & "                 when nGarant = 2 and substring(cCtaCod,9,1) = '2' and x.cTpoCredCod <> '853' AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then nSaldoCap * 0.5 * " & tc
End If

sql = sql & "  end ),0) SaldoCapGarMuyRR,"
sql = sql & " Isnull(sum(case when nGarant = 3 and substring(cCtaCod,9,1) = '1' /*and x.cTpoCredCod <> '853'*/ AND x.dFecVig>(CASE WHEN x.cTpoCredCod='853' THEN '20100630' ELSE '19000101' END) AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then nSaldoCap-nMontoIntRefinan" 'JUEZ 20121212 Se comentaron la no inclusion de los tipos de creditos 853 (MiVivienda)
sql = sql & "                 when nGarant = 3 and substring(cCtaCod,9,1) = '2' /*and x.cTpoCredCod <> '853'*/ AND x.dFecVig>(CASE WHEN x.cTpoCredCod='853' THEN '20100630' ELSE '19000101' END) AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan) * " & tc

If lbIncluyeCF = True Then
    sql = sql & "                 when nGarant = 3 and substring(cCtaCod,9,1) = '1' /*and x.cTpoCredCod <> '853'*/ AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan) * 0.5"
    sql = sql & "                 when nGarant = 3 and substring(cCtaCod,9,1) = '2' /*and x.cTpoCredCod <> '853'*/ AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan) * 0.5 * " & tc
End If

sql = sql & "  end ),0) SaldoCapGarPref,"
sql = sql & " Isnull(sum(case when x.cTpoCredCod = '853' And substring(cCtaCod,9,1) = '1' AND nConsidera=1 then nSaldoCap"
sql = sql & "          when x.cTpoCredCod = '853' And substring(cCtaCod,9,1) = '2' AND nConsidera=1 then nSaldoCap * " & tc
sql = sql & "  end ),0) SaldoCapMiVivienda,"
sql = sql & " Isnull(sum(case when nGarant = 4 and substring(cCtaCod,9,1) = '1' /*and x.cTpoCredCod <> '853'*/ AND x.dFecVig>(CASE WHEN x.cTpoCredCod='853' THEN '20100630' ELSE '19000101' END) AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan)" 'JUEZ 20121212 Se comentaron la no inclusion de los tipos de creditos 853 (MiVivienda)
sql = sql & "                 when nGarant = 4 and substring(cCtaCod,9,1) = '2' /*and x.cTpoCredCod <> '853'*/ AND x.dFecVig>(CASE WHEN x.cTpoCredCod='853' THEN '20100630' ELSE '19000101' END) AND substring(x.cTpoCredCod,2,2)<>'81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan) * " & tc

If lbIncluyeCF = True Then
    sql = sql & "                 when nGarant = 4 and substring(cCtaCod,9,1) = '1' /*and x.cTpoCredCod <> '853'*/ AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan) * 0.5"
    sql = sql & "                 when nGarant = 4 and substring(cCtaCod,9,1) = '2' /*and x.cTpoCredCod <> '853'*/ AND substring(x.cTpoCredCod,2,2)= '81' AND nConsidera=1 then (nSaldoCap-nMontoIntRefinan) * 0.5 * " & tc
End If

sql = sql & "  end ),0) SaldoCapSinG,"
sql = sql & " Isnull(sum(case when substring(cCtaCod,9,1) = '1' AND nGarant>0  /*and x.cTpoCredCod <> '853'*/ then nProvision" 'JUEZ 20121212 Se comentaron la no inclusion de los tipos de creditos 853 (MiVivienda)
sql = sql & "                 when substring(cCtaCod,9,1) = '2' AND nGarant>0  /*and x.cTpoCredCod <> '853'*/ then (nProvision) * " & tc
sql = sql & "  end ),0) ProvConstituida,"
sql = sql & " Isnull(sum(case when nGarant = 0 and substring(cCtaCod,9,1) = '1'  then nSaldoCap"
sql = sql & "   when nGarant = 0 and substring(cCtaCod,9,1) = '2' then nSaldoCap * " & tc
sql = sql & "  end ),0) SaldoCapitalG,"
sql = sql & " Isnull(sum(case when nGarant = 0 and substring(cCtaCod,9,1) = '1'  then nProvision"
sql = sql & "   when nGarant = 0 and substring(cCtaCod,9,1) = '2' then nProvision * " & tc
sql = sql & "  end ),0) ProvisionG"
sql = sql & ",Isnull(sum(case when substring(cCtaCod,9,1) = '1' AND nGarant>0  then ISNULL(nProvisionRCC,0)  when substring(cCtaCod,9,1) = '2' AND nGarant>0  then (ISNULL(nProvisionRCC,0)) * " & tc & "  end ),0) ProvRCC"
'ALPA20140114************************
sql = sql & ",nMontoIntRefinan=sum(case when substring(cCtaCod,9,1) = '1' AND nGarant>0  then ISNULL(nMontoIntRefinan,0)  when substring(cCtaCod,9,1) = '2' AND nGarant>0  then (ISNULL(nMontoIntRefinan,0)) * " & tc & "  end ) "
sql = sql & ",nCapitalREVI=sum(nCapitalREVI),nInteresREVI=sum(nInteresREVI),nTipoGarantia=nTipoGarantia "
If lbMesActual = True Then
    sql = sql & " FROM( SELECT CA.cCtaCod,CA.cPersCod, CA.cCalGen ,CA.nSaldoCap , CA.nGarant, CA.nProvision, CA.nProvisionRCC ,nConsidera= 1, Ca.cTpoCredCod, Ca.dFecVig,nMontoIntRefinan=isnull(ca.nMontoIntRefinan,0) " 'JUEZ 20130118 Se agregó Ca.dFecVig
    sql = sql & " ,nCapitalREVI=isnull(REVI.nCapital,0),nInteresREVI=isnull(REVI.nInteres,0),nTipoGarantia=isnull(REVI.nTipoGarantia,0)"
    'sql = sql & " FROM( SELECT CA.cCtaCod,CA.cPersCod, CA.cCalGen ,CA.nSaldoCap , CA.nGarant, CA.nProvision, CA.nProvisionRCC ,nConsidera= 1, C.cTpoCredCod"
    sql = sql & "       FROM ColocCalifProv ca "
    'ALPA 20150914*************************
    sql = sql & "left join (select MC.cCtaCod,cTpoCredCod=left(C.cTpoCredCod,1),"
    sql = sql & "   CP.cCalgen,"
    sql = sql & "   nCapital =isnull((case MCD.cOpeCod when '101107' then MCD.nMonto else 0 end*Case when substring(MCD.cCtaCod,9,1) = '1' then 1 else " & tc & " end),0),"
    sql = sql & "   nInteres =isnull((case MCD.cOpeCod when '101108' then MCD.nMonto else 0 end*Case when substring(MCD.cCtaCod,9,1) = '1' then 1 else " & tc & " end),0),"
    sql = sql & "   nTipoGarantia = CP.nGarant "
    sql = sql & "from Mov M"
    sql = sql & "    inner join MovCol MC on M.nMovNro=MC.nMovNro"
    sql = sql & "    inner join MovColDet MCD on MC.nMovNro=MCD.nMovNro"
    sql = sql & "                and MC.nMovNro=MCD.nMovNro and MC.cOpeCod=MCD.cOpeCod and MC.cCtaCod=MCD.cCtaCod"
    sql = sql & "    inner join Colocaciones C on MC.cCtaCod=C.cCtaCod"
    sql = sql & "    inner join ColocCalifProv CP on CP.cCtaCod=C.cCtaCod "
    sql = sql & "where left(M.cMovNro,8)='" & Format(dFecha, "YYYY/MM/DD") & "'"
    sql = sql & "   and M.nMovFlag=0"
    sql = sql & " and M.cOpeCod in ('101107','101108')"
    'sql = sql & "group by C.cTpoCredCod, CP.nGarant,CP.cCalgen
    sql = sql & ") REVI on REVI.cCtaCod=ca.cCtaCod "
    '**************************************
    sql = sql & " INNER JOIN Colocaciones C ON C.cCtaCod = ca.cCtaCod WHERE"
Else
    sql = sql & " FROM( SELECT CA.cCtaCod,CA.cPersCod, CA.cCalGen ,CA.nSaldoCap , CA.nGarant, CA.nProvision, CA.nProvisionRCC ,nConsidera= 1,Ca.cTpoCredCod, Ca.dFecVig,nMontoIntRefinan=isnull(ca.nMontoIntRefinan,0) " 'JUEZ 20130118 Se agregó Ca.dFecVig
    sql = sql & " ,nCapitalREVI=isnull(REVI.nCapital,0),nInteresREVI=isnull(REVI.nInteres,0),nTipoGarantia=isnull(REVI.nTipoGarantia,0)" 'ALPA20150915
    sql = sql & " FROM DBConsolidada..ColocCalifProvTotal ca "
     'ALPA 20150914*************************
    sql = sql & "left join (select cTpoCredCod=left(C.cTpoCredCod,1),"
    sql = sql & "   CP.cCalgen,"
    sql = sql & "   nCapital =isnull((case MCD.cOpeCod when '101107' then MCD.nMonto else 0 end*Case when substring(MCD.cCtaCod,9,1) = '1' then 1 else " & tc & " end),0),"
    sql = sql & "   nInteres =isnull((case MCD.cOpeCod when '101108' then MCD.nMonto else 0 end*Case when substring(MCD.cCtaCod,9,1) = '1' then 1 else " & tc & " end),0),"
    sql = sql & "   nTipoGarantia = CP.nGarant,MCD.cCtaCod "
    sql = sql & "from Mov M"
    sql = sql & "    inner join MovCol MC on M.nMovNro=MC.nMovNro"
    sql = sql & "    inner join MovColDet MCD on MC.nMovNro=MCD.nMovNro"
    sql = sql & "                and MC.nMovNro=MCD.nMovNro and MC.cOpeCod=MCD.cOpeCod and MC.cCtaCod=MCD.cCtaCod"
    sql = sql & "    inner join Colocaciones C on MC.cCtaCod=C.cCtaCod"
    sql = sql & "    inner join ColocCalifProv CP on CP.cCtaCod=C.cCtaCod "
    sql = sql & "where left(M.cMovNro,8)='" & Format(dFecha, "YYYY/MM/DD") & "'"
    sql = sql & "   and M.nMovFlag=0"
    sql = sql & " and M.cOpeCod in ('101107','101108') "
    'sql = sql & "group by C.cTpoCredCod, CP.nGarant,CP.cCalgen
    sql = sql & ") REVI on REVI.cCtaCod=ca.cCtaCod "
    '**************************************
    sql = sql & " Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
End If
sql = sql & " nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107,2092) "
sql = sql & ") X    "
sql = sql & " Group by substring(cTpoCredCod,1,1), cCalGen,X.nTipoGarantia   Order by substring(cTpoCredCod,1,1), cCalGen "


Co.AbreConexion
Set rs = Co.CargaRecordSet(sql)

sql = "DROP TABLE ColocCalifRefin"
Co.Ejecutar (sql)

While Not rs.EOF
            'JUEZ 20160210 *********************************************************************************************************************
'            Select Case rs!cCalGen
'                Case "0":
'                            xlHoja1.Range("B" & 15 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("B" & 27 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapEquivRiesCred), 0, rs!SaldoCapEquivRiesCred - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("B" & 37 + rs!Cta) = Format(IIf(IsNull(rs!Deudores), 0, rs!Deudores), "#0")
'                            xlHoja1.Range("B" & 49 + rs!Cta) = Format(0, "#.00") ' Item C
'                            xlHoja1.Range("B" & 61 + rs!Cta) = Format(0, "#.00") ' Item C'
'                            xlHoja1.Range("B" & 74 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarAutoLiq), 0, rs!SaldoCapGarAutoLiq - IIf(rs!nTipoGarantia = 1, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("B" & 100 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarPref), 0, rs!SaldoCapGarPref - IIf(rs!nTipoGarantia = 3, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("B" & 115 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapSinG), 0, rs!SaldoCapSinG - IIf(rs!nTipoGarantia = 4, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("B" & 125 + rs!Cta) = Format(IIf(IsNull(rs!ProvConstituida), 0, rs!ProvConstituida), "#.000")
'
'                Case "1":
'                            xlHoja1.Range("C" & 15 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("C" & 27 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapEquivRiesCred), 0, rs!SaldoCapEquivRiesCred - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("C" & 37 + rs!Cta) = Format(IIf(IsNull(rs!Deudores), 0, rs!Deudores), "#0")
'                            xlHoja1.Range("C" & 49 + rs!Cta) = Format(0, "#.00") ' Item C
'                            xlHoja1.Range("C" & 61 + rs!Cta) = Format(0, "#.00") ' Item C'
'                            xlHoja1.Range("C" & 74 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarAutoLiq), 0, rs!SaldoCapGarAutoLiq - IIf(rs!nTipoGarantia = 1, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("C" & 100 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarPref), 0, rs!SaldoCapGarPref - IIf(rs!nTipoGarantia = 3, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("C" & 115 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapSinG), 0, rs!SaldoCapSinG - IIf(rs!nTipoGarantia = 4, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("C" & 125 + rs!Cta) = Format(IIf(IsNull(rs!ProvConstituida), 0, rs!ProvConstituida), "#.000")
'
'                Case "2":
'                            xlHoja1.Range("D" & 15 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("D" & 27 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapEquivRiesCred), 0, rs!SaldoCapEquivRiesCred - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("D" & 37 + rs!Cta) = Format(IIf(IsNull(rs!Deudores), 0, rs!Deudores), "#0")
'                            xlHoja1.Range("D" & 49 + rs!Cta) = Format(0, "#.00") ' Item C
'                            xlHoja1.Range("D" & 61 + rs!Cta) = Format(0, "#.00") ' Item C'
'                            xlHoja1.Range("D" & 74 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarAutoLiq), 0, rs!SaldoCapGarAutoLiq - IIf(rs!nTipoGarantia = 1, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("D" & 100 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarPref), 0, rs!SaldoCapGarPref - IIf(rs!nTipoGarantia = 3, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("D" & 115 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapSinG), 0, rs!SaldoCapSinG - IIf(rs!nTipoGarantia = 4, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("D" & 125 + rs!Cta) = Format(IIf(IsNull(rs!ProvConstituida), 0, rs!ProvConstituida), "#.000")
'
'                Case "3":
'                            xlHoja1.Range("E" & 15 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("E" & 27 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapEquivRiesCred), 0, rs!SaldoCapEquivRiesCred - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("E" & 37 + rs!Cta) = Format(IIf(IsNull(rs!Deudores), 0, rs!Deudores), "#0")
'                            xlHoja1.Range("E" & 49 + rs!Cta) = Format(0, "#.00") ' Item C
'                            xlHoja1.Range("E" & 61 + rs!Cta) = Format(0, "#.00") ' Item C'
'                            xlHoja1.Range("E" & 74 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarAutoLiq), 0, rs!SaldoCapGarAutoLiq - IIf(rs!nTipoGarantia = 1, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("E" & 100 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarPref), 0, rs!SaldoCapGarPref - IIf(rs!nTipoGarantia = 3, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("E" & 115 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapSinG), 0, rs!SaldoCapSinG - IIf(rs!nTipoGarantia = 4, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("E" & 125 + rs!Cta) = Format(IIf(IsNull(rs!ProvConstituida), 0, rs!ProvConstituida), "#.000")
'
'                Case "4":
'                            xlHoja1.Range("F" & 15 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("F" & 27 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapEquivRiesCred), 0, rs!SaldoCapEquivRiesCred - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
'                            xlHoja1.Range("F" & 37 + rs!Cta) = Format(IIf(IsNull(rs!Deudores), 0, rs!Deudores), "#0")
'                            xlHoja1.Range("F" & 49 + rs!Cta) = Format(0, "#.00") ' Item C
'                            xlHoja1.Range("F" & 61 + rs!Cta) = Format(0, "#.00") ' Item C'
'                            xlHoja1.Range("F" & 74 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarAutoLiq), 0, rs!SaldoCapGarAutoLiq - IIf(rs!nTipoGarantia = 1, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("F" & 100 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarPref), 0, rs!SaldoCapGarPref - IIf(rs!nTipoGarantia = 3, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("F" & 115 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapSinG), 0, rs!SaldoCapSinG - IIf(rs!nTipoGarantia = 4, rs!nInteresREVI, 0)), "#.00")
'                            xlHoja1.Range("F" & 125 + rs!Cta) = Format(IIf(IsNull(rs!ProvConstituida), 0, rs!ProvConstituida), "#.000")
'            End Select
            Select Case rs!cCalGen
                Case "0": lsColumna = "B"
                Case "1": lsColumna = "C"
                Case "2": lsColumna = "D"
                Case "3": lsColumna = "E"
                Case "4": lsColumna = "F"
            End Select
            
            xlHoja1.Range(lsColumna & 15 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
            xlHoja1.Range(lsColumna & 27 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapEquivRiesCred), 0, rs!SaldoCapEquivRiesCred - (rs!nMontoIntRefinan - rs!nInteresREVI)), "#.00")
            xlHoja1.Range(lsColumna & 37 + rs!Cta) = Format(IIf(IsNull(rs!Deudores), 0, rs!Deudores), "#0")
            xlHoja1.Range(lsColumna & 49 + rs!Cta) = Format(0, "#.00") ' Item C
            xlHoja1.Range(lsColumna & 61 + rs!Cta) = Format(0, "#.00") ' Item C'
            xlHoja1.Range(lsColumna & 74 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarAutoLiq), 0, rs!SaldoCapGarAutoLiq - IIf(rs!nTipoGarantia = 1, rs!nInteresREVI, 0)), "#.00")
            xlHoja1.Range(lsColumna & 100 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapGarPref), 0, rs!SaldoCapGarPref - IIf(rs!nTipoGarantia = 3, rs!nInteresREVI, 0)), "#.00")
            xlHoja1.Range(lsColumna & 115 + rs!Cta) = Format(IIf(IsNull(rs!SaldoCapSinG), 0, rs!SaldoCapSinG - IIf(rs!nTipoGarantia = 4, rs!nInteresREVI, 0)), "#.00")
            xlHoja1.Range(lsColumna & 125 + rs!Cta) = Format(IIf(IsNull(rs!ProvConstituida), 0, rs!ProvConstituida), "#.000")
            'END JUEZ **************************************************************************************************************************
    rs.MoveNext
 Wend

'** By BRGO 20100714
sql = " select substring(Ca.cTpoCredCod,1,1), cCalGen ,"
sql = sql & " Isnull(sum(case when substring(ca.cCtaCod,9,1) = '1' then nSaldoCap"
sql = sql & "   when substring(ca.cCtaCod,9,1) = '2' then nSaldoCap * " & tc & " end ),0) SaldoCap,"
sql = sql & " Count( distinct(ca.cPersCod) ) Deudores,"
sql = sql & " Isnull(sum(case when nGarant = 1 and substring(ca.cCtaCod,9,1) = '1'  then nSaldoCap"
sql = sql & "   when nGarant = 1 and substring(ca.cCtaCod,9,1) = '2'  then nSaldoCap * " & tc & " end ),0) SaldoCapGarAutoLiq,"
sql = sql & " Isnull(sum(case when nGarant = 2 and substring(ca.cCtaCod,9,1) = '1'  then nSaldoCap"
sql = sql & "   when nGarant = 2 and substring(ca.cCtaCod,9,1) = '2'  then nSaldoCap * " & tc & " end ),0) SaldoCapGarMuyRR,"
sql = sql & " Isnull(sum(case when nGarant = 3 and substring(ca.cCtaCod,9,1) = '1'  then nSaldoCap"
sql = sql & "   when nGarant = 3 and substring(ca.cCtaCod,9,1) = '2' then nSaldoCap * " & tc & " end ),0) SaldoCapGarPref,"
sql = sql & " Isnull(sum(case when  substring(ca.cCtaCod,9,1) = '1' then nSaldoCap"
sql = sql & "   when  substring(ca.cCtaCod,9,1) = '2' then nSaldoCap * " & tc & " end ),0) SaldoCapMiVivienda,"
sql = sql & " Isnull(sum(case when nGarant = 4 and substring(ca.cCtaCod,9,1) = '1'  then nSaldoCap"
sql = sql & "   when nGarant = 4 and substring(ca.cCtaCod,9,1) = '2' then nSaldoCap * " & tc & " end ),0) SaldoCapSinG,"
sql = sql & " Isnull(sum(case when substring(ca.cCtaCod,9,1) = '1'  then nProvision + ISNULL(nProvisionRCC,0)"
sql = sql & "   when substring(ca.cCtaCod,9,1) = '2'  then ( nProvision + ISNULL(nProvisionRCC,0) ) * " & tc & " end ),0) ProvConstituida, "
sql = sql & " Isnull(sum(case when substring(ca.cCtaCod,9,1) = '1'  then ISNULL(nProvisionRCC,0)"
sql = sql & "   when substring(ca.cCtaCod,9,1) = '2'  then (ISNULL(nProvisionRCC,0) ) * " & tc & " end ),0) ProvRCC "

If lbMesActual = True Then
    sql = sql & " FROM ColocCalifProv ca inner join Colocaciones C on C.cCtaCod = ca.cCtaCod Where "
Else
    sql = sql & " FROM DBConsolidada..ColocCalifProvTotal ca inner join Colocaciones C on C.cCtaCod = ca.cCtaCod Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
End If
sql = sql & " nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107) And "
sql = sql & " Ca.cTpoCredCod = '853' AND Ca.dFecVig<='20100630' " 'JUEZ 20121212 consignar sólo si los créditos hipotecarios fueron otorgados antes del 30 de junio de 2010
sql = sql & " Group by substring(Ca.cTpoCredCod,1,1), cCalGen Order by substring(Ca.cTpoCredCod,1,1), cCalGen "

'** End By

Set rs = Co.CargaRecordSet(sql)

While Not rs.EOF
    Select Case rs!cCalGen
        Case "0": 'JUEZ 20121212
             xlHoja1.Range("B112") = IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap) 'Round(rs!SaldoCap / 3, 2))
        Case "1"
             xlHoja1.Range("C112") = IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap) 'Round(rs!SaldoCap / 3, 2))
        Case "2"
             xlHoja1.Range("D112") = IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap) 'Round(rs!SaldoCap / 3, 2))
        Case "3"
             xlHoja1.Range("E112") = IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap) 'Round(rs!SaldoCap / 3, 2))
        Case "4"
             xlHoja1.Range("F112") = IIf(IsNull(rs!SaldoCap), 0, rs!SaldoCap) 'Round(rs!SaldoCap / 3, 2))
    End Select
    rs.MoveNext
Wend
'--- Cálculos ------------------------------------

 xlHoja1.Range("C136").Formula = "= (C50*5%)+(C75*1%)+(C91*1.25%)+(C101*2.5%)+(C116*5%)"
 xlHoja1.Range("C137").Formula = "= (C51*5%)+(C76*1%)+(C92*1.25%)+(C102*2.5%)+(C117*5%)"
 xlHoja1.Range("C138").Formula = "= (C77*1%)+(C93*1.25%)+(C103*2.5%)+(C118*5%)"
 xlHoja1.Range("C139").Formula = "= (C78*1%)+(C94*1.25%)+(C104*2.5%)+(C119*5%)"
 xlHoja1.Range("C140").Formula = "= (C79*1%)+(C95*1.25%)+(C105*2.5%)+(C120*5%)"
 xlHoja1.Range("C141").Formula = "= (C80*1%)+(C106*2.5%)+(C121*5%)"
 xlHoja1.Range("C142").Formula = "= (C81*1%)+(C86*1.25%)+(C107*2.5%)+(C122*5%)"
 xlHoja1.Range("C143").Formula = "= (C82*1%)+(C87*1.25%)+(C108*2.5%)+(C123*5%)"
 
 xlHoja1.Range("D136").Formula = "= (D62*25%)+(D75*1%)+(D91*6.25%)+(D101*12.5%)+(D116*25%)"
 xlHoja1.Range("D137").Formula = "= (D63*25%)+(D76*1%)+(D92*6.25%)+(D102*12.5%)+(D117*25%)"
 xlHoja1.Range("D138").Formula = "= (D77*1%)+(D93*6.25%)+(D103*12.5%)+(D118*25%)"
 xlHoja1.Range("D139").Formula = "= (D78*1%)+(D94*6.25%)+(D104*12.5%)+(D119*25%)"
 xlHoja1.Range("D140").Formula = "= (D79*1%)+(D95*6.25%)+(D105*12.5%)+(D120*25%)"
 xlHoja1.Range("D141").Formula = "= (D80*1%)+(D106*12.5%)+(D121*25%)"
 xlHoja1.Range("D142").Formula = "= (D81*1%)+(D86*6.25%)+(D107*12.5%)+(D122*25%)"
 xlHoja1.Range("D143").Formula = "= (D82*1%)+(D98*6.25%)+(D108*12.5%)+(D123*25%)"
 
 xlHoja1.Range("E136").Formula = "= (E62*60%)+(E75*1%)+(E91*15%)+(E101*30%)+(E116*60%)"
 xlHoja1.Range("E137").Formula = "= (E63*60%)+(E76*1%)+(E92*15%)+(E102*30%)+(E117*60%)"
 xlHoja1.Range("E138").Formula = "= (E77*1%)+(E93*15%)+(E103*30%)+(E118*60%)"
 xlHoja1.Range("E139").Formula = "= (E78*1%)+(E94*15%)+(E104*30%)+(E119*60%)"
 xlHoja1.Range("E140").Formula = "= (E79*1%)+(E95*15%)+(E105*30%)+(E120*60%)"
 xlHoja1.Range("E141").Formula = "= (E80*1%)+(E106*30%)+(E121*60%)"
 xlHoja1.Range("E142").Formula = "= (E81*1%)+(E86*15%)+(E107*30%)+(E122*60%)"
 xlHoja1.Range("E143").Formula = "= (E82*1%)+(E98*15%)+(E108*30%)+(E123*60%)"

 xlHoja1.Range("F136").Formula = "= (F62*100%)+(F75*1%)+(F91*30%)+(F101*60%)+(F116*100%)"
 xlHoja1.Range("F137").Formula = "= (F63*100%)+(F76*1%)+(F92*30%)+(F102*60%)+(F117*100%)"
 xlHoja1.Range("F138").Formula = "= (F77*1%)+(F93*30%)+(F103*60%)+(F118*100%)"
 xlHoja1.Range("F139").Formula = "= (F78*1%)+(F94*30%)+(F104*60%)+(F119*100%)"
 xlHoja1.Range("F140").Formula = "= (F79*1%)+(F95*30%)+(F105*60%)+(F120*100%)"
 xlHoja1.Range("F141").Formula = "= (F80*1%)+(F106*60%)+(F121*100%)"
 xlHoja1.Range("F142").Formula = "= (F81*1%)+(F86*30%)+(F107*60%)+(F122*100%)"
 xlHoja1.Range("F143").Formula = "= (F82*1%)+(F98*30%)+(F108*60%)+(F123*100%)"

 For i = 147 To 154
    xlHoja1.Range("B" & i).Formula = "= B" & i - 21 & "-B" & i - 11
    xlHoja1.Range("C" & i).Formula = "= C" & i - 21 & "-C" & i - 11
    xlHoja1.Range("D" & i).Formula = "= D" & i - 21 & "-D" & i - 11
    xlHoja1.Range("E" & i).Formula = "= E" & i - 21 & "-E" & i - 11
    xlHoja1.Range("F" & i).Formula = "= F" & i - 21 & "-F" & i - 11
 Next i
'---------
  xlHoja1.Range("B24").Formula = "= B16+B17+B18+B19+B20+B21+B22+B23"
  xlHoja1.Range("C24").Formula = "= C16+C17+C18+C19+C20+C21+C22+C23"
  xlHoja1.Range("D24").Formula = "= D16+D17+D18+D19+D20+D21+D22+D23"
  xlHoja1.Range("E24").Formula = "= E16+E17+E18+E19+E20+E21+E22+E23"
  xlHoja1.Range("B36").Formula = "= B28+B29+B30+B31+B32+B33+B34+B35"
  xlHoja1.Range("C36").Formula = "= C28+C29+C30+C31+C32+C33+C34+C35"
  xlHoja1.Range("D36").Formula = "= D28+D29+D30+D31+D32+D33+D34+D35"
  xlHoja1.Range("E36").Formula = "= E28+E29+E30+E31+E32+E33+E34+E35"

 Call SumaTotalAnxo5(16, 23, xlHoja1)
 Call SumaTotalAnxo5(28, 35, xlHoja1)
 Call SumaTotalAnxo5(38, 45, xlHoja1)
 Call SumaTotalAnxo5(50, 57, xlHoja1)
 Call SumaTotalAnxo5(62, 69, xlHoja1)
 Call SumaTotalAnxo5(75, 82, xlHoja1)
 Call SumaTotalAnxo5(101, 108, xlHoja1)
 Call SumaTotalAnxo5(116, 123, xlHoja1)
 Call SumaTotalAnxo5(126, 133, xlHoja1)
 Call SumaTotalAnxo5(136, 143, xlHoja1)
 Call SumaTotalAnxo5(147, 154, xlHoja1)
 
 Call SumaExcelAnexo5(16, 24, xlHoja1)
 Call SumaExcelAnexo5(28, 36, xlHoja1)
 Call SumaExcelAnexo5(38, 46, xlHoja1)
 Call SumaExcelAnexo5(50, 58, xlHoja1)
 Call SumaExcelAnexo5(62, 70, xlHoja1)
 Call SumaExcelAnexo5(75, 83, xlHoja1)
 Call SumaExcelAnexo5(86, 86, xlHoja1)
 Call SumaExcelAnexo5(91, 97, xlHoja1)
 Call SumaExcelAnexo5(101, 109, xlHoja1)
 Call SumaExcelAnexo5(112, 112, xlHoja1)
 Call SumaExcelAnexo5(116, 124, xlHoja1)
 Call SumaExcelAnexo5(126, 134, xlHoja1)
 Call SumaExcelAnexo5(136, 144, xlHoja1)
 Call SumaExcelAnexo5(147, 155, xlHoja1)
 

 'By Capi 12112007  Para cambiar numero de deudores
 sql = " Select cCalGen Calificacion,Count(distinct cPersCod)Deudores "
 If lbMesActual = True Then
    sql = sql & " From ColocCalifProv ca Where "
 Else
    sql = sql & " From DBConsolidada..ColocCalifProvTotal ca Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
 End If
 sql = sql & " nPrdEstado Not in (2202,2204,2206,2900) Group By cCalGen"
 
 Set rs = Co.CargaRecordSet(sql)
 While Not rs.EOF
    Select Case rs!Calificacion
        Case "0"
            xlHoja1.Range("B46") = Format(rs!Deudores, "#0")
        Case "1"
            xlHoja1.Range("C46") = Format(rs!Deudores, "#0")
        Case "2"
            xlHoja1.Range("D46") = Format(rs!Deudores, "#0")
        Case "3"
            xlHoja1.Range("E46") = Format(rs!Deudores, "#0")
        Case "4"
            xlHoja1.Range("F46") = Format(rs!Deudores, "#0")
    End Select
    rs.MoveNext
 Wend
 xlHoja1.Range("G46").Formula = "= B46+C46+D46+E46+F46"
 
 '***CUADRE DEL BALANCE BRGO
  xlHoja1.Range("B162") = Obtener_Valor("B162", dFecha)
  xlHoja1.Range("B165") = Obtener_Valor("B165", dFecha)
  xlHoja1.Range("D162") = Obtener_Valor("D162", dFecha)
  xlHoja1.Range("D165") = Obtener_Valor("D165", dFecha)
  xlHoja1.Range("E162") = Obtener_Valor("E162", dFecha)
  xlHoja1.Range("E165") = Obtener_Valor("E165", dFecha)
  xlHoja1.Range("C162").Formula = "= B162"
  xlHoja1.Range("C165").Formula = "= B165*50%"
  xlHoja1.Range("B167").Formula = "= B162+B164+B165+B166"
  xlHoja1.Range("C167").Formula = "= C162+C164+C165+C166"
  xlHoja1.Range("D167").Formula = "= D162+D164+D165+D166"
  xlHoja1.Range("E167").Formula = "= E162+E164+E165+E166"
  
  xlHoja1.Range("C170").Formula = "= G36"
  xlHoja1.Range("D170").Formula = "= B134"
  xlHoja1.Range("E170").Formula = "= G134-B134"
  xlHoja1.Range("B147:G154").NumberFormat = "#,###,###.00"
  xlHoja1.Range("B162:G167").NumberFormat = "#,###,###.00"

 Set rs = Nothing
 sql = "exec stp_sel_ObtieneDatosAnexo5A " & CDbl(TxtCambio) & ", '" & Format(dFecha, "YYYY/MM/DD") & "'"
 Set rs = Co.CargaRecordSet(sql)
 pbAdecuacion = True
 While Not rs.EOF
    Select Case rs!nConsValor
        Case 100: 'Corporativo
            MatrixAnexo5A(2, 1) = rs!nValorNor
            MatrixAnexo5A(2, 2) = rs!nValorGaranAuto
        Case 200: 'Grandes empresas
            MatrixAnexo5A(2, 3) = rs!nValorNor
            MatrixAnexo5A(2, 4) = rs!nValorGaranAuto
        Case 300: 'Medianas empresas
            MatrixAnexo5A(2, 5) = rs!nValorNor
        Case 400: 'Pequeñas empresas
            MatrixAnexo5A(2, 6) = rs!nValorNor
        Case 500: 'Microempresas
            MatrixAnexo5A(2, 7) = rs!nValorNor
        Case 600: 'Consumo revolvente
            MatrixAnexo5A(2, 8) = rs!nValorNor
        Case 700: 'Consumo no revolvente
            MatrixAnexo5A(2, 9) = rs!nValorNor
        Case 800: 'Hipotecario
            MatrixAnexo5A(2, 11) = rs!nValorNor
            MatrixAnexo5A(2, 12) = rs!nValorGaranAuto
    End Select
    rs.MoveNext
Wend
'ALPA 20090217********************************************************
Set rs = Nothing
'*********************************************************************
MatrixAnexo5A(2, 10) = 0
MatrixAnexo5A(2, 13) = "=SUM(D10:D28)"
MatrixAnexo5A(1, 1) = IIf(xlHoja1.Range("B28").Formula = "", 0, xlHoja1.Range("B28").Formula) - IIf(xlHoja1.Range("B75").Formula = "", 0, xlHoja1.Range("B75").Formula) - IIf(xlHoja1.Range("B50").Formula = "", 0, xlHoja1.Range("B50").Formula) + IIf(xlHoja1.Range("B62").Formula = "", 0, xlHoja1.Range("B62").Formula)
MatrixAnexo5A(1, 2) = IIf(xlHoja1.Range("B75").Formula = "", 0, xlHoja1.Range("B75").Formula)
MatrixAnexo5A(1, 3) = IIf(xlHoja1.Range("B29").Formula = "", 0, xlHoja1.Range("B29").Formula) - IIf(xlHoja1.Range("B76").Formula = "", 0, xlHoja1.Range("B76").Formula) - IIf(xlHoja1.Range("B51").Formula = "", 0, xlHoja1.Range("B51").Formula) + IIf(xlHoja1.Range("B63").Formula = "", 0, xlHoja1.Range("B63").Formula)
MatrixAnexo5A(1, 4) = IIf(xlHoja1.Range("B76").Formula = "", 0, xlHoja1.Range("B76").Formula)
MatrixAnexo5A(1, 5) = IIf(xlHoja1.Range("B30").Formula = "", 0, xlHoja1.Range("B30").Formula) - IIf(xlHoja1.Range("B77").Formula = "", 0, xlHoja1.Range("B77").Formula) - IIf(xlHoja1.Range("B52").Formula = "", 0, xlHoja1.Range("B52").Formula)
MatrixAnexo5A(1, 6) = IIf(xlHoja1.Range("B31").Formula = "", 0, xlHoja1.Range("B31").Formula) - IIf(xlHoja1.Range("B78").Formula = "", 0, xlHoja1.Range("B78").Formula) - IIf(xlHoja1.Range("B53").Formula = "", 0, xlHoja1.Range("B53").Formula)
MatrixAnexo5A(1, 7) = IIf(xlHoja1.Range("B32").Formula = "", 0, xlHoja1.Range("B32").Formula) - IIf(xlHoja1.Range("B79").Formula = "", 0, xlHoja1.Range("B79").Formula) - IIf(xlHoja1.Range("B54").Formula = "", 0, xlHoja1.Range("B54").Formula)
MatrixAnexo5A(1, 8) = IIf(xlHoja1.Range("B33").Formula = "", 0, xlHoja1.Range("B33").Formula) - IIf(xlHoja1.Range("B80").Formula = "", 0, xlHoja1.Range("B80").Formula) - IIf(xlHoja1.Range("B55").Formula = "", 0, xlHoja1.Range("B55").Formula)
MatrixAnexo5A(1, 9) = IIf(xlHoja1.Range("B34").Formula = "", 0, xlHoja1.Range("B34").Formula) - IIf(xlHoja1.Range("B81").Formula = "", 0, xlHoja1.Range("B81").Formula) - IIf(xlHoja1.Range("B86").Formula = "", 0, xlHoja1.Range("B86").Formula) - IIf(xlHoja1.Range("B56").Formula = "", 0, xlHoja1.Range("B56").Formula)
MatrixAnexo5A(1, 10) = IIf(xlHoja1.Range("B75").Formula = "", 0, xlHoja1.Range("B75").Formula)
MatrixAnexo5A(1, 11) = IIf(xlHoja1.Range("B35").Formula = "", 0, xlHoja1.Range("B35").Formula) - IIf(xlHoja1.Range("B82").Formula = "", 0, xlHoja1.Range("B82").Formula) - IIf(xlHoja1.Range("B112").Formula = "", 0, xlHoja1.Range("B112").Formula) - IIf(xlHoja1.Range("B57").Formula = "", 0, xlHoja1.Range("B57").Formula)
MatrixAnexo5A(1, 12) = IIf(xlHoja1.Range("B82").Formula = "", 0, xlHoja1.Range("B82").Formula)
MatrixAnexo5A(1, 13) = "=SUM(C10:C28)"
 'ALPA 20090217******************************************************************
 sql = "exec stp_ObtenerProcentajeDeProvision "
 Set rs = Co.CargaRecordSet(sql)
 
 Dim x40 As Integer ' Indicar si se encuentra en periodo de adecuación
 Dim x41 As Integer ' Si esta en periodo de adecuación, indicar mes
 x40 = 1 'Asignar valor
 Select Case Format(dFecha, "YYYYMM")
    Case "201006": x41 = 0 'Borrar. Ingresado sólo para efectos de pruebas
    Case "201007": x41 = 0
    Case "201008": x41 = 1
    Case "201009": x41 = 1
    Case "201010": x41 = 2
    Case "201011": x41 = 2
    Case Else: x41 = 2
 End Select
 
While Not rs.EOF
    Select Case rs!nCalCodTab
        Case 1001:
                MatrixAnexo5A(3, 1) = MatrixAnexo5A(1, 1) * rs!nProvisionProciclica / 100
                MatrixAnexo5A(3, 2) = MatrixAnexo5A(1, 2) * rs!nProvisionProciclicaA / 100
        Case 2001:
                MatrixAnexo5A(3, 3) = MatrixAnexo5A(1, 3) * rs!nProvisionProciclica / 100
                MatrixAnexo5A(3, 4) = MatrixAnexo5A(1, 4) * rs!nProvisionProciclicaA / 100
        Case 3001:
                MatrixAnexo5A(3, 5) = MatrixAnexo5A(1, 5) * rs!nProvisionProciclica / 100
        Case 4001:
                MatrixAnexo5A(3, 6) = MatrixAnexo5A(1, 6) * rs!nProvisionProciclica / 100
        Case 5001:
                MatrixAnexo5A(3, 7) = MatrixAnexo5A(1, 7) * rs!nProvisionProciclica / 100
        Case 6001:
                MatrixAnexo5A(3, 8) = MatrixAnexo5A(1, 8) * rs!nProvisionProciclica / 100
        Case 7001:
                MatrixAnexo5A(3, 9) = MatrixAnexo5A(1, 9) * rs!nProvisionProciclica / 100
                MatrixAnexo5A(3, 10) = "=0"
        Case 8001:
                MatrixAnexo5A(3, 11) = MatrixAnexo5A(1, 11) * rs!nProvisionProciclica / 100
                MatrixAnexo5A(3, 12) = MatrixAnexo5A(1, 12) * rs!nProvisionProciclicaA / 100
    End Select
    rs.MoveNext
Wend

MatrixAnexo5A(3, 13) = "=SUM(E10:E28)"
Set rs = Nothing

nValProvProcCorporativo = CDbl(IIf(MatrixAnexo5A(2, 1) = "", 0, MatrixAnexo5A(2, 1))) + CDbl(IIf(MatrixAnexo5A(2, 2) = "", 0, MatrixAnexo5A(2, 2)))
nValProvProcGrandeEmp = CDbl(IIf(MatrixAnexo5A(2, 3) = "", 0, MatrixAnexo5A(2, 3))) + CDbl(IIf(MatrixAnexo5A(2, 4) = "", 0, MatrixAnexo5A(2, 4)))
nValProvProcMedianaEmp = CDbl(IIf(MatrixAnexo5A(2, 5) = "", 0, MatrixAnexo5A(2, 5)))
nValProvProcPequenaEmp = CDbl(IIf(MatrixAnexo5A(2, 6) = "", 0, MatrixAnexo5A(2, 6)))
nValProvProcMicroEmp = CDbl(IIf(MatrixAnexo5A(2, 7) = "", 0, MatrixAnexo5A(2, 7)))
nValProvProcConsumoRev = CDbl(IIf(MatrixAnexo5A(2, 8) = "", 0, MatrixAnexo5A(2, 8)))
nValProvProcConsumoNoRev = CDbl(IIf(MatrixAnexo5A(2, 9) = "", 0, MatrixAnexo5A(2, 9))) + CDbl(IIf(MatrixAnexo5A(2, 10) = "", 0, MatrixAnexo5A(2, 10)))
nValProvProcHipotecario = CDbl(IIf(MatrixAnexo5A(2, 11) = "", 0, MatrixAnexo5A(2, 11))) + CDbl(IIf(MatrixAnexo5A(2, 12) = "", 0, MatrixAnexo5A(2, 12)))

 lsNomHoja = "Anexo5A"
 lbExisteHoja = False
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

Call DefineFormato178208A(psNomCmac, dFecha)

'*** Se verifica si se ha activado la provision procíclica BRGO 2010/09/13
sql = "select SUM(ISNULL(nProvisionProciclica,0)) ProvProc from ColocCalifProv"
Set rs = Co.CargaRecordSet(sql)
'************************************

'Ingresando los valores al Cuadro de provisiones prociclicas
Dim xj As Integer
For i = 1 To 13
    If i <> 13 Then
        xj = i
    Else
        xj = 20
    End If
    xlHoja1.Range("C" & xj + 9) = Format(MatrixAnexo5A(1, i), "###,###,###,##0.00")
    xlHoja1.Range("D" & xj + 9) = Format(MatrixAnexo5A(2, i), "###,###,###,##0.00")
    
    If rs!ProvProc = 0 Then
        xlHoja1.Range("E" & xj + 9) = Format(0, "###,###,###,##0.00")
    Else
        xlHoja1.Range("E" & xj + 9) = Format(MatrixAnexo5A(3, i), "###,###,###,##0.00")
    End If
 
Next
xlHoja1.Range("C22:E28").value = 0
xlHoja1.Range("C10:E28").NumberFormat = "#,###,###.00"
xlHoja1.Range("C31") = x40
xlHoja1.Range("C32") = x41

'Terminar
lsNomHoja = "Anexo5"
 lbExisteHoja = False
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
'BRGO *****************************************************************
 xlHoja1.Range("B136").Formula = "= (B62+B75+B91+B101+B116)*0.7% + Anexo5A!E10 + Anexo5A!E11"
 xlHoja1.Range("B137").Formula = "= (B63+B76+B92+B102+B117)*0.7% + Anexo5A!E12 + Anexo5A!E13"
 xlHoja1.Range("B138").Formula = "= (B77+B93+B103+B118)*1% + Anexo5A!E14"
 xlHoja1.Range("B139").Formula = "= (B78+B94+B104+B119)*1% + Anexo5A!E15"
 xlHoja1.Range("B140").Formula = "= (B79+B95+B105+B120)*1% + Anexo5A!E16"
 xlHoja1.Range("B141").Formula = "= (B80+B106+B121)*1.0% + Anexo5A!E17"
 xlHoja1.Range("B142").Formula = "= (B81+B107+B122)*1.0% + (B86*1.0%) + Anexo5A!E18 + Anexo5A!E19"
 xlHoja1.Range("B143").Formula = "= (B82+B98+B108+B123)*0.7% + Anexo5A!E20 + Anexo5A!E21"
 xlHoja1.Range("B126") = xlHoja1.Range("B126") + nValProvProcCorporativo
 xlHoja1.Range("B127") = xlHoja1.Range("B127") + nValProvProcGrandeEmp
 xlHoja1.Range("B128") = xlHoja1.Range("B128") + nValProvProcMedianaEmp
 xlHoja1.Range("B129") = xlHoja1.Range("B129") + nValProvProcPequenaEmp
 xlHoja1.Range("B130") = xlHoja1.Range("B130") + nValProvProcMicroEmp
 xlHoja1.Range("B131") = xlHoja1.Range("B131") + nValProvProcConsumoRev
 xlHoja1.Range("B132") = xlHoja1.Range("B132") + nValProvProcConsumoNoRev
 xlHoja1.Range("B133") = xlHoja1.Range("B133") + nValProvProcHipotecario
 
Co.CierraConexion
xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo
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
End Sub

Private Sub SumaTotalAnxo5(ByVal nI As Integer, ByVal nF As Integer, ByVal xlHoja1 As Object)
  xlHoja1.Range("B" & nF + 1).Formula = "= B" & nI & "+B" & nI + 1 & "+B" & nI + 2 & "+B" & nI + 3 & "+B" & nI + 4 & "+B" & nI + 5 & "+B" & nI + 6 & "+B" & nI + 7
  xlHoja1.Range("C" & nF + 1).Formula = "= C" & nI & "+C" & nI + 1 & "+C" & nI + 2 & "+C" & nI + 3 & "+C" & nI + 4 & "+C" & nI + 5 & "+C" & nI + 6 & "+C" & nI + 7
  xlHoja1.Range("D" & nF + 1).Formula = "= D" & nI & "+D" & nI + 1 & "+D" & nI + 2 & "+D" & nI + 3 & "+D" & nI + 4 & "+D" & nI + 5 & "+D" & nI + 6 & "+D" & nI + 7
  xlHoja1.Range("E" & nF + 1).Formula = "= E" & nI & "+E" & nI + 1 & "+E" & nI + 2 & "+E" & nI + 3 & "+E" & nI + 4 & "+E" & nI + 5 & "+E" & nI + 6 & "+E" & nI + 7
  xlHoja1.Range("F" & nF + 1).Formula = "= F" & nI & "+F" & nI + 1 & "+F" & nI + 2 & "+F" & nI + 3 & "+F" & nI + 4 & "+F" & nI + 5 & "+F" & nI + 6 & "+F" & nI + 7
End Sub
Private Sub SumaExcelAnexo5(ByVal nInicio As Integer, ByVal nFin As Integer, ByVal xlHoja1 As Object)
 Dim i As Integer
 For i = nInicio To nFin
    xlHoja1.Range("G" & i).Formula = "= B" & i & "+C" & i & "+D" & i & "+E" & i & "+F" & i
 Next i

End Sub


Private Function Obtener_Valor(ByVal psCelda As String, _
                                ByVal pdFecha As Date) As String
Dim oCon As DConecta
Dim sql As String
Dim rs As ADODB.Recordset

'LUCV20180601, Se pasó a script las condiciones
'Select Case psCelda
''    Case "B106"
''        sql = "SELECT nDato= dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','141[13456]%',1) + dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','721[1234]%',1) "
''        sql = sql & " + (dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','142[13456]%',2) + dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','722[1234]%',2) ) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 1, pdFecha), "mm/dd/yyyy") & "') "
''    Case "C106"
''        sql = "SELECT nDato= ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14190[1234]0[23]%',1)) + ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','271102%',1)) "
''        sql = sql & " + (   ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14290[1234]0[23]%',2)) + ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14290[1234]0[5]%',2)) + ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','272102%',2)) ) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 1, pdFecha), "mm/dd/yyyy") & "')"
''    Case "D106"
''        sql = "SELECT nDato= ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14190[1234]01%',1)) + ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','271101%',1)) "
''        sql = sql & " + (ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14290[1234]01%',2)) + ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','272101%',2))  ) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 1, pdFecha), "mm/dd/yyyy") & "')"
'
'     Case "B162"
''         sql = "SELECT nDato= dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','141[13456]%',1)"
''         sql = sql & " + (dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','142[13456]%',2)) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 5, pdFecha), "mm/dd/yyyy") & "') "
''         sql = sql & " - (dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','29110[124]%',1))"
''         sql = sql & " - (dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','29210[124]%',2)) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 5, pdFecha), "mm/dd/yyyy") & "') "
'          sql = "SELECT nDato = ((SELECT sum(nSaldoFinImporte) FROM BalanceEstad "
'          'ALPA20140114****************************
'          'Sql = Sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and cCtaContCod like '140[13456]' and cBalanceCate = 1)" 'BRGO 20120403
'          sql = sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and (cCtaContCod like '140[13456]' OR cCtaContCod like '29010207') and cBalanceCate = 1)" 'BRGO 20120403
'
'          sql = sql & " - (SELECT sum(nSaldoFinImporte) FROM BalanceEstad "
'          'ALPA20140114****************************
'          'Sql = Sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and cCtaContCod like '29010[124]' and cBalanceCate = 1))" 'BRGO 20120403
'          sql = sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and (cCtaContCod like '29010[124]' OR cCtaContCod like '29010801') and cBalanceCate = 1))" 'BRGO 20120403
'
'     Case "B165"
''         sql = "SELECT nDato= dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','721[1234]%',1) "
''         sql = sql & " + (dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','722[1234]%',2)) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 5, pdFecha), "mm/dd/yyyy") & "') "
'          sql = "SELECT nDato = (SELECT sum(nSaldoFinImporte) FROM BalanceEstad "
'          sql = sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and cCtaContCod like '720[1234]'  and cBalanceCate = 1)"
'
'     Case "D162"
''         sql = "SELECT nDato= ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14190[23456789]0[23]%',1)) + ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14191[012]0[23]%',1)) "
''         sql = sql & " +     (ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14290[23456789]0[23]%',2)) + ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14291[012]0[23]%',2))) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 5, pdFecha), "mm/dd/yyyy") & "')"
'          sql = "SELECT nDato = ((SELECT sum(nSaldoFinImporte) FROM BalanceEstad "
'          sql = sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and cCtaContCod like '14090[23456789]0[23]' and cBalanceCate = 1)"  'BRGO 20120403
'          sql = sql & " + (SELECT sum(nSaldoFinImporte) FROM BalanceEstad "
'          sql = sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and cCtaContCod like '14091[0123]0[23]' and cBalanceCate = 1))" 'BRGO 20120403
'     Case "D165"
''         sql = "SELECT nDato= ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','271102%',1)) "
''         sql = sql & " + (ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','272102%',2))) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 5, pdFecha), "mm/dd/yyyy") & "')"
'          sql = "SELECT nDato = (SELECT sum(nSaldoFinImporte) FROM BalanceEstad "
'          sql = sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and cCtaContCod like '270102' and cBalanceCate = 1)" 'BRGO 20120403
'     Case "E162"
''         sql = "SELECT nDato= ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14190[1234]01%',1)) "
''         sql = sql & " + (ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','14290[1234]01%',2))) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 5, pdFecha), "mm/dd/yyyy") & "')"
'          sql = "SELECT nDato = ((SELECT sum(nSaldoFinImporte) FROM BalanceEstad "
'          sql = sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and cCtaContCod like '14090[23456789]01' and cBalanceCate = 1)" 'BRGO 20120403
'          sql = sql & " + (SELECT sum(nSaldoFinImporte) FROM BalanceEstad "
'          sql = sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and cCtaContCod like '14091[0123]01' and cBalanceCate = 1))" 'BRGO 20120403
'
'     Case "E165"
''         sql = "SELECT nDato= ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','271101%',1)) "
''         sql = sql & " + (ABS(dbo.GetSaldoCtaAcumulado('" & Format(pdFecha, "mm/dd/yyyy") & "','272101%',2))  ) * dbo.GetTipoCambio(1,'" & Format(DateAdd("d", 5, pdFecha), "mm/dd/yyyy") & "')"
'          sql = "SELECT nDato = (SELECT sum(nSaldoFinImporte) FROM BalanceEstad "
'          sql = sql & " WHERE cBalanceMes = '" & Format(pdFecha, "mm") & "' and cBalanceAnio = '" & Format(pdFecha, "yyyy") & "' and cCtaContCod like '270101' and cBalanceCate = 1)" 'BRGO 20120403
'End Select
    
    sql = "Exec stp_sel_ObtenerValorBalance '" & Format(pdFecha, "yyyymmdd") & "',  '" & psCelda & "'  "
'Fin LUCV20180601

Set oCon = New DConecta
oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sql)

Obtener_Valor = CStr(Abs(rs!nDato))

oCon.CierraConexion
Set oCon = Nothing
End Function

Public Sub Genera_Reporte178210Antes(ByVal TxtCambio As Double, ByVal dFecha As Date, _
                        ByVal pnCartaFianza As Integer, ByVal psServer As String, _
                        ByVal dFecData As String, _
                        Optional ByVal psNomCmac As String, _
                        Optional ByVal psCodCMAC As String)

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
If pnCartaFianza = 1 Then
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
Co.AbreConexion

lbCambio = False
lsArchivo = "Anexo5_" & Format(dFecha, "yyyymm") & "_HojaW_808" & IIf(lbIncluyeCF = True, "", "_SinCF") & ".XLS"
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & lsArchivo)
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
        If pnCartaFianza = 1 Then
            xlHoja1.Cells(pnFila, 1) = xlHoja1.Cells(pnFila, 1) & " (Incluye Cartas Fianza ) "
        Else
            xlHoja1.Cells(pnFila, 1) = xlHoja1.Cells(pnFila, 1) & " (No Incluye Cartas Fianza )"
        End If
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0C0FF

        pnFila = pnFila + 1
        ' Limpiar
        For lnIndice = 2 To 14
            lsTotSubTipo(lnIndice) = ""
        Next lnIndice

        lsSQL = " Select distinct(Substring(CA.cCtaCod,6,3)) cSubT FROM ColocCalifProv CA " _
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
                lsSubTDesc = rsSubT!cSubT
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

                lsSQL = "SELECT Substring(CA.cLineaCred,7,3) as cSubtipo, " _
                      & "Substring(CA.cLineaCred,1,4) as cFF, substring(CA.cCtaCod,1,2) as cAgencia, " _
                      & "Count (CA.cCtaCod) AS NroCred, " _
                      & "Isnull(Sum (CA.nSaldoCap), 0) AS SaldoCap, " _
                      & "Count( CASE WHEN CA.cCalGen = '0' THEN (CA.cCtaCod) END ) AS NroProv0, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '0' THEN (CA.nSaldoCap) END ), 0) AS Prov0, " _
                      & "Count( CASE WHEN CA.cCalGen = '1' THEN (CA.cCtaCod) END ) AS NroProv1, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '1' THEN (CA.nSaldoCap) END ), 0) AS Prov1, " _
                      & "Count(CASE WHEN CA.cCalGen = '2' THEN (CA.cCtaCod) END ) AS NroProv2, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '2' THEN (CA.nSaldoCap) END ), 0) AS Prov2, " _
                      & "Count(CASE WHEN CA.cCalGen = '3' THEN (CA.cCtaCod) END ) AS NroProv3, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '3' THEN (CA.nSaldoCap) END ), 0) AS Prov3, " _
                      & "Count(CASE WHEN CA.cCalGen = '4' THEN (CA.cCtaCod) END ) AS NroProv4, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '4' THEN (CA.nSaldoCap) END ), 0) AS Prov4, " _
                      & "Sum(IsNull(ca.nProvision, 0)) As Provision "
                If lbMesActual = True Then
                    lsSQL = lsSQL & " FROM ColocCalifProv ca Where "
                Else
                    lsSQL = lsSQL & " FROM ColocCalifProvTotal ca"
                    lsSQL = lsSQL & " Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
                End If
                lsSQL = lsSQL & " " _
                      & " CA.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2101,2104,2106,2107)   " _
                      & "AND Substring(CA.cLineaCred,9,1) = '" & Trim(Str(lnMoneda)) & "' " _
                      & "AND CA.nSaldoCap > 0  " _
                      & "AND Substring(CA.cLineaCred,7,1) ='" & Trim(Str(lnTipo)) & "' " _
                      & "AND Substring(CA.cLineaCred,7,3) ='" & Trim(Str(lnSubTipo)) & "' " _
                      & "AND CA.nGarant = " & lnRango & " " _
                      & lsCFianza
                lsSQL = lsSQL & "GROUP BY  Substring(CA.cLineaCred,7,3), " _
                     & "Substring(CA.cLineaCred,1,4), Substring(CA.cCtaCod,1,2) " _
                     & "ORDER By Substring(CA.cLineaCred,7,3), " _
                      & "Substring(CA.cLineaCred,1,4), Substring(CA.cCtaCod,1,2) "


                Set rs = Co.CargaRecordSet(lsSQL)
                
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
                    lsSubTipo = rs!cSubtipo
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
xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
End Sub


'Modificado x ARCV 26-07-2006
Public Sub Genera_Reporte178210(ByVal TxtCambio As Double, ByVal dFecha As Date, _
                ByVal pCartaFianza As Boolean, ByVal psServer As String, _
                ByVal dFecData As String, _
                Optional ByVal psNomCmac As String, _
                Optional ByVal psCodCMAC As String)

Dim fs As Scripting.FileSystemObject
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
'Rep_HojaWProvision808 = True
lnLinea = 1
Dim lbMesActual As Boolean
Dim lsArchivo As String
Dim lbExcel As Boolean
Dim J As Integer
Dim lsNombre As String

Dim Co As DConecta
Set Co = New DConecta

'ARCV 27-07-2006
Dim lsRangoIni(14) As String
Dim lsRangoFin(14) As String

Dim lnTotFF(14) As Double
Dim lnTotRango(14) As Double
Dim lnTotSubTipo(14) As Double
Dim lnTotTipo(14) As Double
Dim lnTotCartera(14) As Double
'--------------

Co.AbreConexion

lbCambio = False
lsArchivo = "Anexo5_" & Format(dFecha, "yyyymm") & "_HojaW_808" & IIf(pCartaFianza = True, "", "_SinCF") & ".XLS"
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If
lbExcel = True

If dFecha = dFecData Then
    lbMesActual = True
Else
    lbMesActual = False
End If
If pCartaFianza = False Then
    'lsCFianza = " And Substring(CA.cCtaCod,7,2) <> '21' "
    lsCFianza = " And Substring(C.cTpoCredCod,2,2) <> '81' " ' By BRGO 20100609 81:Nuevo codigo CARTA FIANZA
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
    Call ImpHojaWProvisionCab(pnFila, Trim(Str(lnMoneda)), dFecha)
    pnFila = 8
    ' Cuerpo del Reporte
    
    ' Limpiar Tot Cartera / Tipo / SubTipo
    For lnIndice = 2 To 14
        lsTotCartera(lnIndice) = ""
        lsTotTipo(lnIndice) = ""
        lsTotSubTipo(lnIndice) = ""
        
        lnTotCartera(lnIndice) = 0
        lnTotTipo(lnIndice) = 0
        lnTotSubTipo(lnIndice) = 0
    Next lnIndice
    
    For lnTipo = 1 To 8
        
        ' By BRGO 20100609 BASILEA II
        If Trim(Str(lnTipo)) = "1" Then
            lsTipoDesc = "I.- CREDITOS CORPORATIVOS"
        ElseIf Trim(Str(lnTipo)) = "2" Then
            lsTipoDesc = "II.- CREDITOS GRANDE EMPRESA"
        ElseIf Trim(Str(lnTipo)) = "3" Then
            lsTipoDesc = "III.- CREDITOS MEDIANA EMPRESA"
        ElseIf Trim(Str(lnTipo)) = "4" Then
            lsTipoDesc = "IV.- CREDITOS PEQUEÑA EMPRESA"
        ElseIf Trim(Str(lnTipo)) = "5" Then
            lsTipoDesc = "V.- CREDITOS PEQUEÑA MICROEMPRESA"
        ElseIf Trim(Str(lnTipo)) = "6" Then
            lsTipoDesc = "VI.- CREDITOS CONSUMO REVOLVENTE"
        ElseIf Trim(Str(lnTipo)) = "7" Then
            lsTipoDesc = "VII.- CREDITOS CONSUMO NO REVOLVENTE"
        ElseIf Trim(Str(lnTipo)) = "8" Then
            lsTipoDesc = "VIII.- CREDITOS HIPOTECARIO"
        End If
        ' End BRGO
        
        xlHoja1.Cells(pnFila, 1) = lsTipoDesc
        If pCartaFianza = True Then
            xlHoja1.Cells(pnFila, 1) = xlHoja1.Cells(pnFila, 1) & " (Incluye Cartas Fianza ) "
        Else
            xlHoja1.Cells(pnFila, 1) = xlHoja1.Cells(pnFila, 1) & " (No Incluye Cartas Fianza )"
        End If
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0C0FF
        
        pnFila = pnFila + 1
        ' Limpiar
        For lnIndice = 2 To 14
            lsTotSubTipo(lnIndice) = ""
            
            lnTotSubTipo(lnIndice) = 0
        Next lnIndice
        
        'By BRGO 20100609 BASILEA II
        lsSQL = " Select distinct(C.cTpoCredCod) cSubT FROM ColocCalifProv CA INNER JOIN Colocaciones C ON C.cCtaCod = CA.cCtaCod " _
              & " Where substring(C.cTpoCredCod,1,1) = '" & lnTipo & "' AND Substring(CA.cCtaCod,9,1) = '" & Trim(Str(lnMoneda)) & "' " & lsCFianza & " ORDER BY C.cTpoCredCod "
        'End BRGO
        Set rsSubT = Co.CargaRecordSet(lsSQL)
        'rsSubT.CursorLocation = adUseClient
        'rsSubT.Open lsSQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
        'rsSubT.ActiveConnection = Nothing
        
        Do While Not rsSubT.EOF
            lnSubTipo = Trim(Str(rsSubT!cSubT))
        'For lnSubTipo = 1 To 2
            If Mid(rsSubT!cSubT, 2, 2) = "21" Then
                lsSubTDesc = "CARTA FIANZA"
            Else
                'lsSQL = " Select * from " & gcCentralCom & "Tablacod " _
                '      & " Where cCodtab like '6" & lnTipo & "__' And substring(cCodTab,2,3) = '" & Trim(Str(rsSubT!cSubT)) & "' "
                'rsDesc.Open lsSQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                'lsSubTDesc = Trim(IIf(IsNull(rsDesc!cNomTab), "", rsDesc!cNomTab))
                'rsDesc.Close
                lsSubTDesc = "Producto - " & rsSubT!cSubT
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
                
                lnTotRango(lnIndice) = 0
            Next lnIndice
            
            For lnRango = 1 To 4
                'By BRGO 20100609 BASILEA II
                lsSQL = "SELECT Substring(CA.cLineaCred,7,3) as cSubtipo, " _
                      & "Substring(CA.cLineaCred,1,4) as cFF, C.cAgeCodAct as cAgencia, " _
                      & "Count (CA.cCtaCod) AS NroCred, " _
                      & "Isnull(Sum (CA.nSaldoCap), 0) AS SaldoCap, " _
                      & "Count( CASE WHEN CA.cCalGen = '0' THEN (CA.cCtaCod) END ) AS NroProv0, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '0' THEN (CA.nSaldoCap) END ), 0) AS Prov0, " _
                      & "Count( CASE WHEN CA.cCalGen = '1' THEN (CA.cCtaCod) END ) AS NroProv1, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '1' THEN (CA.nSaldoCap) END ), 0) AS Prov1, " _
                      & "Count(CASE WHEN CA.cCalGen = '2' THEN (CA.cCtaCod) END ) AS NroProv2, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '2' THEN (CA.nSaldoCap) END ), 0) AS Prov2, " _
                      & "Count(CASE WHEN CA.cCalGen = '3' THEN (CA.cCtaCod) END ) AS NroProv3, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '3' THEN (CA.nSaldoCap) END ), 0) AS Prov3, " _
                      & "Count(CASE WHEN CA.cCalGen = '4' THEN (CA.cCtaCod) END ) AS NroProv4, " _
                      & "Isnull(Sum (CASE WHEN CA.cCalGen = '4' THEN (CA.nSaldoCap) END ), 0) AS Prov4, " _
                      & "Sum(IsNull(ca.nProvision, 0)) As Provision "
                If lbMesActual = True Then
                    lsSQL = lsSQL & " FROM ColocCalifProv CA INNER JOIN Colocaciones C ON C.cCtaCod = CA.cCtaCod Where "
                Else
                    lsSQL = lsSQL & " FROM ColocCalifProvTotal CA INNER JOIN Colocaciones C ON C.cCtaCod = CA.cCtaCod "
                    lsSQL = lsSQL & " Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
                End If
                lsSQL = lsSQL & " " _
                      & " CA.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107)   " _
                      & "AND Substring(CA.cLineaCred,5,1) = '" & Trim(Str(lnMoneda)) & "' " _
                      & "AND CA.nSaldoCap > 0  " _
                      & "AND Substring(C.cTpoCredCod,1,1) ='" & Trim(Str(lnTipo)) & "' " _
                      & "AND C.cTpoCredCod ='" & Trim(Str(lnSubTipo)) & "' " _
                      & "AND CA.nGarant = " & lnRango & " " _
                      & lsCFianza
                lsSQL = lsSQL & "GROUP BY  Substring(CA.cLineaCred,7,3), " _
                      & "Substring(CA.cLineaCred,1,4), C.cAgeCodAct " _
                      & "ORDER By Substring(CA.cLineaCred,7,3), " _
                      & "Substring(CA.cLineaCred,1,4), C.cAgeCodAct "
                'End BRGO
                Set rs = Co.CargaRecordSet(lsSQL)
                
                'dbCmactCentral.CommandTimeout = 1200
                'rs.CursorLocation = adUseClient
                'rs.Open lsSQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
                'rs.ActiveConnection = Nothing
                
                lsFF = ""
                
                If Not (rs.BOF And rs.EOF) Then
                    ' Imprimir Encabezado
                    If lnRango = 1 Then
                        xlHoja1.Cells(pnFila, 1) = "Garantias Preferida" '"Garantias AutoLiquidable"
                    ElseIf lnRango = 2 Then
                        xlHoja1.Cells(pnFila, 1) = "Garantias AutoLiquidable" '"Garantias Muy Rapida Realizacion"
                    ElseIf lnRango = 3 Then
                        xlHoja1.Cells(pnFila, 1) = "Garantias No Preferida" '"Garantias Preferida"
                    'ElseIf lnRango = 4 Then
                    '    xlHoja1.Cells(pnFila, 1) = "Garantias No Preferida"
                    End If
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 1)).Font.Bold = True
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HFF00C
                    pnFila = pnFila + 1
                    ' Limpiar
                    For lnIndice = 2 To 14
                        lsTotFF(lnIndice) = ""
                        
                        lnTotFF(lnIndice) = 0
                    Next lnIndice
                End If
                
                Do While Not rs.EOF
                    
                    If lsFF <> rs!cFF Then ' Otro Fuente Financ
                        'Imprimir Total de anterior Fuente Fin
                        If lsFF <> "" Then
                            'Imprimir totales FF
                            xlHoja1.Cells(pnFila, 1) = "SubTotal F.F. " & Trim(lsFF)
                            For lnIndice = 2 To 14
                                xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = lnTotFF(lnIndice)  '"=SUM(" & lsTotFF(lnIndice) & ")"
                            Next lnIndice
                            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                            'Asigno los valores al Rango
                            For lnIndice = 2 To 14
                                lsTotRango(lnIndice) = lsTotRango(lnIndice) & IIf(Len(lsTotRango(lnIndice)) > 0, "+", "") & lsTotFF(lnIndice)
                                
                                lnTotRango(lnIndice) = lnTotRango(lnIndice) + lnTotFF(lnIndice) 'ARCV 27-07
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
                            
                            lnTotFF(lnIndice) = 0
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
                        
                        lnTotFF(lnIndice) = lnTotFF(lnIndice) + CDbl(xlHoja1.Cells(pnFila, lnIndice))
                    Next lnIndice
                    pnFila = pnFila + 1
                    'lsTipo = rs!cTipo
                    lsSubTipo = rs!cSubtipo
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
                        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = lnTotFF(lnIndice) '"=SUM(" & lsTotFF(lnIndice) & ")"
                    Next lnIndice
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                    'Asigno los valores al Rango
                    For lnIndice = 2 To 14
                        lsTotRango(lnIndice) = lsTotRango(lnIndice) & IIf(Len(lsTotRango(lnIndice)) > 0, "+", "") & lsTotFF(lnIndice)
                        
                        lnTotRango(lnIndice) = lnTotRango(lnIndice) + lnTotFF(lnIndice)
                    Next lnIndice
                    pnFila = pnFila + 1
                    
                    'Imprimir totales Rango
                    xlHoja1.Cells(pnFila, 1) = "SubTotal Garantia  " & Trim(lsRango)
                    For lnIndice = 2 To 14
                        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = lnTotRango(lnIndice) '"=SUM(" & lsTotRango(lnIndice) & ")"
                    Next lnIndice
                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                    'xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = RGB(10, 100, 10)
                    'Asigno los valores al SubTipo
                    For lnIndice = 2 To 14
                        'lsTotSubTipo(lnIndice) = lsTotSubTipo(lnIndice) + IIf(Len(lsTotSubTipo(lnIndice)) > 0, "+", "") + lsTotRango(lnIndice)
                        lsTotSubTipo(lnIndice) = lsTotSubTipo(lnIndice) + IIf(Len(lsTotSubTipo(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
                        
                        lnTotSubTipo(lnIndice) = lnTotSubTipo(lnIndice) + CDbl(xlHoja1.Cells(pnFila, lnIndice))
                    Next lnIndice
                    pnFila = pnFila + 1
                End If
                
                ' Limpio la variable FF
                For lnIndice = 2 To 14
                    lsTotFF(lnIndice) = ""
                    
                    lnTotFF(lnIndice) = 0
                Next lnIndice
                ' Limpio la variable Rango
                For lnIndice = 2 To 14
                    lsTotRango(lnIndice) = ""
                    
                    lnTotRango(lnIndice) = 0
                Next lnIndice
                
            Next lnRango
            
            'Imprimir totales SubTipo
            If Trim(lsTotSubTipo(2)) <> "" Then
                xlHoja1.Cells(pnFila, 1) = "Total " & lsSubTDesc
                For lnIndice = 2 To 14
                    xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = lnTotSubTipo(lnIndice) '"=SUM(" & lsTotSubTipo(lnIndice) & ")" ' "=SUM(" & lsTotSubTipo(lnIndice) & ")" ' ARCV 26-07-2006
                    
                Next lnIndice
                xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
                'Asigno los valores al tipo
                For lnIndice = 2 To 14
                    'lsTotTipo(lnIndice) = lsTotTipo(lnIndice) & IIf(Len(lsTotTipo(lnIndice)) > 0, "+", "") & lsTotSubTipo(lnIndice)
                    lsTotTipo(lnIndice) = lsTotTipo(lnIndice) & IIf(Len(lsTotTipo(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
                    lnTotTipo(lnIndice) = lnTotTipo(lnIndice) + CDbl(xlHoja1.Cells(pnFila, lnIndice))
                Next lnIndice
                pnFila = pnFila + 1
            End If
            
            ' Limpio la variable Subtipo
            For lnIndice = 2 To 14
                lsTotSubTipo(lnIndice) = ""
                
                lnTotSubTipo(lnIndice) = 0
            Next lnIndice
            
            rsSubT.MoveNext
         Loop
         
        rsSubT.Close
        'Imprimir totales Tipo
        If Trim(lsTotTipo(2)) <> "" Then
            xlHoja1.Cells(pnFila, 1) = "Tot." & Mid(lsTipoDesc, 4)
            For lnIndice = 2 To 14
                xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = lnTotTipo(lnIndice) '"=SUM(" & lsTotTipo(lnIndice) & ")"
            Next lnIndice
            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
            'Asigno los valores a la Cartera
            For lnIndice = 2 To 14
                'lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & lsTotTipo(lnIndice)
                lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
                
                lnTotCartera(lnIndice) = lnTotCartera(lnIndice) + CDbl(xlHoja1.Cells(pnFila, lnIndice))
            Next lnIndice
            pnFila = pnFila + 2
        End If
        
        ' Limpio la variable tipo
        For lnIndice = 2 To 14
            lsTotTipo(lnIndice) = ""
            
            lnTotTipo(lnIndice) = 0
        Next lnIndice
        
    Next lnTipo
    
    'Imprimir totales Tipo
    If Trim(lsTotTipo(2)) <> "" Then
        xlHoja1.Cells(pnFila, 1) = "Total " & lsTipoDesc
        For lnIndice = 2 To 14
            xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = lnTotTipo(lnIndice)  '"=SUM(" & lsTotTipo(lnIndice) & ")"
        Next lnIndice
        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
        'Asigno los valores a la Cartera
        For lnIndice = 2 To 14
            'lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & lsTotTipo(lnIndice)
            lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
            
            lnTotCartera(lnIndice) = lnTotCartera(lnIndice) + CDbl(xlHoja1.Cells(pnFila, lnIndice))
        Next lnIndice
        pnFila = pnFila + 2
    End If
    
    ' Limpio la variable tipo
    For lnIndice = 2 To 14
        lsTotTipo(lnIndice) = ""
        
        lnTotTipo(lnIndice) = 0
    Next lnIndice
    
    
    ' ********************
    
    'Imprime Total Cartera
    pnFila = pnFila + 1
    xlHoja1.Cells(pnFila, 1) = "Total CARTERA"
    For lnIndice = 2 To 14
        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = lnTotCartera(lnIndice)  '"=SUM(" & lsTotCartera(lnIndice) & ")"
    Next lnIndice
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0FFFF
    
    'CuadroExcel 1, 6, 14, pnFila, True
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(pnFila, 14)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(pnFila, 14)).Cells.Borders.LineStyle = xlInside

Next lnMoneda

'*******
xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo
'Cierra el libro de trabajo
xlLibro.Close
' Cierra Microsoft Excel con el método Quit.
xlAplicacion.Quit
'Libera los objetos.
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
lbExcel = False

MsgBox "Se genero el Anexo. Verificar en la carpeta Spooler", vbInformation, "AVISO"

End Sub

Public Sub Genera_Reporte178209(ByVal TxtCambio As Double, ByVal dFecha As Date, _
                        ByVal nCF As Integer, ByVal psServer As String, _
                        ByVal dFecData As String, _
                        ByVal psFFinan As String, _
                        Optional ByVal psNomCmac As String, _
                        Optional ByVal psCodCMAC As String, _
                        Optional ByVal pnTipoConsulta As Integer = 0)


Dim sql As String
Dim rs As New ADODB.Recordset
Dim fs As New Scripting.FileSystemObject
Dim tc As Double
Dim i As Integer
Dim J As Integer
Dim k As Integer
Dim lsNomHoja As String
Dim lbExisteHoja As Boolean
'ALPA 20090217**********************************
Dim MatrixAnexo5D() As String
ReDim Preserve MatrixAnexo5D(1 To 3, 1 To 8)
'***********************************************
Dim lbMesActual As Boolean
Dim lbIncluyeCF As Boolean
Dim lsArchivo As String
Dim lsServConsol As String

Dim RCC As DRCC
Set RCC = New DRCC
lsServConsol = RCC.ServConsol(psServer)
Set RCC = Nothing


Dim Co As DConecta
Set Co = New DConecta
Co.AbreConexion

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
If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & lsArchivo)
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

 xlHoja1.Range("A1") = "ANEXO N° 5 D"
 xlHoja1.Range("A3") = "EMPRESA : " & psNomCmac
 xlHoja1.Range("F3") = "CODIGO : " & psCodCMAC
 xlHoja1.Range("A5") = "INFORME DE CLASIFICACION DE DEUDORES DE LA CARTERA DE CREDITOS, QUE RESPALDAN FINANCIAMIENTOS O LÍNEAS DE CRÉDITOS 1/"

 xlHoja1.Range("A7") = "AL " & Format(dFecha, "DD") & " DE  " & UCase(Format(dFecha, "MMMM")) & " DEL  " & Format(dFecha, "YYYY")
 xlHoja1.Range("A8") = "(En nuevos soles)"

 xlHoja1.Range("A1").ColumnWidth = 42
 xlHoja1.Range("B1").ColumnWidth = 15
 xlHoja1.Range("C1").ColumnWidth = 15
 xlHoja1.Range("D1").ColumnWidth = 15
 xlHoja1.Range("E1").ColumnWidth = 15
 xlHoja1.Range("F1").ColumnWidth = 15
 xlHoja1.Range("G1").ColumnWidth = 16

 xlHoja1.Application.ActiveWindow.Zoom = 75
 xlHoja1.Range("A1:G1").MergeCells = True
 xlHoja1.Range("A5:G5").MergeCells = True
 xlHoja1.Range("A7:G7").MergeCells = True
 xlHoja1.Range("A8:G8").MergeCells = True
 
 xlHoja1.Range("A1:G1").HorizontalAlignment = xlCenter
 xlHoja1.Range("A5:G5").HorizontalAlignment = xlCenter
 xlHoja1.Range("A7:G7").HorizontalAlignment = xlCenter
 xlHoja1.Range("A8:G8").HorizontalAlignment = xlCenter
 
 xlHoja1.Range("A1:G10").Font.Bold = True

 xlHoja1.Range("A10") = "A. MONTO DE LOS CREDITOS Y CONTINGENTES 2/"
 xlHoja1.Range("A10").WrapText = True
 xlHoja1.Range("B10") = "NORMAL"
 xlHoja1.Range("C10") = "CPP"
 xlHoja1.Range("D10") = "DEFICIENTE"
 xlHoja1.Range("E10") = "DUDOSO"
 xlHoja1.Range("F10") = "PERDIDA"
 xlHoja1.Range("G10") = "TOTAL"
 xlHoja1.Range("A10:G10").Font.Bold = True

 xlHoja1.Range("A11") = " Corporativos"
 xlHoja1.Range("A12") = " Grandes empresas"
 xlHoja1.Range("A13") = " Medianas empresas"
 xlHoja1.Range("A14") = " Pequeñas empresas"
 xlHoja1.Range("A15") = " Microempresas"
 xlHoja1.Range("A16") = " Consumo revolvente"
 xlHoja1.Range("A17") = " Consumo no revolvente"
 xlHoja1.Range("A18") = " Hipotecario para vivienda"

 xlHoja1.Range("A19") = "B. PROVISIONES CONSTITUIDAS 3/"
 xlHoja1.Range("B19") = "NORMAL"
 xlHoja1.Range("C19") = "CPP"
 xlHoja1.Range("D19") = "DEFICIENTE"
 xlHoja1.Range("E19") = "DUDOSO"
 xlHoja1.Range("F19") = "PERDIDA"
 xlHoja1.Range("G19") = "TOTAL"
 xlHoja1.Range("A19:G19").Font.Bold = True

 xlHoja1.Range("A20") = " Corporativos"
 xlHoja1.Range("A21") = " Grandes empresas"
 xlHoja1.Range("A22") = " Medianas empresas"
 xlHoja1.Range("A23") = " Pequeñas empresas"
 xlHoja1.Range("A24") = " Microempresas"
 xlHoja1.Range("A25") = " Consumo revolvente"
 xlHoja1.Range("A26") = " Consumo no revolvente"
 xlHoja1.Range("A27") = " Hipotecario para vivienda"

 xlHoja1.Range("A28") = "C. PROVISIONES REQUERIDAS 4/"
 xlHoja1.Range("B28") = "NORMAL"
 xlHoja1.Range("C28") = "CPP"
 xlHoja1.Range("D28") = "DEFICIENTE"
 xlHoja1.Range("E28") = "DUDOSO"
 xlHoja1.Range("F28") = "PERDIDA"
 xlHoja1.Range("G28") = "TOTAL"
 xlHoja1.Range("A28:G28").Font.Bold = True

 xlHoja1.Range("A29") = " Corporativos"
 xlHoja1.Range("A30") = " Grandes empresas"
 xlHoja1.Range("A31") = " Medianas empresas"
 xlHoja1.Range("A32") = " Pequeñas empresas"
 xlHoja1.Range("A33") = " Microempresas"
 xlHoja1.Range("A34") = " Consumo revolvente"
 xlHoja1.Range("A35") = " Consumo no revolvente"
 xlHoja1.Range("A36") = " Hipotecario para vivienda"

 xlHoja1.Range("A37") = "D. SUPERAVIT (DEFICIT) DE PROVISIONES 5/"
 xlHoja1.Range("A37").WrapText = True
 xlHoja1.Range("B37") = "NORMAL"
 xlHoja1.Range("C37") = "CPP"
 xlHoja1.Range("D37") = "DEFICIENTE"
 xlHoja1.Range("E37") = "DUDOSO"
 xlHoja1.Range("F37") = "PERDIDA"
 xlHoja1.Range("G37") = "TOTAL"
 xlHoja1.Range("A37:G37").Font.Bold = True

 xlHoja1.Range("A38") = " Corporativos"
 xlHoja1.Range("A39") = " Grandes empresas"
 xlHoja1.Range("A40") = " Medianas empresas"
 xlHoja1.Range("A41") = " Pequeñas empresas"
 xlHoja1.Range("A42") = " Microempresas"
 xlHoja1.Range("A43") = " Consumo revolvente"
 xlHoja1.Range("A44") = " Consumo no revolvente"
 xlHoja1.Range("A45") = " Hipotecario para vivienda"

 xlHoja1.Range("A46") = "TOTAL"

 xlHoja1.Range("A47") = "E. MONTO DE LOS CREDITOS Y CONTINGENTES CEDIDOS EN EJECUCIÓN 6/"
 xlHoja1.Range("A47").WrapText = True
 
 xlHoja1.Range("B47") = "NORMAL"
 xlHoja1.Range("C47") = "CPP"
 xlHoja1.Range("D47") = "DEFICIENTE"
 xlHoja1.Range("E47") = "DUDOSO"
 xlHoja1.Range("F47") = "PERDIDA"
 xlHoja1.Range("G47") = "TOTAL"
 xlHoja1.Range("A47:G47").Font.Bold = True

 xlHoja1.Range("A48") = " Corporativos"
 xlHoja1.Range("A49") = " Grandes empresas"
 xlHoja1.Range("A50") = " Medianas empresas"
 xlHoja1.Range("A51") = " Pequeñas empresas"
 xlHoja1.Range("A52") = " Microempresas"
 xlHoja1.Range("A53") = " Consumo revolvente"
 xlHoja1.Range("A54") = " Consumo no revolvente"
 xlHoja1.Range("A55") = " Hipotecario para vivienda"


 xlHoja1.Range("B61").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("B61").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("B61:B62").HorizontalAlignment = xlCenter
 xlHoja1.Range("B61") = "GERENTE"
 xlHoja1.Range("B62") = "GENERAL"

 xlHoja1.Range("D61").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("D61").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("D61:D62").HorizontalAlignment = xlCenter
 xlHoja1.Range("D61") = "CONTADOR"
 xlHoja1.Range("D62") = "GENERAL"

 xlHoja1.Range("F61").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("F61").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("F61:F62").HorizontalAlignment = xlCenter
 xlHoja1.Range("F61") = "FUNCIONARIO"
 xlHoja1.Range("F62") = "RESPONSABLE"
 
 xlHoja1.Range("A10:G55").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A10:G55").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A10:G55").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A10:G55").Borders(xlEdgeBottom).Weight = xlThin
 xlHoja1.Range("A10:G55").Borders(xlEdgeLeft).LineStyle = xlContinuous
 xlHoja1.Range("A10:G55").Borders(xlEdgeLeft).Weight = xlThin
 xlHoja1.Range("A10:G55").Borders(xlEdgeRight).LineStyle = xlContinuous
 xlHoja1.Range("A10:G55").Borders(xlEdgeRight).Weight = xlThin
 xlHoja1.Range("A10:G55").Borders(xlInsideVertical).LineStyle = xlContinuous
 xlHoja1.Range("A10:G55").Borders(xlInsideVertical).Weight = xlThin
 xlHoja1.Range("A10:G10").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A10:G10").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A19:G19").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A19:G19").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A19:G19").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A19:G19").Borders(xlEdgeBottom).Weight = xlThin
 xlHoja1.Range("A28:G28").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A28:G28").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A28:G28").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A28:G28").Borders(xlEdgeBottom).Weight = xlThin
 xlHoja1.Range("A37:G37").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A37:G37").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A37:G37").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A37:G37").Borders(xlEdgeBottom).Weight = xlThin
 
 xlHoja1.Range("A46:G46").Borders(xlEdgeTop).LineStyle = xlContinuous
 xlHoja1.Range("A46:G46").Borders(xlEdgeTop).Weight = xlThin
 xlHoja1.Range("A46:G46").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A46:G46").Borders(xlEdgeBottom).Weight = xlThin
 xlHoja1.Range("A47:G47").Borders(xlEdgeBottom).LineStyle = xlContinuous
 xlHoja1.Range("A47:G47").Borders(xlEdgeBottom).Weight = xlThin

'-------- BRGO BASILEA II
sql = " select substring(ca.cTpoCredCod,1,1) cta, cCalGen ,"
sql = sql & " Isnull(sum(case when substring(ca.cCtaCod,9,1) = '1' then ca.nSaldoCap"
sql = sql & "                 when substring(ca.cCtaCod,9,1) = '2' then ca.nSaldoCap * " & tc
sql = sql & "  end ),0) SaldoCap,"
sql = sql & " Isnull(sum(case when substring(ca.cCtaCod,9,1) = '1'  then ca.nProvision"
sql = sql & "                 when substring(ca.cCtaCod,9,1) = '2'  then ca.nProvision * " & tc
sql = sql & "  end ),0) ProvConstituida,"
sql = sql & " Isnull(sum(case when substring(ca.cCtaCod,9,1) = '1'  then ca.nProvisionRCC"
sql = sql & "                 when substring(ca.cCtaCod,9,1) = '2'  then ca.nProvisionRCC * " & tc
sql = sql & "  end ),0) ProvConstituidaRCC,"
sql = sql & " Isnull(sum(case when substring(ca.cCtaCod,9,1) = '1'  then ca.nProvisionProciclica"
sql = sql & "                 when substring(ca.cCtaCod,9,1) = '2'  then ca.nProvisionProciclica * " & tc
sql = sql & "  end ),0) ProvConstituidaProciclica"
'*******************************************************************************

If lbMesActual = True Then
    sql = sql & " FROM ColocCalifProv ca Where "
Else
    sql = sql & " FROM ColocCalifProvTotal ca "
    sql = sql & " Where datediff(d,dfecha,'" & Format(dFecha, "YYYY/MM/DD") & "')=0  And "
End If
sql = sql & " nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107) "
If pnTipoConsulta = 0 Then
    sql = sql & " And Substring(ca.cLineaCred,1,2) IN ('03','02')"
Else
    sql = sql & " And Substring(ca.cLineaCred,1,2) IN ('02')"
End If
'-------------
If lbIncluyeCF = False Then
    sql = sql & " And substring(ca.cTpoCredCod,2,2) <> '81' "
End If
sql = sql & " Group by substring(ca.cTpoCredCod,1,1), cCalGen"
sql = sql & " Order by substring(ca.cTpoCredCod,1,1), cCalGen"

Set rs = Co.CargaRecordSet(sql)

While Not rs.EOF
     Select Case rs!cCalGen
          Case "0": xlHoja1.Range("B" & 10 + rs!Cta) = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("B" & 19 + rs!Cta) = Format(rs!ProvConstituida + rs!ProvConstituidaRCC + rs!ProvConstituidaProciclica, "#,###.00")
                          MatrixAnexo5D(1, rs!Cta) = rs!ProvConstituidaRCC
                          MatrixAnexo5D(2, rs!Cta) = rs!ProvConstituidaProciclica
          Case "1": xlHoja1.Range("C" & 10 + rs!Cta) = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("C" & 19 + rs!Cta) = Format(rs!ProvConstituida, "#,###.00")

          Case "2": xlHoja1.Range("D" & 10 + rs!Cta) = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("D" & 19 + rs!Cta) = Format(rs!ProvConstituida, "#,###.00")

          Case "3": xlHoja1.Range("E" & 10 + rs!Cta) = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("E" & 19 + rs!Cta) = Format(rs!ProvConstituida, "#,###.00")

          Case "4": xlHoja1.Range("F" & 10 + rs!Cta) = Format(rs!SaldoCap, "#,###.00")
                          xlHoja1.Range("F" & 19 + rs!Cta) = Format(rs!ProvConstituida, "#,###.00")
    End Select
    rs.MoveNext
Wend
 
'-----------------------------------------------
' Calculos
'-----------------------------------------------

 For i = 1 To 8
    xlHoja1.Range("G" & 10 + i).Formula = "= SUM(A" & 10 + i & ":F" & 10 + i & ")"
    xlHoja1.Range("G" & 19 + i).Formula = "= SUM(A" & 19 + i & ":F" & 19 + i & ")"
    xlHoja1.Range("G" & 28 + i).Formula = "= SUM(A" & 28 + i & ":F" & 28 + i & ")"
    xlHoja1.Range("G" & 37 + i).Formula = "= SUM(A" & 37 + i & ":F" & 37 + i & ")"
    xlHoja1.Range("G" & 47 + i).Formula = "= SUM(A" & 47 + i & ":F" & 47 + i & ")"
 Next
 xlHoja1.Range("B46").Formula = "= SUM(B38:B45)"
 xlHoja1.Range("C46").Formula = "= SUM(C38:C45)"
 xlHoja1.Range("D46").Formula = "= SUM(D38:D45)"
 xlHoja1.Range("E46").Formula = "= SUM(E38:E45)"
 xlHoja1.Range("F46").Formula = "= SUM(F38:F45)"
 xlHoja1.Range("G46").Formula = "= SUM(B46:F46)"

 xlHoja1.Range("B29").Formula = "= (B11)*0.7% + " & CDbl(IIf(MatrixAnexo5D(2, 1) = "", 0, MatrixAnexo5D(2, 1)))
 xlHoja1.Range("B30").Formula = "= (B12)*0.7% + " & CDbl(IIf(MatrixAnexo5D(2, 1) = "", 0, MatrixAnexo5D(2, 1)))
 xlHoja1.Range("B31").Formula = "= (B13)*1%  + " & CDbl(IIf(MatrixAnexo5D(2, 2) = "", 0, MatrixAnexo5D(2, 2)))
 xlHoja1.Range("B32").Formula = "= (B14)*1%  + " & CDbl(IIf(MatrixAnexo5D(2, 2) = "", 0, MatrixAnexo5D(2, 2)))
 xlHoja1.Range("B33").Formula = "= (B15)*1% + " & CDbl(IIf(MatrixAnexo5D(2, 3) = "", 0, MatrixAnexo5D(2, 3)))
 xlHoja1.Range("B34").Formula = "= (B16)*1% + " & CDbl(IIf(MatrixAnexo5D(2, 3) = "", 0, MatrixAnexo5D(2, 3)))
 xlHoja1.Range("B35").Formula = "= (B17)*1% + " & CDbl(IIf(MatrixAnexo5D(2, 3) = "", 0, MatrixAnexo5D(2, 3)))
 xlHoja1.Range("B36").Formula = "= (B18)*0.7% + " & CDbl(IIf(MatrixAnexo5D(2, 4) = "", 0, MatrixAnexo5D(2, 4)))
 
'***********************************************************

For i = 1 To 8
    xlHoja1.Range("C" & 28 + i & "").Formula = "= (C" & 19 + i & ") "
    xlHoja1.Range("D" & 28 + i & "").Formula = "= (D" & 19 + i & ") "
    xlHoja1.Range("E" & 28 + i & "").Formula = "= (E" & 19 + i & ") "
    xlHoja1.Range("F" & 28 + i & "").Formula = "= (F" & 19 + i & ") "
Next

 For i = 20 To 27
    xlHoja1.Range("B" & 18 + i).Formula = "= B" & i & "-B" & i + 9
    xlHoja1.Range("C" & 18 + i).Formula = "= C" & i & "-C" & i + 9
    xlHoja1.Range("D" & 18 + i).Formula = "= D" & i & "-D" & i + 9
    xlHoja1.Range("E" & 18 + i).Formula = "= E" & i & "-E" & i + 9
    xlHoja1.Range("F" & 18 + i).Formula = "= F" & i & "-F" & i + 9
 Next
 '--- formato
 '----------

 xlHoja1.Range("B11:G18").NumberFormat = "#,###,##0.00"
 xlHoja1.Range("B20:G27").NumberFormat = "#,###,##0.00"
 xlHoja1.Range("B29:G36").NumberFormat = "#,###,##0.00"
 xlHoja1.Range("B38:G46").NumberFormat = "#,###,##0.00"
 xlHoja1.Range("B48:G55").NumberFormat = "#,###,##0.00"
 

xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo
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
                        Optional ByVal psCodCMAC As String = "", Optional ByVal psCodAnexo As String = "")

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
    xlHoja1.Cells(pnFila, 1) = "EMPRESA :  " & psNomCmac: xlHoja1.Cells(pnFila, 11) = "CODIGO : " & psCodCMAC
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
Dim sql As String
sql = "select nConsSisValor from constsistema where nConsSisCod=14"
Set loConecta = New DConecta
Set rs = New ADODB.Recordset
loConecta.AbreConexion
Set rs = loConecta.CargaRecordSet(sql)
loConecta.CierraConexion
GetUltimaFechaCierre = rs!nConsSisValor
End Function
Public Function ObtieneDescripcionFuenteFinanciamiento(ByVal psFuente As String) As String
Dim sql As String
Dim ServerCons As String
Dim rs As New ADODB.Recordset

Dim RCD As nRcdReportes
Set RCD = New nRcdReportes
ServerCons = RCD.GetServerConsol
sql = "select cDescripcion from ColocLineaCredito where clineaCred='" & psFuente & "'"
Set rs = RCD.GetQuery(sql)
If rs.EOF Then
    ObtieneDescripcionFuenteFinanciamiento = ""
Else
    ObtieneDescripcionFuenteFinanciamiento = IIf(IsNull(rs!CDescripcion), " ", rs!CDescripcion)
End If
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
Optional Flag As Boolean = True, Optional pnCalGen As String = "", Optional psCodCMAC As String) As ADODB.Recordset
Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsCadena As String
Dim lsSQL As String


lsCadena = IIf(psMoneda = "1", " and substring(cCtaCod,9,1)=1 ", " and substring(cCtaCod,9,1)=2 ")
'Select Case psCodCmac
'    Case "112"
'        If pnTipo = 0 Then
'            lsCadena = lsCadena & " and substring(cCtaCod,6,3)=101"
'        Else
'            lsCadena = lsCadena & " and substring(cCtaCod,6,3)=201"
'        End If
'    Case "102"
        If pnTipo = 0 Then
            lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('101','102','103')"
        Else
            lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('201','202')"
        End If
'End Select

If pnEstado = 0 Then
    lsCadena = lsCadena & " and cRefinan='N' "
  Else
    lsCadena = lsCadena & " and cRefinan='R' "
End If
lsCadena = lsCadena & " and  substring(cLineaCred,1,4)='" & psLineaCred & "' "


If Flag Then
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( Select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  "
    lsSQL = lsSQL & "   From ColocCalifProv"
    lsSQL = lsSQL & "   where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201) And nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & lsCadena
    lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
    lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
    lsSQL = lsSQL & " ORDER BY cAgeCod"
Else
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan "
    lsSQL = lsSQL & "   From ColocCalifProv"
    lsSQL = lsSQL & "   where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201,2205) And nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & "   and cCalGen=" & pnCalGen
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
    lsSQL = lsSQL & " where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201,2205) And " 'nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & lsCadena
    lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
    lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
    lsSQL = lsSQL & " ORDER BY cAgeCod"

Else
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  from ColocCalifProv"
    lsSQL = lsSQL & " where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201,2205) And " ' nMontoApr" & psSigno & pnMonto
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
    lsSQL = lsSQL & " where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201,2205) And nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & lsCadena
    lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
    lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
    lsSQL = lsSQL & " ORDER BY cAgeCod"
Else
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  from ColocCalifProv"
    lsSQL = lsSQL & " where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201,2205) And nMontoApr" & psSigno & pnMonto
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
Optional Flag As Boolean = True, Optional pnCalGen As String = "", Optional ByVal psCodCMAC As String = "") As ADODB.Recordset
Dim loReg As DConecta
Dim lrRs As New ADODB.Recordset
Dim lsCadena As String
Dim lsSQL As String


lsCadena = IIf(psMoneda = "1", " and substring(cCtaCod,9,1)=1 ", " and substring(cCtaCod,9,1)=2 ")
'If psCodCmac = "112" Then ' Trujillo
    lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('401','423')"
'End If
'If psCodCmac = "102" Then
'    lsCadena = lsCadena & " and substring(cCtaCod,6,3) in ('401','402')"
'End If
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
    lsSQL = lsSQL & " where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201,2205) And nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & lsCadena
    lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
    lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
    lsSQL = lsSQL & " ORDER BY cAgeCod"
Else
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan  from ColocCalifProv"
    lsSQL = lsSQL & " where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201,2205) And nMontoApr" & psSigno & pnMonto
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
    lsSQL = lsSQL & " where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201) And nMontoApr" & psSigno & pnMonto
    lsSQL = lsSQL & lsCadena
    lsSQL = lsSQL & " ) TA on Ag.cAgeCod=substring(TA.cCtaCod,4,2)"
    lsSQL = lsSQL & " group by cAgeDescripcion, cAgeCod"
    lsSQL = lsSQL & " ORDER BY cAgeCod"

Else
    lsSQL = " Select cAgeCod,cAgeDescripcion, isnull(Count(TA.cCtaCod),0) nTotal, isnull(sum(TA.nSaldoCap),0) sk_Total, isnull(sum(TA.nProvision),0) Prov"
    lsSQL = lsSQL & " from Agencias Ag Left outer join"
    lsSQL = lsSQL & " ( select cPersCod,cCtaCod, cCalGen, nSaldoCap, nProvision, nMontoApr, cLineaCred, cRefinan From ColocCalifProv "
    lsSQL = lsSQL & " where nprdestado in (2020,2022,2021,2030,2032,2031,2101,2104,2106,2107,2201) And nMontoApr" & psSigno & pnMonto
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

Public Sub ReporteSaldosProvxProd(ByVal lnTC As Currency)
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Dim vExcelObj  As Excel.Application

'sql = " SELECT  CPRODUCTO, A.cAgeDescripcion, CREC.cPersNombre, "
'sql = sql & "   sum(CASE WHEN nNroCliNor >0 THEN 1 ELSE 0 END) as nNroCliNor,"
'sql = sql & "   SUM(nSaldoNor) as nSaldoNor,"
'sql = sql & "   SUM(nProvNor) as nProvNor,"
'sql = sql & "   sum(CASE WHEN nNroCliCPP >0 THEN 1 ELSE 0 END) as nNroCliCPP,"
'sql = sql & "   SUM(nSaldoCPP) as nSaldoCPP,"
'sql = sql & "   SUM(nProvCPP) as nProvCPP,"
'sql = sql & "   sum(CASE WHEN nNroCliDEF >0 THEN 1 ELSE 0 END) as nNroCliDEF,"
'sql = sql & "   SUM(nSaldoDEF) as nSaldoDEF,"
'sql = sql & "   SUM(nProvDEF) as nProvDEF,"
'sql = sql & "   sum(CASE WHEN nNroCliDUD >0 THEN 1 ELSE 0 END) as nNroCliDUD,"
'sql = sql & "   SUM(nSaldoDUD) as nSaldoDUD,"
'sql = sql & "   SUM(nProvDUD) as nProvDUD,"
'sql = sql & "   sum(CASE WHEN nNroCliPERD >0 THEN 1 ELSE 0 END) as nNroCliPERD,"
'sql = sql & "   SUM(nSaldoPerd) as nSaldoPerd,"
'sql = sql & "   SUM(nProvPERD) as nProvPERD,"
'sql = sql & "   SUM(nSaldoCap) as nSaldoCap,"
'sql = sql & "   SUM(nProvision) As nProvision"
'sql = sql & " From"
'sql = sql & "   (SELECT cRecurso, cAgencia , cProducto, cPersCod,"
'sql = sql & "           SUM(NROCLINOR) AS nNroCliNor, SUM(nSaldoNor) AS nSaldoNor, SUM(nProvNor)  AS nProvNor,"
'sql = sql & "           SUM(NROCLICPP) AS nNroCliCPP, SUM(nSaldoCPP) AS nSaldoCPP, SUM(nProvCPP)  AS nProvCPP,"
'sql = sql & "           SUM(NROCLIDEF) AS nNroCliDEF, SUM(nSaldoDEF) AS nSaldoDEF, SUM(nProvDEF)  AS nProvDEF,"
'sql = sql & "           SUM(NROCLIDUD) AS nNroCliDUD, SUM(nSaldoDUD) AS nSaldoDUD, SUM(nProvDUD)  AS nProvDUD,"
'sql = sql & "           SUM(NROCLIPERD) AS nNroCliPERD, SUM(nSaldoPERD) AS nSaldoPerd, SUM(nProvPERD) AS nProvPERD ,"
'sql = sql & "           SUM(nSaldoCap) AS nSaldoCap, SUM(nProvision) as nProvision"
'sql = sql & "   From"
'sql = sql & "       (SELECT CC.cPersCod, SUBSTRING(CC.cLineaCred,1,2) AS cRecurso,"
'sql = sql & "               CASE SUBSTRING(CC.cCtaCod,4,2) WHEN '05' THEN '04'"
'sql = sql & "                                                   WHEN '06' THEN '03'"
'sql = sql & "                                                       WHEN '13' THEN '03'"
'sql = sql & "                                                           ELSE SUBSTRING(CC.cCtaCod,4,2) END cAgencia,"
'sql = sql & "           CASE WHEN SUBSTRING(CC.cCtaCod,6,1) ='1' AND SUBSTRING(CC.cCtaCod,7,2)<>'21' THEN 'COMERCIAL' ELSE"
'sql = sql & "           CASE WHEN SUBSTRING(CC.cCtaCod,6,1) ='2' AND SUBSTRING(CC.cCtaCod,7,2)<>'21' THEN 'MES' ELSE"
'sql = sql & "           CASE WHEN SUBSTRING(CC.cCtaCod,6,1) ='3' AND SUBSTRING(CC.cCtaCod,6,3)<>'305' THEN 'CONSUMO' ELSE"
'sql = sql & "           CASE WHEN SUBSTRING(CC.cCtaCod,6,3) ='305' THEN 'PRENDARIO' ELSE"
'sql = sql & "           CASE WHEN SUBSTRING(CC.cCtaCod,6,1) ='4' THEN 'HIPOTECARIO' ELSE"
'sql = sql & "           CASE WHEN SUBSTRING(CC.cCtaCod,7,2)='21' THEN 'CARTA FIANZA' ELSE '***********' END END END END END END AS CPRODUCTO,"
'sql = sql & "           CASE WHEN CC.cCalGen = '0' THEN  1 ELSE 0 END AS NROCLINOR,"
'sql = sql & "           CASE WHEN CC.cCalGen = '0' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & lnTC & ",2) END ELSE 0 END AS nSaldoNor,"
'sql = sql & "           CASE WHEN CC.cCalGen = '0' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nProvision ELSE ROUND(CC.nProvision*" & lnTC & ",2) END ELSE 0 END AS nProvNor,"
'sql = sql & "           CASE WHEN CC.cCalGen = '1' THEN  1 ELSE 0 END AS NROCLICPP,"
'sql = sql & "           CASE WHEN CC.cCalGen = '1' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & lnTC & ",2) END ELSE 0 END AS nSaldoCPP,"
'sql = sql & "           CASE WHEN CC.cCalGen = '1' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nProvision ELSE ROUND(CC.nProvision*" & lnTC & ",2) END ELSE 0 END AS nProvCPP,"
'sql = sql & "           CASE WHEN CC.cCalGen = '2' THEN  1 ELSE 0 END AS NROCLIDEF,"
'sql = sql & "           CASE WHEN CC.cCalGen = '2' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & lnTC & ",2) END ELSE 0 END AS nSaldoDEF,"
'sql = sql & "           CASE WHEN CC.cCalGen = '2' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nProvision ELSE ROUND(CC.nProvision*" & lnTC & ",2) END ELSE 0 END AS nProvDEF,"
'sql = sql & "           CASE WHEN CC.cCalGen = '3' THEN  1 ELSE 0 END AS NROCLIDUD,"
'sql = sql & "           CASE WHEN CC.cCalGen = '3' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & lnTC & ",2) END ELSE 0 END AS nSaldoDUD,"
'sql = sql & "           CASE WHEN CC.cCalGen = '3' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nProvision ELSE ROUND(CC.nProvision*" & lnTC & ",2) END ELSE 0 END AS nProvDUD,"
'sql = sql & "           CASE WHEN CC.cCalGen = '4' THEN  1 ELSE 0 END AS NROCLIPERD,"
'sql = sql & "           CASE WHEN CC.cCalGen = '4' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & lnTC & ",2) END ELSE 0 END AS nSaldoPERD,"
'sql = sql & "           CASE WHEN CC.cCalGen = '4' THEN  CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nProvision ELSE ROUND(CC.nProvision*" & lnTC & ",2) END ELSE 0 END AS nProvPERD,"
'sql = sql & "           CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nSaldoCap ELSE ROUND(CC.nSaldoCap*" & lnTC & ",2) END AS nSaldoCap,"
'sql = sql & "           CASE WHEN SUBSTRING(CC.cCtaCod,9,1) ='1' THEN CC.nProvision ELSE ROUND(CC.nProvision*" & lnTC & ",2) END nProvision"
'sql = sql & "       From    ColocCalifProv CC"
'sql = sql & "       WHERE   CC.nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)) AS DATA"
'sql = sql & "       GROUP BY cRecurso, cAgencia , cProducto, cPersCod ) AS X"
'sql = sql & "   Join"
'sql = sql & "       (Select L.cLineaCred, P.cPersNombre, L.cAbrev"
'sql = sql & "       from    ColocLineaCredito L"
'sql = sql & "       Inner Join Persona P ON P.cPersCod = L.cPersCod"
'sql = sql & "       Where Len(cLineaCred) = 2) AS CREC ON CREC.cLineaCred = X.cRecurso"
'sql = sql & "   JOIN AGENCIAS A ON A.cAgeCod = X.cAgencia"
'sql = sql & "   GROUP BY CPRODUCTO, X.cAgencia, A.cAgeDescripcion, CREC.cPersNombre,  X.cRecurso"
'sql = sql & "   ORDER BY CPRODUCTO,  X.cAgencia, X.cRecurso "

' By BRGO 20100609 BASILEA II
sql = " exec stp_sel_RS178211_ResumenSaldosProvisiones " & lnTC

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sql)
Dim lsArchivo As String
lsArchivo = "RepSaldoProvxProd.XLS"
If Not rs.EOF And Not rs.BOF Then
    RaiseEvent ShowProgress
        
    Set vExcelObj = New Excel.Application
    vExcelObj.DisplayAlerts = False
    
    vExcelObj.Workbooks.Add
    vExcelObj.Sheets("Hoja1").Select
    vExcelObj.Sheets("Hoja1").Name = "SALDOSPROVXPROD"
    
    vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
    vExcelObj.Range("A1:IV65536").Font.Size = 8
    vExcelObj.Columns("A:IV").Select
    vExcelObj.Selection.VerticalAlignment = 3
    
    vExcelObj.Columns("D:T").Select
    vExcelObj.Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"
        
    vExcelObj.Range("A1").Select
    vExcelObj.Range("A1").Font.Bold = True
    vExcelObj.Range("A1").ColumnWidth = 25
    vExcelObj.ActiveCell.value = "PRODUCTO"
        
    vExcelObj.Range("B1").Select
    vExcelObj.Range("B1").Font.Bold = True
    vExcelObj.Range("B1").ColumnWidth = 25
    vExcelObj.ActiveCell.value = "AGENCIA"
        
    vExcelObj.Range("C1").Select
    vExcelObj.Range("C1").Font.Bold = True
    vExcelObj.Range("C1").ColumnWidth = 25
    vExcelObj.ActiveCell.value = "RECURSO"
    
    vExcelObj.Range("D1").Select
    vExcelObj.Range("D1").Font.Bold = True
    vExcelObj.Range("D1").HorizontalAlignment = xlCenter
    vExcelObj.Range("D1").ColumnWidth = 15
    vExcelObj.ActiveCell.value = "NORMALES"
    
    vExcelObj.Range("D1:F1").Merge
    
    vExcelObj.Range("D2").Select
    vExcelObj.Range("D2").Font.Bold = True
    vExcelObj.Range("D2").ColumnWidth = 15
    vExcelObj.Range("D2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Nro.Clientes"
    
    vExcelObj.Range("E2").Select
    vExcelObj.Range("E2").Font.Bold = True
    vExcelObj.Range("E2").ColumnWidth = 15
    vExcelObj.Range("E2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Saldos"
    
    vExcelObj.Range("F2").Select
    vExcelObj.Range("F2").Font.Bold = True
    vExcelObj.Range("F2").ColumnWidth = 15
    vExcelObj.Range("F2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Provision"
    
    vExcelObj.Range("G1").Select
    vExcelObj.Range("G1").Font.Bold = True
    vExcelObj.Range("G1").ColumnWidth = 15
    vExcelObj.Range("G1").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "CPP"
    
    vExcelObj.Range("G1:I1").Merge
        
    vExcelObj.Range("G2").Select
    vExcelObj.Range("G2").Font.Bold = True
    vExcelObj.Range("G2").ColumnWidth = 15
    vExcelObj.Range("G2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Nro.Clientes"
    
    vExcelObj.Range("H2").Select
    vExcelObj.Range("H2").Font.Bold = True
    vExcelObj.Range("H2").ColumnWidth = 15
    vExcelObj.Range("H2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Saldos"
    
    vExcelObj.Range("I2").Select
    vExcelObj.Range("I2").Font.Bold = True
    vExcelObj.Range("I2").ColumnWidth = 15
    vExcelObj.Range("I2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Provision"
    
    vExcelObj.Range("J1").Select
    vExcelObj.Range("J1").Font.Bold = True
    vExcelObj.Range("J1").ColumnWidth = 15
    vExcelObj.Range("J1").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "DEFICIENTE"
    
    vExcelObj.Range("J1:L1").Merge
        
    vExcelObj.Range("J2").Select
    vExcelObj.Range("J2").Font.Bold = True
    vExcelObj.Range("J2").ColumnWidth = 15
    vExcelObj.Range("J2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Nro.Clientes"
    
    vExcelObj.Range("K2").Select
    vExcelObj.Range("K2").Font.Bold = True
    vExcelObj.Range("K2").ColumnWidth = 15
    vExcelObj.Range("K2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Saldos"
    
    vExcelObj.Range("L2").Select
    vExcelObj.Range("L2").Font.Bold = True
    vExcelObj.Range("L2").ColumnWidth = 15
    vExcelObj.Range("L2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Provision"
    
    vExcelObj.Range("M1").Select
    vExcelObj.Range("M1").Font.Bold = True
    vExcelObj.Range("M1").ColumnWidth = 15
    vExcelObj.Range("M1").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "DUDOSO"
    
    vExcelObj.Range("M1:O1").Merge
        
    vExcelObj.Range("M2").Select
    vExcelObj.Range("M2").Font.Bold = True
    vExcelObj.Range("M2").ColumnWidth = 15
    vExcelObj.Range("M2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Nro.Clientes"
    
    vExcelObj.Range("N2").Select
    vExcelObj.Range("N2").Font.Bold = True
    vExcelObj.Range("N2").ColumnWidth = 15
    vExcelObj.Range("N2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Saldos"
    
    vExcelObj.Range("O2").Select
    vExcelObj.Range("O2").Font.Bold = True
    vExcelObj.Range("O2").ColumnWidth = 15
    vExcelObj.Range("O2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Provision"
    
    vExcelObj.Range("P1").Select
    vExcelObj.Range("P1").Font.Bold = True
    vExcelObj.Range("P1").ColumnWidth = 15
    vExcelObj.Range("P1").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "PERDIDA"
    
    vExcelObj.Range("P1:R1").Merge
        
    vExcelObj.Range("P2").Select
    vExcelObj.Range("P2").Font.Bold = True
    vExcelObj.Range("P2").ColumnWidth = 15
    vExcelObj.Range("P2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Nro.Clientes"
    
    vExcelObj.Range("Q2").Select
    vExcelObj.Range("Q2").Font.Bold = True
    vExcelObj.Range("Q2").ColumnWidth = 15
    vExcelObj.Range("Q2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Saldos"
    
    vExcelObj.Range("R2").Select
    vExcelObj.Range("R2").Font.Bold = True
    vExcelObj.Range("R2").ColumnWidth = 15
    vExcelObj.Range("R2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "Provision"
    
    vExcelObj.Range("S1").Select
    vExcelObj.Range("S1").Font.Bold = True
    vExcelObj.Range("S1").ColumnWidth = 15
    vExcelObj.Range("S1").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "TOTAL SALDOS"
    
    vExcelObj.Range("S1:T1").Merge
        
    vExcelObj.Range("S2").Select
    vExcelObj.Range("S2").Font.Bold = True
    vExcelObj.Range("S2").ColumnWidth = 15
    vExcelObj.Range("S2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "SALDO.CAPITAL"
    
    vExcelObj.Range("T2").Select
    vExcelObj.Range("T2").Font.Bold = True
    vExcelObj.Range("T2").ColumnWidth = 15
    vExcelObj.Range("T2").HorizontalAlignment = xlCenter
    vExcelObj.ActiveCell.value = "PROVISION"
    
    Dim lsCelda As String
    Dim lnFila  As Long
    lnFila = 2
    
    Do While Not rs.EOF
        lnFila = lnFila + 1
        
        lsCelda = "A" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!cTipoCredito ' BRGO 20100609
        
        lsCelda = "B" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!cAgeDescripcion
        
        lsCelda = "C" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!cPersNombre
        
        lsCelda = "D" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nNroCliNor
        
        lsCelda = "E" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nSaldoNor
        
        lsCelda = "F" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nProvNor
        
        lsCelda = "G" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nNroCliCPP
        
        lsCelda = "H" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nSaldoCPP
        
        lsCelda = "I" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nProvCPP
        
        lsCelda = "J" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nNroCliDEF
        
        lsCelda = "K" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nSaldoDEF
        
        lsCelda = "L" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nProvDEF
        
        lsCelda = "M" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nNroCliDUD
        
        lsCelda = "N" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nSaldoDUD
        
        lsCelda = "O" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nProvDUD
        
        lsCelda = "P" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nNroCliPERD
        
        lsCelda = "Q" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nSaldoPerd
        
        lsCelda = "R" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nProvPERD
        
        lsCelda = "S" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nSaldoCap
        
        lsCelda = "T" & Trim(Str(lnFila))
        vExcelObj.Range(lsCelda).Select
        vExcelObj.ActiveCell.value = rs!nProvision
        
          
        RaiseEvent Progress(rs.Bookmark, rs.RecordCount)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End If
RaiseEvent CloseProgress
oCon.CierraConexion
Set oCon = Nothing

vExcelObj.Range("A1").Select
vExcelObj.ActiveWorkbook.SaveAs (App.Path & "\SPOOLER\" & lsArchivo)
vExcelObj.ActiveWorkbook.Close
MsgBox App.Path & "\SPOOLER\" & lsArchivo, vbInformation, "AVISO"
vExcelObj.Workbooks.Open (App.Path & "\SPOOLER\" & lsArchivo)
vExcelObj.Visible = True
    

End Sub


'Public Sub Genera_Reporte178210(ByVal TxtCambio As Double, ByVal dFecha As Date, _
'                ByVal pCartaFianza As Boolean, ByVal psServer As String, _
'                ByVal dFecData As String, _
'                Optional ByVal psNomCmac As String, _
'                Optional ByVal pscodcmac As String)
'
'Dim fs As Scripting.FileSystemObject
'Dim lbExisteHoja  As Boolean
'Dim lsNomHoja As String
'Dim pnFila As Integer, lnCol As Integer
'Dim lsSQL As String
'Dim rs As New ADODB.Recordset
'Dim rsSubT As New ADODB.Recordset
'Dim rsDesc As New ADODB.Recordset
'Dim lnLinea As Integer
'Dim lsMoneda As String, lsTipo As String, lsSubTipo As String, lsRango As String, lsFF As String
'Dim lsTotFF(14) As String
'Dim lsTotRango(14) As String
'Dim lsTotSubTipo(14) As String
'Dim lsTotTipo(14) As String
'Dim lsTotCartera(14) As String
'Dim lnIndice As Integer
'Dim lbCambio As Boolean ' Indica el cambio condicion
'Dim lnMoneda As Integer, lnTipo As Integer, lnSubTipo As Integer, lnRango As Integer, lnFF As Integer
'Dim lnCuenta As Long
'Dim lsCFianza As String
'Dim lsTipoDesc As String
'Dim lsSubTDesc As String
''Rep_HojaWProvision808 = True
'lnLinea = 1
'Dim lbMesActual As Boolean
'Dim lsArchivo As String
'Dim lbExcel As Boolean
'Dim J As Integer
'Dim lsNombre As String
'
'Dim Co As DConecta
'Set Co = New DConecta
'Co.AbreConexion
'
'lbCambio = False
'lsArchivo = "Anexo5_" & Format(dFecha, "yyyymm") & "_HojaW_808" & IIf(pCartaFianza = True, "", "_SinCF") & ".XLS"
'Set fs = New Scripting.FileSystemObject
'Set xlAplicacion = New Excel.Application
'If fs.FileExists(App.path & "\SPOOLER\" & lsArchivo) Then
'    Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & lsArchivo)
'Else
'    Set xlLibro = xlAplicacion.Workbooks.Add
'End If
'lbExcel = True
'
'If dFecha = dFecData Then
'    lbMesActual = True
'Else
'    lbMesActual = False
'End If
'If pCartaFianza = False Then
'    lsCFianza = " And Substring(CA.cCtaCod,7,2) <> '21' "
'End If
'
'For lnMoneda = 1 To 2 ' Moneda
'
'    lbExisteHoja = False
'    lsNomHoja = "HW_Moneda" & Str(lnMoneda)
'    For Each xlHoja1 In xlLibro.Worksheets
'        If xlHoja1.Name = lsNomHoja Then
'            xlHoja1.Activate
'            xlHoja1.Range("A1", "AZ10000") = ""
'            xlHoja1.Range("A1", "AZ10000").Rows.Delete
'            lbExisteHoja = True
'            Exit For
'        End If
'    Next
'    If lbExisteHoja = False Then
'        Set xlHoja1 = xlLibro.Worksheets.Add
'        xlHoja1.Name = lsNomHoja
'    End If
'    'Me.EstatuBarICC.Panels(2).Text = "Generando Reporte en Excel Por favor espere..."
'    pnFila = 2
'    ' Cabecera del Reporte en Excel
'    Call ImpHojaWProvisionCab(pnFila, Trim(Str(lnMoneda)), dFecha)
'    pnFila = 8
'    ' Cuerpo del Reporte
'
'    ' Limpiar Tot Cartera / Tipo / SubTipo
'    For lnIndice = 2 To 14
'        lsTotCartera(lnIndice) = ""
'        lsTotTipo(lnIndice) = ""
'        lsTotSubTipo(lnIndice) = ""
'    Next lnIndice
'
'    For lnTipo = 1 To 4
'
'        ' Imprimir Encabezado
'        If Trim(Str(lnTipo)) = "1" Then
'            lsTipoDesc = "I.- CREDITOS COMERCIALES"
'        ElseIf Trim(Str(lnTipo)) = "2" Then
'            lsTipoDesc = "II.- CREDITOS MICROEMPRESA"
'        ElseIf Trim(Str(lnTipo)) = "3" Then
'            lsTipoDesc = "III.- CREDITOS CONSUMO"
'        ElseIf Trim(Str(lnTipo)) = "4" Then
'            lsTipoDesc = "IV.- CREDITOS HIPOTECARIOS"
'        End If
'        xlHoja1.Cells(pnFila, 1) = lsTipoDesc
'        If pCartaFianza = True Then
'            xlHoja1.Cells(pnFila, 1) = xlHoja1.Cells(pnFila, 1) & " (Incluye Cartas Fianza ) "
'        Else
'            xlHoja1.Cells(pnFila, 1) = xlHoja1.Cells(pnFila, 1) & " (No Incluye Cartas Fianza )"
'        End If
'        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
'        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
'        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0C0FF
'
'        pnFila = pnFila + 1
'        ' Limpiar
'        For lnIndice = 2 To 14
'            lsTotSubTipo(lnIndice) = ""
'        Next lnIndice
'
'        lsSQL = " Select distinct(Substring(CA.cCtaCod,6,3)) cSubT FROM ColocCalifProv CA " _
'              & " Where substring(Ca.cCtaCod,6,1) = '" & lnTipo & "' AND Substring(CA.cCtaCod,9,1) = '" & Trim(Str(lnMoneda)) & "' " & lsCFianza & " ORDER BY Substring(CA.cCtaCod,6,3) "
'
'        Set rsSubT = Co.CargaRecordSet(lsSQL)
'        'rsSubT.CursorLocation = adUseClient
'        'rsSubT.Open lsSQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
'        'rsSubT.ActiveConnection = Nothing
'
'        Do While Not rsSubT.EOF
'            lnSubTipo = Trim(Str(rsSubT!cSubT))
'        'For lnSubTipo = 1 To 2
'            If Mid(rsSubT!cSubT, 2, 2) = "21" Then
'                lsSubTDesc = "CARTA FIANZA"
'            Else
'                'lsSQL = " Select * from " & gcCentralCom & "Tablacod " _
'                '      & " Where cCodtab like '6" & lnTipo & "__' And substring(cCodTab,2,3) = '" & Trim(Str(rsSubT!cSubT)) & "' "
'                'rsDesc.Open lsSQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
'                'lsSubTDesc = Trim(IIf(IsNull(rsDesc!cNomTab), "", rsDesc!cNomTab))
'                'rsDesc.Close
'                lsSubTDesc = "Producto - " & rsSubT!cSubT
'            End If
'
'            ' Imprimir Encabezado
'            xlHoja1.Cells(pnFila, 1) = lsSubTDesc
'            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
'            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
'            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HFFC0C0
'            pnFila = pnFila + 1
'            ' Limpiar
'            For lnIndice = 2 To 14
'                lsTotRango(lnIndice) = ""
'            Next lnIndice
'
'            For lnRango = 1 To 4   ' Por tipo de Garantia
'
'                lsSQL = "SELECT Substring(CA.cLineaCred,7,3) as cSubtipo, " _
'                      & "Substring(CA.cLineaCred,1,4) as cFF, substring(CA.cCtaCod,4,2) as cAgencia, " _
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
'                      & " CA.nPrdEstado in (2020,2021,2022,2030,2031,2032,2201,2205,2101,2104,2106,2107)   " _
'                      & "AND Substring(CA.cLineaCred,5,1) = '" & Trim(Str(lnMoneda)) & "' " _
'                      & "AND CA.nSaldoCap > 0  " _
'                      & "AND Substring(CA.cLineaCred,7,1) ='" & Trim(Str(lnTipo)) & "' " _
'                      & "AND Substring(CA.cLineaCred,7,3) ='" & Trim(Str(lnSubTipo)) & "' " _
'                      & "AND CA.nGarant = " & lnRango & " " _
'                      & lsCFianza
'                lsSQL = lsSQL & "GROUP BY  Substring(CA.cLineaCred,7,3), " _
'                      & "Substring(CA.cLineaCred,1,4), Substring(CA.cCtaCod,4,2) " _
'                      & "ORDER By Substring(CA.cLineaCred,7,3), " _
'                      & "Substring(CA.cLineaCred,1,4), Substring(CA.cCtaCod,4,2) "
'
'                Set rs = Co.CargaRecordSet(lsSQL)
'
'                'dbCmactCentral.CommandTimeout = 1200
'                'rs.CursorLocation = adUseClient
'                'rs.Open lsSQL, dbCmactCentral, adOpenStatic, adLockReadOnly, adCmdText
'                'rs.ActiveConnection = Nothing
'
'                lsFF = ""
'
'                If Not (rs.BOF And rs.EOF) Then
'                    ' Imprimir Encabezado
'                    If lnRango = 1 Then
'                        xlHoja1.Cells(pnFila, 1) = "Garantias AutoLiquidable"
'                    ElseIf lnRango = 2 Then
'                        xlHoja1.Cells(pnFila, 1) = "Garantias Muy Rapida Realizacion"
'                    ElseIf lnRango = 3 Then
'                        xlHoja1.Cells(pnFila, 1) = "Garantias Preferida"
'                    ElseIf lnRango = 4 Then
'                        xlHoja1.Cells(pnFila, 1) = "Garantias No Preferida"
'                    End If
'                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
'                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 1)).Font.Bold = True
'                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HFF00C
'                    pnFila = pnFila + 1
'                    ' Limpiar
'                    For lnIndice = 2 To 14
'                        lsTotFF(lnIndice) = ""
'                    Next lnIndice
'                End If
'
'                Do While Not rs.EOF
'
'                    If lsFF <> rs!cFF Then ' Otro Fuente Financ
'                        'Imprimir Total de anterior Fuente Fin
'                        If lsFF <> "" Then
'                            'Imprimir totales FF
'                            xlHoja1.Cells(pnFila, 1) = "SubTotal F.F. " & Trim(lsFF)
'                            For lnIndice = 2 To 14
'                                xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotFF(lnIndice) & ")"
'                            Next lnIndice
'                            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
'                            'Asigno los valores al Rango
'                            For lnIndice = 2 To 14
'                                lsTotRango(lnIndice) = lsTotRango(lnIndice) & IIf(Len(lsTotRango(lnIndice)) > 0, "+", "") & lsTotFF(lnIndice)
'                            Next lnIndice
'                            pnFila = pnFila + 1
'
'                        End If
'                        ' Imprimir encabezado
'                        xlHoja1.Cells(pnFila, 1) = "   F.F. " & ObtieneDescripcionFuenteFinanciamiento(Trim(rs!cFF))
'                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 1)).Font.Bold = True
'                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Merge True
'                        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0FFC0
'                        pnFila = pnFila + 1
'                        For lnIndice = 2 To 14
'                            lsTotFF(lnIndice) = ""
'                        Next lnIndice
'
'                    End If ' Fuente F
'
'                    '*****
'
'                    ' Imprimir el detalle
'                    xlHoja1.Cells(pnFila, 1) = "     Agencia " & rs!cAgencia
'                    xlHoja1.Cells(pnFila, 2) = Format(rs!NroCred, "#0")
'                    xlHoja1.Cells(pnFila, 3) = Format(rs!SaldoCap, "#0.00")
'                    xlHoja1.Cells(pnFila, 4) = Format(rs!NroProv0, "#0")
'                    xlHoja1.Cells(pnFila, 5) = Format(rs!Prov0, "#0.00")
'                    xlHoja1.Cells(pnFila, 6) = Format(rs!NroProv1, "#0")
'                    xlHoja1.Cells(pnFila, 7) = Format(rs!Prov1, "#0.00")
'                    xlHoja1.Cells(pnFila, 8) = Format(rs!NroProv2, "#0")
'                    xlHoja1.Cells(pnFila, 9) = Format(rs!Prov2, "#0.00")
'                    xlHoja1.Cells(pnFila, 10) = Format(rs!NroProv3, "#0")
'                    xlHoja1.Cells(pnFila, 11) = Format(rs!Prov3, "#0.00")
'                    xlHoja1.Cells(pnFila, 12) = Format(rs!NroProv4, "#0")
'                    xlHoja1.Cells(pnFila, 13) = Format(rs!Prov4, "#0.00")
'                    xlHoja1.Cells(pnFila, 14) = Format(rs!Provision, "#0.00")
'                    'Acumulo totales
'                    For lnIndice = 2 To 14
'                        lsTotFF(lnIndice) = lsTotFF(lnIndice) & IIf(Len(lsTotFF(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)  'Nro Cred
'                    Next lnIndice
'                    pnFila = pnFila + 1
'                    'lsTipo = rs!cTipo
'                    lsSubTipo = rs!cSubTipo
'                    'lnRango = rs!cRango
'                    lsFF = rs!cFF
'
'                    rs.MoveNext
'                Loop
'
'                rs.Close
'                Set rs = Nothing
'
'                'Imprimir totales FF
'                If Trim(lsTotFF(2)) <> "" Then
'                    xlHoja1.Cells(pnFila, 1) = "SubTotal F.F. " & Trim(lsFF)
'                    For lnIndice = 2 To 14
'                        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotFF(lnIndice) & ")"
'                    Next lnIndice
'                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
'                    'Asigno los valores al Rango
'                    For lnIndice = 2 To 14
'                        lsTotRango(lnIndice) = lsTotRango(lnIndice) & IIf(Len(lsTotRango(lnIndice)) > 0, "+", "") & lsTotFF(lnIndice)
'                    Next lnIndice
'                    pnFila = pnFila + 1
'
'                    'Imprimir totales Rango
'                    xlHoja1.Cells(pnFila, 1) = "SubTotal Garantia  " & Trim(lsRango)
'                    For lnIndice = 2 To 14
'                        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotRango(lnIndice) & ")"
'                    Next lnIndice
'                    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
'                    'xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = RGB(10, 100, 10)
'                    'Asigno los valores al SubTipo
'                    For lnIndice = 2 To 14
'                        'lsTotSubTipo(lnIndice) = lsTotSubTipo(lnIndice) + IIf(Len(lsTotSubTipo(lnIndice)) > 0, "+", "") + lsTotRango(lnIndice)
'                        lsTotSubTipo(lnIndice) = lsTotSubTipo(lnIndice) + IIf(Len(lsTotSubTipo(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
'                    Next lnIndice
'                    pnFila = pnFila + 1
'                End If
'
'                ' Limpio la variable FF
'                For lnIndice = 2 To 14
'                    lsTotFF(lnIndice) = ""
'                Next lnIndice
'                ' Limpio la variable Rango
'                For lnIndice = 2 To 14
'                    lsTotRango(lnIndice) = ""
'                Next lnIndice
'
'            Next lnRango
'
'            'Imprimir totales SubTipo
'            If Trim(lsTotSubTipo(2)) <> "" Then
'                xlHoja1.Cells(pnFila, 1) = "Total " & lsSubTDesc
'                For lnIndice = 2 To 14
'                    xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotSubTipo(lnIndice) & ")"
'                Next lnIndice
'                xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
'                'Asigno los valores al tipo
'                For lnIndice = 2 To 14
'                    'lsTotTipo(lnIndice) = lsTotTipo(lnIndice) & IIf(Len(lsTotTipo(lnIndice)) > 0, "+", "") & lsTotSubTipo(lnIndice)
'                    lsTotTipo(lnIndice) = lsTotTipo(lnIndice) & IIf(Len(lsTotTipo(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
'                Next lnIndice
'                pnFila = pnFila + 1
'            End If
'
'            ' Limpio la variable Subtipo
'            For lnIndice = 2 To 14
'                lsTotSubTipo(lnIndice) = ""
'            Next lnIndice
'
'            rsSubT.MoveNext
'         Loop
'
'        rsSubT.Close
'        'Imprimir totales Tipo
'        If Trim(lsTotTipo(2)) <> "" Then
'            xlHoja1.Cells(pnFila, 1) = "Tot." & Mid(lsTipoDesc, 4)
'            For lnIndice = 2 To 14
'                xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotTipo(lnIndice) & ")"
'            Next lnIndice
'            xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
'            'Asigno los valores a la Cartera
'            For lnIndice = 2 To 14
'                'lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & lsTotTipo(lnIndice)
'                lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
'            Next lnIndice
'            pnFila = pnFila + 2
'        End If
'
'        ' Limpio la variable tipo
'        For lnIndice = 2 To 14
'            lsTotTipo(lnIndice) = ""
'        Next lnIndice
'
'    Next lnTipo
'
'    'Imprimir totales Tipo
'    If Trim(lsTotTipo(2)) <> "" Then
'        xlHoja1.Cells(pnFila, 1) = "Total " & lsTipoDesc
'        For lnIndice = 2 To 14
'            xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotTipo(lnIndice) & ")"
'        Next lnIndice
'        xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
'        'Asigno los valores a la Cartera
'        For lnIndice = 2 To 14
'            'lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & lsTotTipo(lnIndice)
'            lsTotCartera(lnIndice) = lsTotCartera(lnIndice) & IIf(Len(lsTotCartera(lnIndice)) > 0, "+", "") & xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Address(False, False)
'        Next lnIndice
'        pnFila = pnFila + 2
'    End If
'
'    ' Limpio la variable tipo
'    For lnIndice = 2 To 14
'        lsTotTipo(lnIndice) = ""
'    Next lnIndice
'
'
'    ' ********************
'
'    'Imprime Total Cartera
'    pnFila = pnFila + 1
'    xlHoja1.Cells(pnFila, 1) = "Total CARTERA"
'    For lnIndice = 2 To 14
'        xlHoja1.Range(xlHoja1.Cells(pnFila, lnIndice), xlHoja1.Cells(pnFila, lnIndice)).Formula = "=SUM(" & lsTotCartera(lnIndice) & ")"
'    Next lnIndice
'    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Font.Bold = True
'    xlHoja1.Range(xlHoja1.Cells(pnFila, 1), xlHoja1.Cells(pnFila, 14)).Interior.Color = &HC0FFFF
'
'    'CuadroExcel 1, 6, 14, pnFila, True
'    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(pnFila, 14)).Cells.Borders.LineStyle = xlOutside
'    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(pnFila, 14)).Cells.Borders.LineStyle = xlInside
'
'Next lnMoneda
'
''*******
'xlHoja1.SaveAs App.path & "\SPOOLER\" & lsArchivo
''Cierra el libro de trabajo
'xlLibro.Close
'' Cierra Microsoft Excel con el método Quit.
'xlAplicacion.Quit
''Libera los objetos.
'Set xlAplicacion = Nothing
'Set xlLibro = Nothing
'Set xlHoja1 = Nothing
'lbExcel = False
'
'
'
'End Sub
'ALPA 20090703******************************************************
Public Sub Genera_Reporte178213(ByVal dFecha As Date, _
                        ByVal nCF As Integer, ByVal psServer As String, _
                        ByVal dFecData As String, _
                        ByVal psFFinan As String, _
                        Optional ByVal psNomCmac As String, _
                        Optional ByVal psCodCMAC As String)


Dim sql As String
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
Dim lnAnio As Integer
Dim RCC As DRCC
Set RCC = New DRCC
lsServConsol = RCC.ServConsol(psServer)
Set RCC = Nothing


Dim Co As DConecta
Set Co = New DConecta
Co.AbreConexion

'If dFecha = dFecData Then
'    lbMesActual = True
'Else
'    lbMesActual = False
'End If

'If nCF = 1 Then
'    lbIncluyeCF = True
'Else
'    lbIncluyeCF = False
'End If

lsArchivo = "Rep_Anexo4_2C" & Format(dFecha, "YYYYMM") & ".XLS"
Set fs = New Scripting.FileSystemObject
Set xlAplicacion = New Excel.Application
If fs.FileExists(App.Path & "\SPOOLER\" & lsArchivo) Then
    Set xlLibro = xlAplicacion.Workbooks.Open(App.Path & "\SPOOLER\" & lsArchivo)
Else
    Set xlLibro = xlAplicacion.Workbooks.Add
End If

 lsNomHoja = "Anexo4"
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

xlHoja1.Range("A1:AZ1000") = ""
xlHoja1.Range("A2") = "Anexo N° 4"
xlHoja1.Range("A3") = "Reporte Nº 2-C1"
xlHoja1.Range("A4") = "Requerimiento de Patrimonio Efectivo por Riesgo Operacional"
xlHoja1.Range("A5") = "Metodo del Indicador Básico"
xlHoja1.Range("A6") = "Correpondiente al " & Format(dFecha, "DD") & " DE  " & UCase(Format(dFecha, "MMMM")) & " del  " & Format(dFecha, "YYYY")
xlHoja1.Range("A7") = "Entidad : " & psNomCmac

xlHoja1.Range("A1").ColumnWidth = 42
xlHoja1.Range("B1").ColumnWidth = 15
xlHoja1.Range("C1").ColumnWidth = 15
xlHoja1.Range("D1").ColumnWidth = 15
xlHoja1.Range("E1").ColumnWidth = 15
xlHoja1.Range("F1").ColumnWidth = 15
xlHoja1.Range("G1").ColumnWidth = 16

xlHoja1.Application.ActiveWindow.Zoom = 75
xlHoja1.Range("A2:G2").MergeCells = True
xlHoja1.Range("A2:A2").HorizontalAlignment = xlLeft
xlHoja1.Range("A3:A3").HorizontalAlignment = xlLeft
xlHoja1.Range("A3:G3").MergeCells = True
xlHoja1.Range("A4:G4").MergeCells = True
xlHoja1.Range("A8:G8").MergeCells = True
xlHoja1.Range("A8:A8").HorizontalAlignment = xlLeft
xlHoja1.Range("A9:A9").HorizontalAlignment = xlLeft
xlHoja1.Range("A2:G9").Font.Bold = True
lnAnio = CInt(Format(dFecha, "YYYY"))
xlHoja1.Range("A9") = "Periodo de 12 Meses "
xlHoja1.Range("A10") = "Del " & Format(dFecha, "MM") & "/" & Format(dFecha, "YYYY") & " al " & Format(dFecha, "MM") & "/" & CStr(lnAnio - 1)
xlHoja1.Range("A11") = "Del " & Format(dFecha, "MM") & "/" & CStr(lnAnio - 1) & " al " & Format(dFecha, "MM") & "/" & CStr(lnAnio - 2)
xlHoja1.Range("A12") = "Del " & Format(dFecha, "MM") & "/" & CStr(lnAnio - 2) & " al " & Format(dFecha, "MM") & "/" & CStr(lnAnio - 3)
xlHoja1.Range("A13") = "Requerimiento de patrimonio efectivo por riesgo operacional"
xlHoja1.Range("A13:A13").HorizontalAlignment = xlJustify
xlHoja1.Range("A14") = "APR riesgo operacional"
xlHoja1.Range("A16") = "Inversa del limite global"
xlHoja1.Range("A17") = "Factor de Ajuste"

xlHoja1.Range("A9:G9").Font.Bold = True

xlHoja1.Range("B9") = " Ingresos Financieros"
xlHoja1.Range("C9") = " Ingresos por servicios"
xlHoja1.Range("D9") = " Gastos Financieros"
xlHoja1.Range("E9") = " Gastos por servicios"
xlHoja1.Range("F9") = " Margen Operacional Bruto"

xlHoja1.Range("A9:F12").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A9:F12").Borders(xlInsideHorizontal).LineStyle = xlContinuous
xlHoja1.Range("A9:F12").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("A9:F12").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("A9:F12").Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("A9:F12").Borders(xlEdgeLeft).LineStyle = xlContinuous


xlHoja1.Range("A13:A14").Borders(xlInsideHorizontal).LineStyle = xlContinuous
xlHoja1.Range("A13:A14").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("A13:A14").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("A13:A14").Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("A13:A14").Borders(xlEdgeLeft).LineStyle = xlContinuous


xlHoja1.Range("F13:F14").Borders(xlInsideHorizontal).LineStyle = xlContinuous
xlHoja1.Range("F13:F14").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("F13:F14").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("F13:F14").Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("F13:F14").Borders(xlEdgeLeft).LineStyle = xlContinuous
  
xlHoja1.Range("A16:B17").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A16:B17").Borders(xlInsideHorizontal).LineStyle = xlContinuous
xlHoja1.Range("A16:B17").Borders(xlEdgeTop).LineStyle = xlContinuous
xlHoja1.Range("A16:B17").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("A16:B17").Borders(xlEdgeRight).LineStyle = xlContinuous
xlHoja1.Range("A16:B17").Borders(xlEdgeLeft).LineStyle = xlContinuous

sql = " exec stp_sel_ReporteRiesgoCambiario '" & Format(dFecha, "YYYY") & "','" & Format(dFecha, "MM") & "'"

Set rs = Co.CargaRecordSet(sql)
While Not rs.EOF
        xlHoja1.Range("B10") = Format(rs!Saldo_Anualizado1_51, "#,###.00")
        xlHoja1.Range("C10") = Format(rs!Saldo_Anualizado1_52_57, "#,###.00")
        xlHoja1.Range("D10") = Format(rs!Saldo_Anualizado1_41, "#,###.00")
        xlHoja1.Range("E10") = Format(rs!Saldo_Anualizado1_42_49, "#,###.00")
               
        xlHoja1.Range("B17") = CStr(Format(rs!nFacAjuste, "##.#0"))
        xlHoja1.Range("B16") = 20 - rs!nFacPonRequer
        'xlHoja1.Range("F10").
        
        xlHoja1.Range("B11") = Format(rs!Saldo_Anualizado2_51, "#,###.00")
        xlHoja1.Range("C11") = Format(rs!Saldo_Anualizado2_52_57, "#,###.00")
        xlHoja1.Range("D11") = Format(rs!Saldo_Anualizado2_41, "#,###.00")
        xlHoja1.Range("E11") = Format(rs!Saldo_Anualizado2_42_49, "#,###.00")
        'xlHoja1.Range("F11").Formula = "=" & rs!Saldo_Anualizado2 '& "*B17"
        
        xlHoja1.Range("B12") = Format(rs!Saldo_Anualizado3_51, "#,###.00")
        xlHoja1.Range("C12") = Format(rs!Saldo_Anualizado3_52_57, "#,###.00")
        xlHoja1.Range("D12") = Format(rs!Saldo_Anualizado3_41, "#,###.00")
        xlHoja1.Range("E12") = Format(rs!Saldo_Anualizado3_42_49, "#,###.00")
        'xlHoja1.Range("F12").Formula = "=" & rs!Saldo_Anualizado3 '& "*B17"
                
        rs.MoveNext
Wend

xlHoja1.Range("F10").Formula = "= sum(B10:C10)-sum(D10:E10)"
xlHoja1.Range("F11").Formula = "= sum(B11:C11)-sum(D11:E11)"
xlHoja1.Range("F12").Formula = "= sum(B12:C12)-sum(D12:E12)"

xlHoja1.Range("F13").Formula = "= (SUM(F10:F12)*15)/300"
xlHoja1.Range("F14").Formula = "= (((SUM(F10:F12)*15)/300)" & "*(B16))" & "*B17"
xlHoja1.Range("F10:F14").Style = "Comma"
 
xlHoja1.SaveAs App.Path & "\SPOOLER\" & lsArchivo

xlLibro.Close
xlAplicacion.Quit
Set xlAplicacion = Nothing
Set xlLibro = Nothing
Set xlHoja1 = Nothing
MsgBox "Se genero el Archivo. Verifique en la Carpeta Spooler", vbInformation, "AVISO"
End Sub
'******************************************************************************************
'ALPA 20150722*****************************************************************************
Public Function ObtieneSaldosDeRefinanciadoANormal(ByVal pdFecha As Date, ByVal pnTipoCambio As Currency) As ADODB.Recordset
Dim sql As String
Dim loReg As DConecta
Set loReg = New DConecta
    loReg.AbreConexion
    sql = "exec stp_sel_ObtenerSaldosCambiosRefinanciadoANormal '" & Format(pdFecha, "YYYY/MM/DD") & "'," & pnTipoCambio
    Set ObtieneSaldosDeRefinanciadoANormal = loReg.CargaRecordSet(sql)
    loReg.AbreConexion
End Function
'******************************************************************************************

