Attribute VB_Name = "mdlLogProSelReportes"
Option Explicit

Global Const gsCmact = "CMAC-T"
Global Const gsConnectionLOG_DBF = "DSN=DSNCmact_CONSUCODE;UID=;PWD=;Database=;Server="
Global Const gsArchCONVOCA = 1
Global Const gsArchITEM = 2
Global Const gnConvocatoriaElectrónica = 2
'Global Const gsVersionCONSUCODE = "1.2.1"

'Global Const gsCuentaBCredito_S = "570-03976641-0-62"
'Global Const gsCuentaBContinental_S = "0249-05-0200018227"

'Global Const gsDireccion = "Av. España No 2611, Trujillo"
    
Global Const cnConvocatoria = 1
Global Const cnRegistroParticipantes = 2
Global Const cnPresentacionConsultas = 3
Global Const cnAbsolucionConsultas = 4
Global Const cnObservaciones = 5
Global Const cnEvaluacionPropuestasObservaciones = 6
Global Const cnResolucionObservaciones = 7
Global Const cnIntegracionBases = 8
Global Const cnPresentacionPropuestas = 9
Global Const cnEvaluacionPropuestas = 10
Global Const cnOtorgamientoBP = 11
Global Const cnAtencionEvaluacionConsultas = 13
Global Const cnConcentimientoBP = 16
Global Const cnApelaciones = 17

Public Function LimpiaString(ByVal sTexto As String) As String
On Error GoTo LimpiaStringErr
    Dim nLen As Long, i As Long, sNewTexto As String, nCaracter As Integer
    nLen = Len(sTexto)
    For i = 1 To nLen
        nCaracter = Asc(Mid(sTexto, i, 1))
        Select Case nCaracter
            Case 8
                sNewTexto = sNewTexto & " "
            Case 32, 13
                sNewTexto = sNewTexto & Mid(sTexto, i, 1)
            Case 40 To 125
                If nCaracter <> 96 Then sNewTexto = sNewTexto & Chr(nCaracter)
        End Select
    Next i
    LimpiaString = sNewTexto
    Exit Function
LimpiaStringErr:
    MsgBox Err.Number & vbCrLf & Err.Description
End Function


Public Sub Generar_ActaBuenaPro(ByVal pnProSelNro As Integer, _
            ByVal pcProceso As String, ByVal pcDescripcion As String, _
            pcMoneda As String, pnMonto As Currency)

On Error GoTo Generar_ActaBuenaProErr

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1(100) As Excel.Worksheet
    Dim lbExisteHoja  As Boolean
    
    Dim cBSGrupoCod As String
    Dim cBSGrupoDescripcion As String

    Dim lilineas As Integer
    Dim Total As Currency, TotalPos As Currency
    Dim nFactor As Integer
    Dim nCol As Integer, i As Integer, nLetra As Integer
    
    Dim CadTit As String
    Dim CadImp   As String, Ganador As Integer
    Dim rs As New ADODB.Recordset, RsItem As ADODB.Recordset, RsDesestimado As ADODB.Recordset
    Dim cGanadorNom As String, cGanadorRuc As String, _
        cGanadorDir As String, cGanadorTlf As String, _
        nNroHoja As Integer, nProSelItem As Integer, nMontoItem As Currency
    Dim lsNomHoja As String

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    
    nNroHoja = 1
    Set RsItem = ListaItem(pnProSelNro)
    Do While Not RsItem.EOF
        nLetra = 0
        nCol = 5
        Ganador = 0
        cGanadorDir = ""
        cGanadorNom = ""
        cGanadorRuc = ""
        cGanadorTlf = ""
        
        cBSGrupoCod = RsItem!cBSGrupoCod
        cBSGrupoDescripcion = RsItem!cBSGrupoDescripcion
        nProSelItem = RsItem!nProSelItem
        nMontoItem = RsItem!nMonto
        
        ExcelAddHoja "Item " & nNroHoja, xlLibro, xlHoja1(nNroHoja), False
        'Set xlHoja1(nNroHoja) = xlLibro.Worksheets.Item(nNroHoja)
        
        'xlHoja1(nNroHoja).Name = "Item " & nNroHoja 'cBSGrupoDescripcion
        '**************************************************************************

'         xlAplicacion.Range("A1:A1").ColumnWidth = 5
'         xlAplicacion.Range("B1:B1").ColumnWidth = 5
'         xlAplicacion.Range("C1:C1").ColumnWidth = 20
         
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 1), xlHoja1(nNroHoja).Cells(1, 1)).ColumnWidth = 5
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 2), xlHoja1(nNroHoja).Cells(1, 2)).ColumnWidth = 5
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 3), xlHoja1(nNroHoja).Cells(1, 3)).ColumnWidth = 20
         
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 2)).Font.Bold = True
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 4)).Merge True
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(3, 4)).Merge True
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 4)).Merge True
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(5, 2), xlHoja1(nNroHoja).Cells(5, 4)).Merge True
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(5, 7)).Font.Bold = True

    '******************************************************
                
        lilineas = 9

        Screen.MousePointer = 11
        xlHoja1(nNroHoja).Cells(2, 2) = "Caja Municipal de Ahorros y Creditos de Trujillo SA"
        xlHoja1(nNroHoja).Cells(3, 2) = pcProceso
        xlHoja1(nNroHoja).Cells(4, 2) = "ADQUISICION DE " & pcDescripcion
        xlHoja1(nNroHoja).Cells(5, 2) = "Acta de la Buena Pro"
        xlHoja1(nNroHoja).Cells(7, 3) = "ITEM N° " & nNroHoja & ": " & cBSGrupoDescripcion
        
        xlHoja1(nNroHoja).Cells(lilineas, 2) = "ORDEN DE PRELACIÓN"
        xlHoja1(nNroHoja).Cells(lilineas, 4) = "Puntos"
        lilineas = lilineas + 1
        xlHoja1(nNroHoja).Cells(lilineas, 2) = "Total Puntaje = (1) + (2)"
        lilineas = lilineas + 1
            
        Set rs = CargarFactores(pnProSelNro, cBSGrupoCod)
        Total = 0
        lilineas = 11
        TotalPos = -1000
        Do While Not rs.EOF
            If nFactor <> rs!fnFactorNro Then
                nFactor = rs!fnFactorNro
                nLetra = nLetra + 1
                nCol = 5
                lilineas = lilineas + 1
                If rs!xnTipo Then
                    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 1), xlHoja1(nNroHoja).Cells(lilineas, 100)).Font.Bold = True
                    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 3)).Merge True
                    
                    xlHoja1(nNroHoja).Cells(lilineas, 2) = "2. " & rs!cFactorDescripcion
                    xlHoja1(nNroHoja).Cells(lilineas + 1, 3) = Space(10) & rs!cUnidades
                    xlHoja1(nNroHoja).Cells(lilineas, 4) = rs!npuntaje
                    xlHoja1(nNroHoja).Cells(lilineas + 1, 3) = Space(10) & "Valor Referencial " & pcMoneda & Format(nMontoItem, "###,###,##0.00")
                    
                    xlHoja1(nNroHoja).Cells(11, 2) = "1. Puntaje Propuesta Tecnica"
                    xlHoja1(nNroHoja).Cells(11, 4) = Total
                    xlHoja1(nNroHoja).Cells(10, 4) = Val(xlHoja1(nNroHoja).Cells(11, 4)) + Val(xlHoja1(nNroHoja).Cells(lilineas, 4))
                Else
                    xlHoja1(nNroHoja).Cells(lilineas, 3) = "(" & Chr(nLetra + 96) & ") " & rs!cFactorDescripcion
                    xlHoja1(nNroHoja).Cells(lilineas + 1, 3) = Space(10) & rs!cUnidades
                    xlHoja1(nNroHoja).Cells(lilineas, 4) = rs!npuntaje
                    Total = Total + rs!npuntaje
                End If
            Else
                lilineas = lilineas - 1
            End If
            xlHoja1(nNroHoja).Cells(8, nCol) = "RUC: " & rs!RUC
            xlHoja1(nNroHoja).Cells(7, nCol) = rs!cPersNombre
            xlHoja1(nNroHoja).Cells(6, nCol) = rs!cPersDireccDomicilio
            xlHoja1(nNroHoja).Cells(5, nCol) = rs!cPersTelefono
            If rs!bDesestimado Then
                xlHoja1(nNroHoja).Cells(9, nCol) = "Desestimado por " & rs!cdesesdescripcion
                xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, nCol), xlHoja1(nNroHoja).Cells(9, nCol)).WrapText = True
                xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, nCol), xlHoja1(nNroHoja).Cells(9, nCol)).RowHeight = 24
                'xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, nCol), xlHoja1(nNroHoja).Cells(10, nCol)).Merge False
            End If
            If Not rs!xnTipo Then
                xlHoja1(nNroHoja).Cells(lilineas, nCol) = rs!Obtenido
                xlHoja1(nNroHoja).Cells(11, nCol) = Val(xlHoja1(nNroHoja).Cells(11, nCol)) + rs!Obtenido
                lilineas = lilineas + 1
                xlHoja1(nNroHoja).Cells(lilineas, nCol) = rs!nValor
            Else
                If rs!Obtenido > 0 Then
                    xlHoja1(nNroHoja).Cells(lilineas, nCol) = rs!Obtenido
                    xlHoja1(nNroHoja).Cells(10, nCol) = Val(xlHoja1(nNroHoja).Cells(11, nCol)) + Val(xlHoja1(nNroHoja).Cells(lilineas, nCol))
                    lilineas = lilineas + 1
                    xlHoja1(nNroHoja).Cells(lilineas, nCol) = rs!nValor
                Else
                    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, nCol), xlHoja1(nNroHoja).Cells(lilineas + 1, nCol)).Merge False
                    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, nCol), xlHoja1(nNroHoja).Cells(lilineas + 1, nCol)).Interior.Color = &HC0C0C0
                    lilineas = lilineas + 1
                End If
            End If
            
            If rs!bGanador Then
'                    TotalPos = xlHoja1(nNroHoja).Cells(11, nCol)
                Ganador = nCol
                cGanadorRuc = xlHoja1(nNroHoja).Cells(8, nCol)
                cGanadorNom = xlHoja1(nNroHoja).Cells(7, nCol)
                cGanadorDir = xlHoja1(nNroHoja).Cells(6, nCol)
                cGanadorTlf = xlHoja1(nNroHoja).Cells(5, nCol)
            End If
            nCol = nCol + 1
            rs.MoveNext
        Loop
        
        Set RsDesestimado = CargarDescalificado(pnProSelNro, nProSelItem)
        Do While Not RsDesestimado.EOF
            xlHoja1(nNroHoja).Cells(8, nCol) = "RUC: " & RsDesestimado!RUC
            xlHoja1(nNroHoja).Cells(7, nCol) = RsDesestimado!cPersNombre
            xlHoja1(nNroHoja).Cells(6, nCol) = RsDesestimado!cPersDireccDomicilio
            xlHoja1(nNroHoja).Cells(5, nCol) = RsDesestimado!cPersTelefono
            xlHoja1(nNroHoja).Cells(9, nCol) = "Desestimado por " & RsDesestimado!cdesesdescripcion
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, nCol), xlHoja1(nNroHoja).Cells(9, nCol)).WrapText = True
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, nCol), xlHoja1(nNroHoja).Cells(9, nCol)).RowHeight = 24
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(11, nCol), xlHoja1(nNroHoja).Cells(lilineas, nCol)).Merge False
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(11, nCol), xlHoja1(nNroHoja).Cells(lilineas, nCol)).Cells.Interior.Color = &HC0C0C0
            nCol = nCol + 1
            RsDesestimado.MoveNext
        Loop
        
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 1), xlHoja1(nNroHoja).Cells(11, nCol - 1)).Font.Bold = True
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, 2), xlHoja1(nNroHoja).Cells(10, nCol - 1)).Cells.Interior.Color = &HC0FFFF
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 5), xlHoja1(nNroHoja).Cells(9, nCol - 1)).Cells.Interior.Color = &HC0FFFF
        
        If Ganador > 0 Then
            xlHoja1(nNroHoja).Cells(9, Ganador) = "1°"
            'xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, Ganador), xlHoja1(nNroHoja).Cells(9, Ganador)).Font.Size = 16
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, Ganador), xlHoja1(nNroHoja).Cells(lilineas, Ganador)).Cells.Interior.Color = &HFCFFE1
        End If
        
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(11, 2), xlHoja1(nNroHoja).Cells(11, 3)).Merge True
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 2), xlHoja1(nNroHoja).Cells(10, 3)).Merge True
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(9, 2), xlHoja1(nNroHoja).Cells(9, 3)).Merge True
                
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 4), xlHoja1(nNroHoja).Cells(lilineas, nCol - 1)).HorizontalAlignment = xlCenter
        
        ExcelCuadro xlHoja1(nNroHoja), 5, 6, nCol - 1, 9, True, True
        ExcelCuadro xlHoja1(nNroHoja), 2, 9, nCol - 1, lilineas, True, True
        
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, nCol - 1), xlHoja1(nNroHoja).Cells(2, nCol - 1)).HorizontalAlignment = xlRight
        xlHoja1(nNroHoja).Cells(2, IIf(nCol = 5, 5, nCol - 1)) = Format(gdFecSis, "Trujillo, dd - mmmm - yyyy")
        
        lilineas = lilineas + 3
        
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, nCol - 1)).Merge True
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 1), xlHoja1(nNroHoja).Cells(lilineas, nCol - 1)).Font.Bold = True
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 1), xlHoja1(nNroHoja).Cells(lilineas, nCol - 1)).Font.Underline = True
        xlHoja1(nNroHoja).Cells(lilineas, 2) = "Ganador de la Buena Pro:"
        lilineas = lilineas + 1
        If Ganador = 0 Then
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, nCol - 1)).Merge True
            xlHoja1(nNroHoja).Cells(lilineas, 2) = "Proceso Decierto..."
            lilineas = lilineas + 1
        Else
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, nCol - 1)).Merge True
            xlHoja1(nNroHoja).Cells(lilineas, 2) = cGanadorNom
            lilineas = lilineas + 1
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, nCol - 1)).Merge True
            xlHoja1(nNroHoja).Cells(lilineas, 2) = cGanadorRuc
            lilineas = lilineas + 1
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, nCol - 1)).Merge True
            xlHoja1(nNroHoja).Cells(lilineas, 2) = "DOMICILIO LEGAL: " & cGanadorDir
            lilineas = lilineas + 1
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, nCol - 1)).Merge True
            xlHoja1(nNroHoja).Cells(lilineas, 2) = "TELEFONO: " & cGanadorTlf
            lilineas = lilineas + 1
        End If
        
        i = 5
        Do While i < nCol
            xlHoja1(nNroHoja).Cells(5, i) = ""
            xlHoja1(nNroHoja).Cells(6, i) = "Postor " & i - 4
            i = i + 1
        Loop
        
        Screen.MousePointer = 0
    '******************************************************
    
'        xlHoja1(nNroHoja).Cells.Select
        xlHoja1(nNroHoja).Cells.Font.Size = 8
        xlHoja1(nNroHoja).Cells.EntireColumn.AutoFit
        xlHoja1(nNroHoja).Cells.NumberFormat = "###,###,##0.00"
    
        'Libera los objetos.
        Set xlHoja1(nNroHoja) = Nothing
        nNroHoja = nNroHoja + 1
        RsItem.MoveNext
    Loop
    
    xlLibro.SaveAs App.path & "\spooler\Buena_Pro_" & pcProceso
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
    xlAplicacion.Windows(1).Visible = True
    Exit Sub
Generar_ActaBuenaProErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Function CargarFactores(ByVal pnProSelNro As Integer, ByVal pcBSGrupoCod As String) As ADODB.Recordset
    On Error GoTo CargarFactoresErr
    Dim ocon As DConecta, rs As ADODB.Recordset, sSQL As String
    Set ocon = New DConecta
'    sSQL = "select distinct f.cFactorDescripcion, e.nPuntaje " & _
           " from LogProSelFactor f " & _
           " inner join LogProSelEvalFactor e on e.nFactorNro = f.nFactorNro " & _
           " inner join LogProSelEvalFactorValor v on v.nFactorNro = f.nFactorNro " & _
           " where v.nProSelNro=" & pnProSelNro & " and e.cBSGrupoCod='" & pcBSGrupoCod & "'"
    sSQL = "select RUC=isnull(i.cPersIDnro,''), p.cPersNombre, p.cPersDireccDomicilio, p.cPersTelefono, f.nPuntaje, " & _
           "v.nValor, Obtenido=v.nPuntaje, fnFactorNro=f.nFactorNro, x.cFactorDescripcion,  f.cBSGrupoCod , f.nFormula, " & _
           "x.cUnidades, xnTipo=x.nTipo, pp.bGanador, pp.cDesesDescripcion, pp.bDesestimado " & _
           " from LogProSelEvalResultado v " & _
           " inner join LogProSelEvalFactor f on v.nFactorNro = f.nFactorNro and f.cBSGrupoCod = v.cBSGrupoCod and f.nProSelNro = v.nProSelNro  " & _
           " inner join constante c on f.nFormula = c.nConsValor and c.nconscod=9084 " & _
           " inner join LogProSelFactor x on v.nFactorNro = x.nFactorNro " & _
           " inner join Persona p on p.cPersCod = v.cPersCod " & _
           " left outer join PersID i on p.cPersCod = i.cPersCod and cPersIDTpo=2" & _
           " inner join LogProSelPostorPropuesta pp on pp.nProSelNro = v.nProSelNro and pp.nProSelItem = v.nProSelItem and pp.cPersCod = v.cPersCod" & _
           " where nVigente=1 and  v.nProSelNro=" & pnProSelNro & " AND f.cBSGrupoCod='" & pcBSGrupoCod & "' " & _
           " group by x.nTipo, i.cPersIDnro, p.cPersNombre, p.cPersDireccDomicilio, p.cPersTelefono, f.nPuntaje, v.nValor, v.nPuntaje, v.nFactorNro, " & _
           " x.cFactorDescripcion,  f.cBSGrupoCod , f.nFormula , f.nFactorNro, x.cUnidades, pp.bGanador, pp.cDesesDescripcion, pp.bDesestimado"
'    sSQL = sSQL & "union all "
'    sSQL = sSQL & "select RUC=isnull(i.cPersIDnro,''), p.cPersNombre, p.cPersDireccDomicilio, p.cPersTelefono, nPuntaje=0, " & _
           "nValor=0, Obtenido=0, fnFactorNro=0, cFactorDescripcion='',  cBSGrupoCod='" & pcBSGrupoCod & "', nFormula=0, " & _
           "cUnidades='', xnTipo=1, bGanador,pp.cDesesDescripcion " & _
           "from LogProSelPostorPropuesta pp " & _
           "inner join Persona p on p.cPersCod = pp.cPersCod " & _
           "left outer join PersID i on p.cPersCod = i.cPersCod and cPersIDTpo=2 " & _
           "Where nProSelNro = " & pnProSelNro & " And nProSelItem = " & pnProSelItem & " And bDesestimado = 1"
      sSQL = sSQL & " order by xnTipo, fnFactorNro"
           '" where x.nTipo=0 and nVigente=1 and  v.nProSelNro=" & pnProSelNro & " AND f.cBSGrupoCod='" & pcBSGrupoCod & "' "
           '" order by x.nTipo, f.nFactorNro"
    If ocon.AbreConexion Then
        Set rs = ocon.CargaRecordSet(sSQL)
'        If Not Rs.EOF Then
            Set CargarFactores = rs
        'End If
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
CargarFactoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

'Public Sub ExcelCuadro(xlHoja1 As Excel.Worksheet, ByVal X1 As Currency, ByVal Y1 As Currency, ByVal X2 As Currency, ByVal Y2 As Currency, Optional lbLineasVert As Boolean = True, Optional lbLineasHoriz As Boolean = False)
'xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).BorderAround xlContinuous, xlThin
'If lbLineasVert Then
'   If X2 <> X1 Then
'     xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).Borders(xlInsideVertical).LineStyle = xlContinuous
'   End If
'End If
'If lbLineasHoriz Then
'    If Y1 <> Y2 Then
'        xlHoja1.Range(xlHoja1.Cells(Y1, X1), xlHoja1.Cells(Y2, X2)).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'    End If
'End If
'End Sub

'"select nNroItem=count(nProSelItem) from LogProSelItem where nProSelNro=27"

Public Function ListaItem(ByVal pnProSelNro As Integer) As ADODB.Recordset
    On Error GoTo ListaItemErr
    Dim ocon As DConecta, rs As ADODB.Recordset, sSQL As String
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select i.cBSGrupoCod, g.cBSGrupoDescripcion, i.nProSelItem, i.nMonto " & _
                "from LogProSelItem i inner join BSGrupos g on i.cBSGrupoCod = g.cBSGrupoCod " & _
                "where i.nProSelNro=" & pnProSelNro
        Set rs = ocon.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            Set ListaItem = rs
        End If
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
ListaItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Function CargarDescalificado(ByVal pnProSelNro As Integer, ByVal pnProSelItem As Integer) As ADODB.Recordset
On Error GoTo CargarDescalificadoErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    sSQL = "select RUC=isnull(i.cPersIDnro,''), p.cPersNombre, p.cPersDireccDomicilio, p.cPersTelefono, " & _
           "bGanador,pp.cDesesDescripcion " & _
           "from LogProSelPostorPropuesta pp " & _
           "inner join Persona p on p.cPersCod = pp.cPersCod " & _
           "left outer join PersID i on p.cPersCod = i.cPersCod and cPersIDTpo=2 " & _
           "Where nProSelNro = " & pnProSelNro & " And nProSelItem = " & pnProSelItem & " And bDesestimado = 1 " & _
           "and not pp.cPersCod in (select distinct cPersCod from LogProSelEvalResultado where nProSelNro = " & pnProSelNro & " And nProSelItem = " & pnProSelItem & ")"
    Set ocon = New DConecta
    If ocon.AbreConexion() Then
        Set rs = ocon.CargaRecordSet(sSQL)
        Set CargarDescalificado = rs
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
CargarDescalificadoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Public Sub ImprimePostores(ByVal pnProSelNro As Integer, _
            ByVal pcProceso As String, ByVal pcDescripcion As String, _
            pcMoneda As String, pnMonto As Currency)
    
    On Error GoTo ImprimePostoresErr

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1(100) As Excel.Worksheet
    Dim lbExisteHoja  As Boolean
    
    Dim cBSGrupoCod As String
    Dim cBSGrupoDescripcion As String

    Dim lilineas As Integer
    Dim Total As Currency, TotalPos As Currency
    Dim nFactor As Integer
    Dim nCol As Integer, i As Integer, nLetra As Integer
    
    Dim CadTit As String
    Dim CadImp   As String, Ganador As Integer
    Dim rs As New ADODB.Recordset, RsItem As ADODB.Recordset, RsDesestimado As ADODB.Recordset
    Dim cGanadorNom As String, cGanadorRuc As String, _
        cGanadorDir As String, cGanadorTlf As String, _
        nNroHoja As Integer, nProSelItem As Integer, nMontoItem As Currency
    Dim lsNomHoja As String

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    
    nNroHoja = 1
        
    Set xlHoja1(nNroHoja) = xlLibro.Worksheets.Item(nNroHoja)
    
    xlHoja1(nNroHoja).Name = "Item " & nNroHoja 'cBSGrupoDescripcion
    '**************************************************************************

'         xlAplicacion.Range("A1:A1").ColumnWidth = 5
'         xlAplicacion.Range("B1:B1").ColumnWidth = 5
'         xlAplicacion.Range("C1:C1").ColumnWidth = 20
     
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 1), xlHoja1(nNroHoja).Cells(1, 1)).ColumnWidth = 5
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 2), xlHoja1(nNroHoja).Cells(1, 2)).ColumnWidth = 5
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 3), xlHoja1(nNroHoja).Cells(1, 3)).ColumnWidth = 20
     
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 6)).Font.Bold = True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(3, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(5, 2), xlHoja1(nNroHoja).Cells(5, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(5, 7)).Font.Bold = True

'******************************************************
            
    lilineas = 10

    Screen.MousePointer = 11
    xlHoja1(nNroHoja).Cells(2, 2) = "Caja Municipal de Ahorros y Creditos de Trujillo SA"
    xlHoja1(nNroHoja).Cells(3, 2) = pcProceso
    xlHoja1(nNroHoja).Cells(8, 7) = pcMoneda
    xlHoja1(nNroHoja).Cells(8, 8) = Format(pnMonto, "###,###,###.00")
    xlHoja1(nNroHoja).Cells(4, 2) = "ADQUISICION DE " & pcDescripcion
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(7, 8)).Merge False
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 2)).WrapText = True
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(8, 8), xlHoja1(nNroHoja).Cells(8, 5)).HorizontalAlignment = xlRight
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(8, 2), xlHoja1(nNroHoja).Cells(lilineas, 8)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 8)).HorizontalAlignment = xlCenter
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 8)).Cells.Interior.Color = &HC0FFFF
    
    xlHoja1(nNroHoja).Cells(lilineas, 2) = "Nombre"
    xlHoja1(nNroHoja).Cells(lilineas, 3) = "Direccion"
    xlHoja1(nNroHoja).Cells(lilineas, 4) = "Telefono"
    xlHoja1(nNroHoja).Cells(lilineas, 5) = "RUC"
    xlHoja1(nNroHoja).Cells(lilineas, 6) = "E Mail"
    xlHoja1(nNroHoja).Cells(lilineas, 7) = "Nro de Recibo"
    xlHoja1(nNroHoja).Cells(lilineas, 8) = "Fecha"
        
    Set rs = CargarPostores(pnProSelNro)
    Total = 0
    lilineas = lilineas + 2
    TotalPos = -1000
    Do While Not rs.EOF
        If rs!nPresentoProp Then
            xlHoja1(nNroHoja).Cells(lilineas, 2) = rs!cPersNombre
            xlHoja1(nNroHoja).Cells(lilineas, 3) = rs!cPersDireccDomicilio
            xlHoja1(nNroHoja).Cells(lilineas, 4) = rs!cPersTelefono
            xlHoja1(nNroHoja).Cells(lilineas, 5) = rs!RUC
            xlHoja1(nNroHoja).Cells(lilineas, 6) = rs!Email
            xlHoja1(nNroHoja).Cells(lilineas, 7) = rs!cNroRecibo
            xlHoja1(nNroHoja).Cells(lilineas, 8) = Format(rs!dFecha, "dd/mm/yyyy")
            lilineas = lilineas + 1
        End If
        rs.MoveNext
    Loop
    ExcelCuadro xlHoja1(nNroHoja), 2, 10, 8, lilineas - 1, True, True
    
    Screen.MousePointer = 0
'******************************************************

'        xlHoja1(nNroHoja).Cells.Select
    xlHoja1(nNroHoja).Cells.Font.Size = 8
    xlHoja1(nNroHoja).Cells.EntireColumn.AutoFit

    'Libera los objetos.
    Set xlHoja1(nNroHoja) = Nothing
    
    xlLibro.SaveAs App.path & "\spooler\Lista_Postores_" & pcProceso
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
    xlAplicacion.Windows(1).Visible = True
    Exit Sub
ImprimePostoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Err
End Sub

Private Function CargarPostores(ByVal pnProSelNro As Integer, Optional ByVal pcPersCod As String = "") As ADODB.Recordset
On Error GoTo CargarPostoresErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select x.cPersNombre, x.cPersDireccDomicilio, x.cPersTelefono, p.cNroRecibo, p.dFecha, Email=isnull(x.cPersEmail,''), RUC=isnull(cPersIDnro,''), nPresentoProp " & _
               " from LogProSelPostor p " & _
               "    inner join Persona x on p.cPersCod = x.cPersCod " & _
               "    left outer join PersID i on  x.cPersCod = i.cPersCod and len(i.cPersIDnro)=11 " & _
               " where   nProSelNro = " & pnProSelNro
        If pcPersCod <> "" Then sSQL = sSQL & " and x.cPersCod='" & pcPersCod & "'"
        Set rs = ocon.CargaRecordSet(sSQL)
        Set CargarPostores = rs
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
CargarPostoresErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Public Function CargarTipos() As ADODB.Recordset
On Error GoTo CargarTiposErr
    Dim ocon As DConecta, rs As ADODB.Recordset, sSQL As String
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        'sSQL = "select distinct cAbreviatura from LogProSelTpoRangos"
        sSQL = "select distinct cAbreviatura, cProSelTpoDescripcion from LogProSelTpoRangos r inner join LogProSelTpo t on r.nProSelTpoCod = t.nProSelTpoCod"
        Set rs = ocon.CargaRecordSet(sSQL)
        Set CargarTipos = rs
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
CargarTiposErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Private Function CargarProcesos(ByVal pcSintesis As String, ByVal pnAnio As Integer, ByVal pnmesI As Integer, ByVal pnmesF As Integer) As ADODB.Recordset
On Error GoTo CargarProcesosErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    sSQL = "select p.nProSelNro, p.nMoneda, r.cAbreviatura, y.cProSelTpoDescripcion, " & _
           " p.nNroProceso, p.nContratoCompra, p.cSintesis, p.nProSelMonto, p.dProSelFecha, " & _
           " Objeto=c.cConsDescripcion, FuenteFinanciamiento=x.cConsDescripcion, p.nProSelEstado " & _
           " from LogProcesoSeleccion p " & _
           " inner join LogProSelTpo y on p.nProselTpoCod = y.nProselTpoCod " & _
           " inner join LogProSelTpoRangos r on p.nProselTpoCod = r.nProselTpoCod and p.nProSelsubTpo = r.nProSelsubTpo " & _
           " inner join Constante c on p.nObjetoCod = c.nConsValor and c.nconscod=9048 " & _
           " inner join Constante x on p.nFuenteFinanciemiento = x.nConsValor and x.nconscod=9046" & _
           " where r.cAbreviatura='" & pcSintesis & "' and p.nPlanAnualAnio=" & pnAnio & " and nPlanAnualMes between " & pnmesI & " and " & pnmesF
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        Set rs = ocon.CargaRecordSet(sSQL)
        Set CargarProcesos = rs
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
CargarProcesosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function


'Public Sub ExcelAddHoja(psHojName As String, xlLibro As Excel.Workbook, xlHoja1 As Excel.Worksheet, Optional pbActivaHoja As Boolean = True)
'Dim lbExisteHoja As Boolean
'Dim lbBorrarRangos As Boolean
'On Error Resume Next
'lbExisteHoja = False
'lbBorrarRangos = False
'activaHoja:
'For Each xlHoja1 In xlLibro.Worksheets
'    If UCase(xlHoja1.Name) = UCase(psHojName) Then
'        If Not pbActivaHoja Then
'            SendKeys "{ENTER}"
'            xlHoja1.Delete
'        Else
'            xlHoja1.Activate
'            If lbBorrarRangos Then xlHoja1.Range("A1:BZ1").EntireColumn.Delete
'            lbExisteHoja = True
'        End If
'       Exit For
'    End If
'Next
'If Not lbExisteHoja Then
'    Set xlHoja1 = xlLibro.Worksheets.Add
'    xlHoja1.Name = psHojName
'    If Err Then
'        Err.Clear
'        pbActivaHoja = True
'        lbBorrarRangos = True
'        GoTo activaHoja
'    End If
'End If
'End Sub

Public Sub ImprimeListaProceso(ByVal pnAnio As Integer, ByVal pcAbreviatura As String, _
            Optional ByVal pnmesI As Integer = 1, Optional ByVal pnmesF As Integer = 12)

On Error GoTo ImprimeListaProcesoErr

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1(100) As Excel.Worksheet
    Dim lbExisteHoja  As Boolean
    
    Dim cBSGrupoCod As String
    Dim cBSGrupoDescripcion As String

    Dim lilineas As Integer
    Dim Total As Currency, TotalPos As Currency
    Dim nFactor As Integer
    Dim nCol As Integer, i As Integer, j As Integer
    
    Dim CadTit As String, sComite As String, sGanador As String, sPropuestaEconomica As String
    Dim CadImp   As String, sMontoItem As String
    Dim rs As New ADODB.Recordset, RsProceso As ADODB.Recordset, RsTmp As ADODB.Recordset
    Dim lsNomHoja As String, nNroHoja As Integer, sEstado As String

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    
    nNroHoja = 1
'    Set Rs = CargarTipos
'    Do While Not Rs.EOF
        'xlAplicacion.Name = App.Path & "\" & pcAbreviatura
        ExcelAddHoja pcAbreviatura & " - " & pnAnio, xlLibro, xlHoja1(nNroHoja), False
        
        xlHoja1(nNroHoja).Name = pcAbreviatura & " - " & pnAnio
        '**************************************************************************
         
         lilineas = 3
         
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 1), xlHoja1(nNroHoja).Cells(lilineas, 1)).RowHeight = 1
         
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).Font.Bold = True
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).WrapText = True
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).RowHeight = 68.25
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).ColumnWidth = 10.71
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).HorizontalAlignment = xlCenter
         xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).VerticalAlignment = xlCenter
    
    '******************************************************
    
        Screen.MousePointer = 11
        xlHoja1(nNroHoja).Cells(lilineas, 2) = "ITEM"
        xlHoja1(nNroHoja).Cells(lilineas, 3) = "MODALIDAD y NÚMERO DE PROCESO"
        xlHoja1(nNroHoja).Cells(lilineas, 4) = "OBJETO DEL PROCESO"
        xlHoja1(nNroHoja).Cells(lilineas, 5) = "VALOR REFERENCIAL"
        xlHoja1(nNroHoja).Cells(lilineas, 6) = "SISTEMA DE CONTRATACION"
        xlHoja1(nNroHoja).Cells(lilineas, 7) = "FECHA DE CONVOCATORIA"
        xlHoja1(nNroHoja).Cells(lilineas, 8) = "MIEMBROS DE COMITÉ"
        xlHoja1(nNroHoja).Cells(lilineas, 9) = "NOMBRE DEL POSTOR GANADOR"
        xlHoja1(nNroHoja).Cells(lilineas, 10) = "FECHA DE BUENA PRO"
        xlHoja1(nNroHoja).Cells(lilineas, 11) = "FECHA DE REALIZACION / FECHA DE CONSENTIMIENTO BUENA PRO"
        xlHoja1(nNroHoja).Cells(lilineas, 12) = "FUENTE DE FINANCIAMIENTO"
        xlHoja1(nNroHoja).Cells(lilineas, 13) = "MONTO ADJUDICADO"
        xlHoja1(nNroHoja).Cells(lilineas, 14) = "MONTO DEL CONTRATO"
        xlHoja1(nNroHoja).Cells(lilineas, 15) = "TIEMPO DE ENTREGA DEL BIEN / EJECUCION DEL SERVICIO"
        xlHoja1(nNroHoja).Cells(lilineas, 16) = "INICIO"
        xlHoja1(nNroHoja).Cells(lilineas, 17) = "TERMINO"
        xlHoja1(nNroHoja).Cells(lilineas, 18) = "PENALIDAD"
        xlHoja1(nNroHoja).Cells(lilineas, 19) = "COSTO TOTAL"
        xlHoja1(nNroHoja).Cells(lilineas, 20) = "CARTA  FIANZA"
        xlHoja1(nNroHoja).Cells(lilineas, 21) = "DESTINO DEL BIEN"
        xlHoja1(nNroHoja).Cells(lilineas, 22) = "UBICACIÓN"
        xlHoja1(nNroHoja).Cells(lilineas, 23) = "SITUACIÓN ACTUAL"
        xlHoja1(nNroHoja).Cells(lilineas, 24) = "N° DE CONTRATO"
        
        lilineas = lilineas + 1
        i = 1
        Set RsProceso = CargarProcesos(pcAbreviatura, pnAnio, pnmesI, pnmesF)
        Do While Not RsProceso.EOF
            
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).WrapText = True
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).RowHeight = 68.25
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).ColumnWidth = 10.71
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).HorizontalAlignment = xlCenter
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 24)).VerticalAlignment = xlCenter
            
            xlHoja1(nNroHoja).Cells(lilineas, 2) = i 'RsProceso!nNroProceso
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 3), xlHoja1(nNroHoja).Cells(lilineas, 3)).ColumnWidth = 18#
            xlHoja1(nNroHoja).Cells(lilineas, 3) = RsProceso!cProSelTpoDescripcion & "N° " & RsProceso!nNroProceso & " " & pnAnio & " CMAC-T"
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 4), xlHoja1(nNroHoja).Cells(lilineas, 4)).HorizontalAlignment = xlLeft
            xlHoja1(nNroHoja).Cells(lilineas, 4) = RsProceso!cSintesis
            Set RsTmp = CargarMontoItem(RsProceso!nProselNro)
            sMontoItem = ""
            Do While Not RsTmp.EOF
                sMontoItem = sMontoItem & " Item " & RsTmp!nProSelItem & ": " & IIf(RsProceso!nMoneda = 1, "S/. ", "$ ") & FNumero(RsTmp!nMonto) & vbCrLf
                RsTmp.MoveNext
            Loop
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 5), xlHoja1(nNroHoja).Cells(lilineas, 5)).HorizontalAlignment = xlLeft
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 5), xlHoja1(nNroHoja).Cells(lilineas, 5)).ColumnWidth = 17#
            xlHoja1(nNroHoja).Cells(lilineas, 5) = sMontoItem & " Total: " & IIf(RsProceso!nMoneda = 1, "S/. ", "$ ") & FNumero(RsProceso!nProSelMonto)
            xlHoja1(nNroHoja).Cells(lilineas, 6) = IIf(RsProceso!nContratoCompra, "Suma Alzada", "Por Item")
            sComite = ""
            Set RsTmp = CargarConvocatoria_FechaComite(RsProceso!nProselNro)
            If Not RsTmp.EOF Then
                xlHoja1(nNroHoja).Cells(lilineas, 7) = Format(RsTmp!dFechaInicio, "dd/mm/yyyy")
                j = 0
                Do While Not RsTmp.EOF
                    j = j + 1
                    sComite = sComite & "  " & j & "- " & Trim(RsTmp!cPersNombre) & vbCrLf
                    RsTmp.MoveNext
                Loop
                xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 8), xlHoja1(nNroHoja).Cells(lilineas, 8)).HorizontalAlignment = xlLeft
                xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 8), xlHoja1(nNroHoja).Cells(lilineas, 8)).ColumnWidth = 30#
                xlHoja1(nNroHoja).Cells(lilineas, 8) = sComite
            End If
            Set RsTmp = Nothing
            sGanador = "": sPropuestaEconomica = ""
            Set RsTmp = CargarGanador(RsProceso!nProselNro)
            Do While Not RsTmp.EOF
                j = j + 1
                If sGanador = "" Then
                    sGanador = sGanador & "  Item " & RsTmp!nProSelItem & ": " & Trim(RsTmp!cPersNombre)
                Else
                    sGanador = sGanador & vbCrLf & "  Item " & RsTmp!nProSelItem & ": " & Trim(RsTmp!cPersNombre)
                End If
                sPropuestaEconomica = sPropuestaEconomica & "  Item " & RsTmp!nProSelItem & " " & IIf(RsTmp!nMoneda = 1, "S/.", "$") & " " & FNumero(RsTmp!nPropEconomica) & vbCrLf
                RsTmp.MoveNext
            Loop
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 9), xlHoja1(nNroHoja).Cells(lilineas, 9)).HorizontalAlignment = xlLeft
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 9), xlHoja1(nNroHoja).Cells(lilineas, 9)).ColumnWidth = 33#
            xlHoja1(nNroHoja).Cells(lilineas, 9) = sGanador
            Set RsTmp = CargarFechasBuenaPro(RsProceso!nProselNro)
            If Not RsTmp.EOF Then
                xlHoja1(nNroHoja).Cells(lilineas, 10) = Format(RsTmp!dFechaInicio, "dd/mm/yyyy")
                RsTmp.MoveNext
            End If
            If Not RsTmp.EOF Then xlHoja1(nNroHoja).Cells(lilineas, 11) = Format(RsTmp!dFechaInicio, "dd/mm/yyyy")
            Set RsTmp = Nothing
            xlHoja1(nNroHoja).Cells(lilineas, 12) = RsProceso!FuenteFinanciamiento
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 13), xlHoja1(nNroHoja).Cells(lilineas, 14)).HorizontalAlignment = xlLeft
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 13), xlHoja1(nNroHoja).Cells(lilineas, 14)).ColumnWidth = 20#
            xlHoja1(nNroHoja).Cells(lilineas, 13) = sPropuestaEconomica
            xlHoja1(nNroHoja).Cells(lilineas, 14) = sPropuestaEconomica
            xlHoja1(nNroHoja).Cells(lilineas, 15) = "xxx" 'RsProceso!
            xlHoja1(nNroHoja).Cells(lilineas, 16) = "xxx" 'RsProceso!
            xlHoja1(nNroHoja).Cells(lilineas, 17) = "xxx" 'RsProceso!
            xlHoja1(nNroHoja).Cells(lilineas, 18) = "-"
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 19), xlHoja1(nNroHoja).Cells(lilineas, 19)).HorizontalAlignment = xlLeft
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 19), xlHoja1(nNroHoja).Cells(lilineas, 19)).ColumnWidth = 20#
            xlHoja1(nNroHoja).Cells(lilineas, 19) = sPropuestaEconomica
            xlHoja1(nNroHoja).Cells(lilineas, 20) = "xxx" 'RsProceso!
            xlHoja1(nNroHoja).Cells(lilineas, 21) = "xxx" 'RsProceso!
            xlHoja1(nNroHoja).Cells(lilineas, 22) = "xxx" 'RsProceso!
            Select Case RsProceso!nProSelEstado
                Case -1
                    sEstado = "Anulado"
                Case 0
                    sEstado = "Terminado"
                Case 1
                    sEstado = "Proceso"
            End Select
            xlHoja1(nNroHoja).Cells(lilineas, 23) = sEstado
            xlHoja1(nNroHoja).Cells(lilineas, 24) = "xxx" 'RsProceso!
            RsProceso.MoveNext
            lilineas = lilineas + 1
            i = i + 1
        Loop
        
        ExcelCuadro xlHoja1(nNroHoja), 2, 3, 24, lilineas - 1, True, True
'        Rs.MoveNext
        xlHoja1(nNroHoja).Cells.Select
        xlHoja1(nNroHoja).Cells.Font.Size = 8
        xlHoja1(nNroHoja).Cells.EntireColumn.AutoFit
'        nNroHoja = nNroHoja + 1
'    Loop
    Screen.MousePointer = 0
'******************************************************

    'Libera los objetos.
    Set xlHoja1(nNroHoja) = Nothing
    xlLibro.SaveAs App.path & "\spooler\" & pcAbreviatura & " - " & " A " & Format(pnmesF, "00") & " - " & pnAnio
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
    xlAplicacion.Windows(1).Visible = True
    Exit Sub
ImprimeListaProcesoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Private Function CargarMontoItem(ByVal nProselNro As Integer) As ADODB.Recordset
On Error GoTo CargarMontoItemErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    sSQL = "select nProSelItem, nMonto from LogProSelItem where nProSelNro = " & nProselNro
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        Set rs = ocon.CargaRecordSet(sSQL)
        Set CargarMontoItem = rs
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
CargarMontoItemErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Private Function CargarConvocatoria_FechaComite(ByVal pnProSelNro As Integer) As ADODB.Recordset
On Error GoTo CargarConvocatoria_FechaComiteErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select cPersNombre=replace(p.cPersNombre,'/',' '), e.dFechaInicio, e.dFechaTermino from LogProSelComite c " & _
               " inner join LogProSelEtapa e on c.nProSelNro = e.nProSelNro " & _
               " inner join persona p on c.cPersCod = p.cPersCod " & _
               " inner join LogProSelEtapaComite y on c.nProSelNro = y.nProSelNro and p.cPersCod= y.cPersCod and e.nEtapaCod = y.nEtapaCod and e.nEtapaCod = 1 " & _
               " where c.nProSelNro = " & pnProSelNro
        Set rs = ocon.CargaRecordSet(sSQL)
        Set CargarConvocatoria_FechaComite = rs
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
CargarConvocatoria_FechaComiteErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Private Function CargarFechasBuenaPro(ByVal pnProSelNro As Integer) As ADODB.Recordset
On Error GoTo CargarFechasBuenaProErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "SELECT  dFechaInicio FROM LogProSelEtapa WHERE (nEtapaCod = 11 or nEtapaCod = 16) " & _
               " and nProSelNro = " & pnProSelNro & " order by nEtapaCod"
        Set rs = ocon.CargaRecordSet(sSQL)
        Set CargarFechasBuenaPro = rs
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
CargarFechasBuenaProErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function


Private Function CargarGanador(ByVal pnProSelNro As Integer) As ADODB.Recordset
On Error GoTo CargarGanadorErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select p.nProSelItem, cPersNombre=replace(x.cPersNombre,'/',' '), p.nMoneda, p.nPropEconomica " & _
                " from LogProSelPostorPropuesta p " & _
                " inner join Persona x on p.cPersCod = x.cPersCod " & _
                " where nProSelNro=" & pnProSelNro & " and bGanador=1"
        Set rs = ocon.CargaRecordSet(sSQL)
        Set CargarGanador = rs
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
CargarGanadorErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

'*************************************************************************************************
'*************************************************************************************************
'********************************************word*************************************************
'*************************************************************************************************
'*************************************************************************************************

Function ImpConcentimientoBuenaProWORD(ByVal pdFecha As String, ByVal psDir As String, ByVal psProceso As String, _
        ByVal psSintesis As String, ByVal pnValor_Ref As Currency, ByVal psPostor As String, ByVal pnPorcentaje As Currency, _
        ByVal pnMonto As Currency, pdFechaF As String) As Boolean

On Error GoTo ImpConcentimientoBuenaProWORDErr

Dim aLista() As String
Dim vFilas As Integer
Dim vFecAviso As Date
Dim K As Integer
Dim CadenaAna As String

Dim sSQL As String
Dim rsCarta As New ADODB.Recordset

Dim lsModelo As String
Dim vCont As Integer
Dim lrPersRel As ADODB.Recordset
Dim lsRuta As String

lsRuta = ""
lsModelo = App.path & "\Plantillas\PConcentimientoBuenaPro.doc"

lsRuta = Dir(lsModelo)

If Len(lsRuta) = 0 Then
   MsgBox "No se halla " + lsModelo + Space(10), vbInformation, "Aviso"
   Exit Function
End If

'CadenaAna = Mid(CadenaAna, 1, (Len(CadenaAna) - 1))
    
    vFilas = 0
    
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
        
    
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModelo
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
        
    wAppSource.ActiveDocument.Content.Copy
    
    'Crea Nuevo Documento
    wApp.Documents.Add
'   wApp.Application.Selection.TypeParagraph
    wApp.Application.Selection.Paste
'   wApp.Application.Selection.InsertBreak
        
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd
        
        With wApp.Selection.Find
            .Text = "<Fecha>"
            .Replacement.Text = pdFecha
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<Síntesis>"
            .Replacement.Text = psSintesis
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<postor>"
            .Replacement.Text = psPostor
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<Dir>"
            .Replacement.Text = psDir
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<proceso>"
            .Replacement.Text = psProceso
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<tpostor>"
            .Replacement.Text = psPostor ' & " " & psDir
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<valorRef>"
            .Replacement.Text = FNumero(pnValor_Ref)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<Monto>"
            .Replacement.Text = FNumero(pnMonto)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<Por>"
            .Replacement.Text = (pnPorcentaje * 100)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<FechaP>"
            .Replacement.Text = pdFechaF
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
'wapp.Selection.Font
wAppSource.ActiveDocument.Close
wApp.ActiveDocument.SaveAs App.path + "\spooler\CONCENTIMIENTO BUENA PRO" & psProceso + ".doc"
wApp.Visible = True
    Exit Function
ImpConcentimientoBuenaProWORDErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Public Sub ImpConsultasWord(ByVal pnProSelNro As Integer)
On Error GoTo ImpConsultasWordErr
Dim MSWord As New Word.Application
Dim Documento As Word.Document
Dim Parrafo As Paragraph
Dim Parrafo1 As Paragraph
Set Documento = MSWord.Documents.Add
Set Parrafo = Documento.Paragraphs.Add
Set Parrafo1 = Documento.Paragraphs.Add

Dim strCadena As String
Dim rsComite As New ADODB.Recordset
Dim rsFase As New ADODB.Recordset
Dim rsCriterio As New ADODB.Recordset
Dim rsDatosProceso As New ADODB.Recordset
Dim intFila As Integer

'***********************************************

Dim ocon As DConecta, rs As ADODB.Recordset, sSQL As String, sPostor As String, NomArch As String
Set ocon = New DConecta

strCadena = Format(gdFecSis, "Trujillo, dd - mmmm - yyyy") & vbCrLf & vbCrLf

strCadena = strCadena & "Carta No " & vbCrLf & vbCrLf

strCadena = strCadena & "Señores: " & vbCrLf
strCadena = strCadena & "POSTORES EN PROCESO DE SELECCION " & vbCrLf
strCadena = strCadena & "Ciudad.-" & vbCrLf & vbCrLf

If ocon.AbreConexion Then
    sSQL = "select cPersNombre=replace(p.cPersNombre,'/',' '), cProSelTpoDescripcion, nNroProceso, nPlanAnualAnio, " & _
           " c.cConsulta , c.cRespuesta " & _
           " from LogProcesoSeleccion x " & _
           " inner join LogProSelTpo t on x.nProSelTpoCod = t.nProSelTpoCod " & _
           " inner join LogProSelItem b on x.nProSelNro = b.nProSelNro " & _
           " inner join LogProSelConsultas c on x.nProSelNro = c.nProSelNro " & _
           " inner join Persona p  on c.cPersCod = p.cPersCod " & _
           " where  (cResolucion1 <> '' or cResolucion2 <> '' or cResolucion3 <> '') and " & _
           " x.nProSelNro= " & pnProSelNro
    Set rs = ocon.CargaRecordSet(sSQL)
    If Not rs.EOF Then
        NomArch = rs!cProSelTpoDescripcion & " N° " & rs!nNroProceso & "-" & rs!nPlanAnualAnio & "-CMAC-T S.A."
        strCadena = strCadena & "Ref: " & rs!cProSelTpoDescripcion & " N° " & rs!nNroProceso & "-" & rs!nPlanAnualAnio & "-CMAC-T S.A." & vbCrLf
        strCadena = strCadena & vbCrLf
        
        strCadena = strCadena & "De nuestra consideración:" & vbCrLf
        strCadena = strCadena & "Tenemos el agrado de dirigirnos a ustedes con el objeto de saludarlos, y hacer propicia la oportunidad para alcanzarles la absolución de las consultas hechas a las bases del proceso de selección de la referencia. " & vbCrLf

        Do While Not rs.EOF
            If sPostor <> rs!cPersNombre Then
                sPostor = rs!cPersNombre
                strCadena = strCadena & vbCrLf
                strCadena = strCadena & "Postor: " & sPostor & vbCrLf
            End If
            strCadena = strCadena & vbCrLf & "Consulta " & vbCrLf & vbCrLf
            strCadena = strCadena & rs!cConsulta & vbCrLf & vbCrLf
            strCadena = strCadena & vbCrLf & "Absolucion " & vbCrLf & vbCrLf
            strCadena = strCadena & rs!cRespuesta & vbCrLf & vbCrLf
            rs.MoveNext
        Loop
    Else
        MsgBox "No Existen Consultas Absueltas y Aprobadas para Imprimir...", vbInformation, "Aviso"
        Exit Sub
    End If
    ocon.CierraConexion
End If

strCadena = strCadena & vbCrLf & vbCrLf & vbCrLf & "Atentamente,"

Parrafo.Range.Font.Name = "Arial"
Parrafo.Range.Font.Size = 9
Parrafo.Space1
Parrafo.Alignment = wdAlignParagraphJustify
'Parrafo.Range.Font.Underline = wdUnderlineDouble
Parrafo.Range.InsertBefore strCadena

MSWord.ActiveDocument.SaveAs App.path + "\spooler\CONCLUSIONES " + NomArch + " .doc"
MSWord.Visible = True

Set MSWord = Nothing
Set Documento = Nothing
Set Parrafo = Nothing
    Exit Sub
ImpConsultasWordErr:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Public Sub ImpObservacionesWord(ByVal pnProSelNro As Integer)
On Error GoTo ImpObservacionesWordErr
Dim MSWord As New Word.Application
Dim Documento As Word.Document
Dim Parrafo As Paragraph
Dim Parrafo1 As Paragraph
Set Documento = MSWord.Documents.Add
Set Parrafo = Documento.Paragraphs.Add
Set Parrafo1 = Documento.Paragraphs.Add

Dim strCadena As String
Dim rsComite As New ADODB.Recordset
Dim rsFase As New ADODB.Recordset
Dim rsCriterio As New ADODB.Recordset
Dim rsDatosProceso As New ADODB.Recordset
Dim intFila As Integer

'***********************************************

Dim ocon As DConecta, rs As ADODB.Recordset, sSQL As String, sPostor As String
Set ocon = New DConecta

strCadena = Format(gdFecSis, "Trujillo, dd - mmmm - yyyy") & vbCrLf & vbCrLf

strCadena = strCadena & "Carta No " & vbCrLf & vbCrLf

strCadena = strCadena & "Señores: " & vbCrLf
strCadena = strCadena & "POSTORES EN PROCESO DE SELECCION " & vbCrLf
strCadena = strCadena & "Ciudad.-" & vbCrLf & vbCrLf

If ocon.AbreConexion Then
    sSQL = "select distinct cPersNombre=replace(p.cPersNombre,'/',' '), cProSelTpoDescripcion, nNroProceso, nPlanAnualAnio, " & _
           " c.cObservacion , c.cRespuesta " & _
           " from LogProcesoSeleccion x " & _
           " inner join LogProSelTpo t on x.nProSelTpoCod = t.nProSelTpoCod " & _
           " inner join LogProSelItem b on x.nProSelNro = b.nProSelNro " & _
           " inner join LogProSelObsBases c on x.nProSelNro = c.nProSelNro " & _
           " inner join Persona p  on c.cPersCod = p.cPersCod " & _
           " where (cResolucion1 <> '' or cResolucion2 <> '' or cResolucion3 <> '') and " & _
           " x.nProSelNro= " & pnProSelNro
    Set rs = ocon.CargaRecordSet(sSQL)
    If Not rs.EOF Then
        strCadena = strCadena & "Ref: " & rs!cProSelTpoDescripcion & " N° " & rs!nNroProceso & "-" & rs!nPlanAnualAnio & "-CMAC-T S.A." & vbCrLf
        strCadena = strCadena & vbCrLf
        
        strCadena = strCadena & "De nuestra consideración:" & vbCrLf
        strCadena = strCadena & "Tenemos el agrado de dirigirnos a ustedes con el objeto de saludarlos, y hacer propicia la oportunidad para alcanzarles la absolución de las consultas hechas a las bases del proceso de selección de la referencia. " & vbCrLf

        Do While Not rs.EOF
            If sPostor <> rs!cPersNombre Then
                sPostor = rs!cPersNombre
                strCadena = strCadena & vbCrLf
                strCadena = strCadena & "Postor: " & sPostor & vbCrLf
            End If
            strCadena = strCadena & vbCrLf & "Observacion " & vbCrLf & vbCrLf
            strCadena = strCadena & rs!cObservacion & vbCrLf & vbCrLf
            strCadena = strCadena & vbCrLf & "Absolucion " & vbCrLf & vbCrLf
            strCadena = strCadena & rs!cRespuesta & vbCrLf & vbCrLf
            rs.MoveNext
        Loop
    Else
        MsgBox "No Existen Observaciones Absueltas y Aprobadas para Imprimir...", vbInformation, "Aviso"
        Exit Sub
    End If
    ocon.CierraConexion
End If

strCadena = strCadena & vbCrLf & vbCrLf & vbCrLf & "Atentamente,"

Parrafo.Range.Font.Name = "Arial"
Parrafo.Range.Font.Size = 9
Parrafo.Space1
Parrafo.Alignment = wdAlignParagraphJustify
'Parrafo.Range.Font.Underline = wdUnderlineDouble
Parrafo.Range.InsertBefore strCadena
MSWord.ActiveDocument.SaveAs App.path + "\spoller\OBSERVACIONES " + rs!cProSelTpoDescripcion & " N° " & rs!nNroProceso & "-" & rs!nPlanAnualAnio & "-CMAC-T S.A. .doc"
MSWord.Visible = True

Set MSWord = Nothing
Set Documento = Nothing
Set Parrafo = Nothing
Exit Sub
ImpObservacionesWordErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Sub

Function ImpBasesWORD(ByVal psObjeto As String, ByVal psValor_Ref As String, ByVal pnAño As Integer, ByVal psMoneda As String, _
                    ByVal psTipo As String, ByVal psnro_Proceso As String, ByVal psProceso As String, psEtapas As String, _
                    ByVal psPlazo As String, ByVal pnMin As Currency, ByVal pnMax As Currency, ByVal pnProSelNro As Integer, _
                    ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pcMes As String, pnCosto As Currency, _
                    ByVal pnMinPtos As Integer, ByVal pcObj As String, pcMonedaCostoBases As String) As String

On Error GoTo ImpBasesWORDErr

Dim lsModeloPlantilla As String
Dim Nombre As String, nI As Integer, lnPos As Integer, lsEtapa As String, lsPlazo As String, lsFactores As String, lsPuntaje As String

lsModeloPlantilla = App.path & "\spooler\Plantillas\BASES_" & pnProSelTpoCod & pnProSelSubTpo & "_" & pcObj & ".doc"
'lsModeloPlantilla = App.Path & "\spooler\Plantillas\BASES_41_.doc"
    
    'Crea una clase que de Word Object
    Dim wApp As Word.Application
    Dim wAppSource As Word.Application
    'Create a new instance of word
    Set wApp = New Word.Application
    Set wAppSource = New Word.Application
    
    Dim RangeSource As Word.Range
    'Abre Documento Plantilla
    wAppSource.Documents.Open FileName:=lsModeloPlantilla
    Set RangeSource = wAppSource.ActiveDocument.Content
    'Lo carga en Memoria
    wAppSource.ActiveDocument.Content.Copy
    
    'Crea Nuevo Documento
    wApp.Documents.Add

'        wApp.Application.Selection.TypeParagraph
        wApp.Application.Selection.PasteFormat
'        wApp.Application.Selection.InsertBreak
        
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd
        
        With wApp.Selection.Find
            .Text = "<MinPtos>"
            .Replacement.Text = pnMinPtos
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<mes>"
            .Replacement.Text = pcMes
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<Costo>"
            .Replacement.Text = pnCosto
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<MCostoBases>"
            .Replacement.Text = pcMonedaCostoBases
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<Objeto>"
            .Replacement.Text = psObjeto
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<min>"
            .Replacement.Text = pnMin
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<max>"
            .Replacement.Text = pnMax
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        wApp.Selection.CreateAutoTextEntry("yo", "nose") = psEtapas
        
        nI = Right(psPlazo, 2)
                
        '***********************************************************************************
        '****************************************************************************
        
        lnPos = InStr(1, psEtapas, vbCrLf)
        lsEtapa = Mid(psEtapas, 1, lnPos - 1)
        lsEtapa = lsEtapa & Space(50) & "<Etapa>"
        psEtapas = Mid(psEtapas, lnPos + 2)
        
        Do While lnPos > -1
            With wApp.Selection.Find
                .Text = "<Etapa>"
                .Replacement.Text = lsEtapa
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            If lnPos = 0 Then Exit Do
            lnPos = InStr(1, psEtapas, vbCrLf)
            If lnPos = 0 Then
                lsEtapa = ""
            Else
                lsEtapa = Mid(psEtapas, 1, lnPos - 1)
                lsEtapa = lsEtapa & Space(50) & "<Etapa>"
                psEtapas = Mid(psEtapas, lnPos + 2)
            End If
        Loop
        
        '****************************************************************************
        
        lnPos = InStr(1, psPlazo, vbCrLf)
        lsPlazo = Mid(psPlazo, 1, lnPos - 1)
        lsPlazo = lsPlazo & Space(35) & "<Plazo>"
        psPlazo = Mid(psPlazo, lnPos + 2)
        
        Do While lnPos > -1
            With wApp.Selection.Find
                .Text = "<Plazo>"
                .Replacement.Text = lsPlazo
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
              End With
            wApp.Selection.Find.Execute Replace:=wdReplaceAll
            
            If lnPos = 0 Then Exit Do
            lnPos = InStr(1, psPlazo, vbCrLf)
            If lnPos = 0 Then
                lsPlazo = ""
            Else
                lsPlazo = Mid(psPlazo, 1, lnPos - 1)
                lsPlazo = lsPlazo & Space(50) & "<Plazo>"
                psPlazo = Mid(psPlazo, lnPos + 2)
            End If
        Loop
'        wApp.Documents.Item(1).Paragraphs.Item(29).Range.InsertBefore psEtapas '26
'        wApp.Documents.Item(1).Paragraphs.Item(29 + nI).Range.InsertBefore Mid(psPlazo, 1, Len(psPlazo) - 2) '27

        '****************************************************************************
        '*****************************************************************************************
        
        With wApp.Selection.Find
            .Text = "<proceso>"
            .Replacement.Text = psProceso
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<Valor_Ref>"
            .Replacement.Text = psValor_Ref
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<año>"
            .Replacement.Text = pnAño
            .Forward = True
            .Wrap = wdFindContinue
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<moneda>"
            .Replacement.Text = psMoneda
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<Tipo>"
            .Replacement.Text = psTipo
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        With wApp.Selection.Find
            .Text = "<minimo>"
            .Replacement.Text = FNumero(psValor_Ref * pnMin / 100#)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
                
        With wApp.Selection.Find
            .Text = "<maximo>"
            .Replacement.Text = FNumero(psValor_Ref * pnMax / 100#)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<nro_proceso>"
            .Replacement.Text = psnro_Proceso
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        '**********************************************************************************************************
        Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset, sFactores As String, sPuntos As String, _
            Item As Integer, nPar As Integer, maxt As Integer, maxe As Integer
        Set ocon = New DConecta
        If ocon.AbreConexion Then
            sSQL = "select i.nProSelItem, f.cBSGrupoCod, g.cBSGrupoDescripcion, x.nFactorNro, nProSelTpoCod, nProSelSubTpo, " & _
                   "    x.cFactorDescripcion , x.cUnidades, x.nTipo, f.nPuntaje, nFormula, nObjeto, f.cBSGrupoCod, cUnidades " & _
                   " from LogProSelEvalFactor f " & _
                   "    inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
                   "    inner join BSGrupos g on f.cBSGrupoCod = g.cBSGrupoCod " & _
                   "    inner join LogProSelItem i on f.cBSGrupoCod = i.cBSGrupoCod and f.nProSelNro = i. nProSelNro " & _
                   " Where f.nVigente = 1 and i.nProSelNro = " & pnProSelNro & " and f.nProSelTpoCod=" & pnProSelTpoCod & " and f.nProSelSubTpo=" & pnProSelSubTpo & " " & _
                   " order by i.nProSelItem"
            nPar = 86 + nI
            Set rs = ocon.CargaRecordSet(sSQL)
            Do While Not rs.EOF
                If Item <> rs!nProSelItem Then
                    Item = rs!nProSelItem
                    maxt = 0: maxe = 0
                    sFactores = sFactores & rs!cBSGrupoDescripcion & vbCrLf '& vbCrLf
                    sPuntos = sPuntos & "-" & vbCrLf '& "-" & vbCrLf  '& vbCrLf
                    nPar = nPar + 2 '3
                End If
                sFactores = sFactores & Space(2) & rs!cFactorDescripcion & vbCrLf
                Select Case rs!nFormula
                    Case 0
                        'Directamente
                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " '& vbCrLf
                        sFactores = sFactores & "maximo valor" & vbCrLf
                        nPar = nPar + 3
                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & "-" & vbCrLf '& " -" & vbCrLf & " -" & vbCrLf
                    Case 1
                        'Inversamente
                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " '& vbCrLf
                        sFactores = sFactores & "minimo valor" & vbCrLf
                        nPar = nPar + 3
                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & "-" & vbCrLf '& " -" & vbCrLf & " -" & vbCrLf
                    Case 2
                        'Rangos
                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & "-" '& vbCrLf
                        sFactores = sFactores & CargarRangos(rs!nFactorNro, rs!nProSelTpoCod, rs!nProSelSubTpo, rs!nObjeto, rs!cBSGrupoCod, nPar, rs!cUnidades, sPuntos, pnProSelNro)
                        nPar = nPar + 1
                    Case 3
                        'SIno
                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje a los postores que cumplan  " '& vbCrLf
                        sFactores = sFactores & "con lo pedido" & vbCrLf
                        nPar = nPar + 3
                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & "-" & vbCrLf '& " -" & vbCrLf & " -" & vbCrLf
                End Select
'                If Rs!nTipo Then
'                    maxe = maxe + Rs!npuntaje
'                Else
'                    maxt = maxt + Rs!npuntaje
'                End If
                rs.MoveNext
            Loop
            ocon.CierraConexion
            
            lnPos = InStr(1, sFactores, vbCrLf)
            lsFactores = Mid(sFactores, 1, lnPos - 1)
            lsFactores = lsFactores & Space(100) & "<Factores>"
            sFactores = Mid(sFactores, lnPos + 2)
            
            Do While lnPos > -1
                With wApp.Selection.Find
                    .Text = "<Factores>"
                    .Replacement.Text = lsFactores
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                  End With
                wApp.Selection.Find.Execute Replace:=wdReplaceAll
                If lnPos = 0 Then Exit Do
                lnPos = InStr(1, sFactores, vbCrLf)
                If lnPos = 0 Then
                    lsFactores = ""
                Else
                    lsFactores = Mid(sFactores, 1, lnPos - 1)
                    lsFactores = lsFactores & Space(100) & "<Factores>"
                    sFactores = Mid(sFactores, lnPos + 2)
                End If
            Loop
            
            lnPos = InStr(1, sPuntos, vbCrLf)
            lsPuntaje = Mid(sPuntos, 1, lnPos - 1)
            lsPuntaje = lsPuntaje & Space(100) & "<Puntaje>"
            sPuntos = Mid(sPuntos, lnPos + 2)
            If Mid(lsPuntaje, 1, 2) <> "- " Then lsPuntaje = Mid(lsPuntaje, 2)
            
            Do While lnPos > -1
                With wApp.Selection.Find
                    .Text = "<Puntaje>"
                    .Replacement.Text = lsPuntaje
                    .Forward = True
                    .Wrap = wdFindContinue
                    .Format = False
                  End With
                wApp.Selection.Find.Execute Replace:=wdReplaceAll
                If lnPos = 0 Then Exit Do
                lnPos = InStr(1, sPuntos, vbCrLf)
                If lnPos = 0 Then
                    lsPuntaje = ""
                Else
                    lsPuntaje = Mid(sPuntos, 1, lnPos - 1)
                    lsPuntaje = lsPuntaje & Space(100) & "<Puntaje>"
                    sPuntos = Mid(sPuntos, lnPos + 2)
                    If Mid(lsPuntaje, 1, 2) <> "- " And Val(Mid(lsPuntaje, 1, 2)) < 0 Then lsPuntaje = Mid(lsPuntaje, 2)
                End If
            Loop
            
'            wApp.Documents.Item(1).Paragraphs.Item(86 + nI).Range.InsertBefore sFactores
'            wApp.Documents.Item(1).Paragraphs.Item(nPar + 1).Range.InsertBefore sPuntos
            
        End If
        '**********************************************************************************************************

        With wApp.Selection.Find
            .Text = "<maxTecnica>"
            .Replacement.Text = maxt
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<maxEconomica>"
            .Replacement.Text = maxe
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'wapp.Selection.Font
wAppSource.ActiveDocument.Close
Nombre = App.path + "\spooler\BASE " & psnro_Proceso + ".doc"
wApp.ActiveDocument.SaveAs Nombre
ActualizaBases pnProSelNro, Nombre
wApp.Visible = True
ImpBasesWORD = Nombre
Exit Function
ImpBasesWORDErr:
    ImpBasesWORD = ""
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function


'Function ImpBasesWORD_11_Bienes(ByVal psObjeto As String, ByVal psValor_Ref As String, ByVal pnAño As Integer, ByVal psMoneda As String, _
'                    ByVal psTipo As String, ByVal psnro_Proceso As String, ByVal psProceso As String, psEtapas As String, _
'                    ByVal psPlazo As String, ByVal pnMin As Currency, ByVal pnMax As Currency, ByVal pnProSelNro As Integer, _
'                    ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pcMes As String, pnCosto As Currency, _
'                    ByVal pnMinPtos As Integer, ByVal pcObj As String, pcMonedaCostoBases As String) As String
'
'On Error GoTo ImpBasesWORDErr
'
'Dim lsModeloPlantilla As String
'Dim Nombre As String, nI As Integer, lnPos As Integer, lsEtapa As String, lsPlazo As String, lsFactores As String, lsPuntaje As String
'
'lsModeloPlantilla = App.Path & "\spooler\Plantillas\BASES_11_BIENES.doc"
'
'    'Crea una clase que de Word Object
'    Dim wApp As Word.Application
'    Dim wAppSource As Word.Application
'    'Create a new instance of word
'    Set wApp = New Word.Application
'    Set wAppSource = New Word.Application
'
'    Dim RangeSource As Word.Range
'    'Abre Documento Plantilla
'    wAppSource.Documents.Open FileName:=lsModeloPlantilla
'    Set RangeSource = wAppSource.ActiveDocument.Content
'    'Lo carga en Memoria
'    wAppSource.ActiveDocument.Content.Copy
'
'    'Crea Nuevo Documento
'    wApp.Documents.Add
'
''        wApp.Application.Selection.TypeParagraph
'        wApp.Application.Selection.PasteAndFormat wdFormatOriginalFormatting
''        wApp.Application.Selection.InsertBreak
'
'        wApp.Selection.SetRange Start:=wApp.Selection.Start, End:=wApp.ActiveDocument.Content.End
'        wApp.Selection.MoveEnd
'
'        With wApp.Selection.Find
'            .Text = "<MinPtos>"
'            .Replacement.Text = pnMinPtos
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<mes>"
'            .Replacement.Text = pcMes
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Costo>"
'            .Replacement.Text = pnCosto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<MCostoBases>"
'            .Replacement.Text = pcMonedaCostoBases
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Objeto>"
'            .Replacement.Text = psObjeto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<min>"
'            .Replacement.Text = pnMin
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<max>"
'            .Replacement.Text = pnMax
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''        wApp.Selection.CreateAutoTextEntry("yo", "nose") = psEtapas
'
'        nI = Right(psPlazo, 2)
'
'        '***********************************************************************************
'        '****************************************************************************
'
'        lnPos = InStr(1, psEtapas, vbCrLf)
'        lsEtapa = Mid(psEtapas, 1, lnPos - 1)
'        lsEtapa = lsEtapa & Space(50) & "<Etapa>"
'        psEtapas = Mid(psEtapas, lnPos + 2)
'
'        Do While lnPos > -1
'            With wApp.Selection.Find
'                .Text = "<Etapa>"
'                .Replacement.Text = lsEtapa
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'              End With
'            wApp.Selection.Find.Execute Replace:=wdReplaceAll
'            If lnPos = 0 Then Exit Do
'            lnPos = InStr(1, psEtapas, vbCrLf)
'            If lnPos = 0 Then
'                lsEtapa = ""
'            Else
'                lsEtapa = Mid(psEtapas, 1, lnPos - 1)
'                lsEtapa = lsEtapa & Space(50) & "<Etapa>"
'                psEtapas = Mid(psEtapas, lnPos + 2)
'            End If
'        Loop
'
'        '****************************************************************************
'
'        lnPos = InStr(1, psPlazo, vbCrLf)
'        lsPlazo = Mid(psPlazo, 1, lnPos - 1)
'        lsPlazo = lsPlazo & Space(35) & "<Plazo>"
'        psPlazo = Mid(psPlazo, lnPos + 2)
'
'        Do While lnPos > -1
'            With wApp.Selection.Find
'                .Text = "<Plazo>"
'                .Replacement.Text = lsPlazo
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'              End With
'            wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'            If lnPos = 0 Then Exit Do
'            lnPos = InStr(1, psPlazo, vbCrLf)
'            If lnPos = 0 Then
'                lsPlazo = ""
'            Else
'                lsPlazo = Mid(psPlazo, 1, lnPos - 1)
'                lsPlazo = lsPlazo & Space(50) & "<Plazo>"
'                psPlazo = Mid(psPlazo, lnPos + 2)
'            End If
'        Loop
''        wApp.Documents.Item(1).Paragraphs.Item(29).Range.InsertBefore psEtapas '26
''        wApp.Documents.Item(1).Paragraphs.Item(29 + nI).Range.InsertBefore Mid(psPlazo, 1, Len(psPlazo) - 2) '27
'
'        '****************************************************************************
'        '*****************************************************************************************
'
'        With wApp.Selection.Find
'            .Text = "<proceso>"
'            .Replacement.Text = psProceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Valor_Ref>"
'            .Replacement.Text = psValor_Ref
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<año>"
'            .Replacement.Text = pnAño
'            .Forward = True
'            .Wrap = wdFindContinue
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<moneda>"
'            .Replacement.Text = psMoneda
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<Tipo>"
'            .Replacement.Text = psTipo
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<minimo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMin / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maximo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMax / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<nro_proceso>"
'            .Replacement.Text = psnro_Proceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        '**********************************************************************************************************
'        Dim oCon As DConecta, sSQL As String, rs As ADODB.Recordset, sFactores As String, sPuntos As String, _
'            Item As Integer, nPar As Integer, maxt As Integer, maxe As Integer
'        Set oCon = New DConecta
'        If oCon.AbreConexion Then
'            sSQL = "select i.nProSelItem, f.cBSGrupoCod, g.cBSGrupoDescripcion, x.nFactorNro, nProSelTpoCod, nProSelSubTpo, " & _
'                   "    x.cFactorDescripcion , x.cUnidades, x.nTipo, f.nPuntaje, nFormula, nObjeto, f.cBSGrupoCod, cUnidades " & _
'                   " from LogProSelEvalFactor f " & _
'                   "    inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
'                   "    inner join BSGrupos g on f.cBSGrupoCod = g.cBSGrupoCod " & _
'                   "    inner join LogProSelItem i on f.cBSGrupoCod = i.cBSGrupoCod and f.nProSelNro = i. nProSelNro " & _
'                   " Where f.nVigente = 1 and i.nProSelNro = " & pnProSelNro & " and f.nProSelTpoCod=" & pnProSelTpoCod & " and f.nProSelSubTpo=" & pnProSelSubTpo & " " & _
'                   " order by i.nProSelItem"
'            nPar = 86 + nI
'            Set rs = oCon.CargaRecordSet(sSQL)
'            Do While Not rs.EOF
'                If Item <> rs!nProSelItem Then
'                    Item = rs!nProSelItem
'                    maxt = 0: maxe = 0
'                    sFactores = sFactores & rs!cBSGrupoDescripcion & vbCrLf '& vbCrLf
'                    sPuntos = sPuntos & "-" & vbCrLf '& "-" & vbCrLf  '& vbCrLf
'                    nPar = nPar + 2 '3
'                End If
'                sFactores = sFactores & Space(2) & rs!cFactorDescripcion & vbCrLf
'                Select Case rs!nFormula
'                    Case 0
'                        'Directamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " '& vbCrLf
'                        sFactores = sFactores & "maximo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & "-" & vbCrLf '& " -" & vbCrLf & " -" & vbCrLf
'                    Case 1
'                        'Inversamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " '& vbCrLf
'                        sFactores = sFactores & "minimo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & "-" & vbCrLf '& " -" & vbCrLf & " -" & vbCrLf
'                    Case 2
'                        'Rangos
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & "-" '& vbCrLf
'                        sFactores = sFactores & CargarRangos(rs!nFactorNro, rs!nProSelTpoCod, rs!nProSelSubTpo, rs!nObjeto, rs!cBSGrupoCod, nPar, rs!cUnidades, sPuntos, pnProSelNro)
'                        nPar = nPar + 1
'                    Case 3
'                        'SIno
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje a los postores que cumplan  " '& vbCrLf
'                        sFactores = sFactores & "con lo pedido" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & "-" & vbCrLf '& " -" & vbCrLf & " -" & vbCrLf
'                End Select
''                If Rs!nTipo Then
''                    maxe = maxe + Rs!npuntaje
''                Else
''                    maxt = maxt + Rs!npuntaje
''                End If
'                rs.MoveNext
'            Loop
'            oCon.CierraConexion
'
'            lnPos = InStr(1, sFactores, vbCrLf)
'            lsFactores = Mid(sFactores, 1, lnPos - 1)
'            lsFactores = lsFactores & Space(80) & "<Factores>"
'            sFactores = Mid(sFactores, lnPos + 2)
'
'            Do While lnPos > -1
'                With wApp.Selection.Find
'                    .Text = "<Factores>"
'                    .Replacement.Text = lsFactores
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                  End With
'                wApp.Selection.Find.Execute Replace:=wdReplaceAll
'                If lnPos = 0 Then Exit Do
'                lnPos = InStr(1, sFactores, vbCrLf)
'                If lnPos = 0 Then
'                    lsFactores = ""
'                Else
'                    lsFactores = Mid(sFactores, 1, lnPos - 1)
'                    lsFactores = lsFactores & Space(80) & "<Factores>"
'                    sFactores = Mid(sFactores, lnPos + 2)
'                End If
'            Loop
'
'            lnPos = InStr(1, sPuntos, vbCrLf)
'            lsPuntaje = Mid(sPuntos, 1, lnPos - 1)
'            lsPuntaje = lsPuntaje & Space(100) & "<Puntaje>"
'            sPuntos = Mid(sPuntos, lnPos + 2)
'            If Mid(lsPuntaje, 1, 2) <> "- " Then lsPuntaje = Mid(lsPuntaje, 2)
'
'            Do While lnPos > -1
'                With wApp.Selection.Find
'                    .Text = "<Puntaje>"
'                    .Replacement.Text = lsPuntaje
'                    .Forward = True
'                    .Wrap = wdFindContinue
'                    .Format = False
'                  End With
'                wApp.Selection.Find.Execute Replace:=wdReplaceAll
'                If lnPos = 0 Then Exit Do
'                lnPos = InStr(1, sPuntos, vbCrLf)
'                If lnPos = 0 Then
'                    lsPuntaje = ""
'                Else
'                    lsPuntaje = Mid(sPuntos, 1, lnPos - 1)
'                    lsPuntaje = lsPuntaje & Space(100) & "<Puntaje>"
'                    sPuntos = Mid(sPuntos, lnPos + 2)
'                    If Mid(lsPuntaje, 1, 2) <> "- " And Val(Mid(lsPuntaje, 1, 2)) < 0 Then lsPuntaje = Mid(lsPuntaje, 2)
'                End If
'            Loop
'
''            wApp.Documents.Item(1).Paragraphs.Item(86 + nI).Range.InsertBefore sFactores
''            wApp.Documents.Item(1).Paragraphs.Item(nPar + 1).Range.InsertBefore sPuntos
'
'        End If
'        '**********************************************************************************************************
'
'        With wApp.Selection.Find
'            .Text = "<maxTecnica>"
'            .Replacement.Text = maxt
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maxEconomica>"
'            .Replacement.Text = maxe
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''wapp.Selection.Font
'wAppSource.ActiveDocument.Close
'Nombre = App.Path + "\spooler\Plantillas\BASE " & psnro_Proceso + ".doc"
'wApp.ActiveDocument.SaveAs Nombre
'ActualizaBases pnProSelNro, Nombre
'wApp.Visible = True
'ImpBasesWORD_11_Bienes = Nombre
'Exit Function
'ImpBasesWORDErr:
'    ImpBasesWORD_11_Bienes = ""
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
'End Function
'
'Function ImpBasesWORD_11_SERVICIOS(ByVal psObjeto As String, ByVal psValor_Ref As String, ByVal pnAño As Integer, ByVal psMoneda As String, _
'                    ByVal psTipo As String, ByVal psnro_Proceso As String, ByVal psProceso As String, psEtapas As String, _
'                    ByVal psPlazo As String, ByVal pnMin As Currency, ByVal pnMax As Currency, ByVal pnProSelNro As Integer, _
'                    ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pcMes As String, pnCosto As Currency, _
'                    ByVal pnMinPtos As Integer, ByVal pcObj As String) As String
'
'On Error GoTo ImpBasesWORDErr
'
'Dim lsModeloPlantilla As String
'Dim Nombre As String
'Dim nI As Integer
'
''******** NUEVAS VARIABLES ********
'
''lsModeloPlantilla = RecuperaPlantillaBases(pnProSelTpoCod, pnProSelSubTpo)
''If Len(lsModeloPlantilla) = 0 Then
''   ImpBasesWORD = ""
''   Exit Function
''End If
'
''lsModeloPlantilla = App.Path & "\spooler\BASES_" & pnProSelTpoCod & pnProSelSubTpo & "_" & pcObj & ".DOC"
'lsModeloPlantilla = App.Path & "\spooler\Plantillas\BASES_11_SERVICIOS.doc"
''CadenaAna = Mid(CadenaAna, 1, (Len(CadenaAna) - 1))
'
'    'Crea una clase que de Word Object
'    Dim wApp As Word.Application
'    Dim wAppSource As Word.Application
'    'Create a new instance of word
'    Set wApp = New Word.Application
'    Set wAppSource = New Word.Application
'
'    Dim RangeSource As Word.Range
'    'Abre Documento Plantilla
'    wAppSource.Documents.Open FileName:=lsModeloPlantilla
'    Set RangeSource = wAppSource.ActiveDocument.Content
'    'Lo carga en Memoria
'    wAppSource.ActiveDocument.Content.Copy
'
'    'Crea Nuevo Documento
'    wApp.Documents.Add
'
''        wApp.Application.Selection.TypeParagraph
'        wApp.Application.Selection.PasteAndFormat wdFormatOriginalFormatting
''        wApp.Application.Selection.InsertBreak
'
'        wApp.Selection.SetRange Start:=wApp.Selection.Start, End:=wApp.ActiveDocument.Content.End
'        wApp.Selection.MoveEnd
'
'        With wApp.Selection.Find
'            .Text = "<MinPtos>"
'            .Replacement.Text = pnMinPtos
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<mes>"
'            .Replacement.Text = pcMes
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Costo>"
'            .Replacement.Text = pnCosto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Objeto>"
'            .Replacement.Text = psObjeto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<min>"
'            .Replacement.Text = pnMin
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<max>"
'            .Replacement.Text = pnMax
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        nI = Right(psPlazo, 2)
'
'        wApp.Documents.Item(1).Paragraphs.Item(37).Range.InsertBefore psEtapas '26
'        wApp.Documents.Item(1).Paragraphs.Item(37 + nI).Range.InsertBefore Mid(psPlazo, 1, Len(psPlazo) - 2) '27
'
''        With wApp.Selection.Find
''            .Text = "<Etapas>"
''            .Replacement.Text = psEtapas
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
''        With wApp.Selection.Find
''            .Text = "<Plazo>"
''            .Replacement.Text = psPlazo
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<proceso>"
'            .Replacement.Text = psProceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Valor_Ref>"
'            .Replacement.Text = psValor_Ref
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<año>"
'            .Replacement.Text = pnAño
'            .Forward = True
'            .Wrap = wdFindContinue
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<moneda>"
'            .Replacement.Text = psMoneda
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<Tipo>"
'            .Replacement.Text = psTipo
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<minimo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMin / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maximo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMax / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<nro_proceso>"
'            .Replacement.Text = psnro_Proceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        '**********************************************************************************************************
''        Dim oCon As DConecta, sSQL As String, rs As ADODB.Recordset, sFactores As String, sPuntos As String, _
''            Item As Integer, nPar As Integer, maxt As Integer, maxe As Integer
''        Set oCon = New DConecta
''        If oCon.AbreConexion Then
''            sSQL = "select i.nProSelItem, f.cBSGrupoCod, g.cBSGrupoDescripcion, x.nFactorNro, nProSelTpoCod, nProSelSubTpo, " & _
''                   "    x.cFactorDescripcion , x.cUnidades, x.nTipo, f.nPuntaje, nFormula, nObjeto, f.cBSGrupoCod, cUnidades " & _
''                   " from LogProSelEvalFactor f " & _
''                   "    inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
''                   "    inner join BSGrupos g on f.cBSGrupoCod = g.cBSGrupoCod " & _
''                   "    inner join LogProSelItem i on f.cBSGrupoCod = i.cBSGrupoCod and f.nProSelNro = i. nProSelNro " & _
''                   " Where f.nVigente = 1 and i.nProSelNro = " & pnProSelNro & " and f.nProSelTpoCod=" & pnProSelTpoCod & " and f.nProSelSubTpo=" & pnProSelSubTpo & " " & _
''                   " order by i.nProSelItem"
''            nPar = 88
''            Set rs = oCon.CargaRecordSet(sSQL)
''            Do While Not rs.EOF
''                If Item <> rs!nProSelItem Then
''                    Item = rs!nProSelItem
''                    maxt = 0: maxe = 0
''                    sFactores = sFactores & vbCrLf & rs!cBSGrupoDescripcion & vbCrLf '& vbCrLf
''                    sPuntos = sPuntos & vbCrLf & vbCrLf '& vbCrLf
''                    nPar = nPar + 2 '3
''                End If
''                sFactores = sFactores & Space(2) & rs!cFactorDescripcion & vbCrLf
''                Select Case rs!nFormula
''                    Case 0
''                        'Directamente
''                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
''                        sFactores = sFactores & Space(5) & "maximo valor" & vbCrLf
''                        nPar = nPar + 3
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
''                    Case 1
''                        'Inversamente
''                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
''                        sFactores = sFactores & Space(5) & "minimo valor" & vbCrLf
''                        nPar = nPar + 3
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
''                    Case 2
''                        'Rangos
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf
''                        sFactores = sFactores & CargarRangos(rs!nFactorNro, rs!nProSelTpoCod, rs!nProSelSubTpo, rs!nObjeto, rs!cBSGrupoCod, nPar, rs!cUnidades, sPuntos, pnProSelNro)
''                        nPar = nPar + 1
''                    Case 3
''                        'SIno
''                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje a los postores que cumplan  " & vbCrLf
''                        sFactores = sFactores & Space(5) & "con lo pedido" & vbCrLf
''                        nPar = nPar + 3
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
''                End Select
'''                If Rs!nTipo Then
'''                    maxe = maxe + Rs!npuntaje
'''                Else
'''                    maxt = maxt + Rs!npuntaje
'''                End If
''                rs.MoveNext
''            Loop
''            oCon.CierraConexion
''
''            wApp.Documents.Item(1).Paragraphs.Item(88).Range.InsertBefore sFactores
''            wApp.Documents.Item(1).Paragraphs.Item(nPar + 1).Range.InsertBefore sPuntos
''
''        End If
'        '**********************************************************************************************************
'
''        With wApp.Selection.Find
''            .Text = "<maxTecnica>"
''            .Replacement.Text = maxt
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
''        With wApp.Selection.Find
''            .Text = "<maxEconomica>"
''            .Replacement.Text = maxe
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''wapp.Selection.Font
'wAppSource.ActiveDocument.Close
'Nombre = App.Path + "\spooler\Plantillas\BASE " & psnro_Proceso + ".doc"
'wApp.ActiveDocument.SaveAs Nombre
'ActualizaBases pnProSelNro, Nombre
'wApp.Visible = True
'ImpBasesWORD_11_SERVICIOS = Nombre
'Exit Function
'ImpBasesWORDErr:
'    ImpBasesWORD_11_SERVICIOS = ""
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
'End Function
'
'Function ImpBasesWORD_12_OBRAS(ByVal psObjeto As String, ByVal psValor_Ref As String, ByVal pnAño As Integer, ByVal psMoneda As String, _
'                    ByVal psTipo As String, ByVal psnro_Proceso As String, ByVal psProceso As String, psEtapas As String, _
'                    ByVal psPlazo As String, ByVal pnMin As Currency, ByVal pnMax As Currency, ByVal pnProSelNro As Integer, _
'                    ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pcMes As String, pnCosto As Currency, _
'                    ByVal pnMinPtos As Integer, ByVal pcObj As String, ByVal pcMonedaCostoBases As String) As String
'
'On Error GoTo ImpBasesWORDErr
'
'Dim lsModeloPlantilla As String
'Dim Nombre As String
'Dim nI As Integer
'
''******** NUEVAS VARIABLES ********
'
''lsModeloPlantilla = RecuperaPlantillaBases(pnProSelTpoCod, pnProSelSubTpo)
''If Len(lsModeloPlantilla) = 0 Then
''   ImpBasesWORD = ""
''   Exit Function
''End If
'
''lsModeloPlantilla = App.Path & "\spooler\BASES_" & pnProSelTpoCod & pnProSelSubTpo & "_" & pcObj & ".DOC"
'lsModeloPlantilla = App.Path & "\spooler\Plantillas\BASES_12_OBRAS.doc"
''CadenaAna = Mid(CadenaAna, 1, (Len(CadenaAna) - 1))
'
'    'Crea una clase que de Word Object
'    Dim wApp As Word.Application
'    Dim wAppSource As Word.Application
'    'Create a new instance of word
'    Set wApp = New Word.Application
'    Set wAppSource = New Word.Application
'
'    Dim RangeSource As Word.Range
'    'Abre Documento Plantilla
'    wAppSource.Documents.Open FileName:=lsModeloPlantilla
'    Set RangeSource = wAppSource.ActiveDocument.Content
'    'Lo carga en Memoria
'    wAppSource.ActiveDocument.Content.Copy
'
'    'Crea Nuevo Documento
'    wApp.Documents.Add
'
''        wApp.Application.Selection.TypeParagraph
'        wApp.Application.Selection.PasteAndFormat wdFormatOriginalFormatting
''        wApp.Application.Selection.InsertBreak
'
'        wApp.Selection.SetRange Start:=wApp.Selection.Start, End:=wApp.ActiveDocument.Content.End
'        wApp.Selection.MoveEnd
'
'        With wApp.Selection.Find
'            .Text = "<MinPtos>"
'            .Replacement.Text = pnMinPtos
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<mes>"
'            .Replacement.Text = pcMes
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Costo>"
'            .Replacement.Text = pnCosto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<MCostoBases>"
'            .Replacement.Text = pcMonedaCostoBases
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<MCostoBases>"
'            .Replacement.Text = pcMonedaCostoBases
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Objeto>"
'            .Replacement.Text = psObjeto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<min>"
'            .Replacement.Text = pnMin
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<max>"
'            .Replacement.Text = pnMax
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        nI = Right(psPlazo, 2)
'
'        wApp.Documents.Item(1).Paragraphs.Item(24).Range.InsertBefore psEtapas '26
'        wApp.Documents.Item(1).Paragraphs.Item(24 + nI).Range.InsertBefore Mid(psPlazo, 1, Len(psPlazo) - 2) '27
'
''        With wApp.Selection.Find
''            .Text = "<Etapas>"
''            .Replacement.Text = psEtapas
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
''        With wApp.Selection.Find
''            .Text = "<Plazo>"
''            .Replacement.Text = psPlazo
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<proceso>"
'            .Replacement.Text = psProceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Valor_Ref>"
'            .Replacement.Text = psValor_Ref
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<año>"
'            .Replacement.Text = pnAño
'            .Forward = True
'            .Wrap = wdFindContinue
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<moneda>"
'            .Replacement.Text = psMoneda
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<Tipo>"
'            .Replacement.Text = psTipo
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<minimo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMin / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maximo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMax / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<nro_proceso>"
'            .Replacement.Text = psnro_Proceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        '**********************************************************************************************************
''        Dim oCon As DConecta, sSQL As String, rs As ADODB.Recordset, sFactores As String, sPuntos As String, _
''            Item As Integer, nPar As Integer, maxt As Integer, maxe As Integer
''        Set oCon = New DConecta
''        If oCon.AbreConexion Then
''            sSQL = "select i.nProSelItem, f.cBSGrupoCod, g.cBSGrupoDescripcion, x.nFactorNro, nProSelTpoCod, nProSelSubTpo, " & _
''                   "    x.cFactorDescripcion , x.cUnidades, x.nTipo, f.nPuntaje, nFormula, nObjeto, f.cBSGrupoCod, cUnidades " & _
''                   " from LogProSelEvalFactor f " & _
''                   "    inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
''                   "    inner join BSGrupos g on f.cBSGrupoCod = g.cBSGrupoCod " & _
''                   "    inner join LogProSelItem i on f.cBSGrupoCod = i.cBSGrupoCod and f.nProSelNro = i. nProSelNro " & _
''                   " Where f.nVigente = 1 and i.nProSelNro = " & pnProSelNro & " and f.nProSelTpoCod=" & pnProSelTpoCod & " and f.nProSelSubTpo=" & pnProSelSubTpo & " " & _
''                   " order by i.nProSelItem"
''            nPar = 88
''            Set rs = oCon.CargaRecordSet(sSQL)
''            Do While Not rs.EOF
''                If Item <> rs!nProSelItem Then
''                    Item = rs!nProSelItem
''                    maxt = 0: maxe = 0
''                    sFactores = sFactores & vbCrLf & rs!cBSGrupoDescripcion & vbCrLf '& vbCrLf
''                    sPuntos = sPuntos & vbCrLf & vbCrLf '& vbCrLf
''                    nPar = nPar + 2 '3
''                End If
''                sFactores = sFactores & Space(2) & rs!cFactorDescripcion & vbCrLf
''                Select Case rs!nFormula
''                    Case 0
''                        'Directamente
''                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
''                        sFactores = sFactores & Space(5) & "maximo valor" & vbCrLf
''                        nPar = nPar + 3
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
''                    Case 1
''                        'Inversamente
''                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
''                        sFactores = sFactores & Space(5) & "minimo valor" & vbCrLf
''                        nPar = nPar + 3
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
''                    Case 2
''                        'Rangos
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf
''                        sFactores = sFactores & CargarRangos(rs!nFactorNro, rs!nProSelTpoCod, rs!nProSelSubTpo, rs!nObjeto, rs!cBSGrupoCod, nPar, rs!cUnidades, sPuntos, pnProSelNro)
''                        nPar = nPar + 1
''                    Case 3
''                        'SIno
''                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje a los postores que cumplan  " & vbCrLf
''                        sFactores = sFactores & Space(5) & "con lo pedido" & vbCrLf
''                        nPar = nPar + 3
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
''                End Select
'''                If Rs!nTipo Then
'''                    maxe = maxe + Rs!npuntaje
'''                Else
'''                    maxt = maxt + Rs!npuntaje
'''                End If
''                rs.MoveNext
''            Loop
''            oCon.CierraConexion
''
''            wApp.Documents.Item(1).Paragraphs.Item(88).Range.InsertBefore sFactores
''            wApp.Documents.Item(1).Paragraphs.Item(nPar + 1).Range.InsertBefore sPuntos
''
''        End If
''        '**********************************************************************************************************
''
''        With wApp.Selection.Find
''            .Text = "<maxTecnica>"
''            .Replacement.Text = maxt
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''
''        With wApp.Selection.Find
''            .Text = "<maxEconomica>"
''            .Replacement.Text = maxe
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''wapp.Selection.Font
'wAppSource.ActiveDocument.Close
'Nombre = App.Path + "\spooler\Plantillas\BASE " & psnro_Proceso + ".doc"
'wApp.ActiveDocument.SaveAs Nombre
'ActualizaBases pnProSelNro, Nombre
'wApp.Visible = True
'ImpBasesWORD_12_OBRAS = Nombre
'Exit Function
'ImpBasesWORDErr:
'    ImpBasesWORD_12_OBRAS = ""
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
'End Function
'
'Function ImpBasesWORD_21_BIENES(ByVal psObjeto As String, ByVal psValor_Ref As String, ByVal pnAño As Integer, ByVal psMoneda As String, _
'                    ByVal psTipo As String, ByVal psnro_Proceso As String, ByVal psProceso As String, psEtapas As String, _
'                    ByVal psPlazo As String, ByVal pnMin As Currency, ByVal pnMax As Currency, ByVal pnProSelNro As Integer, _
'                    ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pcMes As String, pnCosto As Currency, _
'                    ByVal pnMinPtos As Integer, ByVal pcObj As String) As String
'
'On Error GoTo ImpBasesWORDErr
'
'Dim lsModeloPlantilla As String
'Dim Nombre As String
'Dim nI As Integer
'
''******** NUEVAS VARIABLES ********
'
''lsModeloPlantilla = RecuperaPlantillaBases(pnProSelTpoCod, pnProSelSubTpo)
''If Len(lsModeloPlantilla) = 0 Then
''   ImpBasesWORD = ""
''   Exit Function
''End If
'
''lsModeloPlantilla = App.Path & "\spooler\BASES_" & pnProSelTpoCod & pnProSelSubTpo & "_" & pcObj & ".DOC"
'lsModeloPlantilla = App.Path & "\spooler\PLANTILLAS\BASES_21_BIENES.doc"
''CadenaAna = Mid(CadenaAna, 1, (Len(CadenaAna) - 1))
'
'    'Crea una clase que de Word Object
'    Dim wApp As Word.Application
'    Dim wAppSource As Word.Application
'    'Create a new instance of word
'    Set wApp = New Word.Application
'    Set wAppSource = New Word.Application
'
'    Dim RangeSource As Word.Range
'    'Abre Documento Plantilla
'    wAppSource.Documents.Open FileName:=lsModeloPlantilla
'    Set RangeSource = wAppSource.ActiveDocument.Content
'    'Lo carga en Memoria
'    wAppSource.ActiveDocument.Content.Copy
'
'    'Crea Nuevo Documento
'    wApp.Documents.Add
'
''        wApp.Application.Selection.TypeParagraph
'        wApp.Application.Selection.PasteAndFormat wdFormatOriginalFormatting
''        wApp.Application.Selection.InsertBreak
'
'        wApp.Selection.SetRange Start:=wApp.Selection.Start, End:=wApp.ActiveDocument.Content.End
'        wApp.Selection.MoveEnd
'
'        With wApp.Selection.Find
'            .Text = "<MinPtos>"
'            .Replacement.Text = pnMinPtos
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<mes>"
'            .Replacement.Text = pcMes
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Costo>"
'            .Replacement.Text = pnCosto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Objeto>"
'            .Replacement.Text = psObjeto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<min>"
'            .Replacement.Text = pnMin
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<max>"
'            .Replacement.Text = pnMax
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        nI = Right(psPlazo, 2)
'
'        wApp.Documents.Item(1).Paragraphs.Item(30).Range.InsertBefore psEtapas '26
'        wApp.Documents.Item(1).Paragraphs.Item(30 + nI).Range.InsertBefore Mid(psPlazo, 1, Len(psPlazo) - 2) '27
'
''        With wApp.Selection.Find
''            .Text = "<Etapas>"
''            .Replacement.Text = psEtapas
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
''        With wApp.Selection.Find
''            .Text = "<Plazo>"
''            .Replacement.Text = psPlazo
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<proceso>"
'            .Replacement.Text = psProceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Valor_Ref>"
'            .Replacement.Text = psValor_Ref
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<año>"
'            .Replacement.Text = pnAño
'            .Forward = True
'            .Wrap = wdFindContinue
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<moneda>"
'            .Replacement.Text = psMoneda
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<Tipo>"
'            .Replacement.Text = psTipo
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<minimo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMin / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maximo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMax / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<nro_proceso>"
'            .Replacement.Text = psnro_Proceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        '**********************************************************************************************************
'        Dim oCon As DConecta, sSQL As String, rs As ADODB.Recordset, sFactores As String, sPuntos As String, _
'            Item As Integer, nPar As Integer, maxt As Integer, maxe As Integer
'        Set oCon = New DConecta
'        If oCon.AbreConexion Then
'            sSQL = "select i.nProSelItem, f.cBSGrupoCod, g.cBSGrupoDescripcion, x.nFactorNro, nProSelTpoCod, nProSelSubTpo, " & _
'                   "    x.cFactorDescripcion , x.cUnidades, x.nTipo, f.nPuntaje, nFormula, nObjeto, f.cBSGrupoCod, cUnidades " & _
'                   " from LogProSelEvalFactor f " & _
'                   "    inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
'                   "    inner join BSGrupos g on f.cBSGrupoCod = g.cBSGrupoCod " & _
'                   "    inner join LogProSelItem i on f.cBSGrupoCod = i.cBSGrupoCod and f.nProSelNro = i. nProSelNro " & _
'                   " Where f.nVigente = 1 and i.nProSelNro = " & pnProSelNro & " and f.nProSelTpoCod=" & pnProSelTpoCod & " and f.nProSelSubTpo=" & pnProSelSubTpo & " " & _
'                   " order by i.nProSelItem"
'            nPar = 96 + nI
'            Set rs = oCon.CargaRecordSet(sSQL)
'            Do While Not rs.EOF
'                If Item <> rs!nProSelItem Then
'                    Item = rs!nProSelItem
'                    maxt = 0: maxe = 0
'                    sFactores = sFactores & vbCrLf & rs!cBSGrupoDescripcion & vbCrLf '& vbCrLf
'                    sPuntos = sPuntos & vbCrLf & vbCrLf '& vbCrLf
'                    nPar = nPar + 2 '3
'                End If
'                sFactores = sFactores & Space(2) & rs!cFactorDescripcion & vbCrLf
'                Select Case rs!nFormula
'                    Case 0
'                        'Directamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
'                        sFactores = sFactores & Space(5) & "maximo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                    Case 1
'                        'Inversamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
'                        sFactores = sFactores & Space(5) & "minimo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                    Case 2
'                        'Rangos
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf
'                        sFactores = sFactores & CargarRangos(rs!nFactorNro, rs!nProSelTpoCod, rs!nProSelSubTpo, rs!nObjeto, rs!cBSGrupoCod, nPar, rs!cUnidades, sPuntos, pnProSelNro)
'                        nPar = nPar + 1
'                    Case 3
'                        'SIno
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje a los postores que cumplan  " & vbCrLf
'                        sFactores = sFactores & Space(5) & "con lo pedido" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                End Select
''                If Rs!nTipo Then
''                    maxe = maxe + Rs!npuntaje
''                Else
''                    maxt = maxt + Rs!npuntaje
''                End If
'                rs.MoveNext
'            Loop
'            oCon.CierraConexion
'
'            wApp.Documents.Item(1).Paragraphs.Item(96 + nI).Range.InsertBefore sFactores
'            wApp.Documents.Item(1).Paragraphs.Item(nPar + 1).Range.InsertBefore sPuntos
'
'        End If
'        '**********************************************************************************************************
'
'        With wApp.Selection.Find
'            .Text = "<maxTecnica>"
'            .Replacement.Text = maxt
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maxEconomica>"
'            .Replacement.Text = maxe
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''wapp.Selection.Font
'wAppSource.ActiveDocument.Close
'Nombre = App.Path + "\spooler\BASE " & psnro_Proceso + ".doc"
'wApp.ActiveDocument.SaveAs Nombre
'ActualizaBases pnProSelNro, Nombre
'wApp.Visible = True
'ImpBasesWORD_21_BIENES = Nombre
'Exit Function
'ImpBasesWORDErr:
'    ImpBasesWORD_21_BIENES = ""
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
'End Function
'
'Function ImpBasesWORD_21_OBRAS(ByVal psObjeto As String, ByVal psValor_Ref As String, ByVal pnAño As Integer, ByVal psMoneda As String, _
'                    ByVal psTipo As String, ByVal psnro_Proceso As String, ByVal psProceso As String, psEtapas As String, _
'                    ByVal psPlazo As String, ByVal pnMin As Currency, ByVal pnMax As Currency, ByVal pnProSelNro As Integer, _
'                    ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pcMes As String, pnCosto As Currency, _
'                    ByVal pnMinPtos As Integer, ByVal pcObj As String, ByVal pcMonedaCostoBases As String) As String
'
'On Error GoTo ImpBasesWORDErr
'
'Dim lsModeloPlantilla As String
'Dim Nombre As String
'Dim nI As Integer
'
''******** NUEVAS VARIABLES ********
'
''lsModeloPlantilla = RecuperaPlantillaBases(pnProSelTpoCod, pnProSelSubTpo)
''If Len(lsModeloPlantilla) = 0 Then
''   ImpBasesWORD = ""
''   Exit Function
''End If
'
'lsModeloPlantilla = App.Path & "\spooler\BASES_" & pnProSelTpoCod & pnProSelSubTpo & "_" & pcObj & ".DOC"
'lsModeloPlantilla = App.Path & "\spooler\PLANTILLAS\BASES_21_OBRAS.doc"
''CadenaAna = Mid(CadenaAna, 1, (Len(CadenaAna) - 1))
'
'    'Crea una clase que de Word Object
'    Dim wApp As Word.Application
'    Dim wAppSource As Word.Application
'    'Create a new instance of word
'    Set wApp = New Word.Application
'    Set wAppSource = New Word.Application
'
'    Dim RangeSource As Word.Range
'    'Abre Documento Plantilla
'    wAppSource.Documents.Open FileName:=lsModeloPlantilla
'    Set RangeSource = wAppSource.ActiveDocument.Content
'    'Lo carga en Memoria
'    wAppSource.ActiveDocument.Content.Copy
'
'    'Crea Nuevo Documento
'    wApp.Documents.Add
'
''        wApp.Application.Selection.TypeParagraph
'        wApp.Application.Selection.PasteAndFormat wdFormatOriginalFormatting
''        wApp.Application.Selection.InsertBreak
'
'        wApp.Selection.SetRange Start:=wApp.Selection.Start, End:=wApp.ActiveDocument.Content.End
'        wApp.Selection.MoveEnd
'
'        With wApp.Selection.Find
'            .Text = "<MinPtos>"
'            .Replacement.Text = pnMinPtos
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<mes>"
'            .Replacement.Text = pcMes
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Costo>"
'            .Replacement.Text = pnCosto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<MCostoBases>"
'            .Replacement.Text = pcMonedaCostoBases
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Objeto>"
'            .Replacement.Text = psObjeto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<min>"
'            .Replacement.Text = pnMin
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<max>"
'            .Replacement.Text = pnMax
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        nI = Right(psPlazo, 2)
'
'        wApp.Documents.Item(1).Paragraphs.Item(24).Range.InsertBefore psEtapas '26
'        wApp.Documents.Item(1).Paragraphs.Item(24 + nI).Range.InsertBefore Mid(psPlazo, 1, Len(psPlazo) - 2) '27
'
''        With wApp.Selection.Find
''            .Text = "<Etapas>"
''            .Replacement.Text = psEtapas
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
''        With wApp.Selection.Find
''            .Text = "<Plazo>"
''            .Replacement.Text = psPlazo
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<proceso>"
'            .Replacement.Text = psProceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Valor_Ref>"
'            .Replacement.Text = psValor_Ref
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<año>"
'            .Replacement.Text = pnAño
'            .Forward = True
'            .Wrap = wdFindContinue
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<moneda>"
'            .Replacement.Text = psMoneda
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<Tipo>"
'            .Replacement.Text = psTipo
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<minimo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMin / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maximo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMax / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<nro_proceso>"
'            .Replacement.Text = psnro_Proceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        '**********************************************************************************************************
''        Dim oCon As DConecta, sSQL As String, rs As ADODB.Recordset, sFactores As String, sPuntos As String, _
''            Item As Integer, nPar As Integer, maxt As Integer, maxe As Integer
''        Set oCon = New DConecta
''        If oCon.AbreConexion Then
''            sSQL = "select i.nProSelItem, f.cBSGrupoCod, g.cBSGrupoDescripcion, x.nFactorNro, nProSelTpoCod, nProSelSubTpo, " & _
''                   "    x.cFactorDescripcion , x.cUnidades, x.nTipo, f.nPuntaje, nFormula, nObjeto, f.cBSGrupoCod, cUnidades " & _
''                   " from LogProSelEvalFactor f " & _
''                   "    inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
''                   "    inner join BSGrupos g on f.cBSGrupoCod = g.cBSGrupoCod " & _
''                   "    inner join LogProSelItem i on f.cBSGrupoCod = i.cBSGrupoCod and f.nProSelNro = i. nProSelNro " & _
''                   " Where f.nVigente = 1 and i.nProSelNro = " & pnProSelNro & " and f.nProSelTpoCod=" & pnProSelTpoCod & " and f.nProSelSubTpo=" & pnProSelSubTpo & " " & _
''                   " order by i.nProSelItem"
''            nPar = 88
''            Set rs = oCon.CargaRecordSet(sSQL)
''            Do While Not rs.EOF
''                If Item <> rs!nProSelItem Then
''                    Item = rs!nProSelItem
''                    maxt = 0: maxe = 0
''                    sFactores = sFactores & vbCrLf & rs!cBSGrupoDescripcion & vbCrLf '& vbCrLf
''                    sPuntos = sPuntos & vbCrLf & vbCrLf '& vbCrLf
''                    nPar = nPar + 2 '3
''                End If
''                sFactores = sFactores & Space(2) & rs!cFactorDescripcion & vbCrLf
''                Select Case rs!nFormula
''                    Case 0
''                        'Directamente
''                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
''                        sFactores = sFactores & Space(5) & "maximo valor" & vbCrLf
''                        nPar = nPar + 3
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
''                    Case 1
''                        'Inversamente
''                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
''                        sFactores = sFactores & Space(5) & "minimo valor" & vbCrLf
''                        nPar = nPar + 3
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
''                    Case 2
''                        'Rangos
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf
''                        sFactores = sFactores & CargarRangos(rs!nFactorNro, rs!nProSelTpoCod, rs!nProSelSubTpo, rs!nObjeto, rs!cBSGrupoCod, nPar, rs!cUnidades, sPuntos, pnProSelNro)
''                        nPar = nPar + 1
''                    Case 3
''                        'SIno
''                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje a los postores que cumplan  " & vbCrLf
''                        sFactores = sFactores & Space(5) & "con lo pedido" & vbCrLf
''                        nPar = nPar + 3
''                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
''                End Select
'''                If Rs!nTipo Then
'''                    maxe = maxe + Rs!npuntaje
'''                Else
'''                    maxt = maxt + Rs!npuntaje
'''                End If
''                rs.MoveNext
''            Loop
''            oCon.CierraConexion
''
''            wApp.Documents.Item(1).Paragraphs.Item(88).Range.InsertBefore sFactores
''            wApp.Documents.Item(1).Paragraphs.Item(nPar + 1).Range.InsertBefore sPuntos
''
''        End If
''        '**********************************************************************************************************
''
''        With wApp.Selection.Find
''            .Text = "<maxTecnica>"
''            .Replacement.Text = maxt
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''
''        With wApp.Selection.Find
''            .Text = "<maxEconomica>"
''            .Replacement.Text = maxe
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''wapp.Selection.Font
'wAppSource.ActiveDocument.Close
'Nombre = App.Path + "\spooler\BASE " & psnro_Proceso + ".doc"
'wApp.ActiveDocument.SaveAs Nombre
'ActualizaBases pnProSelNro, Nombre
'wApp.Visible = True
'ImpBasesWORD_21_OBRAS = Nombre
'Exit Function
'ImpBasesWORDErr:
'    ImpBasesWORD_21_OBRAS = ""
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
'End Function
'
'Function ImpBasesWORD_21_SERVICIOS(ByVal psObjeto As String, ByVal psValor_Ref As String, ByVal pnAño As Integer, ByVal psMoneda As String, _
'                    ByVal psTipo As String, ByVal psnro_Proceso As String, ByVal psProceso As String, psEtapas As String, _
'                    ByVal psPlazo As String, ByVal pnMin As Currency, ByVal pnMax As Currency, ByVal pnProSelNro As Integer, _
'                    ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pcMes As String, pnCosto As Currency, _
'                    ByVal pnMinPtos As Integer, ByVal pcObj As String) As String
'
'On Error GoTo ImpBasesWORDErr
'
'Dim lsModeloPlantilla As String
'Dim Nombre As String
'Dim nI As Integer
'
''******** NUEVAS VARIABLES ********
'
''lsModeloPlantilla = RecuperaPlantillaBases(pnProSelTpoCod, pnProSelSubTpo)
''If Len(lsModeloPlantilla) = 0 Then
''   ImpBasesWORD = ""
''   Exit Function
''End If
'
''lsModeloPlantilla = App.Path & "\spooler\BASES_" & pnProSelTpoCod & pnProSelSubTpo & "_" & pcObj & ".DOC"
'lsModeloPlantilla = App.Path & "\spooler\PLANTILLAS\BASES_21_SERVICIOS.doc"
''CadenaAna = Mid(CadenaAna, 1, (Len(CadenaAna) - 1))
'
'    'Crea una clase que de Word Object
'    Dim wApp As Word.Application
'    Dim wAppSource As Word.Application
'    'Create a new instance of word
'    Set wApp = New Word.Application
'    Set wAppSource = New Word.Application
'
'    Dim RangeSource As Word.Range
'    'Abre Documento Plantilla
'    wAppSource.Documents.Open FileName:=lsModeloPlantilla
'    Set RangeSource = wAppSource.ActiveDocument.Content
'    'Lo carga en Memoria
'    wAppSource.ActiveDocument.Content.Copy
'
'    'Crea Nuevo Documento
'    wApp.Documents.Add
'
''        wApp.Application.Selection.TypeParagraph
'        wApp.Application.Selection.PasteAndFormat wdFormatOriginalFormatting
''        wApp.Application.Selection.InsertBreak
'
'        wApp.Selection.SetRange Start:=wApp.Selection.Start, End:=wApp.ActiveDocument.Content.End
'        wApp.Selection.MoveEnd
'
'        With wApp.Selection.Find
'            .Text = "<MinPtos>"
'            .Replacement.Text = pnMinPtos
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<mes>"
'            .Replacement.Text = pcMes
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Costo>"
'            .Replacement.Text = pnCosto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Objeto>"
'            .Replacement.Text = psObjeto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<min>"
'            .Replacement.Text = pnMin
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<max>"
'            .Replacement.Text = pnMax
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        nI = Right(psPlazo, 2)
'
'        wApp.Documents.Item(1).Paragraphs.Item(30).Range.InsertBefore psEtapas '26
'        wApp.Documents.Item(1).Paragraphs.Item(30 + nI).Range.InsertBefore Mid(psPlazo, 1, Len(psPlazo) - 2) '27
'
''        With wApp.Selection.Find
''            .Text = "<Etapas>"
''            .Replacement.Text = psEtapas
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
''        With wApp.Selection.Find
''            .Text = "<Plazo>"
''            .Replacement.Text = psPlazo
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<proceso>"
'            .Replacement.Text = psProceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Valor_Ref>"
'            .Replacement.Text = psValor_Ref
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<año>"
'            .Replacement.Text = pnAño
'            .Forward = True
'            .Wrap = wdFindContinue
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<moneda>"
'            .Replacement.Text = psMoneda
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<Tipo>"
'            .Replacement.Text = psTipo
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<minimo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMin / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maximo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMax / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<nro_proceso>"
'            .Replacement.Text = psnro_Proceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        '**********************************************************************************************************
'        Dim oCon As DConecta, sSQL As String, rs As ADODB.Recordset, sFactores As String, sPuntos As String, _
'            Item As Integer, nPar As Integer, maxt As Integer, maxe As Integer
'        Set oCon = New DConecta
'        If oCon.AbreConexion Then
'            sSQL = "select i.nProSelItem, f.cBSGrupoCod, g.cBSGrupoDescripcion, x.nFactorNro, nProSelTpoCod, nProSelSubTpo, " & _
'                   "    x.cFactorDescripcion , x.cUnidades, x.nTipo, f.nPuntaje, nFormula, nObjeto, f.cBSGrupoCod, cUnidades " & _
'                   " from LogProSelEvalFactor f " & _
'                   "    inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
'                   "    inner join BSGrupos g on f.cBSGrupoCod = g.cBSGrupoCod " & _
'                   "    inner join LogProSelItem i on f.cBSGrupoCod = i.cBSGrupoCod and f.nProSelNro = i. nProSelNro " & _
'                   " Where f.nVigente = 1 and i.nProSelNro = " & pnProSelNro & " and f.nProSelTpoCod=" & pnProSelTpoCod & " and f.nProSelSubTpo=" & pnProSelSubTpo & " " & _
'                   " order by i.nProSelItem"
'            nPar = 97 + nI
'            Set rs = oCon.CargaRecordSet(sSQL)
'            Do While Not rs.EOF
'                If Item <> rs!nProSelItem Then
'                    Item = rs!nProSelItem
'                    maxt = 0: maxe = 0
'                    sFactores = sFactores & vbCrLf & rs!cBSGrupoDescripcion & vbCrLf '& vbCrLf
'                    sPuntos = sPuntos & vbCrLf & vbCrLf '& vbCrLf
'                    nPar = nPar + 2 '3
'                End If
'                sFactores = sFactores & Space(2) & rs!cFactorDescripcion & vbCrLf
'                Select Case rs!nFormula
'                    Case 0
'                        'Directamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
'                        sFactores = sFactores & Space(5) & "maximo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                    Case 1
'                        'Inversamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
'                        sFactores = sFactores & Space(5) & "minimo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                    Case 2
'                        'Rangos
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf
'                        sFactores = sFactores & CargarRangos(rs!nFactorNro, rs!nProSelTpoCod, rs!nProSelSubTpo, rs!nObjeto, rs!cBSGrupoCod, nPar, rs!cUnidades, sPuntos, pnProSelNro)
'                        nPar = nPar + 1
'                    Case 3
'                        'SIno
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje a los postores que cumplan  " & vbCrLf
'                        sFactores = sFactores & Space(5) & "con lo pedido" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                End Select
''                If Rs!nTipo Then
''                    maxe = maxe + Rs!npuntaje
''                Else
''                    maxt = maxt + Rs!npuntaje
''                End If
'                rs.MoveNext
'            Loop
'            oCon.CierraConexion
'
'            wApp.Documents.Item(1).Paragraphs.Item(97 + nI).Range.InsertBefore sFactores
'            wApp.Documents.Item(1).Paragraphs.Item(nPar + 1).Range.InsertBefore sPuntos
'
'        End If
'        '**********************************************************************************************************
'
'        With wApp.Selection.Find
'            .Text = "<maxTecnica>"
'            .Replacement.Text = maxt
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maxEconomica>"
'            .Replacement.Text = maxe
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''wapp.Selection.Font
'wAppSource.ActiveDocument.Close
'Nombre = App.Path + "\spooler\BASE " & psnro_Proceso + ".doc"
'wApp.ActiveDocument.SaveAs Nombre
'ActualizaBases pnProSelNro, Nombre
'wApp.Visible = True
'ImpBasesWORD_21_SERVICIOS = Nombre
'Exit Function
'ImpBasesWORDErr:
'    ImpBasesWORD_21_SERVICIOS = ""
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
'End Function
'
'Function ImpBasesWORD_31_(ByVal psObjeto As String, ByVal psValor_Ref As String, ByVal pnAño As Integer, ByVal psMoneda As String, _
'                    ByVal psTipo As String, ByVal psnro_Proceso As String, ByVal psProceso As String, psEtapas As String, _
'                    ByVal psPlazo As String, ByVal pnMin As Currency, ByVal pnMax As Currency, ByVal pnProSelNro As Integer, _
'                    ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pcMes As String, pnCosto As Currency, _
'                    ByVal pnMinPtos As Integer, ByVal pcObj As String) As String
'
'On Error GoTo ImpBasesWORDErr
'
'Dim lsModeloPlantilla As String
'Dim Nombre As String
'Dim nI As Integer
'
''******** NUEVAS VARIABLES ********
'
''lsModeloPlantilla = RecuperaPlantillaBases(pnProSelTpoCod, pnProSelSubTpo)
''If Len(lsModeloPlantilla) = 0 Then
''   ImpBasesWORD = ""
''   Exit Function
''End If
'
''lsModeloPlantilla = App.Path & "\spooler\BASES_" & pnProSelTpoCod & pnProSelSubTpo & "_" & pcObj & ".DOC"
'lsModeloPlantilla = App.Path & "\spooler\PLANTILLAS\BASES_31_.doc"
''CadenaAna = Mid(CadenaAna, 1, (Len(CadenaAna) - 1))
'
'    'Crea una clase que de Word Object
'    Dim wApp As Word.Application
'    Dim wAppSource As Word.Application
'    'Create a new instance of word
'    Set wApp = New Word.Application
'    Set wAppSource = New Word.Application
'
'    Dim RangeSource As Word.Range
'    'Abre Documento Plantilla
'    wAppSource.Documents.Open FileName:=lsModeloPlantilla
'    Set RangeSource = wAppSource.ActiveDocument.Content
'    'Lo carga en Memoria
'    wAppSource.ActiveDocument.Content.Copy
'
'    'Crea Nuevo Documento
'    wApp.Documents.Add
'
''        wApp.Application.Selection.TypeParagraph
'        wApp.Application.Selection.PasteAndFormat wdFormatOriginalFormatting
''        wApp.Application.Selection.InsertBreak
'
'        wApp.Selection.SetRange Start:=wApp.Selection.Start, End:=wApp.ActiveDocument.Content.End
'        wApp.Selection.MoveEnd
'
'        With wApp.Selection.Find
'            .Text = "<MinPtos>"
'            .Replacement.Text = pnMinPtos
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<mes>"
'            .Replacement.Text = pcMes
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Costo>"
'            .Replacement.Text = pnCosto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Objeto>"
'            .Replacement.Text = psObjeto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<min>"
'            .Replacement.Text = pnMin
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<max>"
'            .Replacement.Text = pnMax
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        nI = Right(psPlazo, 2)
'
'        wApp.Documents.Item(1).Paragraphs.Item(119).Range.InsertBefore psEtapas '26
'        wApp.Documents.Item(1).Paragraphs.Item(119 + nI).Range.InsertBefore Mid(psPlazo, 1, Len(psPlazo) - 2) '27
'
''        With wApp.Selection.Find
''            .Text = "<Etapas>"
''            .Replacement.Text = psEtapas
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
''        With wApp.Selection.Find
''            .Text = "<Plazo>"
''            .Replacement.Text = psPlazo
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<proceso>"
'            .Replacement.Text = psProceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Valor_Ref>"
'            .Replacement.Text = psValor_Ref
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<año>"
'            .Replacement.Text = pnAño
'            .Forward = True
'            .Wrap = wdFindContinue
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<moneda>"
'            .Replacement.Text = psMoneda
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<Tipo>"
'            .Replacement.Text = psTipo
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<minimo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMin / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maximo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMax / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<nro_proceso>"
'            .Replacement.Text = psnro_Proceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        '**********************************************************************************************************
'        Dim oCon As DConecta, sSQL As String, rs As ADODB.Recordset, sFactores As String, sPuntos As String, _
'            Item As Integer, nPar As Integer, maxt As Integer, maxe As Integer
'        Set oCon = New DConecta
'        If oCon.AbreConexion Then
'            sSQL = "select i.nProSelItem, f.cBSGrupoCod, g.cBSGrupoDescripcion, x.nFactorNro, nProSelTpoCod, nProSelSubTpo, " & _
'                   "    x.cFactorDescripcion , x.cUnidades, x.nTipo, f.nPuntaje, nFormula, nObjeto, f.cBSGrupoCod, cUnidades " & _
'                   " from LogProSelEvalFactor f " & _
'                   "    inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
'                   "    inner join BSGrupos g on f.cBSGrupoCod = g.cBSGrupoCod " & _
'                   "    inner join LogProSelItem i on f.cBSGrupoCod = i.cBSGrupoCod and f.nProSelNro = i. nProSelNro " & _
'                   " Where f.nVigente = 1 and i.nProSelNro = " & pnProSelNro & " and f.nProSelTpoCod=" & pnProSelTpoCod & " and f.nProSelSubTpo=" & pnProSelSubTpo & " " & _
'                   " order by i.nProSelItem"
'            nPar = 195 + nI
'            Set rs = oCon.CargaRecordSet(sSQL)
'            Do While Not rs.EOF
'                If Item <> rs!nProSelItem Then
'                    Item = rs!nProSelItem
'                    maxt = 0: maxe = 0
'                    sFactores = sFactores & vbCrLf & rs!cBSGrupoDescripcion & vbCrLf '& vbCrLf
'                    sPuntos = sPuntos & vbCrLf & vbCrLf '& vbCrLf
'                    nPar = nPar + 2 '3
'                End If
'                sFactores = sFactores & Space(2) & rs!cFactorDescripcion & vbCrLf
'                Select Case rs!nFormula
'                    Case 0
'                        'Directamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
'                        sFactores = sFactores & Space(5) & "maximo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                    Case 1
'                        'Inversamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
'                        sFactores = sFactores & Space(5) & "minimo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                    Case 2
'                        'Rangos
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf
'                        sFactores = sFactores & CargarRangos(rs!nFactorNro, rs!nProSelTpoCod, rs!nProSelSubTpo, rs!nObjeto, rs!cBSGrupoCod, nPar, rs!cUnidades, sPuntos, pnProSelNro)
'                        nPar = nPar + 1
'                    Case 3
'                        'SIno
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje a los postores que cumplan  " & vbCrLf
'                        sFactores = sFactores & Space(5) & "con lo pedido" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                End Select
''                If Rs!nTipo Then
''                    maxe = maxe + Rs!npuntaje
''                Else
''                    maxt = maxt + Rs!npuntaje
''                End If
'                rs.MoveNext
'            Loop
'            oCon.CierraConexion
'
'            wApp.Documents.Item(1).Paragraphs.Item(195 + nI).Range.InsertBefore sFactores
'            wApp.Documents.Item(1).Paragraphs.Item(nPar + 1).Range.InsertBefore sPuntos
'
'        End If
'        '**********************************************************************************************************
'
'        With wApp.Selection.Find
'            .Text = "<maxTecnica>"
'            .Replacement.Text = maxt
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maxEconomica>"
'            .Replacement.Text = maxe
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''wapp.Selection.Font
'wAppSource.ActiveDocument.Close
'Nombre = App.Path + "\spooler\BASE " & psnro_Proceso + ".doc"
'wApp.ActiveDocument.SaveAs Nombre
'ActualizaBases pnProSelNro, Nombre
'wApp.Visible = True
'ImpBasesWORD_31_ = Nombre
'Exit Function
'ImpBasesWORDErr:
'    ImpBasesWORD_31_ = ""
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
'End Function
'
'Function ImpBasesWORD_41_(ByVal psObjeto As String, ByVal psValor_Ref As String, ByVal pnAño As Integer, ByVal psMoneda As String, _
'                    ByVal psTipo As String, ByVal psnro_Proceso As String, ByVal psProceso As String, psEtapas As String, _
'                    ByVal psPlazo As String, ByVal pnMin As Currency, ByVal pnMax As Currency, ByVal pnProSelNro As Integer, _
'                    ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, ByVal pcMes As String, pnCosto As Currency, _
'                    ByVal pnMinPtos As Integer, ByVal pcObj As String) As String
'
'On Error GoTo ImpBasesWORDErr
'
'Dim lsModeloPlantilla As String
'Dim Nombre As String
'Dim nI As Integer
'
''******** NUEVAS VARIABLES ********
'
''lsModeloPlantilla = RecuperaPlantillaBases(pnProSelTpoCod, pnProSelSubTpo)
''If Len(lsModeloPlantilla) = 0 Then
''   ImpBasesWORD = ""
''   Exit Function
''End If
'
''lsModeloPlantilla = App.Path & "\spooler\BASES_" & pnProSelTpoCod & pnProSelSubTpo & "_" & pcObj & ".DOC"
'lsModeloPlantilla = App.Path & "\spooler\PLANTILLAS\BASES_41_.doc"
''CadenaAna = Mid(CadenaAna, 1, (Len(CadenaAna) - 1))
'
'    'Crea una clase que de Word Object
'    Dim wApp As Word.Application
'    Dim wAppSource As Word.Application
'    'Create a new instance of word
'    Set wApp = New Word.Application
'    Set wAppSource = New Word.Application
'
'    Dim RangeSource As Word.Range
'    'Abre Documento Plantilla
'    wAppSource.Documents.Open FileName:=lsModeloPlantilla
'    Set RangeSource = wAppSource.ActiveDocument.Content
'    'Lo carga en Memoria
'    wAppSource.ActiveDocument.Content.Copy
'
'    'Crea Nuevo Documento
'    wApp.Documents.Add
'
''        wApp.Application.Selection.TypeParagraph
'        wApp.Application.Selection.PasteAndFormat wdFormatOriginalFormatting
''        wApp.Application.Selection.InsertBreak
'
'        wApp.Selection.SetRange Start:=wApp.Selection.Start, End:=wApp.ActiveDocument.Content.End
'        wApp.Selection.MoveEnd
'
'        With wApp.Selection.Find
'            .Text = "<MinPtos>"
'            .Replacement.Text = pnMinPtos
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<mes>"
'            .Replacement.Text = pcMes
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Costo>"
'            .Replacement.Text = pnCosto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Objeto>"
'            .Replacement.Text = psObjeto
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'        End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<min>"
'            .Replacement.Text = pnMin
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<max>"
'            .Replacement.Text = pnMax
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        nI = Right(psPlazo, 2)
'
'        wApp.Documents.Item(1).Paragraphs.Item(30).Range.InsertBefore psEtapas '26
'        wApp.Documents.Item(1).Paragraphs.Item(30 + nI).Range.InsertBefore Mid(psPlazo, 1, Len(psPlazo) - 2) '27
'
''        With wApp.Selection.Find
''            .Text = "<Etapas>"
''            .Replacement.Text = psEtapas
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
''        With wApp.Selection.Find
''            .Text = "<Plazo>"
''            .Replacement.Text = psPlazo
''            .Forward = True
''            .Wrap = wdFindContinue
''            .Format = False
''          End With
''        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<proceso>"
'            .Replacement.Text = psProceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<Valor_Ref>"
'            .Replacement.Text = psValor_Ref
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<año>"
'            .Replacement.Text = pnAño
'            .Forward = True
'            .Wrap = wdFindContinue
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<moneda>"
'            .Replacement.Text = psMoneda
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<Tipo>"
'            .Replacement.Text = psTipo
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'        With wApp.Selection.Find
'            .Text = "<minimo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMin / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maximo>"
'            .Replacement.Text = FNumero(psValor_Ref * pnMax / 100#)
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<nro_proceso>"
'            .Replacement.Text = psnro_Proceso
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        '**********************************************************************************************************
'        Dim oCon As DConecta, sSQL As String, rs As ADODB.Recordset, sFactores As String, sPuntos As String, _
'            Item As Integer, nPar As Integer, maxt As Integer, maxe As Integer
'        Set oCon = New DConecta
'        If oCon.AbreConexion Then
'            sSQL = "select i.nProSelItem, f.cBSGrupoCod, g.cBSGrupoDescripcion, x.nFactorNro, nProSelTpoCod, nProSelSubTpo, " & _
'                   "    x.cFactorDescripcion , x.cUnidades, x.nTipo, f.nPuntaje, nFormula, nObjeto, f.cBSGrupoCod, cUnidades " & _
'                   " from LogProSelEvalFactor f " & _
'                   "    inner join LogProSelFactor x on f.nFactorNro = x.nFactorNro " & _
'                   "    inner join BSGrupos g on f.cBSGrupoCod = g.cBSGrupoCod " & _
'                   "    inner join LogProSelItem i on f.cBSGrupoCod = i.cBSGrupoCod and f.nProSelNro = i. nProSelNro " & _
'                   " Where f.nVigente = 1 and i.nProSelNro = " & pnProSelNro & " and f.nProSelTpoCod=" & pnProSelTpoCod & " and f.nProSelSubTpo=" & pnProSelSubTpo & " " & _
'                   " order by i.nProSelItem"
'            nPar = 100 + nI
'            Set rs = oCon.CargaRecordSet(sSQL)
'            Do While Not rs.EOF
'                If Item <> rs!nProSelItem Then
'                    Item = rs!nProSelItem
'                    maxt = 0: maxe = 0
'                    sFactores = sFactores & vbCrLf & rs!cBSGrupoDescripcion & vbCrLf '& vbCrLf
'                    sPuntos = sPuntos & vbCrLf & vbCrLf '& vbCrLf
'                    nPar = nPar + 2 '3
'                End If
'                sFactores = sFactores & Space(2) & rs!cFactorDescripcion & vbCrLf
'                Select Case rs!nFormula
'                    Case 0
'                        'Directamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
'                        sFactores = sFactores & Space(5) & "maximo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                    Case 1
'                        'Inversamente
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje al postor que ofresca el " & vbCrLf
'                        sFactores = sFactores & Space(5) & "minimo valor" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                    Case 2
'                        'Rangos
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf
'                        sFactores = sFactores & CargarRangos(rs!nFactorNro, rs!nProSelTpoCod, rs!nProSelSubTpo, rs!nObjeto, rs!cBSGrupoCod, nPar, rs!cUnidades, sPuntos, pnProSelNro)
'                        nPar = nPar + 1
'                    Case 3
'                        'SIno
'                        sFactores = sFactores & Space(5) & "Se Otorgara el maximo Puntaje a los postores que cumplan  " & vbCrLf
'                        sFactores = sFactores & Space(5) & "con lo pedido" & vbCrLf
'                        nPar = nPar + 3
'                        sPuntos = sPuntos & rs!npuntaje & vbCrLf & vbCrLf & vbCrLf
'                End Select
''                If Rs!nTipo Then
''                    maxe = maxe + Rs!npuntaje
''                Else
''                    maxt = maxt + Rs!npuntaje
''                End If
'                rs.MoveNext
'            Loop
'            oCon.CierraConexion
'
'            wApp.Documents.Item(1).Paragraphs.Item(100 + nI).Range.InsertBefore sFactores
'            wApp.Documents.Item(1).Paragraphs.Item(nPar + 1).Range.InsertBefore sPuntos
'
'        End If
'        '**********************************************************************************************************
'
'        With wApp.Selection.Find
'            .Text = "<maxTecnica>"
'            .Replacement.Text = maxt
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
'
'        With wApp.Selection.Find
'            .Text = "<maxEconomica>"
'            .Replacement.Text = maxe
'            .Forward = True
'            .Wrap = wdFindContinue
'            .Format = False
'          End With
'        wApp.Selection.Find.Execute Replace:=wdReplaceAll
''wapp.Selection.Font
'wAppSource.ActiveDocument.Close
'Nombre = App.Path + "\spooler\BASE " & psnro_Proceso + ".doc"
'wApp.ActiveDocument.SaveAs Nombre
'ActualizaBases pnProSelNro, Nombre
'wApp.Visible = True
'ImpBasesWORD_41_ = Nombre
'Exit Function
'ImpBasesWORDErr:
'    ImpBasesWORD_41_ = ""
'    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
'End Function

Private Sub ActualizaBases(ByVal pnProSelNro As Integer, ByVal pcBases As String)
On Error GoTo ActualizaBasesErr
    Dim ocon As DConecta, sSQL As String
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "update LogProcesoSeleccion set cArchivoBases='" & pcBases & "' where nProSelNro= " & pnProSelNro
        ocon.Ejecutar sSQL
        ocon.CierraConexion
    End If
Exit Sub
ActualizaBasesErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub


Private Function CargarRangos(ByVal pnFactorNro As Integer, _
                            ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer, _
                            ByVal pnObjeto As Integer, ByVal pcBSGrupoCod As String, _
                            ByRef pnPar As Integer, ByVal pcUnidades As String, _
                            ByRef psPuntos As String, ByVal pnProSelNro As Integer) As String
On Error GoTo CargarRangosErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset, Cadena As String
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select nRangoItem, nRangoMin, nRangoMax, nPuntaje " & _
               "    From LogProSelEvalFactorRangos " & _
               "    where nFactorNro = " & pnFactorNro & " and nProSelTpoCod = " & pnProSelTpoCod & _
               "    and nProSelSubTpo = " & pnProSelSubTpo & " and " & _
               "    nObjeto = " & pnObjeto & " and cBSGrupoCod ='" & pcBSGrupoCod & "' and nVigente =1 " & _
               "    and nProSelNro = " & pnProSelNro
        Set rs = ocon.CargaRecordSet(sSQL)
        Do While Not rs.EOF
            Cadena = Cadena & Space(5) & "- Mayor o Igual a " & Format(rs!nRangoMin, "00") & " y Menor o Igual a " & Format(rs!nRangoMax, "00") & " " & pcUnidades & Space(5) & rs!npuntaje & vbCrLf
            pnPar = pnPar + 1
            psPuntos = psPuntos & vbCrLf & "-" '& vbCrLf
            rs.MoveNext
        Loop
        Set rs = Nothing
        CargarRangos = Cadena
        ocon.CierraConexion
    End If
    Exit Function
CargarRangosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Public Function RecuperaPlantillaBases(ByVal pnProSelTpoCod As Integer, ByVal pnProSelSubTpo As Integer) As String
Dim oConn As New DConecta, rs As New ADODB.Recordset, sSQL As String

RecuperaPlantillaBases = ""
sSQL = "select cPlantillaBases from LogProSelTpoRangos where nProSelTpoCod = " & pnProSelTpoCod & " and nProSelSubTpo = " & pnProSelSubTpo & " "
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      RecuperaPlantillaBases = rs!cPlantillaBases
   End If
End If
End Function

Public Sub ImpRegParticipantesWord(ByVal pnProSelNro As Integer, ByVal pcTitulo As String, ByVal pcPersona As String, _
                                   ByVal pcPersRUC As String, ByVal pcPersDom As String, ByVal pcPersTel As String, _
                                   ByVal pcPersEmail As String, _
                                   ByVal pcPersRepNom As String, ByVal pcPersRepDNI As String, ByVal pcFecha As String)
On Error GoTo ImpConsultasWordErr
Dim MSWord As New Word.Application
Dim Documento As Word.Document
Dim Parrafo As Paragraph
Dim Parrafo1 As Paragraph
Set Documento = MSWord.Documents.Add
Set Parrafo = Documento.Paragraphs.Add
Set Parrafo1 = Documento.Paragraphs.Add

Dim strCadena As String
Dim rsComite As New ADODB.Recordset
Dim rsFase As New ADODB.Recordset
Dim rsCriterio As New ADODB.Recordset
Dim rsDatosProceso As New ADODB.Recordset
Dim intFila As Integer
Dim sPostor As String, NomArch As String
'***********************************************

'Dim ocon As DConecta, rs As ADODB.Recordset, sSQL As String
'Set ocon = New DConecta

strCadena = Space(50) & "REGISTRO DE PARTICIPANTES DE PROCESO" & vbCrLf & vbCrLf
strCadena = strCadena & Space(30) & "PROCESO DE SELECCION DE MENOR CUANTIA Nro " + Format(pnProSelNro, "00") + "-" + CStr(Year(Date)) + "-CMAC-T" & vbCrLf & vbCrLf
strCadena = strCadena & Space(50) & pcTitulo & vbCrLf & vbCrLf & vbCrLf & vbCrLf
'Set rs = CargarPostores(pnProSelNro, pcPersCod)
'    If Not rs.EOF Then
        NomArch = CStr(pnProSelNro)
        strCadena = strCadena & "RAZON SOCIAL               : " & pcPersona & vbCrLf & vbCrLf
        strCadena = strCadena & "DIRECCION                      : " & pcPersDom & vbCrLf & vbCrLf
        strCadena = strCadena & "TELEFONO / FAX             : " & pcPersTel & vbCrLf & vbCrLf
        strCadena = strCadena & "RUC                                   : " & pcPersRUC & vbCrLf & vbCrLf
        strCadena = strCadena & "E-MAIL                               : " & pcPersEmail & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        strCadena = strCadena & "RECIBO CONFORME" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        strCadena = strCadena & "NOMBRE PERSONA          : " & pcPersRepNom & vbCrLf & vbCrLf
        strCadena = strCadena & "DNI                                      : " & pcPersRepDNI & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        strCadena = strCadena & "FIRMA                                 :________________________________________________________________" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        strCadena = strCadena & "FECHA                              :" & pcFecha & vbCrLf & vbCrLf
'    Else
'        MsgBox "No Existe El Postor...", vbInformation, "Aviso"
'        Exit Sub
'    End If

Parrafo.Range.Font.Name = "Arial"
Parrafo.Range.Font.Size = 9
Parrafo.Space1
Parrafo.Alignment = wdAlignParagraphJustify
'Parrafo.Range.Font.Underline = wdUnderlineDouble
Parrafo.Range.InsertBefore strCadena

'MSWord.ActiveDocument.SaveAs App.Path + "\spooler\REGISTRO " + " .doc"
MSWord.Visible = True

Set MSWord = Nothing
Set Documento = Nothing
Set Parrafo = Nothing
    Exit Sub
ImpConsultasWordErr:
MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Sub

Public Function CierraEtapa(ByVal pnProSelNro As Integer, ByVal pnEtapaCod As Integer) As Boolean
On Error GoTo CierraEtapaErr
    Dim ocon As DConecta, sSQL As String
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "update LogProSelEtapa set nEstado = 2 where nProSelNro = " & pnProSelNro & " and nEtapaCod = " & pnEtapaCod
        ocon.Ejecutar sSQL
        ocon.CierraConexion
        CierraEtapa = True
    Else
        CierraEtapa = False
    End If
    Exit Function
CierraEtapaErr:
    CierraEtapa = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Public Function CierraEtapaAutomatico() As Boolean
On Error GoTo CierraEtapaErr
    Dim ocon As DConecta, sSQL As String
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        'sSQL = "update LogProSelEtapa set nEstado = 2 where datediff(d,dFechaTermino,'" & Format(gdFecSis, "yyyymmdd") & "') > 0  "
        'sSQL = "Log_sp_CierreAutomaticoEtapasProcesoSeleccion '" & Format(gdFecSis, "yyyymmdd") & "'"
        sSQL = "Log_sp_CierreAutomaticoEtapasProcesoSeleccion "
        ocon.Ejecutar sSQL
        ocon.CierraConexion
        CierraEtapaAutomatico = True
    Else
        CierraEtapaAutomatico = False
    End If
    Exit Function
CierraEtapaErr:
    CierraEtapaAutomatico = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Public Function VerificaEtapa(ByVal pnProSelNro As Integer, ByVal pnEtapaCod As Integer) As Boolean
On Error GoTo CierraEtapaErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select nEstado from LogProSelEtapa where nProSelNro = " & pnProSelNro & " and nEtapaCod = " & pnEtapaCod
        Set rs = ocon.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            If rs!nEstado > 0 Then
                VerificaEtapa = True
            Else
                VerificaEtapa = False
            End If
        Else
            VerificaEtapa = False
        End If
        ocon.CierraConexion
    Else
        VerificaEtapa = False
    End If
    Exit Function
CierraEtapaErr:
    VerificaEtapa = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Public Function VerificaEtapaCerrada(ByVal pnProSelNro As Integer, ByVal pnEtapaCod As Integer) As Boolean
On Error GoTo VerificaEtapaCerradaErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select nEstado from LogProSelEtapa where nProSelNro = " & pnProSelNro & " and nEtapaCod = " & pnEtapaCod
        Set rs = ocon.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            If rs!nEstado = 2 Then
                VerificaEtapaCerrada = True
            Else
                VerificaEtapaCerrada = False
            End If
        Else
            VerificaEtapaCerrada = False
        End If
        ocon.CierraConexion
    Else
        VerificaEtapaCerrada = False
    End If
    Exit Function
VerificaEtapaCerradaErr:
    VerificaEtapaCerrada = False
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Public Function RequerimientosNoAprobadosPS(ByVal pnAnio As Integer) As Integer
Dim oConn As New DConecta, sSQL As String, rs As New ADODB.Recordset

RequerimientosNoAprobadosPS = 0

'sSQL = "select nPlanReq=count(*) from " & _
       " (select r.nPlanReqNro,nNro=count(*) from LogPlanAnualAprobacion a " & _
       "  inner join LogPlanAnualReq r on a.nPlanReqNro = r.nPlanReqNro " & _
       "  Where r.nAnio = " & pnAnio & " and r.nEstado=1 group by r.nPlanReqNro) x left join " & _
       " (select nPlanReqNro,nApro=count(*) from LogPlanAnualAprobacion where nEstadoAprobacion=1 group by nPlanReqNro) y on x.nPlanReqNro = y.nPlanReqNro " & _
       "  Where x.nNro <> Y.nApro  "
sSQL = "select nProSelReq=count(*) from  " & _
       " (select r.nProSelReqNro,nNro=count(*) from LogProSelAprobacion a " & _
       " inner join LogProSelReq r on a.nProSelReqNro = r.nProSelReqNro Where r.nAnio = 2005 and r.nEstado=1 group by r.nProSelReqNro) x " & _
       " left join  (select nProSelReqNro,nApro=count(*) from LogProSelAprobacion where nEstadoAprobacion=1 group by nProSelReqNro) y on x.nProSelReqNro = y.nProSelReqNro " & _
       " Where x.nNro <> Y.nApro"

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      RequerimientosNoAprobadosPS = rs!nProSelReq
   End If
End If
End Function

Public Function RequerimientoAprobado(ByVal pnProSelReqNro As Integer) As Boolean
On Error GoTo RequerimientoAprobadoErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select nAprobaciones=count(*) from LogProSelAprobacion a where nEstadoAprobacion=0 and nProSelReqNro = " & pnProSelReqNro
        Set rs = ocon.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            If rs!nAprobaciones = 0 Then
                RequerimientoAprobado = True
            Else
                RequerimientoAprobado = False
            End If
        Else
            RequerimientoAprobado = False
        End If
        rs.Close
        ocon.CierraConexion
    End If
    Set rs = Nothing
    Set ocon = Nothing
    Exit Function
RequerimientoAprobadoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Public Function ValidaRango(ByVal pnTpo As Integer, ByVal pnSubTpo As Integer, ByVal pnMonto As Currency, ByVal pnObjeto As Integer, ByVal pnMoneda As Integer) As Boolean
On Error GoTo ValidaRangoErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset, nTipoCambio As Currency
    Set ocon = New DConecta
    Select Case pnMoneda
        Case 1
            nTipoCambio = 1
        Case 2
            nTipoCambio = TipoCambio(1, gdFecSis)
    End Select
    If ocon.AbreConexion Then
        sSQL = "select * from LogProSelTpoRangos where nProSelTpoCod = " & pnTpo & " and nProSelSubTpo = " & pnSubTpo
        Set rs = ocon.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            Select Case pnObjeto
                Case 1
                    If rs!nBienesMin <= (pnMonto * nTipoCambio) And (pnMonto * nTipoCambio) <= rs!nBienesMax Then
                        ValidaRango = True
                    Else
                        ValidaRango = False
                    End If
                Case 2
                    If rs!nServiMin <= (pnMonto * nTipoCambio) And (pnMonto * nTipoCambio) <= rs!nServiMax Then
                        ValidaRango = True
                    Else
                        ValidaRango = False
                    End If
                Case 3
                    If rs!nObrasMin <= (pnMonto * nTipoCambio) And (pnMonto * nTipoCambio) <= rs!nObrasMax Then
                        ValidaRango = True
                    Else
                        ValidaRango = False
                    End If
                Case Else
                    ValidaRango = True
            End Select
        End If
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
ValidaRangoErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation
End Function

Public Function TipoCambio(ByVal pnTipo As Integer, ByVal pdFecha As Date) As Currency
On Error GoTo TipoCambioErr
    Dim ocon As DConecta, sSQL As String, rs As ADODB.Recordset
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select isnull(dbo.GetTipoCambio(" & pnTipo & ",'" & Format(pdFecha, "yyyymmdd") & "'),3.26)"
        Set rs = ocon.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            TipoCambio = rs(0)
        End If
        rs.Close
        Set rs = Nothing
        ocon.CierraConexion
    End If
    Exit Function
TipoCambioErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Public Function CargarValorRef(ByVal pcBSCod As String)
On Error GoTo CargarValorRefErr
    Dim ocon As DConecta, rs As ADODB.Recordset, sSQL As String
    Set ocon = New DConecta
    If ocon.AbreConexion Then
        sSQL = "select top 1 nMonto, nStock, dsaldo from BSSaldos where datediff(m,dsaldo,getdate())<=6 and cBSCod = '" & pcBSCod & "' order by dsaldo desc"
        Set rs = ocon.CargaRecordSet(sSQL)
        If Not rs.EOF Then
            CargarValorRef = rs!nMonto / IIf(rs!nsTock = 0, 1, rs!nsTock)
        End If
        ocon.CierraConexion
    End If
    Exit Function
CargarValorRefErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function


'Public Sub LimpiarPDT()
'    Dim ocon As New DConecta, rs As ADODB.Recordset
'    If ocon.AbreConexion Then
'        Set rs = ocon.CargaRecordSet("select * from RHBasePDTDetalle where cTipo = 'P' and cPeriodo = '200508'")
'        Do While Not rs.EOF
'            ocon.Ejecutar "ps_mht " & rs!nRemuneracion & "," & rs!nImpuesto & "," & rs!Dias & "," & rs!cRHCod
'            rs.MoveNext
'        Loop
'        ocon.CierraConexion
'    End If
'End Sub
