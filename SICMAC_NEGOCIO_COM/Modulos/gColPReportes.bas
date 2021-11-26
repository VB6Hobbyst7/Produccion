Attribute VB_Name = "gColPReportes"
'Option Explicit
'
'Public Function nImpAvisoVencRemate(ByVal pnTpoSalida As Integer, ByVal psAgencia As String, ByVal psFecRemate As String) As String
''    Dim loProcesa As dColPFunciones
''    Dim lrs As ADODB.Recordset
''    Dim laLista() As String, lsSQL As String
''    Dim lnFilas As Integer
''    Dim ldFecAviso As Date
''
''
''    ldFecAviso = DateAdd("d", -pDiasCartaVencimiento, psFecRemate)
''
''    lsSQL = "SELECT p.cnompers, p.cdirpers, p.ctelpers, cp.ccodcta , cp.dfecvenc, cp.cestado, p.cCodZon " & _
''        " FROM CredPrenda CP JOIN PersCuenta PC ON cp.ccodcta = pc.ccodcta " & _
''        " JOIN " & gcCentralPers & "Persona P ON p.ccodpers = pc.ccodpers " & _
''        " WHERE DATEDIFF(dd, cp.dfecvenc, '" & Format(vFecAviso, "mm/dd/yyyy") & "') >= 0  AND " & _
''        " cp.cestado IN ('4','6') " & _
''        " ORDER BY p.cnompers "
''
''    Set loProcesa = New dColPFunciones
''        Set lrs = loProcesa.dObtieneRecordSet(lsSQL)
''    Set loProcesa = Nothing
''    If lrs Is Nothing Then
''        nImpAvisoVencRemate = "VACIO"
''        Exit Function
''    Else
''        lnFilas = 0
''        'prgList.Min = 0: vCont = 0
''        'prgList.Max = RegProcesar.RecordCount * 2
''        If pnTpoSalida = 0 Then  ' Previo
''            If prgList.Max > pPrevioMax Then
''                RegProcesar.Close
''                Set RegProcesar = Nothing
''                MsgBox " Cantidad muy grande para ser cargada en el Previo " & vbCr & _
''                    " se recomienda enviar directo a impresión ", vbInformation, " ! Aviso ! "
''                MuestraImpresion = False
''                MousePointer = 0
''                Exit Function
''            End If
''        ElseIf optImpresion(1).Value = True Then
''            ImpreBegin False, pHojaFiMax
''        Else
''
''        End If
''        prgList.Visible = True
''        ReDim aLista(RegProcesar.RecordCount, 5)
''        With RegProcesar
''            Do While Not RegProcesar.EOF
''                vFilas = vFilas + 1
''                If !cNomPers = aLista(vFilas - 1, 0) Then
''                    aLista(vFilas - 1, 3) = aLista(vFilas - 1, 3) & " - " & !cCodCta
''                    aLista(vFilas - 1, 4) = aLista(vFilas - 1, 4) & " - " & Format(!dFecVenc, "dd/mm/yyyy")
''                    vFilas = vFilas - 1
''                Else
''                    aLista(vFilas, 0) = !cNomPers
''                    aLista(vFilas, 1) = !cDirPers & " - " & ClienteZona(!cCodZon) & " - " & ClienteCiudad(!cCodZon)
''                    aLista(vFilas, 2) = !cTelPers & ""
''                    aLista(vFilas, 3) = !cCodCta
''                    aLista(vFilas, 4) = Format(!dFecVenc, "dd/mm/yyyy")
''                End If
''                vCont = vCont + 1
''                prgList.Value = vCont
''                .MoveNext
''            Loop
''        End With
''        RegProcesar.Close
''        Set RegProcesar = Nothing
''    End If
''    vCont = vCont + (vCont - vFilas)
''    prgList.Value = vCont
''    'Llena cartas
''    Dim X As Integer
''    Dim vNom As String * 40
''    Dim vDir As String * 50
''    Dim vCue As String
''    Dim vFecVen As Date
''    Dim vTel As String
''    Dim RTFTmp As String
''    RTFTmp = ""
''    vRTFImp = ""
''    'vBuffer = ""
''    For X = 1 To vFilas
''        vNom = PstaNombre(aLista(X, 0), False)
''        vDir = aLista(X, 1)
''        'vTel = aLista(X, 2)
''        vCue = aLista(X, 3)
''        'vFecVen = aLista(X, 4)
''        RTFTmp = rtfCartas.Text
''        RTFTmp = Replace(RTFTmp, "<<FECHA>>", Format(gdFecSis, "dddd,d mmmm yyyy"), , 1, vbTextCompare)
''        RTFTmp = Replace(RTFTmp, "<<NOMBRE>>", vNom, , 1, vbTextCompare)
''        RTFTmp = Replace(RTFTmp, "<<CONTRATO>>", vCue, , 2, vbTextCompare)
''        RTFTmp = Replace(RTFTmp, "<<FECREMATE>>", Format(gdFecSis, "mmmm"), , 1, vbTextCompare)
''        RTFTmp = Replace(RTFTmp, "<<DIRECCION>>", vDir, , 1, vbTextCompare)
''        If optImpresion(0).Value = True Or optImpresion(2).Value = True Then
''            RTFTmp = RTFTmp & Chr(12)
''            vRTFImp = vRTFImp & RTFTmp
''            If X Mod 50 = 0 Then
''                vBuffer = vBuffer & vRTFImp
''                vRTFImp = ""
''            End If
''        Else
''            Print #ArcSal, ImpreCarEsp(RTFTmp)
''            If X Mod 5 = 0 Then
''                ImpreEnd
''                ImpreBegin False, pHojaFiMax
''            Else
''                ImpreNewPage
''            End If
''        End If
''        vNom = "":     vDir = "":      vCue = ""
''        'vFecVen = Format("01/01/1990", "dd/mm/yyyy"): 'vTel = ""
''        vCont = vCont + 1
''        prgList.Value = vCont
''    Next X
''    If optImpresion(0).Value = True Then
''        vBuffer = vBuffer & vRTFImp
''        rtfImp.Text = vBuffer
''        vRTFImp = ""
''    ElseIf optImpresion(2).Value = True Then
''        vBuffer = vBuffer & vRTFImp
''        vRTFImp = ""
''    Else
''        ImpreEnd
''    End If
''    prgList.Visible = False
''    prgList.Value = 0
''    Erase aLista
''    MousePointer = 0
'End Function
'
'
'
