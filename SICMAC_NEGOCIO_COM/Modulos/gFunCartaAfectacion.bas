Attribute VB_Name = "gFunCartaAfectacion"
Public Sub ImprimeCartaAfectacion(ByVal cCtaCod As String, ByVal nTipoCredito As Integer, ByVal nMontoCol As Currency, Optional ByVal nPoliza As Long)
    'ByVal cPersCod As String, ByVal cPersGarantia As String, ByVal cDoi As String, ByVal nDoi As String, Optional ByVal cEstadoCivil As String, Optional ByVal cDomicilio As String,

    Dim cPersCod As String
    Dim cPersGarantia As String
    Dim cDoi As String
    Dim nDoi As String
    Dim cEstadoCivil As String
    Dim cDomicilio As String
    Dim PersonaGarantia As ADODB.Recordset
    Dim PersonaTitular As ADODB.Recordset 'EAAS 20170831
    Dim obj As COMNCredito.NCOMCredito
    Set obj = New COMNCredito.NCOMCredito
    Set PersonaGarantia = obj.ObtenerGarantiaPersona(cCtaCod)
    Set PersonaTitular = obj.ObtenerTitularCredito(cCtaCod) 'EAAS 20170831
    Set obj = Nothing
   
    If nTipoCredito = 504 Or PersonaGarantia.RecordCount = 1 Then
    Do While Not PersonaGarantia.EOF
        cPersCod = PersonaGarantia!cPersCod
        cPersGarantia = PersonaGarantia!cPersNombre
        cDoi = PersonaGarantia!cDoi
        nDoi = PersonaGarantia!nDoi
        If nTipoCredito = 703 Then
            cEstadoCivil = PersonaGarantia!cEstadoCivil
            cDomicilio = PersonaGarantia!cDireccion
        End If
        
    PersonaGarantia.MoveNext
    Loop
    Else 'EAAS 20170831
    cPersCod = PersonaGarantia!cPersCod
    End If
    


'cf

'    Set obj = New COMNCredito.NCOMCredito
'    Set PersonaGarantia = obj.ObtenerGarantiaPersona(vCodCta)
'    Set obj = Nothing
'
'    Do While Not PersonaGarantia.EOF
'    cPersCod = PersonaGarantia!cPersCod
'    cPersGarantia = PersonaGarantia!cPersNombre
'    cDoi = PersonaGarantia!cDoi
'    nDoi = PersonaGarantia!nDoi
'    PersonaGarantia.MoveNext
'
'    Loop
    Dim oDCOMCartaFianza As COMNCredito.NCOMCredito
    Dim rsGarantias As ADODB.Recordset
    
    On Error GoTo ErrorImprimirPDF
    Dim sParrafo1 As String
    Dim sParrafo1a As String
    Dim sParrafo1b As String
    Dim sParrafo1c As String
    Dim sParrafo2 As String
    Dim sParrafo3 As String
    Dim sParrafo4 As String
    Dim sParrafo5a As String
    Dim sParrafo5b As String
    Dim sParrafo5c As String
    Dim oDoc  As cPDF
    Set oDoc = New cPDF
    
    Dim sMontoColocado As String
    Dim sMontoCol As String
    sMontoCol = Format(nMontoCol, "#,###0.00")
    sMontoColocado = IIf(Mid(cCtaCod, 9, 1) = "1", "S/ ", "$. ") & sMontoCol & " " & "(" & UCase(NumLet(sMontoCol)) & IIf(Mid(cCtaCod, 9, 1) = "2", "", " Y " & IIf(InStr(1, sMontoCol, ".") = 0, "00", Mid(sMontoCol, InStr(1, sMontoCol, ".") + 1, 2)) & "/100 ") & IIf(Mid(cCtaCod, 9, 1) = "1", " SOLES)", " DOLARES)")
    sMontoColocado = IIf(Mid(cCtaCod, 9, 1) = "1", "S/ ", "$. ") & Format(sMontoCol, "#,###0.00") & " " & "(" & UCase(NumLet(sMontoCol)) & IIf(Mid(cCtaCod, 9, 1) = "2", "", " Y " & IIf(InStr(1, str(Format(sMontoCol, "#,###0.00")), ".") = 0, "00", Mid(str(Format(sMontoCol, "#,###0.00")), InStr(1, str(Format(sMontoCol, "#,###0.00")), ".") + 1, 2)) & "/100 ") & IIf(Mid(cCtaCod, 9, 1) = "1", "SOLES)", " US DOLARES)")
    sMontoColocado = IIf(Mid(cCtaCod, 9, 1) = "1", "S/ ", "$. ") & Format(sMontoCol, "#,###0.00") & " " & "(" & UCase(NumLet(sMontoCol)) & IIf(Mid(cCtaCod, 9, 1) = "2", "", " Y " & IIf(InStr(1, sMontoCol, ".") = 0, "00", Mid(sMontoCol, InStr(1, sMontoCol, ".") + 1, 2)) & "/100 ") & IIf(Mid(cCtaCod, 9, 1) = "1", "SOLES)", " US DOLARES)")

    If nTipoCredito = 514 Then
        Dim rsCartaFianza As ADODB.Recordset
        Set rsCartaFianza = New ADODB.Recordset
        Set oDCOMCartaFianza = New COMNCredito.NCOMCredito
        Set rsCartaFianza = oDCOMCartaFianza.ObtenerCFRelacion(cCtaCod)
        Set oDCOMCartaFianza = Nothing
        Dim nTpoRelac As Integer
        Dim cAcreedor As String
        Dim cAval As String
        Dim cDoiAval As String
        Dim nDoiAval As String
    
        Do While Not rsCartaFianza.EOF
        If rsCartaFianza!nAval = 1 Then
            If rsCartaFianza!nPrdPersRelac = 38 Then
                cAval = rsCartaFianza!cPersNombre
                cDoiAval = rsCartaFianza!cDoi
                nDoiAval = rsCartaFianza!nDoi
            Else
                cAcreedor = rsCartaFianza!cPersNombre
            End If
        Else 'CONSORCIOS
            If rsCartaFianza!nPrdPersRelac = 20 Then
                cAval = rsCartaFianza!cPersNombre
                cDoiAval = rsCartaFianza!cDoi
                nDoiAval = rsCartaFianza!nDoi
            Else
                cAcreedor = rsCartaFianza!cPersNombre
            End If
        End If
        rsCartaFianza.MoveNext
        Loop
        
          sParrafo2 = "Incluyendo sus intereses compensatorios, frutos y demás bienes que produzca, con la finalidad de honrar y cumplir " & _
                      "las obligaciones contraídas, por " & cAval & " con " & cDoiAval & "° " & nDoiAval & ", con su representada por la " & _
                      "CARTA FIANZA N° " & Format(nPoliza, "0000000") & " (crédito N° " & cCtaCod & ") por la suma de " & sMontoColocado & " " & _
                      ", a favor  de " & cAcreedor & " en caso de incumplimiento de la deuda; de conformidad a lo dispuesto en el Art. 132 " & _
                      "inc. 11 de la Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de " & _
                      "Banca y Seguros - Ley N° 26702, que permite a las entidades del Sector Financiero la compensación de sus acreencias " & _
                      "por cobrar con los activos de los deudores."
                
                If Not oDoc.PDFCreate(App.Path & "\Spooler\CartaAfectacion_" & nPoliza & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
                    Exit Sub
                End If
        
    ElseIf PersonaGarantia.RecordCount = 1 And cPersCod = PersonaTitular!cPersCod Then 'EAAS 20170831
          sParrafo2 = "Incluyendo sus respectivos intereses y demás bienes que produzca, AFECTEN y RETIREN el importe de " & _
                    "las obligaciones pendientes de pago que mantenga el crédito N° " & cCtaCod & " aprobado y/o otorgado " & _
                    "a mi persona;     afectación que se hace, de conformidad a lo dispuesto en el Art. 132 inc. 11 de la " & _
                    "Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de " & _
                    "Banca y Seguros - Ley N° 26702, que permite a las entidades del Sector Financiero, realizar la " & _
                    "compensación de sus acreencias por cobrar con los activos de los deudores, hasta por el monto de " & _
                    "aquellas. Así mismo, manifiesto que la presente autorización es irrevocable y se mantendrá hasta " & _
                    "que haya cumplido con cancelar todas las obligaciones pendientes de pago con Caja Maynas."
                    
                    If Not oDoc.PDFCreate(App.Path & "\Spooler\CartaAfectacion_" & cCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
                    Exit Sub
                    End If
    Else 'EAAS 20170831
    sParrafo2 = "Incluyendo sus respectivos intereses y demás bienes que produzca, AFECTEN y RETIREN el importe de " & _
                    "las obligaciones pendientes de pago que mantenga el crédito N° " & cCtaCod & " aprobado y/o otorgado " & _
                    "a favor del cliente " & PersonaTitular!cPersNombre & " con DNI." & PersonaTitular!cPersIDnro & " en el cual intervengo/intervenimos en calidad de garante(s) y fiador(es) solidario(s) ;     afectación que se hace, de conformidad a lo dispuesto en el Art. 132 inc. 11 de la " & _
                    "Ley General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de " & _
                    "Banca y Seguros - Ley N° 26702, que permite a las entidades del Sector Financiero, realizar la " & _
                    "compensación de sus acreencias por cobrar con los activos de los deudores, hasta por el monto de " & _
                    "aquellas. Así mismo, manifiesto que la presente autorización es irrevocable y se mantendrá hasta " & _
                    "que haya cumplido con cancelar todas las obligaciones pendientes de pago con Caja Maynas."
         'END EAAS 20170831

'            sParrafo2a = "Incluyendo sus respectivos intereses y demás bienes que produzca, AFECTEN y RETIREN el importe de "
'            sParrafo2b = "las obligaciones pendientes de pago que mantenga el crédito N° 109011011000110124 aprobado y/o otorgado a "
'            sParrafo2c = "mi persona; afectación que se hace, de conformidad a lo dispuesto en el Art. 132 inc. 11 de la Ley"
'            sParrafo2d = "General del Sistema Financiero y del Sistema de Seguros y Orgánica de la Superintendencia de Banca "
'            sParrafo2e = "y Seguros - Ley N° 26702, que permite a las entidades del Sector Financiero, realizar la compensación "
'            sParrafo2f = "de sus acreencias por cobrar con los activos de los deudores, hasta por el monto de aquellas."
'            sParrafo2g = "Así mismo, manifiesto que la presente autorización es irrevocable y se mantendrá hasta que haya"
'            sParrafo2h = "cumplido con cancelar todas las obligaciones pendientes de pago con Caja Maynas."
'            sParrafo2i = ""
        
          If Not oDoc.PDFCreate(App.Path & "\Spooler\CartaAfectacion_" & cCtaCod & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format(Time, "hhmmss") & ".pdf") Then
                    Exit Sub
            End If
    End If
    

    Dim nTamano As Integer
    Dim nValidar As Double
    Dim nTop As Integer
    Dim sFechaActual As String
    
    Dim cNroDoc As String
    Dim nSaldo As String
    Dim sMontoGravado As String
    Dim nNro As Integer
    nNro = 0
    
    Set oDCOMCartaFianza = New COMNCredito.NCOMCredito
    Set rsGarantias = oDCOMCartaFianza.ObtenerGarantias(cCtaCod)
    Set oDCOMCartaFianza = Nothing

    
    'Creacion de Archivo
    oDoc.Author = gsCodUser
    oDoc.Creator = "SICMACT - Negocio"
    oDoc.Producer = "Caja Municipal de Ahorros y Crédito de Maynas S.A."
    oDoc.Subject = "Carta Afectacion Nº " & cCtaCod
    oDoc.Title = "Carta Afectacion Nº " & cCtaCod
    
    
    oDoc.Fonts.Add "F1", "Times New Roman", TrueType, Normal, WinAnsiEncoding
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType, Bold, WinAnsiEncoding
    
    oDoc.NewPage A4_Vertical

    sFechaActual = Format(gdFecSis, "dd") & " de " & Format(gdFecSis, "mmmm") & " del " & Format(gdFecSis, "yyyy")

    oDoc.WTextBox 70, 70, 10, 450, "CARTA DE AUTORIZACIÓN DE AFECTACIÓN", "F1", 13, vMiddle
    oDoc.WTextBox 95, 50, 10, 450, "Iquitos, " & sFechaActual, "F2", 11, hLeft
    oDoc.WTextBox 120, 50, 10, 450, "Señores:", "F1", 11, hLeft
    oDoc.WTextBox 140, 50, 10, 450, "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS", "F2", 11, hLeft '
    oDoc.WTextBox 160, 50, 10, 450, "Presente.-", "F1", 11, hLeft
    oDoc.WTextBox 200, 50, 10, 450, "De mi consideración:", "F1", 11, hLeft
    
    
    If nTipoCredito = 514 Then
            sParrafo1 = "Por medio de la presente me dirijo a ustedes para saludarlos y principalmente para AUTORIZAR a la " & _
            "CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS S.A., a que pueda constituir como garantía mobiliaria, " & _
            "de ser el caso, así como AFECTAR Y RETIRAR DE MI(s) CUENTA(s) DE DEPOSITO A PLAZO FIJO:"
    Else
            If PersonaGarantia.RecordCount = 1 Then
                sParrafo1 = "Por el presente documento Yo,   " & cPersGarantia & "  , identificado con   " & cDoi & "   N°   " & nDoi & "  , estado civil   " & cEstadoCivil & "   y   " & _
                    "domiciliado en   " & cDomicilio & "  ; me dirijo a ustedes para saludarlos y manifestar expresamente mi voluntad de AUTORIZARLOS para " & _
                    "que de mi(s) CUENTA(s) DE DEPÓSITO A PLAZO FIJO:"
             Else
                    sParrafo1 = "Por el presente documento "
                    
                    Dim nTotal As Integer
                     Set obj = New COMNCredito.NCOMCredito
                    Set PersonaGarantia = obj.ObtenerGarantiaPersona(cCtaCod)
                    Set obj = Nothing
                    nTotal = PersonaGarantia.RecordCount
                    
                     Do While Not PersonaGarantia.EOF
                        cPersCod = PersonaGarantia!cPersCod
                        cPersGarantia = PersonaGarantia!cPersNombre
                        cDoi = PersonaGarantia!cDoi
                        nDoi = PersonaGarantia!nDoi
                        cEstadoCivil = PersonaGarantia!cEstadoCivil
                        cDomicilio = PersonaGarantia!cDireccion
                        
                        If nTotal = PersonaGarantia.RecordCount Then
                              sParrafo1a = "Yo,   " & cPersGarantia & "  , identificado con   " & cDoi & "   N°   " & nDoi & "  , estado civil   " & cEstadoCivil & "   y   " & _
                            "domiciliado en   " & cDomicilio & "  "
                        ElseIf nTotal = 1 Then
                            sParrafo1a = "y Yo,   " & cPersGarantia & "  , identificado con   " & cDoi & "   N°   " & nDoi & "  , estado civil   " & cEstadoCivil & "   y   " & _
                            "domiciliado en   " & cDomicilio & "  "
                        Else
                            sParrafo1a = ", Yo,   " & cPersGarantia & "  , identificado con   " & cDoi & "   N°   " & nDoi & "  , estado civil   " & cEstadoCivil & "   y   " & _
                            "domiciliado en   " & cDomicilio & "  "
                        End If
                        
                        sParrafo1 = sParrafo1 + sParrafo1a
                        nTotal = nTotal - 1
                    PersonaGarantia.MoveNext
                    Loop
                    
                    sParrafo1b = "; nos dirigimos a ustedes para saludarlos y manifestar expresamente nuestra voluntad de AUTORIZARLOS para "
                    
                    If PersonaGarantia.RecordCount = 1 Then
                    sParrafo1c = "que de mi(s) CUENTA(s) DE DEPÓSITO A PLAZO FIJO:"
                    Else
                    sParrafo1c = "que de nuestra(s) CUENTA(S) DE DEPÓSITO A PLAZO FIJO:"
                    End If
                   
                    sParrafo1 = sParrafo1 + sParrafo1b + sParrafo1c
             End If
                
    End If
    
    'String(20, "-") & " " &
    'oDoc.WTextBox 220, 50, 50, 580, sParrafo1, "F1", 11, hjustify, , , , , , 50
   
                
    nTamano = Len(sParrafo1)
    nValidar = nTamano / 75
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    nTop = 180
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo1, "F1", 12, hjustify
    oDoc.WTextBox nTop, 0, nTamano * 8, 580, sParrafo1, "F1", 10, hjustify, , , , , , 50 'esto
    'oDoc.WText nTop, 0, sParrafo1, "F1", 11
    'oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True

    nTop = nTop + (nTamano * 8) + 10

     Dim counter As Integer
    counter = nTop
    Dim sSaldo As String

            Do While Not rsGarantias.EOF
                cNroDoc = rsGarantias!cNroDoc
                nSaldo = rsGarantias!nSaldo
                nNro = nNro + 1
                
                sSaldo = Format(nSaldo, "#,###0.00")
                sMontoGravado = IIf(Mid(cNroDoc, 9, 1) = "1", "S/ ", "$. ") & sSaldo & " " & "(" & UCase(NumLet(sSaldo)) & IIf(Mid(cNroDoc, 9, 1) = "2", "", " Y " & IIf(InStr(1, sSaldo, ".") = 0, "00", Mid(sSaldo, InStr(1, sSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(cNroDoc, 9, 1) = "1", "SOLES)", " DOLARES)")
                sMontoGravado = IIf(Mid(cNroDoc, 9, 1) = "1", "S/ ", "$. ") & Format(sSaldo, "#,###0.00") & " " & "(" & UCase(NumLet(sSaldo)) & IIf(Mid(cNroDoc, 9, 1) = "2", "", " Y " & IIf(InStr(1, str(Format(sSaldo, "#,###0.00")), ".") = 0, "00", Mid(str(Format(sSaldo, "#,###0.00")), InStr(1, str(Format(sSaldo, "#,###0.00")), ".") + 1, 2)) & "/100 ") & IIf(Mid(cNroDoc, 9, 1) = "1", "SOLES)", " US DOLARES)")
                sMontoGravado = IIf(Mid(cNroDoc, 9, 1) = "1", "S/ ", "$. ") & Format(sSaldo, "#,###0.00") & " " & "(" & UCase(NumLet(sSaldo)) & IIf(Mid(cNroDoc, 9, 1) = "2", "", " Y " & IIf(InStr(1, sSaldo, ".") = 0, "00", Mid(sSaldo, InStr(1, sSaldo, ".") + 1, 2)) & "/100 ") & IIf(Mid(cNroDoc, 9, 1) = "1", "SOLES)", " US DOLARES)")
    
    
                oDoc.WTextBox counter, 0, 10, 580, nNro & ") " & cNroDoc & " por un monto de " & sMontoGravado & ".", "F1", 10, hjustify, , , , , , 50
                '485
                counter = counter + 20
                rsGarantias.MoveNext
                
            Loop
    
    'PARRAFO 2
                
    nTop = counter + 10
    nTamano = Len(sParrafo2)
    nValidar = nTamano / 80
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))

    oDoc.WTextBox nTop, 0, nTamano * 10, 580, sParrafo2, "F1", 10, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 10, 10, nTamano * 10, 580, sParrafo2b, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 20, 10, nTamano * 10, 580, sParrafo2c, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 30, 10, nTamano * 10, 580, sParrafo2d, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 40, 10, nTamano * 10, 580, sParrafo2e, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 50, 10, nTamano * 10, 580, sParrafo2f, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 60, 10, nTamano * 10, 580, sParrafo2g, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 70, 10, nTamano * 10, 580, sParrafo2h, "F1", 11, hjustify, , , , , , 50
'    oDoc.WTextBox nTop + 80, 10, nTamano * 10, 580, sParrafo2i, "F1", 11, hjustify, , , , , , 50

    nTop = nTop + (nTamano * 10)
                      
     If nTipoCredito = 514 Then
                    sParrafo3 = "Así mismo DECLARO BAJO JURAMENTO que el dinero depositado en dicha cuenta(s), la misma que " & _
                    "ofrezco en garantía y autorizo su afectación por medio del apresente carta, no es un bien " & _
                    "inerbagable, según lo establecido en el Art. 648° del Código Procesal Civil, y no pertenece a " & _
                    "sociedad conyugal, puesto que es un bien propio y por lo tanto tengo plena facultad para disponer " & _
                    "del crédito depositado."
     Else
        
                    sParrafo3 = "Sin otro particular, me despido agradeciendo la atención a la presente."
        End If
        
    nTamano = Len(sParrafo3)
    nValidar = nTamano / 90
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo3, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, sParrafo3, "F1", 10, hjustify, , , , , , 50
    'oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    nTop = nTop + (nTamano * 12)
    'FRHU20131126
    If nTipoCredito = 514 Then
     sParrafo4 = "" & _
                        "Sin otro particular."
    Else
    sParrafo4 = ""
    End If
    nTamano = Len(sParrafo4)
    nValidar = nTamano / 90
    nTamano = CInt(IIf(CInt(nValidar) > nValidar, CInt(nValidar), CInt(nValidar) + 1))
    'oDoc.WTextBox nTop, 50, nTamano * 10, 450, sParrafo4, "F1", 12, hjustify, vMiddle, , , , False
    oDoc.WTextBox nTop, 0, nTamano * 10, 580, sParrafo4, "F1", 10, hjustify, , , , , , 50
    'oDoc.WTextBox nTop + 50, 50, 10, 75, "", "F1", 10, hjustify, , vbWhite, 1, vbWhite, True
    'nTop = nTop + 50
    
    oDoc.WTextBox nTop + 80, 50, 10, 580, "Atentamente,", "F1", 10, hLeft, , , , , 0
    
    
    
    
'
'    oDoc.WTextBox nTop + 100, 60, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '1
'    oDoc.WTextBox nTop + 100, 300, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '2
'    oDoc.WTextBox nTop + 170, 60, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '3
'    oDoc.WTextBox nTop + 170, 300, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '4
'    oDoc.WTextBox nTop + 240, 60, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '5
'    oDoc.WTextBox nTop + 240, 300, 10, 210, sParrafo5, "F1", 10, hLeft, , , 1, , False '6
     If PersonaGarantia.RecordCount = 1 Then
'        sParrafo5 = "............................................................................" & _
'                "   " & cPersGarantia & "  " & _
'                "  " & cDoi & " N° " & nDoi & " "
            sParrafo5a = "............................................................................"
            sParrafo5b = "" & cPersGarantia & ""
            sParrafo5c = "" & cDoi & " N° " & nDoi & ""
        oDoc.WTextBox nTop + 150, 50, 10, 220, sParrafo5a, "F1", 10, hLeft, , , 0, , False '1
        oDoc.WTextBox nTop + 160, 50, 10, 220, sParrafo5b, "F1", 10, hLeft, , , 0, , False '1
        oDoc.WTextBox nTop + 170, 50, 10, 180, sParrafo5c, "F1", 10, hLeft, , , 0, , False '1
     Else
         Dim nCounter As Integer
         nCounter = 1
         nCounterb = 0
         Set obj = New COMNCredito.NCOMCredito
         Set PersonaGarantia = obj.ObtenerGarantiaPersona(cCtaCod)
         Set obj = Nothing
         Dim a As Integer
    
         a = 140 '90
    
         Do While Not PersonaGarantia.EOF
            cPersCod = PersonaGarantia!cPersCod
            cPersGarantia = PersonaGarantia!cPersNombre
            cDoi = PersonaGarantia!cDoi
            nDoi = PersonaGarantia!nDoi
    
            sParrafo5a = "............................................................................"
            sParrafo5b = "" & cPersGarantia & ""
            sParrafo5c = "" & cDoi & " N° " & nDoi & ""
    
            If nCounterb = 2 Then
                a = a + 70
                nCounterb = 0
            End If
    
            If nCounter Mod 2 <> 0 Then
                oDoc.WTextBox nTop + a, 50, 10, 220, sParrafo5a, "F1", 10, hLeft, , , 0, , False '1
                oDoc.WTextBox nTop + a + 10, 50, 10, 220, sParrafo5b, "F1", 10, hLeft, , , 0, , False '1
                oDoc.WTextBox nTop + a + 20, 50, 10, 180, sParrafo5c, "F1", 10, hLeft, , , 0, , False '1
            Else
                oDoc.WTextBox nTop + a, 330, 10, 220, sParrafo5a, "F1", 10, hLeft, , , 0, , False '2
                oDoc.WTextBox nTop + a + 10, 330, 10, 220, sParrafo5b, "F1", 10, hLeft, , , 0, , False '2
                oDoc.WTextBox nTop + a + 20, 330, 10, 180, sParrafo5c, "F1", 10, hLeft, , , 0, , False '2
            End If
    
    
            nCounter = nCounter + 1
            nCounterb = nCounterb + 1
        PersonaGarantia.MoveNext
        Loop
     End If

    oDoc.PDFClose
    oDoc.Show
    Exit Sub
ErrorImprimirPDF:
    MsgBox Err.Description, vbInformation, "Aviso"
   

End Sub

