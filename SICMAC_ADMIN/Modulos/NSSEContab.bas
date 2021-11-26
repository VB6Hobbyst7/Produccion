Attribute VB_Name = "gNSSEContab"
Private Type tArbol
    cDirec As String
    bOper As Boolean
    cDigito As String
    ApDer As String
    ApIzq As String
End Type
Dim Arbol() As tArbol
Dim ContArbol As Integer
Dim sCadArbol() As String
Dim nCadArbol As Integer
Private Type TDigito
    sCadArbolTemp() As String
    nCadArbolTemp As Integer
End Type
Dim ContArbolLec As Integer

'Public Function FactorAjuste(ByVal dFecAdq As Date, ByVal dFecRep As Date, bOk As Boolean) As Double
'Dim sSQL As String
'Dim R As New ADODB.Recordset
'Dim FactAFecha As Double
'Dim FacRep As Double
'    sSQL = "Select * from IPM Where Month(dFecha) = " & Trim(Str(Month(dFecRep))) & " And Year(dFecha) = " & Trim(Str(Year(dFecRep)))
'    R.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'        If R.BOF And R.EOF Then
'            MsgBox "No Existe Factor IPM para la Fecha " & Format(dFecRep, gsFormatoFechaView), vbInformation, "Aviso"
'            bOk = False
'            FactorAjuste = 0
'            Exit Function
'        End If
'    R.Close
'
'    sSQL = "Select * from IPM Where Month(dFecha) = " & Trim(Str(Month(dFecAdq))) & " And Year(dFecha) = " & Trim(Str(Year(dFecAdq)))
'    R.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'        If R.BOF And R.EOF Then
'            MsgBox "No Existe Factor IPM para la Fecha " & Format(dFecAdq, gsFormatoFechaView), vbInformation, "Aviso"
'            bOk = False
'            FactorAjuste = 0
'            Exit Function
'        End If
'    R.Close
'    bOk = True
'
'        sSQL = "Select * from IPM Where MONTH(dFecha) = " & Trim(Str(Month(dFecRep))) & " And YEAR(dFecha) = " & Trim(Str(Year(dFecRep)))
'        R.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'            If Not R.BOF And Not R.EOF Then
'                FacRep = R!nFactor1
'            End If
'        R.Close
'        sSQL = "Select * from IPM Where MONTH(dFecha) = " & Trim(Str(Month(dFecAdq))) & " And YEAR(dFecha) = " & Trim(Str(Year(dFecAdq)))
'        R.Open sSQL, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
'            If Not R.BOF And Not R.EOF Then
'                FactAFecha = R!nFactor1
'            End If
'        R.Close
'
'    FactorAjuste = CDbl(Format(FacRep / FactAFecha, "#0.0000"))
'
'End Function

Function EsOperador(ByVal sCad As String) As Boolean
    If sCad = "^" Or sCad = "*" Or sCad = "/" Or sCad = "+" Or sCad = "-" Or sCad = "(" Or sCad = ")" Then
        EsOperador = True
    Else
        EsOperador = False
    End If
End Function

Function IndCad(ByVal sCad As String, ByVal PosIni As Integer, ByVal cCadCom As String) As Integer
Dim I As Integer
Dim Contpar As Integer
    Contpar = 0
    For I = PosIni To Len(sCad)
        If Mid(sCad, I, 1) = cCadCom And Contpar = 0 Then
            IndCad = I
            Exit Function
        End If
        If Mid(sCad, I, 1) = "(" Then
            Contpar = Contpar + 1
        End If
        If Mid(sCad, I, 1) = ")" And Contpar > 0 Then
            Contpar = Contpar - 1
        End If
    Next I
    IndCad = 0
End Function

Function InterpretaFormula(ByVal CadFor As String) As String
Dim I As Integer
Dim j As Integer
Dim cOpeComp As String
Dim sOperad As String
Dim sDigito() As String
Dim nDigito As Integer
'Dim nOperad As Integer
Dim cDigito As String
Dim cFormula As String
Dim CadTemp As String
Dim bSal As Boolean
Dim sCad As String
Dim sCadArbolTemp() As String
Dim nCadArbolTemp As Integer
Dim d() As TDigito
Dim ND As Integer
Dim Pos As Integer
Dim k As Integer
Dim BajaDigito As Integer
Dim IndOper As Integer
Dim ContBajar As Integer
Dim Seguir As Boolean
Dim X As Integer
Dim UltPosOper As Integer
Dim Pos2 As Integer
Dim nIndpos2 As Integer
    'nOperad = 0
    ND = 0
    ReDim d(0)
    cFormula = CadFor
    sOperad = ""
    cDigito = ""
    nCadArbol = 0
    ReDim sCadArbol(0)
    nDigito = 0
    ReDim sDigito(0)
    I = 1
    
    Do While I <= Len(CadFor)
                
        If EsOperador(Mid(cFormula, I, 1)) Then
            If Mid(cFormula, I, 1) = "(" Then
                nIndpos2 = IndCad(CadFor, I + 1, ")") + 1
                If (cOpeComp = "+" Or cOpeComp = "-") And ((Mid(cFormula, nIndpos2, 1) = "*" Or Mid(cFormula, nIndpos2, 1) = "/") Or Mid(cFormula, nIndpos2, 1) = "^") Then
                    Pos = 0
                    IndOper = Len(sOperad)
                                'baja operador
                                For X = IndOper To 1 Step -1
                                    nCadArbol = nCadArbol + 1
                                    ReDim Preserve sCadArbol(nCadArbol)
                                    sCadArbol(nCadArbol - 1) = Mid(sOperad, X, 1)
                                Next X
                                
                                Pos = 0
                                nDigito = UBound(sDigito)
                                For X = 1 To nDigito
                                        If sDigito(X - 1) = "D" Then
                                            For k = 0 To d(Pos).nCadArbolTemp - 1
                                                nCadArbol = nCadArbol + 1
                                                ReDim Preserve sCadArbol(nCadArbol)
                                                sCadArbol(nCadArbol - 1) = d(Pos).sCadArbolTemp(k)
                                            Next k
                                            Pos = Pos + 1
                                        Else
                                            nCadArbol = nCadArbol + 1
                                            ReDim Preserve sCadArbol(nCadArbol)
                                            sCadArbol(nCadArbol - 1) = sDigito(X - 1)
                                        End If
                                Next X

                            nDigito = 0
                            ReDim sDigito(0)
                            ND = 0
                            ReDim d(0)
                            Pos = 0
                            sOperad = ""
                            cOpeComp = ""
                End If
                
                ReDim sCadArbolTemp(nCadArbol)
                nCadArbolTemp = nCadArbol
                For j = 0 To nCadArbol - 1
                    sCadArbolTemp(j) = sCadArbol(j)
                Next j
                
                cDigito = InterpretaFormula(Mid(cFormula, I + 1, IndCad(CadFor, I + 1, ")") - (I + 1)))
                nDigito = nDigito + 1
                ReDim Preserve sDigito(nDigito)
                sDigito(nDigito - 1) = "D"
                I = IndCad(CadFor, I + 1, ")")
                
                'Salva Detalle de Digitos
                ND = ND + 1
                ReDim Preserve d(ND)
                d(ND - 1).nCadArbolTemp = nCadArbol
                ReDim d(ND - 1).sCadArbolTemp(nCadArbol)
                For j = 0 To nCadArbol - 1
                    d(ND - 1).sCadArbolTemp(j) = sCadArbol(j)
                Next j
                
                'Restaura Cadena Principal
                ReDim sCadArbol(nCadArbolTemp)
                For j = 0 To nCadArbolTemp - 1
                    sCadArbol(j) = sCadArbolTemp(j)
                Next j
                nCadArbol = nCadArbolTemp
                
            Else
                If cOpeComp = "" Then
                    cOpeComp = Mid(cFormula, I, 1)
                    'Añado Operador
                    sOperad = sOperad + Mid(cFormula, I, 1)
                    If nDigito = 0 Then
                        'Añado Digito
                        nDigito = nDigito + 1
                        ReDim Preserve sDigito(nDigito)
                        sDigito(nDigito - 1) = "0"
                    End If
                    
                Else
                    'Comparacion de Prioridad
                        If ((cOpeComp = "*" Or cOpeComp = "/" Or cOpeComp = "^") And (Mid(cFormula, I, 1) = "+" Or Mid(cFormula, I, 1) = "-")) Then
                            sOperad = sOperad + Mid(cFormula, I, 1)
                            Pos = 0
                            IndOper = Len(sOperad)
                            
                                'baja operador
                                For X = IndOper To 1 Step -1
                                    nCadArbol = nCadArbol + 1
                                    ReDim Preserve sCadArbol(nCadArbol)
                                    sCadArbol(nCadArbol - 1) = Mid(sOperad, X, 1)
                                Next X
                                
                                Pos = 0
                                nDigito = UBound(sDigito)
                                For X = 1 To nDigito
                                        If sDigito(X - 1) = "D" Then
                                            For k = 0 To d(Pos).nCadArbolTemp - 1
                                                nCadArbol = nCadArbol + 1
                                                ReDim Preserve sCadArbol(nCadArbol)
                                                sCadArbol(nCadArbol - 1) = d(Pos).sCadArbolTemp(k)
                                            Next k
                                            Pos = Pos + 1
                                        Else
                                            nCadArbol = nCadArbol + 1
                                            ReDim Preserve sCadArbol(nCadArbol)
                                            sCadArbol(nCadArbol - 1) = sDigito(X - 1)
                                        End If
                                Next X

                            nDigito = 0
                            ReDim sDigito(0)
                            ND = 0
                            ReDim d(0)
                            Pos = 0
                            sOperad = ""
                            cOpeComp = ""
                        Else
                            'Añado Operador
                            sOperad = sOperad + Mid(cFormula, I, 1)
                            cOpeComp = Mid(cFormula, I, 1)
                        End If
                End If
            End If
        Else
            
            j = I
            bSal = False
            CadTemp = ""
            Do While Not bSal
                If (Mid(cFormula, j, 1) >= "0" And Mid(cFormula, j, 1) <= "9") Or (Mid(cFormula, j, 1) = ".") Then
                    CadTemp = CadTemp + Mid(cFormula, j, 1)
                Else
                    bSal = True
                    Exit Do
                End If
                j = j + 1
                If j > Len(cFormula) Then
                    bSal = True
                End If
            Loop
               
            If (cOpeComp = "+" Or cOpeComp = "-") And ((Mid(cFormula, j, 1) = "*" Or Mid(cFormula, j, 1) = "/") Or Mid(cFormula, j, 1) = "^") Then
                    Pos = 0
                    IndOper = Len(sOperad)
                                'baja operador
                                For X = IndOper To 1 Step -1
                                    nCadArbol = nCadArbol + 1
                                    ReDim Preserve sCadArbol(nCadArbol)
                                    sCadArbol(nCadArbol - 1) = Mid(sOperad, X, 1)
                                Next X
                                
                                Pos = 0
                                nDigito = UBound(sDigito)
                                For X = 1 To nDigito
                                        If sDigito(X - 1) = "D" Then
                                            For k = 0 To d(Pos).nCadArbolTemp - 1
                                                nCadArbol = nCadArbol + 1
                                                ReDim Preserve sCadArbol(nCadArbol)
                                                sCadArbol(nCadArbol - 1) = d(Pos).sCadArbolTemp(k)
                                            Next k
                                            Pos = Pos + 1
                                        Else
                                            nCadArbol = nCadArbol + 1
                                            ReDim Preserve sCadArbol(nCadArbol)
                                            sCadArbol(nCadArbol - 1) = sDigito(X - 1)
                                        End If
                                Next X

                            nDigito = 0
                            ReDim sDigito(0)
                            ND = 0
                            ReDim d(0)
                            Pos = 0
                            sOperad = ""
                            cOpeComp = ""
                End If
                
            'Añado Digito
            nDigito = nDigito + 1
            ReDim Preserve sDigito(nDigito)
            sDigito(nDigito - 1) = CadTemp
            I = j - 1
        End If
        
        I = I + 1
    Loop
    
    
    Pos = 0
    BajaDig = 0
    IndOper = 1
    j = 0
    If nCadArbol > 0 Then
        CadTemp = sCadArbol(0)
    Else
        CadTemp = ""
    End If
        
IndOper = Len(sOperad)

If IndOper > 0 Or UBound(sDigito) > 0 Then
        'baja operador
        For X = IndOper To 1 Step -1
            nCadArbol = nCadArbol + 1
            ReDim Preserve sCadArbol(nCadArbol)
            sCadArbol(nCadArbol - 1) = Mid(sOperad, X, 1)
        Next X
        
        Pos = 0
        For X = 1 To UBound(sDigito)
                If sDigito(X - 1) = "D" Then
                    For k = 0 To d(Pos).nCadArbolTemp - 1
                        nCadArbol = nCadArbol + 1
                        ReDim Preserve sCadArbol(nCadArbol)
                        sCadArbol(nCadArbol - 1) = d(Pos).sCadArbolTemp(k)
                    Next k
                    Pos = Pos + 1
                Else
                    nCadArbol = nCadArbol + 1
                    ReDim Preserve sCadArbol(nCadArbol)
                    sCadArbol(nCadArbol - 1) = sDigito(X - 1)
                End If
        Next X
    
        nDigito = 0
        ReDim sDigito(0)
        ND = 0
        ReDim d(0)
        Pos = 0
        sOperad = ""
        cOpeComp = ""
End If

    For I = 0 To nCadArbol - 1
        sCad = sCad + sCadArbol(I)
    Next I
    InterpretaFormula = sCad
End Function
Sub LlenaArbol()
Dim I As Integer
Dim Pos As Integer
    Pos = ContArbol - 1
    Arbol(Pos).cDigito = sCadArbol(Pos)
    If ContArbol <= (nCadArbol - 1) Then
        If EsOperador(sCadArbol(Pos)) Then
            'Por Izquierda
            Arbol(Pos).ApIzq = Trim(Str(ContArbol + 1))
            ContArbol = ContArbol + 1
            ReDim Preserve Arbol(ContArbol)
            Arbol(ContArbol - 1).cDirec = ContArbol
            LlenaArbol
            
            If ContArbol <= (nCadArbol - 1) Then
                'Por Derecha
                Arbol(Pos).ApDer = Trim(Str(ContArbol + 1))
                ContArbol = ContArbol + 1
            ReDim Preserve Arbol(ContArbol)
                Arbol(ContArbol - 1).cDirec = ContArbol
                LlenaArbol
            Else
                Arbol(Pos).ApDer = "0"
            End If
        Else
            Arbol(Pos).ApDer = "0"
            Arbol(Pos).ApIzq = "0"
        End If
    Else
        Arbol(Pos).ApDer = "0"
        Arbol(Pos).ApIzq = "0"
    End If
    
        
End Sub
Function LeeArbol(ByVal Direc As String) As Double
Dim N As Double
Dim NI As Double
Dim ND As Double
Dim Pos As Integer
Dim I As Integer
    For I = 0 To ContArbol - 1
        If Direc = Arbol(I).cDirec Then
            Pos = I
            Exit For
        End If
    Next I
    
    If EsOperador(Arbol(Pos).cDigito) Then
            Select Case Arbol(Pos).cDigito
                Case "*"
                    NI = LeeArbol(Arbol(Pos).ApIzq)
                    ND = LeeArbol(Arbol(Pos).ApDer)
                    LeeArbol = NI * ND
                Case "/"
                    NI = LeeArbol(Arbol(Pos).ApIzq)
                    ND = LeeArbol(Arbol(Pos).ApDer)
                    LeeArbol = NI / ND
                Case "^"
                    NI = LeeArbol(Arbol(Pos).ApIzq)
                    ND = LeeArbol(Arbol(Pos).ApDer)
                    LeeArbol = NI ^ ND
                Case "+"
                    NI = LeeArbol(Arbol(Pos).ApIzq)
                    If Arbol(Pos).ApDer = "0" Then
                        LeeArbol = NI
                    Else
                        ND = LeeArbol(Arbol(Pos).ApDer)
                        LeeArbol = NI + ND
                    End If
                Case "-"
                    NI = LeeArbol(Arbol(Pos).ApIzq)
                    If Arbol(Pos).ApDer = "0" Then
                        LeeArbol = -1 * NI
                    Else
                        ND = LeeArbol(Arbol(Pos).ApDer)
                        LeeArbol = NI - ND
                    End If
                    
            End Select
    Else
            LeeArbol = Val(Arbol(Pos).cDigito)
    End If
    
End Function
Function ExprANum(ByVal Expr As String) As Double
Dim Cad As String
Dim CadRes As String

ContArbol = 0
ReDim Arbol(0)
    'Proceso de Llenado ded Arbol
    Cad = Expr
    If Cad = "" Then
      Exit Function
    End If
    If Not ExprValida(Cad) Then
        ExprANum = 0
        Exit Function
    End If
    
    CadRes = InterpretaFormula(Cad)
    
    ContArbol = ContArbol + 1
    ReDim Preserve Arbol(ContArbol)
    Arbol(ContArbol - 1).cDirec = ContArbol
    Arbol(ContArbol - 1).bOper = True
    Arbol(ContArbol - 1).cDigito = sCadArbol(ContArbol - 1)
    
    If nCadArbol = 1 Then
        Arbol(0).ApIzq = "0"
        Arbol(0).ApDer = "0"
    Else
        'Por Izquierda
        Arbol(0).ApIzq = Trim(Str(ContArbol + 1))
        ContArbol = ContArbol + 1
        ReDim Preserve Arbol(ContArbol)
        Arbol(ContArbol - 1).cDirec = ContArbol
        LlenaArbol
        
        If ContArbol <= (nCadArbol - 1) Then
        'Por Derecha
        Arbol(0).ApDer = Trim(Str(ContArbol + 1))
        ContArbol = ContArbol + 1
        ReDim Preserve Arbol(ContArbol)
        Arbol(ContArbol - 1).cDirec = ContArbol
        LlenaArbol
        Else
            Arbol(0).ApDer = "0"
        End If
    End If
    'Lectura de Arbol
    ExprANum = LeeArbol("1")
End Function

Private Function ExprValida(CadFor As String) As Boolean
Dim k As Integer
Dim I As Integer
Dim bEnc As Boolean
Dim cPar As Integer
Dim OpAnt As String
Dim Contpar As Integer
Dim Pos As Integer
Dim Cad As String
Dim bEnc2 As Boolean
    CadFor = Trim(CadFor)
        
    'Eliminar Espacios en Blanco
    bEnc = True
    Do While bEnc
        bEnc = False
        For I = 1 To Len(CadFor)
            If Mid(CadFor, I, 1) = " " Then
                CadFor = Mid(CadFor, 1, I - 1) + Mid(CadFor, I + 1, Len(CadFor) - I)
                bEnc = True
                Exit For
            End If
        Next I
    Loop
        
    'Que no existe un caracter no valido
    For I = 1 To Len(CadFor)
        If Not ((Mid(CadFor, I, 1) >= "0" And Mid(CadFor, I, 1) <= "9") Or (Mid(CadFor, I, 1) = "+" Or Mid(CadFor, I, 1) = "-" Or Mid(CadFor, I, 1) = "*" Or Mid(CadFor, I, 1) = "/" Or Mid(CadFor, I, 1) = "^" Or Mid(CadFor, I, 1) = "(" Or Mid(CadFor, I, 1) = ")" Or Mid(CadFor, I, 1) = ".")) Then
            MsgBox "Existe un Caracter no Valido en la expresion", vbInformation, "Aviso"
            ExprValida = False
            Exit Function
        End If
    Next I
    
    'Valida Expresion
    For I = 1 To Len(CadFor)
        Select Case Mid(CadFor, I, 1)
            Case "/", "*", "+", "-", "^"
                If Not ((Mid(CadFor, I + 1, 1) >= "0" And Mid(CadFor, I + 1, 1) <= "9") Or (Mid(CadFor, I + 1, 1) = "(") Or (Mid(CadFor, I + 1, 1) = "-")) Then
                    MsgBox "Expresion tiene posiblemente dos operadores seguidos", vbInformation, "Aviso"
                    ExprValida = False
                    Exit Function
                End If
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                If Mid(CadFor, I + 1, 1) = "(" Then
                    MsgBox "Expresion tiene posiblemente un digito y un ( juntos", vbInformation, "Aviso"
                    ExprValida = False
                    Exit Function
                End If
            Case "."
                If Not (Mid(CadFor, I + 1, 1) >= "0" And Mid(CadFor, I + 1, 1) <= "9") Then
                    MsgBox "Expresion tiene un operador no valido despues de un punto decimal", vbInformation, "Aviso"
                    ExprValida = False
                    Exit Function
                End If
            Case "("
                If Not ((Mid(CadFor, I + 1, 1) >= "0" And Mid(CadFor, I + 1, 1) <= "9") Or (Mid(CadFor, I + 1, 1) = "(") Or (Mid(CadFor, I + 1, 1) = "-") Or (Mid(CadFor, I + 1, 1) = "+")) Then
                    MsgBox "Expresion tiene un operador no valido despues de un perentesis abierto ", vbInformation, "Aviso"
                    ExprValida = False
                    Exit Function
                End If
            Case ")"
                If (Mid(CadFor, I + 1, 1) >= "0" And Mid(CadFor, I + 1, 1) <= "9") Then
                    MsgBox "Expresion tiene un operador no valido despues de un perentesis Cerrado ", vbInformation, "Aviso"
                    ExprValida = False
                    Exit Function
                End If
        End Select
    Next I
    
    'valida parentesis abiertos y cerrados
    cPar = 0
    For I = 1 To Len(CadFor)
        If Mid(CadFor, I, 1) = "(" Then
            cPar = cPar + 1
        End If
        If Mid(CadFor, I, 1) = ")" Then
            cPar = cPar - 1
        End If
    Next I
    If cPar <> 0 Then
        MsgBox "El numero de parentesis abiertos no es igual al numero de parentesis cerrados", vbInformation, "Aviso"
        ExprValida = False
        Exit Function
    End If
    
    'Valida casos como --
    OpAnt = ""
    I = 1
    Do While I < Len(CadFor)
        If (OpAnt = "+" Or OpAnt = "-") And (Mid(CadFor, I, 1) = "+" Or Mid(CadFor, I, 1) = "-") Then
            CadFor = Mid(CadFor, 1, I - 1) + "(0" + Mid(CadFor, I, Len(CadFor) - (I - 1))
            bEnc = False
            Contpar = 0
            For k = I + 3 To Len(CadFor)
                If EsOperador(Mid(CadFor, k, 1)) Then
                    CadFor = Mid(CadFor, 1, k - 1) + ")" + Mid(CadFor, k, Len(CadFor) - (k - 1))
                    bEnc = True
                    Exit For
                End If
            Next k
            If bEnc = False Then
                If (Contpar = 0) And ((Mid(CadFor, Len(CadFor), 1) >= "0" And Mid(CadFor, Len(CadFor), 1) <= "9") Or Mid(CadFor, Len(CadFor), 1) = ")") Then
                    bEnc = True
                    CadFor = CadFor + ")"
                End If
            End If
            I = I + 4
        End If
        OpAnt = Mid(CadFor, I, 1)
        I = I + 1
Loop
    bEnc = False
    If Mid(CadFor, 1, 1) = "-" Or Mid(CadFor, 1, 1) = "+" Then
        CadFor = "(0" + CadFor
        Contpar = 0
        For k = 4 To Len(CadFor)
            If Mid(CadFor, k, 1) = "(" Then
                Contpar = Contpar + 1
            End If
            If Mid(CadFor, k, 1) = ")" Then
                Contpar = Contpar - 1
                nPos = k
            End If
            If Contpar = 0 And EsOperador(Mid(CadFor, k, 1)) Then
                CadFor = Mid(CadFor, 1, k) + ")" + Mid(CadFor, k + 1, Len(CadFor) - k)
                bEnc = True
                Exit For
            End If
        Next k
        If Not bEnc Then
            CadFor = CadFor + ")"
        End If
    End If
    ExprValida = True
End Function
