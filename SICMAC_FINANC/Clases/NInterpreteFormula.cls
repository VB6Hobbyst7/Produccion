VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NInterpreteFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Dim J As Integer
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
Dim K As Integer
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
                                            For K = 0 To d(Pos).nCadArbolTemp - 1
                                                nCadArbol = nCadArbol + 1
                                                ReDim Preserve sCadArbol(nCadArbol)
                                                sCadArbol(nCadArbol - 1) = d(Pos).sCadArbolTemp(K)
                                            Next K
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
                For J = 0 To nCadArbol - 1
                    sCadArbolTemp(J) = sCadArbol(J)
                Next J
                
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
                For J = 0 To nCadArbol - 1
                    d(ND - 1).sCadArbolTemp(J) = sCadArbol(J)
                Next J
                
                'Restaura Cadena Principal
                ReDim sCadArbol(nCadArbolTemp)
                For J = 0 To nCadArbolTemp - 1
                    sCadArbol(J) = sCadArbolTemp(J)
                Next J
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
                                            For K = 0 To d(Pos).nCadArbolTemp - 1
                                                nCadArbol = nCadArbol + 1
                                                ReDim Preserve sCadArbol(nCadArbol)
                                                sCadArbol(nCadArbol - 1) = d(Pos).sCadArbolTemp(K)
                                            Next K
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
            
            J = I
            bSal = False
            CadTemp = ""
            Do While Not bSal
                If (Mid(cFormula, J, 1) >= "0" And Mid(cFormula, J, 1) <= "9") Or (Mid(cFormula, J, 1) = ".") Then
                    CadTemp = CadTemp + Mid(cFormula, J, 1)
                Else
                    bSal = True
                    Exit Do
                End If
                J = J + 1
                If J > Len(cFormula) Then
                    bSal = True
                End If
            Loop
               
            If (cOpeComp = "+" Or cOpeComp = "-") And ((Mid(cFormula, J, 1) = "*" Or Mid(cFormula, J, 1) = "/") Or Mid(cFormula, J, 1) = "^") Then
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
                                            For K = 0 To d(Pos).nCadArbolTemp - 1
                                                nCadArbol = nCadArbol + 1
                                                ReDim Preserve sCadArbol(nCadArbol)
                                                sCadArbol(nCadArbol - 1) = d(Pos).sCadArbolTemp(K)
                                            Next K
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
            I = J - 1
        End If
        
        I = I + 1
    Loop
    
    
    Pos = 0
    BajaDigito = 0
    IndOper = 1
    J = 0
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
                    For K = 0 To d(Pos).nCadArbolTemp - 1
                        nCadArbol = nCadArbol + 1
                        ReDim Preserve sCadArbol(nCadArbol)
                        sCadArbol(nCadArbol - 1) = d(Pos).sCadArbolTemp(K)
                    Next K
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
                    If ND <> 0 Then
                       LeeArbol = NI / ND
                    Else
                       LeeArbol = 0
                    End If
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
Function ExprANum(ByVal Expr As String, Optional psCodOpe As String = "") As Double
Dim Cad As String
Dim CadRes As String

ContArbol = 0
ReDim Arbol(0)
    'Proceso de Llenado ded Arbol
    Cad = Expr
    If Cad = "" Then
      ExprANum = 0
      Exit Function
    End If
    If Not ExprValida(Cad, psCodOpe) Then
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

Private Function ExprValida(CadFor As String, Optional psCodOpe As String = "") As Boolean
Dim K As Integer
Dim I As Integer
Dim bEnc As Boolean
Dim cPar As Integer
Dim OpAnt As String
Dim Contpar As Integer
Dim nPos As Integer
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
            MsgBox "Expresión " & psCodOpe & ": " & CadFor & ". Existe un Caracter no Válido en la expresion", vbInformation, "Aviso"
            ExprValida = False
            Exit Function
        End If
    Next I
    
    'Valida Expresion
    For I = 1 To Len(CadFor)
        Select Case Mid(CadFor, I, 1)
            Case "/", "*", "+", "-", "^"
                If Not ((Mid(CadFor, I + 1, 1) >= "0" And Mid(CadFor, I + 1, 1) <= "9") Or (Mid(CadFor, I + 1, 1) = "(") Or (Mid(CadFor, I + 1, 1) = "-")) Then
                    MsgBox "Expresion " & psCodOpe & ": " & CadFor & ". Tiene posiblemente dos operadores seguidos", vbInformation, "Aviso"
                    ExprValida = False
                    Exit Function
                End If
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                If Mid(CadFor, I + 1, 1) = "(" Then
                    MsgBox "Expresion " & psCodOpe & ": Tiene posiblemente un digito y un ( juntos", vbInformation, "Aviso"
                    ExprValida = False
                    Exit Function
                End If
            Case "."
                If Not (Mid(CadFor, I + 1, 1) >= "0" And Mid(CadFor, I + 1, 1) <= "9") Then
                    MsgBox "Expresion " & psCodOpe & ": Tiene un operador no valido despues de un punto decimal", vbInformation, "Aviso"
                    ExprValida = False
                    Exit Function
                End If
            Case "("
                If Not ((Mid(CadFor, I + 1, 1) >= "0" And Mid(CadFor, I + 1, 1) <= "9") Or (Mid(CadFor, I + 1, 1) = "(") Or (Mid(CadFor, I + 1, 1) = "-") Or (Mid(CadFor, I + 1, 1) = "+")) Then
                    MsgBox "Expresion " & psCodOpe & ": Tiene un operador no valido despues de un perentesis abierto ", vbInformation, "Aviso"
                    ExprValida = False
                    Exit Function
                End If
            Case ")"
                If (Mid(CadFor, I + 1, 1) >= "0" And Mid(CadFor, I + 1, 1) <= "9") Then
                    MsgBox "Expresion " & psCodOpe & ": Tiene un operador no valido despues de un perentesis Cerrado ", vbInformation, "Aviso"
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
            For K = I + 3 To Len(CadFor)
                If EsOperador(Mid(CadFor, K, 1)) Then
                    CadFor = Mid(CadFor, 1, K - 1) + ")" + Mid(CadFor, K, Len(CadFor) - (K - 1))
                    bEnc = True
                    Exit For
                End If
            Next K
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
        For K = 4 To Len(CadFor)
            If Mid(CadFor, K, 1) = "(" Then
                Contpar = Contpar + 1
            End If
            If Mid(CadFor, K, 1) = ")" Then
                Contpar = Contpar - 1
                nPos = K
            End If
            If Contpar = 0 And EsOperador(Mid(CadFor, K, 1)) Then
                CadFor = Mid(CadFor, 1, K) + ")" + Mid(CadFor, K + 1, Len(CadFor) - K)
                bEnc = True
                Exit For
            End If
        Next K
        If Not bEnc Then
            CadFor = CadFor + ")"
        End If
    End If
    ExprValida = True
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

End Sub

Public Function EjecutaFuncion(psFuncion As String) As Currency
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    On Error GoTo ERROR
    
    oCon.AbreConexion
    
    sql = "Select " & psFuncion & " As Valor "
    Set rs = oCon.CargaRecordSet(sql)
    
    EjecutaFuncion = rs!Valor
    
    oCon.CierraConexion
    Set oCon = Nothing
    rs.Close
    Set rs = Nothing
    Exit Function
ERROR:
    EjecutaFuncion = 0
    Set oCon = Nothing
    Set rs = Nothing
End Function



 
