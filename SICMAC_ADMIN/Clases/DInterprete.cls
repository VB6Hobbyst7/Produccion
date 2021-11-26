VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DInterprete"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9EA7630289"
Option Base 0
Option Explicit

Private Type Operacion
    lsOpeNom As String
    lsOpeTip As String
    lnOpeArg As Integer
End Type

Private Type Tabla_Valores
    lsConCod As String
    lsValor As String
    lbValorN As Boolean
End Type

Private Interprete_Ope(100) As Operacion
Private lnInterprete_OpeTop As Integer
Private TValor(100) As Tabla_Valores
Private lnTValorTop As Integer
Private PosOpe(100) As Integer
Private lcTipOpe(100) As String
Private VSQL As String
Private pdFecSis As Date

Dim oCon As DConecta

Dim cadena As String

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB9030C01C6
    Private mlClassDebugID As Long
#End If

Public Sub Reinicia()
    lnTValorTop = 1
End Sub

'Tipos
'01: Operadores Aritmeticos
'02: Agrupacion
'03: Funciones Internas
'04: Operadores Logicos
'05: Operadores Fecha
Public Sub Interprete_InI()
    Interprete_Ope(1).lsOpeNom = "+"
    Interprete_Ope(1).lsOpeTip = "01"
    Interprete_Ope(1).lnOpeArg = 2
    
    Interprete_Ope(2).lsOpeNom = "-"
    Interprete_Ope(2).lsOpeTip = "01"
    Interprete_Ope(2).lnOpeArg = 2
    
    Interprete_Ope(3).lsOpeNom = "*"
    Interprete_Ope(3).lsOpeTip = "01"
    Interprete_Ope(3).lnOpeArg = 2
    
    Interprete_Ope(4).lsOpeNom = "/"
    Interprete_Ope(4).lsOpeTip = "01"
    Interprete_Ope(4).lnOpeArg = 2
    
    Interprete_Ope(5).lsOpeNom = "("
    Interprete_Ope(5).lsOpeTip = "02"
    Interprete_Ope(5).lnOpeArg = 1
    
    Interprete_Ope(6).lsOpeNom = ")"
    Interprete_Ope(6).lsOpeTip = "02"
    Interprete_Ope(6).lnOpeArg = 1
    
    Interprete_Ope(7).lsOpeNom = "SI("
    Interprete_Ope(7).lsOpeTip = "03"
    Interprete_Ope(7).lnOpeArg = 1
    
    Interprete_Ope(8).lsOpeNom = "SUMA("
    Interprete_Ope(8).lsOpeTip = "03"
    Interprete_Ope(8).lnOpeArg = 1
    
    Interprete_Ope(9).lsOpeNom = "GET("
    Interprete_Ope(9).lsOpeTip = "03"
    Interprete_Ope(9).lnOpeArg = 1
    
    Interprete_Ope(10).lsOpeNom = "CUENTA("
    Interprete_Ope(10).lsOpeTip = "03"
    Interprete_Ope(10).lnOpeArg = 1
    
    Interprete_Ope(11).lsOpeNom = "SUMAGRUPO("
    Interprete_Ope(11).lsOpeTip = "03"
    Interprete_Ope(11).lnOpeArg = 1
    
    Interprete_Ope(12).lsOpeNom = "DIA("
    Interprete_Ope(12).lsOpeTip = "03"
    Interprete_Ope(12).lnOpeArg = 1
   
    Interprete_Ope(13).lsOpeNom = "MES("
    Interprete_Ope(13).lsOpeTip = "03"
    Interprete_Ope(13).lnOpeArg = 1
    
    Interprete_Ope(14).lsOpeNom = "AÑO("
    Interprete_Ope(14).lsOpeTip = "03"
    Interprete_Ope(14).lnOpeArg = 1
    
    Interprete_Ope(15).lsOpeNom = "OBTCONCEP("
    Interprete_Ope(15).lsOpeTip = "03"
    Interprete_Ope(15).lnOpeArg = 1
   
    'Operadores Logicos
    Interprete_Ope(16).lsOpeNom = "<="
    Interprete_Ope(16).lsOpeTip = "04"
    Interprete_Ope(16).lnOpeArg = 1
    
    Interprete_Ope(17).lsOpeNom = ">="
    Interprete_Ope(17).lsOpeTip = "04"
    Interprete_Ope(17).lnOpeArg = 1
    
    Interprete_Ope(18).lsOpeNom = "<>"
    Interprete_Ope(18).lsOpeTip = "04"
    Interprete_Ope(18).lnOpeArg = 1
    
    Interprete_Ope(19).lsOpeNom = "><"
    Interprete_Ope(19).lsOpeTip = "04"
    Interprete_Ope(19).lnOpeArg = 1
    
    Interprete_Ope(20).lsOpeNom = "="
    Interprete_Ope(20).lsOpeTip = "04"
    Interprete_Ope(20).lnOpeArg = 1
    
    Interprete_Ope(21).lsOpeNom = ">"
    Interprete_Ope(21).lsOpeTip = "04"
    Interprete_Ope(21).lnOpeArg = 1
    
    Interprete_Ope(22).lsOpeNom = "<"
    Interprete_Ope(22).lsOpeTip = "04"
    Interprete_Ope(22).lnOpeArg = 1
    
    'Operadores Fecha
    Interprete_Ope(23).lsOpeNom = "NOW"
    Interprete_Ope(23).lsOpeTip = "05"
    Interprete_Ope(23).lnOpeArg = 1
    
    Interprete_Ope(24).lsOpeNom = "DATE"
    Interprete_Ope(24).lsOpeTip = "05"
    Interprete_Ope(24).lnOpeArg = 1
    
    Interprete_Ope(25).lsOpeNom = "FINI"
    Interprete_Ope(25).lsOpeTip = "05"
    Interprete_Ope(25).lnOpeArg = 1
    
    Interprete_Ope(26).lsOpeNom = "FFIN"
    Interprete_Ope(26).lsOpeTip = "05"
    Interprete_Ope(26).lnOpeArg = 1
    
    Interprete_Ope(27).lsOpeNom = "TIME"
    Interprete_Ope(27).lsOpeTip = "05"
    Interprete_Ope(27).lnOpeArg = 1
    
    Interprete_Ope(28).lsOpeNom = "VALORORIGINAL("
    Interprete_Ope(28).lsOpeTip = "03"
    Interprete_Ope(28).lnOpeArg = 1
    
    lnInterprete_OpeTop = 28
End Sub


Public Sub IniTValor(pbTodos As Boolean)
    Dim I As Integer
    Dim lnContador As Integer
    
    lnContador = 0
    
    If Not pbTodos Then
        For I = 1 To lnTValorTop
            If Left(TValor(I).lsConCod, 2) = "G_" Then
                lnContador = lnContador + 1
                TValor(lnContador).lbValorN = TValor(I).lbValorN
                TValor(lnContador).lsConCod = TValor(I).lsConCod
                TValor(lnContador).lsValor = TValor(I).lsValor
            End If
        Next I
    End If
    
    lnTValorTop = lnContador
End Sub

Public Function GetTValor(psNemo As String)
    Dim I As Integer
    Dim lsResultado As String
    
    lsResultado = ""
    For I = 1 To lnTValorTop
       If TValor(I).lsConCod = psNemo Then
            lsResultado = TValor(I).lsValor
            I = lnTValorTop
       End If
    Next I
    
    GetTValor = lsResultado
End Function

Public Sub SetTValor(psNemo As String, psValor As String, pbValorN As Boolean)
    Dim lnValor As Double
    
    If lnTValorTop < 100 Then
        lnTValorTop = lnTValorTop + 1
        TValor(lnTValorTop).lsConCod = psNemo
        TValor(lnTValorTop).lbValorN = pbValorN
        
        psValor = Replace(psValor, ",", "")
        If Trim(psValor) = "" Then
            lnValor = "0"
        Else
            lnValor = ExprANum(psValor)
        End If
        
        If lnValor = "0" Then
            TValor(lnTValorTop).lsValor = "0.00"
        ElseIf lnValor < 5 Then
            TValor(lnTValorTop).lsValor = Format(ExprANum(psValor), "#0.000000")
        Else
            TValor(lnTValorTop).lsValor = Format(ExprANum(psValor), "#0.0000")
        End If
            
        If pbValorN And Left(psNemo, 2) <> "U_" Then
            TValor(lnTValorTop).lsValor = Format(TValor(lnTValorTop).lsValor, "#0.00000")
        End If
        
        If pbValorN And Left(psNemo, 2) = "I_" Or pbValorN And Left(psNemo, 2) = "D_" Or pbValorN And Left(psNemo, 2) = "A_" Then
            TValor(lnTValorTop).lsValor = Format(TValor(lnTValorTop).lsValor, "#0.00")
        End If

    End If
End Sub

Public Function EsConceptoEmpleado(psConcepto As String, psPersCod As String, psCodPla As String) As Boolean
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    Dim sqlC As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    sqlC = " Select cPersCod from rhconcepto RHC" _
         & " Inner Join RHConceptoTabla RHCT On RHC.cRHConceptoCod = RHCT.cRHConceptoCod" _
         & " And cRHconceptoMeno = '" & psConcepto & "'" _
         & " Where cPersCod = '" & psPersCod & "' And cRHPlanillaCod = '" & psCodPla & "'"
    
    'If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        If rs.EOF And rs.BOF Then
            EsConceptoEmpleado = False
        Else
            EsConceptoEmpleado = True
        End If
    'End If
    
    rs.Close
    Set rs = Nothing
    'Set oCon = Nothing
End Function

Public Function FunFiltro(psFuncion As String) As String
    Dim lsFun As String
    Dim lsCadena As String
    Dim I As Integer
    Dim lnTop As Integer
    Dim lbBan As Boolean
    
    lsFun = Replace(psFuncion, "  ", " ", , , vbTextCompare)
    lsFun = Replace(lsFun, Chr(13), "", , , vbTextCompare)
    lsFun = Replace(lsFun, oImpresora.gPrnSaltoLinea, "", , , vbTextCompare)
    lnTop = Len(lsFun)
    lbBan = True
    lsFun = Replace(lsFun, Chr(34), Chr(39))
    
    FunFiltro = lsFun
End Function

Public Function GetValorFunLog(psFuncion As String, psPersCod As String, pdFecIni As Date, pdFecFin As Date, pbValorN As Boolean, ByVal psPlanilla As String, ByVal psPeriodo As String) As String
    Dim lnPosT As Integer
    Dim lnPos As Integer
    Dim lnPosLog As Integer
    Dim lsFun As String
    Dim lsCadena As String
    Dim RESULTADO As String
    Dim lsSubCadena As String
    Dim lsCadAnt As String
    Dim lsOpeInt As String
    
    On Error GoTo ERROR
    
    lsFun = GetFunConcepto(psFuncion)
    lnPos = FunGetPosOpe(lsFun, True, True)
       
    lsCadena = ""
    lsOpeInt = Interprete_Ope(lnPos).lsOpeNom
    lnPosT = PosOpe(lnPos)
    lnPosLog = GetPosParFunLog(lsFun, lnPosT)
    
    If lnPos = 0 Then
        lsCadena = psFuncion
    Else
    While lnPos <> 0
        lsSubCadena = Mid(lsFun, lnPosT + Len(Interprete_Ope(lnPos).lsOpeNom), lnPosLog - lnPosT - Len(Interprete_Ope(lnPos).lsOpeNom) + 1)
        lsOpeInt = Interprete_Ope(lnPos).lsOpeNom
        lsCadAnt = Mid(lsFun, 1, lnPosT - 1)
        
        RESULTADO = ""
        Select Case lsOpeInt
            Case "GET("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaGET(lsSubCadena, psPersCod, pdFecIni, pdFecFin)
            Case "SUMA("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaSuma(lsSubCadena, psPersCod, pdFecIni, pdFecFin)
            Case "CUENTA("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaCuenta(lsSubCadena, psPersCod, pdFecIni, pdFecFin)
            Case "SI("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaSI(lsSubCadena, psPersCod, pdFecIni, pdFecFin, psPeriodo)
            Case "SUMAGRUPO("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaSumaGrupo(lsSubCadena, psPersCod, pdFecIni, pdFecFin)
            Case "DIA("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaDIAMESAÑO(lsSubCadena, psPersCod, pdFecIni, pdFecFin, "1")
            Case "MES("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaDIAMESAÑO(lsSubCadena, psPersCod, pdFecIni, pdFecFin, "2")
            Case "AÑO("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaDIAMESAÑO(lsSubCadena, psPersCod, pdFecIni, pdFecFin, "3")
            Case "OBTCONCEP("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaObtConcep(lsSubCadena, psPersCod, pdFecIni, pdFecFin)
            Case "VALORORIGINAL("
                lsSubCadena = GetValorFunLog(lsSubCadena, psPersCod, pdFecIni, pdFecFin, False, psPlanilla, psPeriodo)
                RESULTADO = FuncionInternaValorOriginal(lsSubCadena, psPersCod, pdFecIni, pdFecFin)
        End Select
        
        lsCadena = lsCadena & lsCadAnt & RESULTADO
        
        If EsConceptoValido(psFuncion) Then SetTValor psFuncion, lsCadena, False
        
        lsFun = Mid(lsFun, lnPosLog + 1)
        lnPos = FunGetPosOpe(lsFun, True, True)
        lnPosT = PosOpe(lnPos)
        lnPosLog = GetPosParFunLog(lsFun, lnPosT)
    Wend
        lsCadena = lsCadena & lsFun
    End If
    'lsCadena = lsCadena & lsCadAnt & Resultado & lsFun
    If EsConceptoValido(psFuncion) And psFuncion <> lsCadena And FunGetPosOpe(lsFun) = 0 Then SetTValor psFuncion, lsCadena, pbValorN
    GetValorFunLog = lsCadena
    Exit Function
ERROR:
    MsgBox Err.Description, vbInformation, "Aviso"
End Function

Public Function GetFunConcepto(psConcepto As String) As String
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    Set rsC = New ADODB.Recordset
    Dim lsDato As String
    
    lsDato = Replace(psConcepto, "'", Chr(34))
    sqlC = "Select cRHConceptoFormula cFormula from rhconceptotabla Where cRHconceptoMeno = '" & lsDato & "'"
        
    'If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If RSVacio(rsC) Then
            GetFunConcepto = psConcepto
        Else
            lsDato = rsC!cFormula
            lsDato = Replace(lsDato, Chr(34), Chr(39))
            GetFunConcepto = Trim(lsDato)
        End If
        'oCon.CierraConexion
    'End If
    
    rsC.Close
    Set rsC = Nothing
    'Set oCon = Nothing
End Function

Public Function FunGetPosOpe(psFuncion As String, Optional pbLogicos As Boolean = False, Optional pbFunLogicos As Boolean = False, Optional pbFunFecha As Boolean = False) As Integer
    Dim I As Integer
    Dim j As Integer
    Dim lnPosMenor As Integer
    Dim lnMenor As Integer
    Dim lbBan As Boolean
    
    For I = 1 To lnInterprete_OpeTop
        PosOpe(I) = 0
        If Not pbLogicos Then
            If Interprete_Ope(I).lsOpeTip <> "04" And Interprete_Ope(I).lsOpeTip <> "03" And Interprete_Ope(I).lsOpeTip <> "05" Then
                PosOpe(I) = InStr(1, psFuncion, Interprete_Ope(I).lsOpeNom, vbTextCompare)
                lcTipOpe(I) = Interprete_Ope(I).lsOpeTip
                If PosOpe(I) <> 0 Then lnPosMenor = I
            End If
        Else
            If pbFunLogicos Then
                If Interprete_Ope(I).lsOpeTip = "03" Then
                    PosOpe(I) = InStr(1, psFuncion, Interprete_Ope(I).lsOpeNom, vbTextCompare)
                    If PosOpe(I) <> 0 Then lnPosMenor = I
                End If
            Else
                If Interprete_Ope(I).lsOpeTip = "04" Then
                    PosOpe(I) = InStr(1, psFuncion, Interprete_Ope(I).lsOpeNom, vbTextCompare)
                    If PosOpe(I) <> 0 Then lnPosMenor = I
                End If
            End If
        End If
    Next I
   
    lbBan = True
    
    For I = 1 To lnInterprete_OpeTop
        If PosOpe(I) <> 0 Then
            lnMenor = PosOpe(I)
            lnPosMenor = I
            I = lnInterprete_OpeTop
        End If
    Next I
    
    For I = 1 To lnInterprete_OpeTop
        If PosOpe(I) <> 0 Then
            If PosOpe(I) < lnMenor Then
                lnMenor = PosOpe(I)
                lnPosMenor = I
            End If
        End If
    Next I
    FunGetPosOpe = lnPosMenor
End Function

Public Function FunEvalua(psFuncion As String, psPersCod As String, psFecIni As Date, psFecFin As Date, pbValorN As Boolean, ByVal psPlanilla As String, ByVal psPeriodo As String) As String
    Dim lsV1 As String
    Dim lsV2 As String
    Dim lsCadena As String
    Dim lsFun As String
    Dim lsFunFor As String
    Dim lnPos As Integer
    Dim lsConcep As String
    Dim lnPosT As Integer
    Dim lnPosLog As Integer
    
    Dim lsResult As String
    
    Dim lbControl As Boolean
    Dim lsCadAux As String
    
    lsFun = psFuncion
    lnPos = FunGetPosOpe(lsFun)
    lnPosT = PosOpe(lnPos)
    
    If lnPosT = 0 Then
    
        If IsNumeric(lsFun) Then
            FunEvalua = lsFun
        Else
            'lsError = 0
            lsConcep = FunEsConcepto(lsFun)
            If lsConcep <> "X" Then
                lsResult = GetTValor(lsFun)
                If lsResult <> "" Then
                    FunEvalua = lsResult
                    Exit Function
                End If
            End If
            
            If lsConcep = "1" Then
                lsConcep = FunGetConstante(lsFun)
            ElseIf lsConcep = "2" Then
                lsConcep = FunGetValorConstante(lsFun, psPersCod, psFecIni, psFecFin, psPlanilla, psPeriodo)
            ElseIf lsConcep = "3" Then
                'lsConcep = FunGetValorConstanteSQL(lsFun, psPersCod, psPersCod, psFecIni, psFecFin)
                lsConcep = FunGetValorPreDef(lsFun, psPersCod, psFecIni, psFecFin, psPlanilla)
                lsConcep = FunEvalua(lsConcep, psPersCod, psFecIni, psFecFin, False, psPlanilla, psPeriodo)
                'MsgBox "aa"
            ElseIf lsConcep = "4" Then
                lsFunFor = GetFomulaCon(lsFun)
                lsFunFor = GetValorFunLog(lsFunFor, psPersCod, psFecIni, psFecFin, False, psPlanilla, psPeriodo)
                lsConcep = FunEvalua(lsFunFor, psPersCod, psFecIni, psFecFin, False, psPlanilla, psPeriodo)
            ElseIf lsConcep = "5" Then
                lsConcep = ""
            Else
            
            End If
            
            If lsFun = "" Then
                FunEvalua = ""
            ElseIf lsConcep = "X" Then
                MsgBox "Error de Sintaxis." & lsFun, vbInformation, "Aviso"
            Else
                If psPlanilla = gsRHPlanillaSueldos Then
                    If EsConceptoValido(psFuncion) Then
                        lsCadAux = GetMesTrabajado(psFuncion)
                        If lsCadAux <> "" Then
                            lsConcep = lsConcep & "*" & FunEvalua(lsCadAux, psPersCod, psFecIni, psFecFin, False, psPlanilla, psPeriodo)
                        End If
                        psPlanilla = ""
                    End If
                End If
                SetTValor lsFun, lsConcep, pbValorN
                FunEvalua = lsConcep
            End If
        End If
    Else
        lsCadena = ""
        While lsFun <> ""
            lsCadena = Trim(lsCadena) & Trim(FunEvalua(Trim(Left(lsFun, lnPosT - 1)), psPersCod, psFecIni, psFecFin, lbControl, psPlanilla, psPeriodo)) & Mid(lsFun, lnPosT, 1)
            lbControl = False
            lsFun = Mid(lsFun, lnPosT + 1)
            lnPos = FunGetPosOpe(lsFun)
            
            If lnPos = 0 Then
               lnPos = 1000
               lnPosT = 1000
            Else
                lnPosT = PosOpe(lnPos)
            End If
        Wend
        FunEvalua = lsCadena
    End If
    
End Function

Public Function EsTotalPlanilla(psConcepto As String) As Boolean
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    sqlC = "Select cRHConceptoCod From RHConceptoTabla Where nRHConceptoOrden = " & gnRHTotalTpo & " And cRHconceptoMeno = '" & psConcepto & "'"
    'If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC) ', adLockReadOnly, adOpenForwardOnly)
        If RSVacio(rsC) Then
            EsTotalPlanilla = False
        Else
            EsTotalPlanilla = True
        End If
        'oCon.CierraConexion
    'End If
    
    Set rsC = Nothing
    'Set oCon = Nothing
End Function

Private Function GetPosParFunLog(psFuncion As String, pnPosIni As Integer) As Integer
    Dim I As Integer
    Dim lnTope As Integer
    Dim lnPar As Integer
    Dim lbBan As Boolean
    Dim lnResul As Integer
    
    lnPar = 0
    lnTope = Len(psFuncion)
    lbBan = False
    
    For I = pnPosIni + 1 To lnTope
        If Mid(psFuncion, I, 1) = "(" Then
            lnPar = lnPar + 1
        ElseIf Mid(psFuncion, I, 1) = ")" Then
            lnPar = lnPar - 1
        End If
        
        If lnPar = 0 Then
            If lbBan Then
                lnResul = I
                I = lnTope
            End If
        Else
            lnResul = I
            lbBan = True
        End If
    Next I
    
    GetPosParFunLog = lnResul
End Function


Public Function FuncionInternaGET(psFuncionGET As String, psPersCod As String, pdFecIni As Date, pdFecFin As Date) As String
    Dim lnPos As Integer
    Dim lnPosCor As Integer
    Dim lnPosLim As Integer
    Dim lnPosPar As Integer
    Dim lsNomTab As String
    Dim lsNomCam As String
    Dim lsEmpCod As String
    Dim lsParm As String
    Dim sqlF As String
    Dim rsF As ADODB.Recordset
    Set rsF = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    On Error GoTo ERROR
    
    lnPos = InStr(1, psFuncionGET, ")", vbTextCompare)
    
    If lnPos <> 0 Then
        lnPosCor = InStr(1, psFuncionGET, ".", vbTextCompare)
        lsNomTab = Mid(psFuncionGET, 1, lnPosCor - 1)
        lsNomTab = GetNomTablaCampo(lsNomTab)
        
        lnPosPar = InStr(1, psFuncionGET, ",", vbTextCompare)
        
        If lnPosPar = 0 Then
            lsNomCam = Mid(psFuncionGET, lnPosCor + 1)
            lsNomCam = GetNomTablaCampo(Left(lsNomCam, Len(lsNomCam) - 1))
            
            sqlF = "Select Resultado = " & lsNomCam & " From " & lsNomTab & " Where cPersCod = '" & psPersCod & "'"
        Else
            lsNomCam = Mid(psFuncionGET, lnPosCor + 1, lnPosPar - lnPosCor - 1)
            lsNomCam = GetNomTablaCampo(lsNomCam)
            
            lsParm = Mid(psFuncionGET, lnPosPar + 1, lnPos - lnPosPar - 1)
            lnPosPar = InStr(1, lsParm, ",", vbTextCompare)
            
            If lnPosPar = 0 Then
                sqlF = "Select Resultado = " & lsNomCam & " From " & lsNomTab & " Where cPersCod = '" & psPersCod & "' And " & lsParm
            Else
                lsEmpCod = Mid(lsParm, lnPosPar + 1)
                lsParm = Mid(lsParm, 1, lnPosPar - 1)
                
                If lsEmpCod = "TRUE" Then
                    sqlF = "Select Resultado = " & lsNomCam & " From " & lsNomTab & " Where cPersCod = '" & psPersCod & "' And " & lsParm
                Else
                    sqlF = "Select Resultado = " & lsNomCam & " From " & lsNomTab & " Where " & lsParm
                End If
            End If
        End If
        
        'If oCon.AbreConexion Then
            'If rsF.State = 0 Then
            '    If rsF.EOF And rsF.BOF Then
            '        FuncionInternaGET = "X"
            '    Else
                    Set rsF = oCon.CargaRecordSet(sqlF)
                    FuncionInternaGET = rsF!RESULTADO
            '    End If
            '    FuncionInternaGET = "X"
            'End If
        'End If
        
        rsF.Close
        Set rsF = Nothing
    Else
        FuncionInternaGET = "X"
    End If
    
    
    Exit Function
ERROR:
    FuncionInternaGET = "X"
    
    'If rsF.State = 1 Then
    '    rsF.Close
    '    Set rsF = Nothing
    'End If
End Function

Public Function FuncionInternaSuma(psFuncionSuma As String, psPersCod As String, pdFecIni As Date, pdFecFin As Date) As String
    Dim lnPos As Integer
    Dim lnPosCor As Integer
    Dim lnPosLim As Integer
    Dim lnPosPar1 As Integer
    Dim lnPosPar2 As Integer
    Dim lnPosPar3 As Integer
    Dim lnPosPar4 As Integer
    Dim lnPosPar5 As Integer
    Dim lnPosPar6 As Integer
    Dim lnPosOpe As Integer
    Dim lsNomTab As String
    Dim lsNomCam As String
    Dim lsParm1 As String
    Dim lsParm2 As String
    Dim lsParm3 As String
    Dim lsParm4 As String
    Dim lsParm5 As String
    Dim lsParm6 As String
    Dim sqlF As String
    Dim lsEmp As String
    Dim rsF As ADODB.Recordset
    Set rsF = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    On Error GoTo ERROR
        
    lnPos = InStr(1, psFuncionSuma, ")", vbTextCompare)
    
    If lnPos <> 0 Then
        lnPosCor = InStr(1, psFuncionSuma, ".", vbTextCompare)
        lsNomTab = Mid(psFuncionSuma, 1, lnPosCor - 1)
        lsNomTab = GetNomTablaCampo(lsNomTab)
        If lsNomTab = "X" Then GoTo ERROR
        
        lnPosPar1 = InStr(1, psFuncionSuma, ",", vbTextCompare)
        If lnPosPar1 = 0 Then
            lsNomCam = Mid(psFuncionSuma, lnPosCor + 1)
            lsNomCam = GetNomTablaCampo(Left(lsNomCam, Len(lsNomCam) - 1))
            
            sqlF = "Select Resultado = Sum(" & lsNomCam & ") From " & lsNomTab & " Where cPersCod = '" & psPersCod & "'"
        Else
            lsNomCam = Mid(psFuncionSuma, lnPosCor + 1, lnPosPar1 - lnPosCor - 1)
            lsNomCam = GetNomTablaCampo(lsNomCam)
            
            lnPosPar2 = InStr(lnPosPar1 + 1, psFuncionSuma, ",", vbTextCompare)
            
            If lnPosPar2 = 0 Or lnPos < lnPosPar2 Then
                lsParm1 = Mid(psFuncionSuma, lnPosPar1 + 1, lnPos - lnPosPar1 - 1)
                lsParm1 = GetNomTablaCampo(lsParm1)
                sqlF = "Select Resultado = Sum(" & lsNomCam & ") From " & lsNomTab & " Where cPersCod = '" & psPersCod & "' and " & lsParm1 & " between '" & Format(pdFecIni, gsFormatoFecha) & "' and '" & Format(pdFecFin, gsFormatoFecha) & "'"
            Else
                lsParm2 = Mid(psFuncionSuma, lnPosPar1 + 1, lnPosPar2 - lnPosPar1 - 1)
                If lsParm2 <> "" Then
                    lsParm2 = GetNomTablaCampo(lsParm2)
                    
                    lnPosPar3 = InStr(lnPosPar2 + 1, psFuncionSuma, ",", vbTextCompare)
                    
                    
                    If lnPosPar3 = 0 Then GoTo ERROR
                    
                    lsParm3 = Mid(psFuncionSuma, lnPosPar2 + 1, lnPosPar3 - lnPosPar2 - 1)
                    
                    lnPosOpe = FunGetPosOpe(lsParm3, True)
                    lnPosOpe = PosOpe(lnPosOpe)
                    
                    lsParm4 = Mid(lsParm3, lnPosOpe)
                    lsParm3 = Left(lsParm3, lnPosOpe - 1)
                    lsParm3 = GetNomTablaCampo(lsParm3)
                    
                    lnPosPar4 = InStr(lnPosPar3 + 1, psFuncionSuma, ",")
                    
                    lsEmp = "cPersCod = '" & psPersCod & "' and"
                    
                    If lnPosPar4 <> 0 Then
                        lsParm5 = Mid(psFuncionSuma, lnPosPar3 + 1, lnPos - lnPosPar3 - 1)
                    End If
                    
                    If lsParm5 <> "" Then
                        lnPosPar5 = InStr(1, lsParm5, ",")
                        If lnPosPar5 = 0 Then
                            If UCase(lsParm5) = "FALSE" Then lsEmp = ""
                        Else
                            lsParm5 = Left(lsParm5, lnPosPar5 - 1)
                            If UCase(lsParm5) = "FALSE" Then lsEmp = ""
                            
                            lsNomCam = Mid(psFuncionSuma, lnPosPar3 + lnPosPar5 + 1, lnPos - lnPosPar3 - lnPosPar5 - 1)
                        End If
                    End If
                    
                    'Para agregar las validaciones Adicionales
                    sqlF = "Select Resultado = Sum(" & lsNomCam & ") From " & lsNomTab & " Where " & lsEmp & "  " & lsParm2 & " between '" & Format(pdFecIni, gsFormatoFecha) & "' and '" & Format(pdFecFin, gsFormatoFecha) & "' and " & lsParm3 & lsParm4
                Else
                    lnPosPar3 = InStr(lnPosPar2 + 1, psFuncionSuma, ",", vbTextCompare)
                    
                    If lnPosPar3 = 0 Then GoTo ERROR
                    
                    lsParm3 = Mid(psFuncionSuma, lnPosPar2 + 1, lnPosPar3 - lnPosPar2 - 1)
                    
                    lnPosOpe = FunGetPosOpe(lsParm3, True)
                    lnPosOpe = PosOpe(lnPosOpe)
                    
                    lsParm4 = Mid(lsParm3, lnPosOpe)
                    lsParm3 = Left(lsParm3, lnPosOpe - 1)
                    lsParm3 = GetNomTablaCampo(lsParm3)
                    
                    'lsParm4 = Mid(psFuncionSuma, lnPosPar3 + 1, lnPos - lnPosPar3 - 1)
                    
                    lnPosPar4 = InStr(lnPosPar3 + 1, psFuncionSuma, ",")
                    
                    lsEmp = "cPersCod = '" & psPersCod & "' and"
                    
                    If lnPosPar4 <> 0 Then
                        lsParm5 = Mid(psFuncionSuma, lnPosPar3 + 1, lnPos - lnPosPar3 - 1)
                    End If
                    
                    If lsParm5 <> "" Then
                        lnPosPar5 = InStr(1, lsParm5, ",")
                        If lnPosPar5 = 0 Then
                            If UCase(lsParm5) = "FALSE" Then lsEmp = ""
                        Else
                            lsParm5 = Left(lsParm5, lnPosPar5 - 1)
                            If UCase(lsParm5) = "FALSE" Then lsEmp = ""
                            
                            lsNomCam = Mid(psFuncionSuma, lnPosPar3 + lnPosPar5 + 1, lnPos - lnPosPar3 - lnPosPar5 - 1)
                        End If
                    End If
                    
                    sqlF = "Select Resultado = Sum(" & lsNomCam & ") From " & lsNomTab & " Where " & lsEmp & " " & lsParm3 & lsParm4
                End If
                
            End If
        End If
            
        'If oCon.AbreConexion Then
           Set rsF = oCon.CargaRecordSet(sqlF)
           'oCon.CierraConexion
        'End If
        
        If Not IsNull(rsF!RESULTADO) Then
            FuncionInternaSuma = rsF!RESULTADO
        Else
            FuncionInternaSuma = "0"
        End If
        
        rsF.Close
        Set rsF = Nothing
        'Set oCon = Nothing
    Else
        FuncionInternaSuma = "X"
    End If
    Exit Function
ERROR:
    FuncionInternaSuma = "X"
    
    If rsF.State = 1 Then
        rsF.Close
        Set rsF = Nothing
    End If
End Function


'FuncionInternaCuenta
'Es muy parecida a la funcion interna suma la unica direncia es que cuenta la cantidad
'de registro en lugar de sumar el contenido de ellos
Public Function FuncionInternaCuenta(psFuncionSuma As String, psPersCod As String, pdFecIni As Date, pdFecFin As Date) As String
    Dim lnPos As Integer
    Dim lnPosCor As Integer
    Dim lnPosLim As Integer
    Dim lnPosPar1 As Integer
    Dim lnPosPar2 As Integer
    Dim lnPosPar3 As Integer
    Dim lnPosPar4 As Integer
    Dim lnPosPar5 As Integer
    
    Dim lsNomTab As String
    Dim lsNomCam As String
    Dim lsParm1 As String
    Dim lsParm2 As String
    Dim lsParm3 As String
    Dim lsParm4 As String
    Dim lsParm5 As String
    
    Dim lnPosOpe As Integer
    
    Dim sqlF As String
    Dim rsF As ADODB.Recordset
    Set rsF = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    On Error GoTo ERROR
    
    lnPos = InStr(1, psFuncionSuma, ")", vbTextCompare)
    
    If lnPos <> 0 Then
        lnPosCor = InStr(1, psFuncionSuma, ".", vbTextCompare)
        lsNomTab = Mid(psFuncionSuma, 1, lnPosCor - 1)
        lsNomTab = GetNomTablaCampo(lsNomTab)
        If lsNomTab = "X" Then GoTo ERROR
        
        lnPosPar1 = InStr(1, psFuncionSuma, ",", vbTextCompare)
        If lnPosPar1 = 0 Then
            lsNomCam = Mid(psFuncionSuma, lnPosCor + 1)
            lsNomCam = GetNomTablaCampo(Left(lsNomCam, Len(lsNomCam) - 1))
            
            sqlF = "Select Resultado = Count(" & lsNomCam & ") From " & lsNomTab & " Where cEmpCod = '" & psPersCod & "'"
        Else
            lsNomCam = Mid(psFuncionSuma, lnPosCor + 1, lnPosPar1 - lnPosCor - 1)
            lsNomCam = GetNomTablaCampo(lsNomCam)
            
            lnPosPar2 = InStr(lnPosPar1 + 1, psFuncionSuma, ",", vbTextCompare)
            
            If lnPosPar2 = 0 Then
                
                lsParm1 = Mid(psFuncionSuma, lnPosPar1 + 1, lnPos - lnPosPar1 - 1)
                If InStr(1, lsParm1, ".") <> 0 Then
                    lsParm1 = GetNomTablaCampo(Mid(lsParm1, InStr(1, lsParm1, ".") + 1))
                Else
                    lsParm1 = GetNomTablaCampo(lsParm1)
                End If
                
                sqlF = "Select Resultado = Count(" & lsNomCam & ") From " & lsNomTab & " Where cEmpCod = '" & psPersCod & "' and " & lsParm1 & " between '" & Format(pdFecIni, gsFormatoFecha) & "' and '" & Format(pdFecFin, gsFormatoFecha) & "'"
            Else
                lsParm2 = Mid(psFuncionSuma, lnPosPar1 + 1, lnPosPar2 - lnPosPar1 - 1)
                If lsParm2 <> "" Then
                    lsParm2 = GetNomTablaCampo(lsParm2)
                    
                    lnPosPar3 = InStr(lnPosPar2 + 1, psFuncionSuma, ",", vbTextCompare)
                    
                    lsParm3 = Mid(psFuncionSuma, lnPosPar2 + 1, lnPosPar3 - lnPosPar2 - 1)
                    
                    lnPosPar3 = InStr(lnPosPar2 + 1, psFuncionSuma, ",")
                    
                    lnPosOpe = FunGetPosOpe(lsParm3, True)
                    lnPosOpe = PosOpe(lnPosOpe)
                    
                    lsParm4 = Replace(Mid(lsParm3, lnPosOpe), Chr(34), Chr(39))
                    
                    lsParm3 = Left(lsParm3, lnPosOpe - 1)
                    
                    lsParm3 = GetNomTablaCampo(lsParm3)
                    
                    lnPosPar4 = InStr(lnPosPar2 + 1 + lnPosOpe, psFuncionSuma, ",")
                    
                    If lnPosPar4 = 0 Then
                        sqlF = "Select Resultado = Count(" & lsNomCam & ") From " & lsNomTab & " Where cEmpCod = '" & psPersCod & "' and " & lsParm2 & " Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(pdFecFin, gsFormatoFecha) & "' And " & lsParm3 & lsParm4
                    Else
                        lsParm5 = Mid(psFuncionSuma, lnPosPar3 + 1)
                        
                        lnPosPar5 = InStr(1, lsParm5, ",")
                        
                        If lnPosPar5 <> 0 Then
                            lsParm5 = Mid(psFuncionSuma, lnPosPar3 + 1, lnPosPar5 - 1)
                            lsNomCam = Mid(psFuncionSuma, lnPosPar5 + lnPosPar3 + 1, Len(psFuncionSuma) - lnPosPar3 - lnPosPar5 - 1)
                        End If
                        
                        If UCase(lsParm5) = "TRUE" Then
                            sqlF = "Select Resultado = Count(" & lsNomCam & ") From " & lsNomTab & " Where cEmpCod = '" & psPersCod & "' and " & lsParm2 & " Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(pdFecFin, gsFormatoFecha) & "' And " & lsParm3 & lsParm4
                        Else
                            sqlF = "Select Resultado = Count(" & lsNomCam & ") From " & lsNomTab & " Where " & lsParm2 & " Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(pdFecFin, gsFormatoFecha) & "' And " & lsParm3 & lsParm4
                        End If
                    End If
                    'Para agregar las validaciones Adicionales
                Else
                
                    lnPosPar3 = InStr(lnPosPar2 + 1, psFuncionSuma, ",", vbTextCompare)
                    
                    lsParm3 = Mid(psFuncionSuma, lnPosPar2 + 1, lnPosPar3 - lnPosPar2 - 1)
                    
                    lnPosOpe = FunGetPosOpe(lsParm3, True)
                    lnPosOpe = PosOpe(lnPosOpe)
                    
                    lsParm4 = Replace(Mid(lsParm3, lnPosOpe), Chr(34), Chr(39))
                    
                    lsParm3 = Left(lsParm3, lnPosOpe - 1)
                    
                    lsParm3 = GetNomTablaCampo(lsParm3)
                    
                    lnPosPar4 = InStr(lnPosPar2 + 1 + lnPosOpe, psFuncionSuma, ",")
                    
                    If lnPosPar4 = 0 Then
                        If Right(lsParm4, 1) = ")" Then lsParm4 = Mid(lsParm4, 1, Len(lsParm4) - 1)
                        sqlF = "Select Resultado = Count(" & lsNomCam & ") From " & lsNomTab & " Where cEmpCod = '" & psPersCod & "' and " & lsParm3 & lsParm4
                    Else
                        lsParm5 = Mid(psFuncionSuma, lnPosPar3 + 1)
                        
                        lnPosPar5 = InStr(1, lsParm5, ",")
                        
                        If lnPosPar5 <> 0 Then
                            lsParm5 = Mid(psFuncionSuma, lnPosPar3 + 1, lnPosPar5 - 1)
                            lsNomCam = Mid(psFuncionSuma, lnPosPar5 + lnPosPar3 + 1, Len(psFuncionSuma) - lnPosPar3 - lnPosPar5 - 1)
                        End If
                        
                        If UCase(lsParm5) = "TRUE" Then
                            sqlF = "Select Resultado = Count(" & lsNomCam & ") From " & lsNomTab & " Where cEmpCod = '" & psPersCod & "' And " & lsParm3 & lsParm4
                        Else
                            sqlF = "Select Resultado = Count(" & lsNomCam & ") From " & lsNomTab & " Where " & lsParm3 & lsParm4
                        End If
                    End If
                End If
            End If
        End If
        
        'If oCon.AbreConexion Then
            Set rsF = oCon.CargaRecordSet(sqlF)
            'oCon.CierraConexion
        'End If
        
        If Not IsNull(rsF!RESULTADO) Then
            FuncionInternaCuenta = rsF!RESULTADO
        Else
            FuncionInternaCuenta = "0"
        End If
        
        rsF.Close
        Set rsF = Nothing
    Else
        FuncionInternaCuenta = "X"
    End If
    
    Exit Function
ERROR:
    FuncionInternaCuenta = "X"
    'MsgBox Err.Description, vbInformation, "Aviso"
    
    If rsF.State = 1 Then
        rsF.Close
        Set rsF = Nothing
    End If
End Function

' Se tenda en cuenta que:
' Solo habra un operador logico en la expresión
' Se permiten en las sos expresiones oepraciones artimeticas que se evaluaran
' con el FunEvaluea
Public Function FuncionInternaSI(psFuncionSi As String, psPersCod As String, pdFecIni As Date, pdFecFin As Date, ByVal psPeriodo As String) As String
    Dim lnPos As Integer
    Dim lnPosOpe As Integer
    Dim lnLenOpe As Integer
    Dim lsExpLog As String
    Dim lsOpe As String
    
    Dim lnExp1 As Integer
    Dim lnExp2 As Integer
    Dim lsExp1 As String
    Dim lsExp2 As String
    
    Dim lnResV As Integer
    Dim lnResF As Integer
    Dim lsResV As String
    Dim lsResF As String
    
    On Error GoTo ERROR
    
    'If Mid(psFuncionSi, 1, 3) = "SI(" Then
        lnPos = InStr(1, psFuncionSi, ")", vbTextCompare)
        If lnPos <> 0 Then
            lsExpLog = Mid(psFuncionSi, 1, Len(psFuncionSi))
            lsOpe = GetOpeLogico(lsExpLog, lnPosOpe, lnLenOpe)
            
            lnExp1 = InStr(1, lsExpLog, ",", vbTextCompare)
            lnExp2 = InStr(lnExp1 + 1, lsExpLog, ",", vbTextCompare)
            
            If lnPosOpe = 0 Then
                If lnExp1 = 0 Then GoTo ERROR
                
                If lnExp2 = 0 Then
                    lsResV = Mid(lsExpLog, lnExp1 + 1, lnPos - lnExp1 - 1)
                Else
                    lsResV = Mid(lsExpLog, lnExp1 + 1, lnExp2 - lnExp1 - 1)
                End If
                
                lsResV = ExprANum(FunEvalua(lsResV, psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodo))
                
                FuncionInternaSI = lsResV
                
            Else
                lsExp1 = Mid(lsExpLog, 1, lnPosOpe - 1)
                lsExp2 = Mid(lsExpLog, lnPosOpe + lnLenOpe, lnExp1 - lnLenOpe - lnPosOpe)
                
                lsExp1 = ExprANum(FunEvalua(lsExp1, psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodo))
                lsExp2 = ExprANum(FunEvalua(lsExp2, psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodo))
                
                lsResV = Mid(lsExpLog, lnExp1 + 1, lnExp2 - lnExp1 - 1)
                lsResF = Mid(lsExpLog, lnExp2 + 1, lnPos - lnExp2 - 1)
                
                lsResV = ExprANum(FunEvalua(lsResV, psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodo))
                lsResF = ExprANum(FunEvalua(lsResF, psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodo))
                
                If EvaluaOpeLogica(lsOpe, lsExp1, lsExp2) Then
                    FuncionInternaSI = lsResV
                Else
                    FuncionInternaSI = lsResF
                End If
                
            End If
        End If
        
    'End If
    
    Exit Function
ERROR:
    FuncionInternaSI = "X"
    
End Function

Public Function FuncionInternaSumaGrupo(psFuncionSumaGrupo As String, psPersCod As String, pdFecIni As Date, pdFecFin As Date) As String
    Dim lnPos As Integer
    Dim lnContenedor  As Currency
    Dim lsFiltro As String
    Dim I As Integer
    
    On Error GoTo ERROR
    
    lnPos = InStr(1, psFuncionSumaGrupo, ")", vbTextCompare)
    lnContenedor = 0
    VSQL = ""
    If lnPos <> 0 Then
        lsFiltro = Left(psFuncionSumaGrupo, lnPos - 1) & "_"
        For I = 1 To lnTValorTop
            If Left(TValor(I).lsConCod, 2) = lsFiltro And TValor(I).lbValorN Then
                lnContenedor = lnContenedor + Round(CCur(TValor(I).lsValor), 2)
                'VSQL = VSQL & TValor(i).lsConCod & "==" & TValor(i).lsValor & oImpresora.gPrnSaltoLinea
            End If
        Next I
        
        FuncionInternaSumaGrupo = lnContenedor
    Else
        FuncionInternaSumaGrupo = "X"
    End If
    
    Exit Function
ERROR:
    FuncionInternaSumaGrupo = "X"

End Function

Public Function FuncionInternaDIAMESAÑO(psFuncionDIA As String, psPersCod As String, pdFecIni As Date, pdFecFin As Date, Optional psDiaMesAño As Integer = "1") As String
    Dim lsNomTab As String
    Dim lsNomCam As String
    Dim lsParametro As String
    Dim ldFecha As Date
    Dim lnPos As Integer
    Dim lnPosCor As Integer
    Dim sqlF As String
    Dim rsF As ADODB.Recordset
    Set rsF = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    On Error GoTo ERROR
    
    lnPos = InStr(1, psFuncionDIA, ")", vbTextCompare)
    lsParametro = Mid(psFuncionDIA, 1, lnPos - 1)
    
    If lnPos <> 0 Then
        If EsOperador("05", lsParametro) Then
            ldFecha = GetValorParametroFecha(lsParametro, pdFecIni, pdFecFin)
        Else
            lnPosCor = InStr(1, psFuncionDIA, ".", vbTextCompare)
            If lnPosCor <> 0 Then
                lsNomTab = Mid(psFuncionDIA, 1, lnPosCor - 1)
                lsNomTab = GetNomTablaCampo(lsNomTab)
            
                lsNomCam = Mid(psFuncionDIA, lnPosCor + 1)
                lsNomCam = GetNomTablaCampo(Left(lsNomCam, Len(lsNomCam) - 1))
                
                sqlF = "Select Resultado = " & lsNomCam & " From " & lsNomTab & " Where cPersCod = '" & psPersCod & "'"
                
                'If oCon.AbreConexion Then
                    Set rsF = oCon.CargaRecordSet(sqlF)
                    'oCon.CierraConexion
                'End If
                ldFecha = Format(rsF!RESULTADO, gsFormatoFechaView)
                
                rsF.Close
                Set rsF = Nothing
            Else
                If IsDate(lsParametro) Then
                    ldFecha = CDate(lsParametro)
                Else
                    FuncionInternaDIAMESAÑO = "X"
                    Exit Function
                End If
            End If
        End If
        Select Case psDiaMesAño
            Case "1"
                FuncionInternaDIAMESAÑO = Day(ldFecha)
            Case "2"
                FuncionInternaDIAMESAÑO = Month(ldFecha)
            Case "3"
                FuncionInternaDIAMESAÑO = Year(ldFecha)
        End Select
    Else
        FuncionInternaDIAMESAÑO = "X"
    End If
    Exit Function
ERROR:
    FuncionInternaDIAMESAÑO = "X"
    
    If rsF.State = 1 Then
        rsF.Close
        Set rsF = Nothing
    End If
End Function

Public Function FuncionInternaObtConcep11(psFuncionObtieneConcepto As String, psPersCod As String, pdFecIni As Date, pdFecFin As Date) As String
    Dim lnPos As Integer
    Dim lnPosCor As Integer
    Dim lnPosLim As Integer
    Dim lsPosLim As String
    Dim lnPosPar As Integer
    Dim lsNomCam As String
    Dim lsParm As String
    Dim I As Integer
    Dim sqlF As String
    Dim rsF As ADODB.Recordset
    Set rsF = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    On Error GoTo ERROR
        lnPos = InStr(1, psFuncionObtieneConcepto, ")", vbTextCompare)
        
        If lnPos <> 0 Then
            lnPosCor = InStr(1, psFuncionObtieneConcepto, ",", vbTextCompare)
            lsNomCam = Mid(psFuncionObtieneConcepto, 1, lnPosCor - 1)
            
            lnPosPar = InStr(lnPosCor + 1, psFuncionObtieneConcepto, ",", vbTextCompare)
            lsParm = Mid(psFuncionObtieneConcepto, lnPosCor + 1, lnPosPar - lnPosCor - 1)
            
            lnPosLim = InStr(lnPosPar + 1, psFuncionObtieneConcepto, ")", vbTextCompare)
            lsPosLim = Mid(psFuncionObtieneConcepto, lnPosPar + 1, lnPosLim - lnPosPar - 1)
            
            sqlF = " Select RESULTADO = nMonto from RHPlanillaDetCon RHP" _
                 & " Inner Join RHConceptoTabla RHCT On RHCT.cRHCOnceptoCod = RHP.cRHCOnceptoCod  And cRHconceptoMeno = '" & lsNomCam & "'" _
                 & " where cPlanillaCod = '" & lsParm & "' and cPersCod = '" & psPersCod & "' and nMonto <> 0" _
                 & " order by cRRHHPeriodo Desc"
                
            Set rsF = oCon.CargaRecordSet(sqlF)
            
            lnPosLim = CCur(lsPosLim)
            
            If Not RSVacio(rsF) Then
                I = 0
                While Not rsF.EOF
                    I = I + 1
                    If I = lnPosLim Then
                        FuncionInternaObtConcep11 = Trim(Str(rsF!RESULTADO))
                        rsF.Close
                        Set rsF = Nothing
                        Exit Function
                    End If
                    rsF.MoveNext
                Wend
            End If
            
            FuncionInternaObtConcep11 = "0"
            
            rsF.Close
            Set rsF = Nothing
        Else
            FuncionInternaObtConcep11 = "X"
        End If
    
    Exit Function
ERROR:
    FuncionInternaObtConcep11 = "X"
    
    If rsF.State = 1 Then
        rsF.Close
        Set rsF = Nothing
    End If
End Function

'***********************************
Public Function FuncionInternaObtConcep(psFuncionObtieneConcepto As String, psCodPers As String, pdFecIni As Date, pdFecFin As Date) As String
    Dim lnPos As Integer
    Dim lnPosCor As Integer
    Dim lnPosLim As Integer
    Dim lsPosLim As String
    Dim lnPosPar As Integer
    Dim lsNomCam As String
    Dim lsParm As String
    Dim I As Integer
    Dim sqlF As String
    Dim rsF As New ADODB.Recordset
    
    On Error GoTo ERROR
        lnPos = InStr(1, psFuncionObtieneConcepto, ")", vbTextCompare)
        
        If lnPos <> 0 Then
            
            lnPosCor = InStr(1, psFuncionObtieneConcepto, ",", vbTextCompare)
            
            If lnPosCor = 0 Then
                lsNomCam = Mid(psFuncionObtieneConcepto, 1, lnPos - 1)
                
                sqlF = " Select RESULTADO = nRHConceptoValor from RHConcepto RHP" _
                     & " Inner Join RHConceptoTabla RHCT On RHCT.cRHConceptoCod = RHP.cRHConceptoCod  And cRHconceptoMeno = '" & lsNomCam & "'" _
                     & " where RHP.cRHPlanillaCod = 'E01' and cPersCod = '" & psCodPers & "' and nRHConceptoValor <> 0 " _
                     & " "
                Set rsF = oCon.CargaRecordSet(sqlF)

                If Not RSVacio(rsF) Then
                    FuncionInternaObtConcep = Format(rsF!RESULTADO, "#0.00")
                    rsF.Close
                    Set rsF = Nothing
                    Exit Function
                Else
                    FuncionInternaObtConcep = "0"
                    rsF.Close
                    Set rsF = Nothing
                    Exit Function
                End If
            
            Else
                lsNomCam = Mid(psFuncionObtieneConcepto, 1, lnPosCor - 1)
                lnPosPar = InStr(lnPosCor + 1, psFuncionObtieneConcepto, ",", vbTextCompare)
                
                If lnPosPar = 0 Then
                    
                    lsParm = Mid(psFuncionObtieneConcepto, lnPosCor + 1, lnPos - lnPosCor - 1)
                    lnPosLim = InStr(lnPosPar + 1, psFuncionObtieneConcepto, ")", vbTextCompare)
                    lsPosLim = Mid(psFuncionObtieneConcepto, lnPosPar + 1, lnPosLim - lnPosPar - 1)
                    
                    sqlF = " Select RESULTADO = nMonto from RHPlanillaDetCon RHP" _
                         & " Inner Join RHConceptoTabla RHCT On RHCT.cRHCOnceptoCod = RHP.cRHCOnceptoCod  And cRHconceptoMeno = '" & lsNomCam & "'" _
                         & " where cPlanillaCod = '" & lsParm & "' and cPersCod = '" & psCodPers & "' and nMonto <> 0 And cRRHHPeriodo Like '" & Format(pdFecFin, "yyyymm") & "%'" _
                         & " order by cRRHHPeriodo Desc"
                    Set rsF = oCon.CargaRecordSet(sqlF)
                
                    If Not RSVacio(rsF) Then
                        FuncionInternaObtConcep = Format(rsF!RESULTADO, "#0.00")
                        rsF.Close
                        Set rsF = Nothing
                        Exit Function
                    Else
                        FuncionInternaObtConcep = "0"
                        rsF.Close
                        Set rsF = Nothing
                        Exit Function
                    End If
                Else
                    lsParm = Mid(psFuncionObtieneConcepto, lnPosCor + 1, lnPosPar - lnPosCor - 1)
                
                    lnPosLim = InStr(lnPosPar + 1, psFuncionObtieneConcepto, ")", vbTextCompare)
                    lsPosLim = Mid(psFuncionObtieneConcepto, lnPosPar + 1, lnPosLim - lnPosPar - 1)
                    
                    sqlF = " Select RESULTADO = nMonto from RHPlanillaDetCon RHP" _
                         & " Inner Join RHConceptoTabla RHCT On RHCT.cRHCOnceptoCod = RHP.cRHCOnceptoCod  And cRHconceptoMeno = '" & lsNomCam & "'" _
                         & " where cPlanillaCod = '" & lsParm & "' and cPersCod = '" & psCodPers & "' and nMonto <> 0 And cRRHHPeriodo < '" & Format(pdFecFin, "yyyymm") & "'" _
                         & " order by cRRHHPeriodo Desc"
                    Set rsF = oCon.CargaRecordSet(sqlF)
                    
                    lnPosLim = CCur(lsPosLim)
                    
                    If Not RSVacio(rsF) Then
                        I = 0
                        While Not rsF.EOF
                            I = I + 1
                            If I = lnPosLim Then
                                FuncionInternaObtConcep = Format(rsF!RESULTADO, "#0.00")
                                rsF.Close
                                Set rsF = Nothing
                                Exit Function
                            End If
                            rsF.MoveNext
                        Wend
                    End If
                    
                    FuncionInternaObtConcep = "0"
                    
                    rsF.Close
                    Set rsF = Nothing
                End If
            End If
        Else
            FuncionInternaObtConcep = "X"
        End If
    
    Exit Function
ERROR:
    FuncionInternaObtConcep = "X"
    
    If rsF.State = 1 Then
        rsF.Close
        Set rsF = Nothing
    End If
End Function
'***********************************



'***********************************
Public Function FuncionInternaValorOriginal(psFuncionObtieneConcepto As String, psCodPers As String, pdFecIni As Date, pdFecFin As Date) As String
    Dim lnPos As Integer
    Dim lnPosCor As Integer
    Dim lnPosLim As Integer
    Dim lsPosLim As String
    Dim lnPosPar As Integer
    Dim lsNomCam As String
    Dim lsParm As String
    Dim I As Integer
    Dim sqlF As String
    Dim rsF As New ADODB.Recordset
    
    On Error GoTo ERROR
        lnPos = InStr(1, psFuncionObtieneConcepto, ")", vbTextCompare)
        
        If lnPos <> 0 Then
            lnPosCor = InStr(1, psFuncionObtieneConcepto, ",", vbTextCompare)
            
            If lnPosCor = 0 Then
                lsNomCam = Mid(psFuncionObtieneConcepto, 1, lnPos - 1)
                
                sqlF = " Select RESULTADO = nRHConceptoValor from RHConcepto RHP" _
                     & " Inner Join RHConceptoTabla RHCT On RHCT.cRHCOnceptoCod = RHP.cRHCOnceptoCod And cRHconceptoMeno = '" & lsNomCam & "'" _
                     & " where cRHPlanillaCod = 'E01' and cPersCod = '" & psCodPers & "' and nRHConceptoValor <> 0 "
                Set rsF = oCon.CargaRecordSet(sqlF)

                If Not RSVacio(rsF) Then
                    FuncionInternaValorOriginal = Format(rsF!RESULTADO, "#0.00")
                    rsF.Close
                    Set rsF = Nothing
                    Exit Function
                Else
                    FuncionInternaValorOriginal = "0"
                    rsF.Close
                    Set rsF = Nothing
                    Exit Function
                End If
            
            Else
                lsNomCam = Mid(psFuncionObtieneConcepto, 1, lnPosCor - 1)
                lnPosPar = InStr(lnPosCor + 1, psFuncionObtieneConcepto, ",", vbTextCompare)
                
                If lnPosPar = 0 Then
                    
                    lsParm = Mid(psFuncionObtieneConcepto, lnPosCor + 1, lnPos - lnPosCor - 1)
                    lnPosLim = InStr(lnPosPar + 1, psFuncionObtieneConcepto, ")", vbTextCompare)
                    lsPosLim = Mid(psFuncionObtieneConcepto, lnPosPar + 1, lnPosLim - lnPosPar - 1)
                    
                    sqlF = " Select RESULTADO = nMonto from RHPlanillaDetCon RHP" _
                         & " Inner Join RHConceptoTabla RHCT On RHCT.cRHCOnceptoCod = RHP.cRHCOnceptoCod  And cRHconceptoMeno = '" & lsNomCam & "'" _
                         & " where cPlanillaCod = '" & lsParm & "' and cPersCod = '" & psCodPers & "' and nMonto <> 0 And cRRHHPeriodo Like '" & Format(pdFecFin, "yyyymm") & "%'" _
                         & " order by cRRHHPeriodo Desc"
                    Set rsF = oCon.CargaRecordSet(sqlF)
                
                    If Not RSVacio(rsF) Then
                        FuncionInternaValorOriginal = Format(rsF!RESULTADO, "#0.00")
                        rsF.Close
                        Set rsF = Nothing
                        Exit Function
                    Else
                        FuncionInternaValorOriginal = "0"
                        rsF.Close
                        Set rsF = Nothing
                        Exit Function
                    End If
                Else
                    lsParm = Mid(psFuncionObtieneConcepto, lnPosCor + 1, lnPosPar - lnPosCor - 1)
                
                    lnPosLim = InStr(lnPosPar + 1, psFuncionObtieneConcepto, ")", vbTextCompare)
                    lsPosLim = Mid(psFuncionObtieneConcepto, lnPosPar + 1, lnPosLim - lnPosPar - 1)
                    
                    sqlF = " Select RESULTADO = nMonto from RHPlanillaDetCon RHP" _
                         & " Inner Join RHConceptoTabla RHCT On RHCT.cRHCOnceptoCod = RHP.cRHCOnceptoCod  And cRHconceptoMeno = '" & lsNomCam & "'" _
                         & " where cPlanillaCod = '" & lsParm & "' and cPersCod = '" & psCodPers & "' and nMonto <> 0 And cRRHHPeriodo < '" & Format(pdFecFin, "yyyymm") & "'" _
                         & " order by cRRHHPeriodo Desc"
                    Set rsF = oCon.CargaRecordSet(sqlF)
                    
                    lnPosLim = CCur(lsPosLim)
                    
                    If Not RSVacio(rsF) Then
                        I = 0
                        While Not rsF.EOF
                            I = I + 1
                            If I = lnPosLim Then
                                FuncionInternaValorOriginal = Format(rsF!RESULTADO, "#0.00")
                                rsF.Close
                                Set rsF = Nothing
                                Exit Function
                            End If
                            rsF.MoveNext
                        Wend
                    End If
                    
                    FuncionInternaValorOriginal = "0"
                    
                    rsF.Close
                    Set rsF = Nothing
                End If
            End If
        Else
            FuncionInternaValorOriginal = "X"
        End If
    
    Exit Function
ERROR:
    FuncionInternaValorOriginal = "X"
    
    If rsF.State = 1 Then
        rsF.Close
        Set rsF = Nothing
    End If
End Function
'***********************************



Public Function EsConceptoValido(psConcepto As String) As Boolean
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    On Error GoTo ERROR
    
    'sqlC = "Select cRHConceptoCod From RHConceptoTabla Where cRHconceptoMeno = '" & Replace(psConcepto, Chr(39), Chr(34)) & "' and cRHConceptoEstado = '1'"
    sqlC = "Select  dbo.RHEsConceptoValido('" & Replace(psConcepto, Chr(39), Chr(34)) & "') cRHConceptoCod"
    
    'If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If IsNull(rsC!cRHConceptoCod) Then
            EsConceptoValido = False
        Else
            EsConceptoValido = True
        End If
        'oCon.CierraConexion
    'End If
    
    rsC.Close
    'Set oCon = Nothing
    Set rsC = Nothing
    Exit Function
ERROR:
    EsConceptoValido = False
End Function

Public Function FunEsConcepto(psFun As String) As String
    Dim lnPos As Integer
    Dim sqlG As String
    Dim rsG As ADODB.Recordset
    Set rsG = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    sqlG = "Select cRHConceptoTpo from rhconceptotabla Where cRHconceptoMeno = '" & psFun & "'"
    
    'If oCon.AbreConexion Then
        Set rsG = oCon.CargaRecordSet(sqlG)
        If RSVacio(rsG) Then
            FunEsConcepto = "X"
        Else
            FunEsConcepto = Trim(rsG!cRHConceptoTpo)
        End If
        rsG.Close
        Set rsG = Nothing
    'End If
    
    Set rsG = Nothing
End Function

Public Function FunGetConstante(psFun As String) As String
    Dim lnPos As Integer
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    Dim lsDato As String
    
    sqlC = "Select dbo.RHFunGetConstante('" & psFun & "') cFormula"
    
    'If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If IsNull(rsC!cFormula) Then
            FunGetConstante = "X"
        Else
            lsDato = rsC!cFormula
            If IsNumeric(Trim(lsDato)) Then
                FunGetConstante = Trim(lsDato)
            Else
                FunGetConstante = "X"
            End If
        End If
    'End If
    rsC.Close
    Set rsC = Nothing
    'Set oCon = Nothing
End Function

Public Function FunGetValorPreDef(psFun As String, psPersCod As String, psFecIni As Date, psFecFin As Date, psPlanillas As String) As String
    Dim sqlS As String
    Dim rsS As ADODB.Recordset
    Set rsS = New ADODB.Recordset
    Dim lsCad As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    On Error GoTo ERROR
    
    'sqlS = "Select nRHConceptoValor nMonto From RHConcepto RHC Inner Join RHConceptoTabla RHCT On RHC.cRHConceptoCod = RHCT.cRHConceptoCod Where cRHPlanillaCod = '" & psPlanillas & "' And cPersCod  = '" & psPersCod & "' And cRHconceptoMeno = '" & psFun & "'"
    If psFun = "I_MOVILI" Then
        sqlS = " Select dbo.RHFunGetValorPreDef('E01','" & psPersCod & "', '" & psFun & "') nMonto"
    'ALPA 20091106*************************************
    ElseIf psFun = "I_GRAT PRO_HOR EXT" Then
        sqlS = " Select dbo.RHFunGetValorPreDef('E02','" & psPersCod & "', '" & psFun & "') nMonto"
    '**************************************************
    Else
        sqlS = " Select dbo.RHFunGetValorPreDef('" & psPlanillas & "','" & psPersCod & "', '" & psFun & "') nMonto"
    End If
    
    'If oCon.AbreConexion Then
        Set rsS = oCon.CargaRecordSet(sqlS)
        'If rsS.EOF And rsS.BOF Then
        If IsNull(rsS!nMonto) And psFun <> "I_MOVILI" Then 'Add By Gitu
            sqlS = " Select dbo.RHFunGetValorPreDef('" & Left(psPlanillas, 1) & Right(gsRHPlanillaSueldos, 2) & "','" & psPersCod & "', '" & psFun & "') nMonto"
            'sqlS = "Select nRHConceptoValor nMonto From RHConcepto RHC Inner Join RHConceptoTabla RHCT On RHC.cRHConceptoCod = RHCT.cRHConceptoCod Where cPersCod  = '" & psPersCod & "' And cRHPlanillaCod = '" & gsRHPlanillaSueldos & "' And cRHconceptoMeno = '" & psFun & "'"
            rsS.Close
            Set rsS = oCon.CargaRecordSet(sqlS)
            If IsNull(rsS!nMonto) Then
                FunGetValorPreDef = "0"
            Else
                If IsNumeric(rsS!nMonto) Then
                    lsCad = Trim(rsS!nMonto)
                    lsCad = lsCad & GetFormulaConcepto(psFun)
                    FunGetValorPreDef = lsCad
                ElseIf Trim(rsS!nMonto) = "" Or IsNull(rsS!nMonto) Then
                    FunGetValorPreDef = "0"
                Else
                    FunGetValorPreDef = "X"
                End If
            End If
        Else
            If IsNumeric(rsS!nMonto) Then
                lsCad = Trim(rsS!nMonto)
                lsCad = lsCad & GetFormulaConcepto(psFun)
                FunGetValorPreDef = lsCad
            ElseIf IsNull(rsS!nMonto) Then
                FunGetValorPreDef = "0"
            Else
                FunGetValorPreDef = "X"
            End If
        End If
        'oCon.CierraConexion
    'End If
    rsS.Close
    Set rsS = Nothing
    Exit Function
ERROR:
    FunGetValorPreDef = "X"
End Function

'SUELDO BASICO                  1
'INGRESO FIJO APF 10.23%        2
'INGRESO FIJO APF 3.00%         3
'REINTEGROS VARIOS              4
'PAGO EN EXCESO                51
'DESCUENTO JUDICIAL FIJO       52
'DESCUENTO JUDICIAL PORCENTUAL 53
Public Function FunGetValorConstante(psFun As String, psPersCod As String, pdFecIni As Date, pdFecFin As Date, Optional cPlanilla As String = "", Optional psPeriodoProc As String = "") As String
    Dim lsFun As String
    Dim lnPos As Integer
    Dim ldFecIni As Date
    Dim ldFecFin As Date
    
    Dim ldFecIniGra As Date
    Dim ldFecFinGra As Date
    
    Dim ldFecRIni As Date
    Dim ldFecRFin As Date
    
    Dim lsCadIngAfecto As String
    Dim lsCadTotIng As String
    
    Dim lsCadUIT7 As String
    Dim lsCadUIT27 As String
    Dim lsCadUIT54 As String
    Dim lsCadPorHasta27 As String
    Dim lsCadPorHasta54 As String
    Dim lsCadPorMas54 As String
    
    ObtFechasAsistencia ldFecIni, ldFecFin, ldFecIniGra, ldFecFinGra, pdFecIni
     
    ldFecRIni = CDate("01/" & FillNum(Trim(Str(Month(ldFecFin))), 2, "0") & "/" & Trim(Str(Year(ldFecFin))))
    If Month(ldFecFin) <> 2 Then
        ldFecRFin = CDate("30/" & FillNum(Trim(Str(Month(ldFecFin))), 2, "0") & "/" & Trim(Str(Year(ldFecFin))))
    Else
        ldFecRFin = CDate("28/" & FillNum(Trim(Str(Month(ldFecFin))), 2, "0") & "/" & Trim(Str(Year(ldFecFin))))
    End If
    lsFun = IIf(Mid(psFun, 2, 1) = "_", Mid(psFun, 3), psFun)
    
    Select Case (lsFun)
    Case "GET_SUELDO_BASICO"
        FunGetValorConstante = Str(GetConceptoEmp(psPersCod, gRHConcIngSueBasico))
    Case "GET_INGRESO_AFP_1023"
        FunGetValorConstante = Str(GetConceptoEmp(psPersCod, gRHConcIngAFP1023))
    Case "GET_INGRESO_AFP_300"
        FunGetValorConstante = Str(GetConceptoEmp(psPersCod, gRHConcIngAFP0300))
    Case "GET_REINTEGROS"
        FunGetValorConstante = Str(GetConceptoEmp(psPersCod, gRHConcIngReintegro))
    Case "GET_PAGO_EXCESO"
        FunGetValorConstante = Str(GetConceptoEmp(psPersCod, "51"))
    Case "GET_DECUENTO_JUDICIAL_FIJO"
        FunGetValorConstante = Str(GetConceptoEmp(psPersCod, gRHConcDctoJudFijo))
    Case "GET_DECUENTO_JUDICIAL_POR"
        FunGetValorConstante = Str(GetConceptoEmp(psPersCod, gRHConcDctoJudPorc))
    Case "GET_MINUTOS_TARDANZA"
        FunGetValorConstante = Str(GetMinTarEmp(psPersCod, ldFecIni, ldFecFin))
    Case "GET_MINUTOS_SALIDA"
        FunGetValorConstante = Str(GetMinPerPerEmp(psPersCod, ldFecIni, ldFecFin, ldFecIni, ldFecFin))
    Case "GET_SUELDO_BASICO_NIVEL"
        FunGetValorConstante = "0" 'Str(GetSueBasNivEmp(psPersCod))
    Case "GET_SUELDO_CONTRATO"
        FunGetValorConstante = "0" ' Str(GetSueConEmp(psPersCod))
    Case "GET_AFP_COM_VAR"
        FunGetValorConstante = Format(GetDesAPFEmp(psPersCod, "1"), "#0.0000")
    Case "GET_AFP_PRIMA"
        FunGetValorConstante = Format(GetDesAPFEmp(psPersCod, "2"), "#0.0000")
    Case "GET_NUM_MES_TRAB"
        FunGetValorConstante = Format(GetNumMesTrab(psPersCod, ldFecIniGra, ldFecFinGra, pdFecIni, pdFecFin, cPlanilla), "#0.0000")
        'FunGetValorConstante = Format(GetNumMesTrab(psPersCod, cdate("30/01/2001"), cdate("30/04/2001")), "#0.0000")
    Case "GET_NUM_MES_TRAB_TERCIO"
        FunGetValorConstante = Format(GetNumMesTrabCTS(psPersCod, pdFecIni, pdFecFin), "#0.0000")
    Case "GET_NUM_DIAS_NO_TRAB"
        FunGetValorConstante = Format(GetNumDiasPermiso(psPersCod, pdFecIni, pdFecFin, "'" & RHEstadosRRHH.gRHEstadosRRHHLICSINGOCE & "','" & RHEstadosRRHH.gRHEstadosRRHHSUSPENDIDO & "','" & RHEstadosRRHH.gRHEstadosRRHHSUBSIDIADO & "'", cPlanilla), "#0.0000")
    Case "GET_IMP_5TA_EMP"
'        lsCadAux1 = ExprANum(FunEvalua("I_REM_NO_AFEC", psPersCod, pdFecIni, pdFecFin, False, "VVVV"))
'        lsCadAux2 = ExprANum(FunEvalua("V_UIT_7", psPersCod, pdFecIni, pdFecFin, False, "VVVV"))
'        lsCadAux3 = FunEvalua("V_POR_IMP_5TA", psPersCod, pdFecIni, pdFecFin, False, "VVVV")
'        lsCadAux4 = FunEvalua(gsRHConceptoITOTING, psPersCod, pdFecIni, pdFecFin, False, "VVVV") & "-" & FunEvalua("I_REINTEGRO", psPersCod, pdFecIni, pdFecFin, False, "VVVV")
'
'        FunGetValorConstante = Format(GetImp5taEmp(psPersCod, CCur(lsCadAux1), pdFecIni, pdFecFin, lsCadAux2, lsCadAux3, lsCadAux4), "#0.0000")
        
        
        lsCadIngAfecto = ExprANum(FunEvalua("I_REM_NO_AFEC", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadUIT7 = ExprANum(FunEvalua("V_UIT_7", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadUIT27 = ExprANum(FunEvalua("V_UIT_27", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadUIT54 = ExprANum(FunEvalua("V_UIT_54", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadPorHasta27 = FunEvalua("V_POR_IMP_5TA", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc)
        lsCadPorHasta54 = FunEvalua("V_POR_5TA_H54", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc)
        lsCadPorMas54 = FunEvalua("V_POR_5TA_M54", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc)
        
        lsCadTotIng = ExprANum(FunEvalua("I_TOT_ING", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc) & "-" & FunEvalua("I_REINTEGRO", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        
        FunGetValorConstante = Format(GetImp5taEmpNew(psPersCod, CCur(lsCadTotIng), pdFecIni, pdFecFin, lsCadUIT7, lsCadUIT27, lsCadUIT54, lsCadPorHasta27, lsCadPorHasta54, lsCadPorMas54, lsCadTotIng, cPlanilla, psPeriodoProc), "#0.0000")
    
    Case "GET_REV_5TA"
        lsCadIngAfecto = ExprANum(FunEvalua("I_REM_NO_AFEC", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadUIT7 = ExprANum(FunEvalua("V_UIT_7", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadUIT27 = ExprANum(FunEvalua("V_UIT_27", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadUIT54 = ExprANum(FunEvalua("V_UIT_54", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadPorHasta27 = FunEvalua("V_POR_IMP_5TA", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc)
        lsCadPorHasta54 = FunEvalua("V_POR_5TA_H54", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc)
        lsCadPorMas54 = FunEvalua("V_POR_5TA_M54", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc)
        'AFE5TA
        'lsCadTotIng  = "0"
        lsCadTotIng = ExprANum(FunEvalua("V_AFE5TA", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        
        FunGetValorConstante = Format(GetImp5taEmpRev(psPersCod, CCur(lsCadTotIng), pdFecIni, pdFecFin, lsCadUIT7, lsCadUIT27, lsCadUIT54, lsCadPorHasta27, lsCadPorHasta54, lsCadPorMas54, lsCadTotIng, cPlanilla, psPeriodoProc, False), "#0.0000")
    
    Case "GET_REV_5TA_NEG"
        lsCadIngAfecto = ExprANum(FunEvalua("I_REM_NO_AFEC", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadUIT7 = ExprANum(FunEvalua("V_UIT_7", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadUIT27 = ExprANum(FunEvalua("V_UIT_27", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadUIT54 = ExprANum(FunEvalua("V_UIT_54", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        lsCadPorHasta27 = FunEvalua("V_POR_IMP_5TA", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc)
        lsCadPorHasta54 = FunEvalua("V_POR_5TA_H54", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc)
        lsCadPorMas54 = FunEvalua("V_POR_5TA_M54", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc)
        'AFE5TA
        'lsCadTotIng  = "0"
        lsCadTotIng = ExprANum(FunEvalua("V_AFE5TA", psPersCod, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
        
        FunGetValorConstante = Format(GetImp5taEmpRev(psPersCod, CCur(lsCadTotIng), pdFecIni, pdFecFin, lsCadUIT7, lsCadUIT27, lsCadUIT54, lsCadPorHasta27, lsCadPorHasta54, lsCadPorMas54, lsCadTotIng, cPlanilla, psPeriodoProc, True), "#0.0000")
        
    Case "GET_MES_TRAB_CTS"
        FunGetValorConstante = Format(GetMesTrabCTS(psPersCod, pdFecIni, pdFecFin), "#0.00")
    Case "GET_ULT_BONIF_VAC"
        FunGetValorConstante = Format(GetUltGratVaca(psPersCod, pdFecFin, psPeriodoProc), "#0.0000")
    Case "GET_ULT_GRAT"
        FunGetValorConstante = Format(GetUltGrat(psPersCod, pdFecFin), "#0.0000")
    'Case "GET_PROM_BON"
        'FunGetValorConstante = Format(GetPromBon(psPersCod, pdFecFin), "#0.0000")
    Case "GET_ULT_TERCIO"
        FunGetValorConstante = Format(GetUltTercio(psPersCod, pdFecIni, pdFecFin, False), "#0.0000")
    Case "GET_MINUTOS_SALIDA_UTIL"
        FunGetValorConstante = Str(GetDiasNoLabUtilEmp(psPersCod, pdFecIni, pdFecFin))
    Case "GET_MINUTOS_TOTAL_UTIL"
        FunGetValorConstante = Str(GetTotalDiasNoLabUtilEmp(psPersCod, pdFecIni, pdFecFin))
    Case "GET_MONTO_TOTAL_UTIL"
        FunGetValorConstante = Str(GetTotalMontoUtilEmp(psPersCod, pdFecIni, pdFecFin))
    Case "GET_MONTO_EMP_UTIL"
        FunGetValorConstante = Str(GetMontoUtilEmp(psPersCod, pdFecIni, pdFecFin))
    Case "GET_DIAS_VAC"
        FunGetValorConstante = Str(GetDiasVacaciones(psPersCod, pdFecIni, pdFecFin))
    Case "GET_SUEL_BAS_COM"
        FunGetValorConstante = Str(GetCopceptoCompleto(psPersCod, 119))
    Case "GET_BON_CARFAMCOM"
        FunGetValorConstante = Str(GetCopceptoCompleto(psPersCod, 101))
    Case "GET_RIECAJA_COM"
        FunGetValorConstante = Str(GetCopceptoCompleto(psPersCod, 164))
    Case "GET_MOVILI_COM"
        FunGetValorConstante = Str(GetCopceptoCompleto(psPersCod, 105))
    Case "GET_ULTIMA_GRATI" 'EJVG 20110813
        FunGetValorConstante = Format(GetUltimaGratificacion(psPersCod), "#0.0000")
    Case "GET_COD_ASIST_MED" 'EJVG 20110813
        FunGetValorConstante = GetCodAsistenciaMedica(psPersCod)
    Case Else
        FunGetValorConstante = "X"
    End Select
End Function

Public Function GetImp5taEmpNew(psCodEmp As String, pnMontSueBasAfecto As Currency, pdFecIni As Date, pdFecFin As Date, psValor7UIT As String, psValor27UIT As String, psValor54UIT As String, psPorDes5taHasta27 As String, psPorDes5taHasta54 As String, psPorDes5taMasDe54 As String, psIngActual As String, psPlanillaCod As String, psPeriodoProc As String) As Currency
    Dim lnNumMes As Integer
    Dim lnFacIng As Integer
    Dim lnFacRet As Integer
    Dim lnMontoImp As Currency
    Dim lnMontoRetAnu As Currency
    Dim lnMontoRetOtrosMes As Currency
    
    Dim lnMontoIngAnu As Currency
    Dim lnMontoRes As Double
    Dim lnMontoAddAguinaldo As Currency
    
    Dim lnMontoAfectacionPlaGratif As Currency
    Dim lnMontoAfectacionPlaSueldo As Currency
    lnNumMes = Month(CDate(pdFecFin))
    
    If lnNumMes < gnMes1 Then
        lnFacIng = gnNumRem - (lnNumMes - 1)
    ElseIf lnNumMes >= gnMes1 And lnNumMes < gnMes2 Then
        lnFacIng = gnNumRem - lnNumMes
    ElseIf lnNumMes = gnMes2 Then
        'ALPA 20101223***********
        'lnFacIng = 1
        If psPlanillaCod = "E02" Then
            lnFacIng = 2
        Else
            lnFacIng = 1
        End If
        '************************
    End If
    'ALPA 20101223****************************
    'lnFacRet = gnNumRem - 2 - (lnNumMes - 1)
    If lnNumMes >= 1 And lnNumMes <= 3 Then
        lnFacRet = 12
    ElseIf lnNumMes = 4 Then
        lnFacRet = 9
    ElseIf lnNumMes >= 5 And lnNumMes <= 7 Then
        lnFacRet = 8
    ElseIf lnNumMes = 8 Then
        lnFacRet = 5
    ElseIf lnNumMes >= 9 And lnNumMes <= 11 Then
        lnFacRet = 4
    ElseIf lnNumMes = 12 Then
        lnFacRet = 1
    End If
    '*****************************************
    
    lnMontoIngAnu = GetSumIngAnuEmp(psCodEmp, pdFecFin, Trim(psPeriodoProc) & Trim(psPlanillaCod))
    
    lnMontoAddAguinaldo = 0
    lnMontoAfectacionPlaGratif = 1
    lnMontoAfectacionPlaSueldo = 1
    'If psPlanillaCod = gsRHPlanillaGratificacion Then
    '    If lnNumMes = gnMes2 Then
    '        lnFacIng = lnFacIng + 1
    '        lnFacRet = lnFacRet + 1
    '        lnMontoAddAguinaldo = CCur(FunEvalua("G_AGUIPROYECT", psCodEmp, pdFecIni, pdFecFin, False, "VVVV", psPeriodoProc))
    '    Else
    '        'lnMontoAfectacionPlaGratif = 0.5
            'pedido cusco
    '        lnMontoAfectacionPlaGratif = 1
    '    End If
    'Else
    If psPlanillaCod = gsRHPlanillaSueldos Then
        If lnNumMes = gnMes1 Then
            'lnMontoAfectacionPlaSueldo = 0.5
            'pedido cusco
            lnMontoAfectacionPlaSueldo = 1
        End If
    Else
        lnMontoAddAguinaldo = 0
    End If
    
    lnMontoImp = ((pnMontSueBasAfecto * lnFacIng) + lnMontoIngAnu + lnMontoAddAguinaldo)
    lnMontoImp = lnMontoImp - CCur(psValor7UIT)
    If lnMontoImp <= 0 Then 'EXONERADO DE IMPUESTO
        lnMontoImp = 0
        
    ElseIf lnMontoImp <= CCur(psValor27UIT) Then 'IMPUESTO SOLO DEL 15%
        
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, pdFecFin, Trim(psPeriodoProc) & Trim(psPlanillaCod))
        If lnNumMes = gnMes2 Then
        If psPlanillaCod = "E02" Then
            lnMontoRes = (((lnMontoImp * CCur(psPorDes5taHasta27)) - lnMontoRetAnu) / lnFacRet) / 2
        Else
            lnMontoRes = ((lnMontoImp * CCur(psPorDes5taHasta27)) - lnMontoRetAnu) / lnFacRet
        End If
        '************************
        Else 'EJVG20120126
            lnMontoRes = ((lnMontoImp * CCur(psPorDes5taHasta27)) - lnMontoRetAnu) / lnFacRet
        End If
        If lnMontoRes < 0 Then
            GetImp5taEmpNew = "0.00"
        Else
            If psPlanillaCod = gsRHPlanillaGratificacion Then
                GetImp5taEmpNew = Str(Round(lnMontoRes * lnMontoAfectacionPlaGratif, 2))
            ElseIf psPlanillaCod = gsRHPlanillaSueldos Then
                GetImp5taEmpNew = Str(Round(lnMontoRes * lnMontoAfectacionPlaSueldo, 2))
            Else
                GetImp5taEmpNew = Str(Round(lnMontoRes, 2))
            End If
        End If
    ElseIf lnMontoImp <= CCur(psValor54UIT) Then 'IMPUESTO DEL 15% Y DEL 21%
        
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, pdFecFin, Trim(psPeriodoProc) & Trim(psPlanillaCod))
        lnMontoRes = (((CCur(psValor27UIT) * CCur(psPorDes5taHasta27)) + (lnMontoImp - CCur(psValor27UIT)) * CCur(psPorDes5taHasta54)) - lnMontoRetAnu) / lnFacRet
        
        If lnMontoRes < 0 Then
            GetImp5taEmpNew = "0.00"
        Else
            If psPlanillaCod = gsRHPlanillaGratificacion Then
                GetImp5taEmpNew = Str(Round(lnMontoRes * lnMontoAfectacionPlaGratif, 2))
            ElseIf psPlanillaCod = gsRHPlanillaSueldos Then
                GetImp5taEmpNew = Str(Round(lnMontoRes * lnMontoAfectacionPlaSueldo, 2))
            Else
                GetImp5taEmpNew = Str(Round(lnMontoRes, 2))
            End If
        End If
    Else                                         'IMPUESTO DEL 15%, DEL 21% Y DEL 27%
        
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, pdFecFin, Trim(psPeriodoProc) & Trim(psPlanillaCod))
        lnMontoRes = (((CCur(psValor27UIT)) * CCur(psPorDes5taHasta27)) + ((CCur(psValor54UIT) - CCur(psValor27UIT)) * CCur(psPorDes5taHasta54)) + ((lnMontoImp - CCur(psValor54UIT)) * CCur(psPorDes5taMasDe54)) - lnMontoRetAnu) / lnFacRet
        
        If lnMontoRes < 0 Then
            GetImp5taEmpNew = "0.00"
        Else
            If psPlanillaCod = gsRHPlanillaGratificacion Then
                GetImp5taEmpNew = Str(Round(lnMontoRes * lnMontoAfectacionPlaGratif, 2))
            ElseIf psPlanillaCod = gsRHPlanillaSueldos Then
                GetImp5taEmpNew = Str(Round(lnMontoRes * lnMontoAfectacionPlaSueldo, 2))
            Else
                GetImp5taEmpNew = Str(Round(lnMontoRes, 2))
            End If
        End If
    End If
End Function



Public Function GetImp5taEmpRev(psCodEmp As String, pnMontSueBasAfecto As Currency, pdFecIni As Date, pdFecFin As Date, psValor7UIT As String, psValor27UIT As String, psValor54UIT As String, psPorDes5taHasta27 As String, psPorDes5taHasta54 As String, psPorDes5taMasDe54 As String, psIngActual As String, psPlanillaCod As String, psPeriodoProc As String, pbNegativo As Boolean) As Currency
    Dim lnNumMes As Integer
    Dim lnFacIng As Integer
    Dim lnFacRet As Integer
    Dim lnMontoImp As Currency
    Dim lnMontoRetAnu As Currency
    Dim lnMontoRetOtrosMes As Currency
    
    Dim lnMontoIngAnu As Currency
    Dim lnMontoRes As Currency
    Dim lnMontoImputoMes As Currency
    
    lnNumMes = Month(CDate("31/12/" & Year(pdFecIni)))
    
    If lnNumMes < gnMes1 Then
        lnFacIng = gnNumRem - (lnNumMes - 1)
    ElseIf lnNumMes >= gnMes1 And lnNumMes < gnMes2 Then
        lnFacIng = gnNumRem - lnNumMes
    ElseIf lnNumMes = gnMes2 Then
        lnFacIng = 1
    End If
    
    lnFacRet = gnNumRem - 2 - (lnNumMes - 1)
    
    lnMontoIngAnu = GetSumIngAnuEmp(psCodEmp, CDate("31/12/" & Year(pdFecFin)), Trim(psPeriodoProc) & Trim(psPlanillaCod))
    lnMontoImp = ((pnMontSueBasAfecto * lnFacIng) + lnMontoIngAnu)
    lnMontoImp = lnMontoImp - CCur(psValor7UIT)
     
    If lnMontoImp <= 0 Then 'EXONERADO DE IMPUESTO
        lnMontoImp = 0
        
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, "31/12/" & Year(pdFecFin), Trim(psPeriodoProc) & Trim(psPlanillaCod))
        lnMontoRes = -lnMontoRetAnu
        
    ElseIf lnMontoImp <= CCur(psValor27UIT) Then 'IMPUESTO SOLO DEL 15%
        
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, "31/12/" & Year(pdFecFin), Trim(psPeriodoProc) & Trim(psPlanillaCod))
        lnMontoRes = (((lnMontoImp) * CCur(psPorDes5taHasta27)) - lnMontoRetAnu) / lnFacRet
        
        If lnMontoRes < 0 Then
            lnMontoImputoMes = "0.00"
        Else
            lnMontoImputoMes = Str(Round(lnMontoRes, 2))
        End If
    ElseIf lnMontoImp <= CCur(psValor54UIT) Then 'IMPUESTO DEL 15% Y DEL 21%
        
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, "31/12/" & Year(pdFecFin), Trim(psPeriodoProc) & Trim(psPlanillaCod))
        lnMontoRes = ((((CCur(psValor27UIT)) * CCur(psPorDes5taHasta27)) + (lnMontoImp - CCur(psValor27UIT)) * CCur(psPorDes5taHasta54)) - lnMontoRetAnu) / lnFacRet
        
        If lnMontoRes < 0 Then
            lnMontoImputoMes = "0.00"
        Else
            lnMontoImputoMes = Str(Round(lnMontoRes, 2))
        End If
    Else                                         'IMPUESTO DEL 15%, DEL 21% Y DEL 27%
        
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, CDate("31/12/" & Year(pdFecFin)), Trim(psPeriodoProc) & Trim(psPlanillaCod))
        lnMontoRes = (((CCur(psValor27UIT)) * CCur(psPorDes5taHasta27)) + ((CCur(psValor54UIT) - CCur(psValor27UIT)) * CCur(psPorDes5taHasta54)) + ((lnMontoImp - CCur(psValor54UIT)) * CCur(psPorDes5taMasDe54)) - lnMontoRetAnu) / lnFacRet
        
        If lnMontoRes < 0 Then
            lnMontoImputoMes = "0.00"
        Else
            lnMontoImputoMes = Str(Round(lnMontoRes, 2))
        End If
    End If
    
    If lnMontoRes < 0 Then
        If pbNegativo Then
            GetImp5taEmpRev = 0
        Else
            GetImp5taEmpRev = Round(Abs(lnMontoRes), 2)
        End If
    Else
        If Not pbNegativo Then
            GetImp5taEmpRev = 0
        Else
            GetImp5taEmpRev = Round(Abs(lnMontoRes), 2)
        End If
    End If
End Function



Public Function GetFomulaCon(psFun As String) As String
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim lsFun As String
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    sqlC = "Select cRHConceptoFormula cFormula from RHConceptoTabla where cRHconceptoMeno = '" & psFun & "'"
    
    'If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If RSVacio(rsC) Then
            GetFomulaCon = ""
        Else
            lsFun = rsC!cFormula
            GetFomulaCon = FunFiltro(lsFun)
        End If
    'End If
    rsC.Close
    Set rsC = Nothing
    'Set oCon = Nothing
End Function

Public Function GetMesTrabajado(psConcepto As String) As String
    Dim sqlM As String
    Dim rsM As ADODB.Recordset
    Set rsM = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    'sqlM = "Select cRHConceptoCod From RHConceptoTabla Where cRHconceptoMeno = '" & psConcepto & "' And bRHConceptoMesTrab = 1"
    sqlM = "Select dbo.RHGetMesTrabajado('" & psConcepto & "') cRHConceptoCod"
    
    'If oCon.AbreConexion Then
        Set rsM = oCon.CargaRecordSet(sqlM)
        If IsNull(rsM!cRHConceptoCod) Then
            GetMesTrabajado = "1"
        Else
            GetMesTrabajado = gsRHConceptoUMESTRAB
        End If
    'End If
    
    rsM.Close
    Set rsM = Nothing
    'Set oCon = Nothing
End Function

Private Function GetFormulaConcepto(psConcepto As String) As String
    Dim sqlF As String
    Dim rsF As ADODB.Recordset
    Set rsF = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    sqlF = "Select cRHConceptoFormula cFormula From RHConceptoTabla Where cRHconceptoMeno = '" & psConcepto & "'"
        
    'If oCon.AbreConexion Then
        Set rsF = oCon.CargaRecordSet(sqlF)
        If rsF.BOF And rsF.EOF Then
            GetFormulaConcepto = ""
        Else
            GetFormulaConcepto = Trim(rsF!cFormula)
        End If
        
    'End If
    rsF.Close
    Set rsF = Nothing
End Function

Public Sub ObtFechasAsistencia(pdFecIni As Date, pdFecFin As Date, pdFecIniGra As Date, pdFecFinGra As Date, pdFecRefIni As Date)
    Dim lsDiaIni As String
    Dim lsDiaFin As String
    Dim lsMesIni As String
    Dim lsMesFin As String
    Dim lsAñoIni As String
    Dim lsAñoFin As String
    Dim lnDiaIni As Integer
    Dim lnDiaFin As Integer
    Dim oGen As DLogGeneral
    Set oGen = New DLogGeneral
    
    lnDiaIni = oGen.CargaParametro(6000, 1002)
    lnDiaFin = oGen.CargaParametro(6000, 1003)
    
    lsDiaIni = FillNum(Trim(Str(lnDiaIni)), 2, "0")
    lsDiaFin = FillNum(Trim(Str(lnDiaFin)), 2, "0")
    
    lsMesFin = FillNum(Trim(Str(Month(pdFecRefIni))), 2, "0")
    lsMesIni = FillNum(Trim(Str(Month(DateAdd("m", -1, pdFecRefIni)))), 2, "0")
    
    lsAñoFin = FillNum(Trim(Str(Year(pdFecRefIni))), 4, "0")
    lsAñoIni = FillNum(Trim(Str(Year(DateAdd("m", -1, pdFecRefIni)))), 4, "0")
    
    pdFecIni = CDate(lsDiaIni & "/" & lsMesIni & "/" & lsAñoIni)
    pdFecFin = CDate(lsDiaFin & "/" & lsMesFin & "/" & lsAñoFin)
    
    If Month(pdFecRefIni) = 1 Then
        pdFecIniGra = CDate(lsDiaIni & "/" & "01" & "/" & lsAñoFin)
        pdFecFinGra = CDate(lsDiaFin & "/" & "07" & "/" & lsAñoFin)
    ElseIf Month(pdFecRefIni) = 7 Then
        pdFecIniGra = CDate(lsDiaIni & "/" & "06" & "/" & lsAñoFin)
        pdFecFinGra = CDate(lsDiaFin & "/" & "12" & "/" & lsAñoFin)
    End If
End Sub

'SUELDO BASICO                  1
'INGRESO FIJO APF 10.23%        2
'INGRESO FIJO APF 3.00%         3
'REINTEGROS VARIOS              4
'PAGO EN EXCESO                51
'DESCUENTO JUDICIAL FIJO       52
'DESCUENTO JUDICIAL PORCENTUAL 53
Public Function GetConceptoEmp(psPersCod As String, psRHConceptoCod As String, Optional psRHPlanillaCod As String = gsRHPlanillaSueldos) As Currency
    Dim sqlC As String
    Dim rsC As New ADODB.Recordset
    
    sqlC = " Select IsNull(nRHConceptoValor,0) nMonto From RHConcepto EM" _
         & " Where cPersCod = '" & psPersCod & "' And cRHConceptoCod =  '" & psRHConceptoCod & "'"
    
    Set rsC = oCon.CargaRecordSet(sqlC)
    If RSVacio(rsC) Then
        GetConceptoEmp = "0"
    Else
        GetConceptoEmp = Str(rsC!nMonto)
    End If
    
    rsC.Close
    Set rsC = Nothing
End Function

Public Function GetMinTarEmp(psPersCod As String, pdFecIni As Date, pdFecFin As Date) As Long
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    sqlC = " Select Sum(nRHAsistenciaTardanzasMin) Dato From RHAsistencia Where cPersCod = '" & psPersCod & "' And dRHAsistenciaFechaRef Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 0, pdFecFin), gsFormatoFecha) & "'"
    'If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If IsNull(rsC!Dato) Then
            GetMinTarEmp = 0
        Else
            GetMinTarEmp = rsC!Dato
        End If
    'End If
    rsC.Close
    Set rsC = Nothing
    'Set oCon = Nothing
End Function

Public Function GetMinPerPerEmp(psPersCod As String, pdFecIni As Date, pdFecFin As Date, pdFecRealIni As Date, pdFecRealFin As Date) As Long
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim ldFecIni As Date
    Dim ldFecFin As Date
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    Dim lnPermiso As Long
    Dim lnDifDias As Long
    
    ldFecIni = GetFecIngEmp(psPersCod)
    ldFecFin = pdFecFin
    If pdFecIni > ldFecIni Then
        ldFecIni = pdFecIni
    End If
    
    'sqlC = " Select Sum(nRHAsistenciaPermisoMin) Suma From RHAsistencia where cPersCod = '" & psPersCod & "' And dRHAsistenciaFechaRef Between '" & Format(ldFecIni, gsFormatoFecha) & "' And '" & Format(ldFecFin, gsFormatoFecha) & "'"
    sqlC = " Select dbo.RHGetMinPerPerEmp('" & psPersCod & "','" & Format(ldFecIni, gsFormatoFecha) & "','" & Format(ldFecFin, gsFormatoFecha) & "') Suma "
    
    'If oCon.AbreConexion Then
        Set rsC = oCon.CargaRecordSet(sqlC)
        If IsNull(rsC!Suma) Then
            lnPermiso = 0
        Else
            lnPermiso = rsC!Suma
        End If
        'oCon.CierraConexion
    'End If
    
    If DateDiff("d", ldFecIni, pdFecIni) <> 0 And Month(ldFecIni) = Month(ldFecFin) Then
        lnDifDias = DateDiff("d", CDate("01/" & Format(ldFecIni, "mm/yyyy ")), ldFecIni)
    Else
        lnDifDias = 0
    End If
    
    
    If lnDifDias < 0 Then lnDifDias = 0
    
    'GetMinPerPerEmp = Abs(lnAcumulador) + Abs(GetMinSalEmp(psCodEmp, ldFecIni, ldFecFin)) + Abs(GetNumDiasDesEmp(psCodEmp, pdFecRealIni, pdFecRealFin, "6','7','8','9")) * 480 + lnDifDias * 480
    GetMinPerPerEmp = Abs(lnPermiso) + lnDifDias * 480
    
    rsC.Close
    Set rsC = Nothing
    'Set oCon = Nothing
End Function

'Public Function GetSueBasNivEmp(psPersCod As String) As Currency
'    Dim sqlC As String
'    Dim rsC As New ADODB.Recordset
'    'Dim oCon As DConecta
'
'    sqlC = " Select cCarSue from Empleado EM" _
'         & " Inner Join Cargos CA On EM.cNivCod = CA.cCarCod" _
'         & " Where cEmpCod = '" & psPersCod & "'"
'
'    rsC.Open sqlC, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    If RSVacio(rsC) Then
'        GetSueBasNivEmp = 0
'    Else
'        GetSueBasNivEmp = rsC!cCarSue
'    End If
'
'    rsC.Close
'    Set rsC = Nothing
'End Function


'1 Prima
'2 Comisión variable
Public Function GetDesAPFEmp(psPersCod As String, psParametro As String) As Currency
    Dim lsCodAFP As String
    Dim sqlA As String
    Dim rsA As ADODB.Recordset
    Set rsA = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    'sqlA = " Select Case " & psParametro & " when 1 then nRHAFPPrima else nRHAFPVariable end Dato from RHEmpleado A " _
         & " Inner Join RHAFP B On A.cRHEmplAFPPersCod = B.cRHAFPPersCod" _
         & " Where cPersCod = '" & psPersCod & "'"
    
    sqlA = "Select dbo.RHGetDesAPFEmp('" & psPersCod & "','" & psParametro & "') Dato"
    
    'If oCon.AbreConexion Then
        Set rsA = oCon.CargaRecordSet(sqlA)
        If IsNull(rsA!Dato) Then
            GetDesAPFEmp = 0
        Else
            GetDesAPFEmp = rsA!Dato / 100
        End If
        'oCon.CierraConexion
    'End If
    
    rsA.Close
    Set rsA = Nothing
    'Set oCon = Nothing
End Function

Public Function GetNumMesTrab(psPersCod As String, pdFecIni As Date, pdFecFin As Date, pdFecIniProceso As Date, pdFecFinProceso As Date, psPlanillaCod As String) As Currency
    Dim sqlE As String
    Dim rsE As New ADODB.Recordset
    Dim ldFecha As Date
    Dim ldFec As Date
    Dim ldFecFin As Date
    Dim lnContador As Currency
    Dim lnContadorDias As Currency
    Dim lnNumDiasPermiso As Currency
    Dim lnDiasRestar As Integer
    Dim lnContadorDiasReal As Integer
    Dim lsCondic As String
    
    ldFecha = GetFecIngEmp(psPersCod)
    'Determinar Fecha de Trabajo
    If pdFecIni > ldFecha Then
        ldFec = pdFecIniProceso
    ElseIf pdFecFin < ldFecha Then
        ldFec = pdFecFin
    Else
        ldFec = ldFecha
    End If
    
    If Year(pdFecFinProceso) = Year(ldFecha) And Month(pdFecIniProceso) = Month(ldFecha) Then
        lnDiasRestar = Day(ldFecha)
    Else
        lnDiasRestar = Day(ldFec)
    End If
    ldFecFin = pdFecFinProceso
    lnContador = 0
    
    'ldFecFin = DateAdd("m", -1, ldFecFin)
    'ldFecFin = DateAdd("d", 1, ldFecFin)
    
'    Do While ldFecFin >= ldFec
'        ldFecFin = DateAdd("d", -1, ldFecFin)
'        ldFecFin = DateAdd("m", -1, ldFecFin)
'        ldFecFin = DateAdd("d", 1, ldFecFin)
'        lnContador = lnContador + 1
'    Loop
    
    If lnDiasRestar = 1 Then
        lnContador = DateDiff("m", ldFec, pdFecFinProceso) + 1
        If ldFecha > pdFecIniProceso Then
            lnContadorDias = DateDiff("d", ldFecha, pdFecFinProceso) + 1
        Else
            lnContadorDias = DateDiff("d", ldFec, pdFecFinProceso) + 1
        End If
        lnContadorDiasReal = 0
    Else
        lnContador = DateDiff("m", ldFec, pdFecFinProceso)
        If ldFecha > pdFecIniProceso Then
            lnContadorDias = DateDiff("d", ldFecha, pdFecFinProceso)
            lnContadorDiasReal = 30 - Day(ldFecha) + 1
        Else
            lnContadorDias = DateDiff("d", ldFec, pdFecFinProceso)
            lnContadorDiasReal = 30 - Day(ldFec) + 1
        End If
    End If
    
    If psPlanillaCod = gsRHPlanillaGratificacion Then
         lsCondic = "'" & RHEstadosRRHH.gRHEstadosRRHHSUBSIDIADO & "'"
     Else
         lsCondic = "'" & RHEstadosRRHH.gRHEstadosRRHHLICSINGOCE & "','" & RHEstadosRRHH.gRHEstadosRRHHSUSPENDIDO & "'"
    End If
    'lnNumDiasPermiso = Int(GetNumDiasPermiso(psPersCod, pdFecIni, pdFecFin, "'" & RHEstadosRRHH.gRHEstadosRRHHLICSINGOCE & "'") / 15)
    'GetNumMesTrab = lnContador - (lnNumDiasPermiso * 0.5)
    
    lnNumDiasPermiso = GetNumDiasPermiso(psPersCod, pdFecIni, pdFecFin, lsCondic, psPlanillaCod)
        
    If lnContadorDias >= 30 Then
        If lnContador = 6 Then
            GetNumMesTrab = ((lnContador * 30) - lnNumDiasPermiso) / 30
        Else
            GetNumMesTrab = (((lnContador * 30) + lnContadorDiasReal Mod 30) - lnNumDiasPermiso) / 30
        End If
    Else
        GetNumMesTrab = 0
    End If
End Function

Public Function GetNumMesTrabCTS(psPersCod As String, pdFecIni As Date, pdFecFin As Date) As Currency
    Dim ldFecha As Date
    Dim ldFec As Date
    Dim ldFecFin As Date
    Dim lnContador As Currency
    Dim lnNumDiasPermiso As Currency
    
    ldFecha = GetFecIngEmp(psPersCod)
    
    'Determinar Fecha de Trabajo
    If pdFecIni > ldFecha Then
        ldFec = pdFecIni
    ElseIf pdFecFin < ldFecha Then
        ldFec = pdFecFin
    Else
        ldFec = ldFecha
    End If
    
    ldFecFin = pdFecFin
    lnContador = 0
    
    lnContador = Abs(DateDiff("d", ldFecFin, ldFec)) + 1
    
    lnNumDiasPermiso = Int(GetNumDiasPermiso(psPersCod, pdFecIni, pdFecFin, "'" & RHEstadosRRHH.gRHEstadosRRHHLICSINGOCE & "','" & RHEstadosRRHH.gRHEstadosRRHHSUSPENDIDO & "'", "'" & gsRHPlanillaCTS & "'"))
    GetNumMesTrabCTS = (lnContador - lnNumDiasPermiso) / 30
End Function

Public Function GetNumDiasPermiso(psPersCod As String, pdFecIni As Date, pdFecFin As Date, lsEstados As String, psPlanilla) As Integer
    Dim sqlP As String
    Dim lsCondicTpo As String
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
     If psPlanilla = gsRHPlanillaGratificacion Then
         lsCondicTpo = "'" & RHEstadosTpo.RHEstadosTpoSubsidiado & "'"
     Else
         lsCondicTpo = "'" & RHEstadosTpo.RHEstadosTpoSuspendido & "'"
     End If

    sqlP = " Select Sum(Abs(Datediff(Day,dEjecutadoInicio,dEjecutadoFin))) Suma From RHPeriodosNoLaborados" _
         & " Where cPersCod = '" & psPersCod & "' And (nRHPeriodoTpo in (" & lsEstados & ") Or Left(nRHPeriodoTpo,1) = " & lsCondicTpo & " ) And (dEjecutadoInicio Between '" & Format(pdFecIni, gsFormatoFecha) & "' and '" & Format(pdFecFin, gsFormatoFecha) & "' Or dEjecutadoFin Between '" & Format(pdFecIni, gsFormatoFecha) & "' and '" & Format(pdFecFin, gsFormatoFecha) & "') and DateDiff(day,dEjecutadoInicio,dEjecutadoFin) <> 0 And nRHPeriodoEstado =1"
    
    'If oCon.AbreConexion Then
        Set rsP = oCon.CargaRecordSet(sqlP)
        If IsNull(rsP!Suma) Then
            GetNumDiasPermiso = 0
        Else
            GetNumDiasPermiso = rsP!Suma
        End If
        'oCon.CierraConexion
    'End If
    
    rsP.Close
    Set rsP = Nothing
    'Set oCon = Nothing
End Function

Public Function GetImp5taEmp(psPersCod As String, pnMontSueBasNoAfecto As Currency, pdFecIni As Date, pdFecFin As Date, psValorUIT As String, psPorDes5ta As String, psIngActual As String, psPlanillaCod As String, psPeriodoProc As String) As Currency
    Dim lnNumMes As Integer
    Dim lnFacIng As Integer
    Dim lnFacRet As Integer
    Dim lnMontoImp As Currency
    Dim lnMontoRetAnu As Currency
    Dim lnMontoRetOtrosMes As Currency
    
    Dim lnMontoIngAnu As Currency
    Dim lnMontoOtrosIngMes As Currency
    Dim lnMontoRes As Double
    
    lnNumMes = Month(CDate(pdFecIni))
    
    If lnNumMes < gnMes1 Then
        lnFacIng = gnNumRem - (lnNumMes - 1)
    ElseIf lnNumMes >= gnMes1 And lnNumMes < gnMes2 Then
        lnFacIng = gnNumRem - lnNumMes
    ElseIf lnNumMes = gnMes2 Then
        lnFacIng = 1
    End If
    
    lnFacRet = gnNumRem - 2 - (lnNumMes - 1)
    
    lnMontoIngAnu = GetSumIngAnuEmp(psPersCod, pdFecIni, Trim(psPeriodoProc) & Trim(psPlanillaCod))
    lnMontoImp = ((pnMontSueBasNoAfecto * lnFacIng) + lnMontoIngAnu) - CCur(psValorUIT)
    
    If lnMontoImp > 0 Then
        lnMontoRetAnu = GetSumRet5taAnuEmp(psPersCod, pdFecIni, Trim(psPeriodoProc) & Trim(psPlanillaCod))
        lnMontoRes = ((lnMontoImp * CCur(psPorDes5ta)) - lnMontoRetAnu) / lnFacRet
        If lnMontoRes < 0 Then
            GetImp5taEmp = "0.00"
        Else
            GetImp5taEmp = Str(lnMontoRes)
        End If
    Else
        GetImp5taEmp = 0
    End If
End Function

Public Function GetMesTrabCTS(ByVal psPersCod As String, ByVal pdFIni As Date, ByVal pdFFin As Date) As Single
    Dim rsDias As ADODB.Recordset
    Set rsDias = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    Dim sqlDias As String
    Dim nMes As Single
    Dim nDias As Single
    Dim lsFIni As String
    Dim lsFFin As String
    Dim ldFIni As Date
    Dim ldFFin As Date
    Dim ldIng As Date, ldCese As Date
    
    ldFIni = GetFecIngEmp(psPersCod)
    ldFFin = GetFecCeseEmp(psPersCod)
    
    'Determinar el #Meses que el trabajador ha estado en la empresa
    'basado en su fecha de ingreso y si cumple el periodo para el calculo
    ldIng = CDate(Format$(ldFIni, gsFormatoFechaView))
    If Not IsNull(ldFFin) Then ldCese = CDate(Format$(ldFFin, gsFormatoFechaView))
    
    If ldIng > pdFIni And Year(ldFFin) = 1900 Then
        nMes = GetMeses(ldIng, pdFFin)
    Else
        nMes = GetMeses(pdFIni, pdFFin)
    End If
    
    'Se deben calcular los dias que no trabajo por inasistencias
    'injustificadas o licencias sin goce de haber para el rango establecido
    lsFIni = Format$(pdFIni, gsFormatoFecha)
    lsFFin = Format$(pdFFin, gsFormatoFecha)
    
    sqlDias = " Select SUM(DateDiff(dd,dIni,dFin)) nDias FROM ( " _
            & " Select dIni = CASE WHEN dEjecutadoInicio < '" & lsFIni & "' THEN '" & lsFIni & "' ELSE dEjecutadoInicio END, " _
            & " dFin = CASE WHEN dEjecutadoFin > '" & lsFFin & "' THEN '" & lsFFin & "' ELSE dEjecutadoFin END " _
            & " From RHPeriodosNoLaborados " _
            & " WHERE (dEjecutadoInicio BETWEEN '" & lsFIni & "' AND '" & lsFFin & "' OR " _
            & " dEjecutadoFin Between '" & lsFIni & "' AND '" & lsFFin & "' OR " _
            & " (dEjecutadoInicio < '" & lsFIni & "' AND dEjecutadoFin > '" & lsFFin & "')) " _
            & " AND cPersCod = '" & psPersCod & "' And nRHPeriodoTpo In ('401','601') And nRHPeriodoEstado = 1) T "

    'If oCon.AbreConexion Then
        Set rsDias = oCon.CargaRecordSet(sqlDias)
        If rsDias.EOF And rsDias.BOF Then
            nDias = 0
            Else
                nDias = IIf(IsNull(rsDias!nDias), 0, rsDias!nDias)
            End If
            rsDias.Close
            Set rsDias = Nothing
            nMes = nMes - nDias / 30
            If nMes < 1 Then nMes = 0
            GetMesTrabCTS = Round(nMes, 2)
     '   End If
End Function

Public Function GetUltGratVaca(ByVal psPersCod As String, pdFecha As Date, ByVal psPeriodoProc As String) As Double
    Dim rsVac As ADODB.Recordset
    Set rsVac = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    Dim dIni As Date
    Dim sValor As String
    Dim nTotRem As Double
    Dim sqlV As String
    dIni = CDate("01/01/" & Year(pdFecha))
    
    sqlV = " Set RowCount 1 " _
         & " Select cRRHHPeriodo, Sum(Abs(nMonto)) Monto from RHPlanillaDetCon Where " _
         & " cPlanillaCod = '" & gsRHPlanillaBonificacionVacacinal & "' And cRHConceptoCod = '" & gsRHConceptoITOTINGCOD & "' And cPersCod = '" & psPersCod & "' " _
         & " And cRRHHPeriodo Between '" & Format$(dIni, "yyyymm") & "' And '" & Format$(pdFecha, "yyyymm") & "' Group by cRRHHPeriodo" _
         & " Order by cRRHHPeriodo Desc " _
         & " Set RowCount 0"
    
    'If oCon.AbreConexion Then
        Set rsVac = oCon.CargaRecordSet(sqlV)
        If rsVac.EOF And rsVac.BOF Then
            'Verificamos si tiene la suerte de tener bonificacion vacacional,
            'y si la tiene la determinamos en base a sus ingresos (lecherazo)
            sqlV = " Select cPersCod, cRHConceptoCod, cRHPlanillaCod from RHConcepto" _
                 & " Where cRHPlanillaCod = '" & gsRHPlanillaBonificacionVacacinal & "' And cPersCod = '" & psPersCod & "'"
            rsVac.Close
            Set rsVac = oCon.CargaRecordSet(sqlV)
            
            If rsVac.EOF And rsVac.BOF Then
                GetUltGratVaca = 0
            Else
                sValor = FunEvalua(gsRHConceptoVTOTREM, psPersCod, pdFecha, pdFecha, False, gsRHPlanillaBonificacionVacacinal, psPeriodoProc)
                nTotRem = ExprANum(sValor)
                GetUltGratVaca = nTotRem
            End If
        Else
            GetUltGratVaca = IIf(IsNull(rsVac!Monto), 0, rsVac!Monto)
        End If
        'oCon.CierraConexion
    'End If
    
    rsVac.Close
    Set rsVac = Nothing
    'Set oCon = Nothing
End Function

Public Function GetUltTercio(ByVal psPersCod As String, ByVal pdFIni As Date, ByVal pdFFin As Date, Optional pbProv As Boolean = True) As Double
    Dim rsTer As ADODB.Recordset
    Set rsTer = New ADODB.Recordset
    Dim nMonto As Double
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    Dim sqlV As String
    
    'If oCon.AbreConexion Then
        If pbProv Then
            sqlV = " Set RowCount 2 " _
                 & " Select cRRHHPeriodo, Sum(Abs(nMonto)) Monto From RHPlanillaDetCon  Where " _
                 & " cRHPlanillaCod = '" & gsRHPlanillaTercio & "' And cRHConceptoCod = '" & gsRHConceptoITOTINGCOD & "' And cPersCod = '" & psPersCod & "' " _
                 & " And cRRHHPeriodo <= '" & Format$(pdFFin, "yyyymm") & "' Group by cRRHHPeriodo " _
                 & " Order by cRRHHPeriodo Desc " _
                 & " Set RowCount 0"
            Set rsTer = oCon.CargaRecordSet(sqlV)
            If rsTer.EOF And rsTer.BOF Then
                GetUltTercio = 0
            Else
                nMonto = 0
                Do While Not rsTer.EOF
                    nMonto = nMonto + rsTer!Monto
                    rsTer.MoveNext
                Loop
                GetUltTercio = nMonto
            End If
        Else
            sqlV = "Select Sum(Abs(nMonto)) Monto from RHPlanillaDetCon Where " _
                 & "cPlanillaCod = '" & gsRHPlanillaTercio & "' And cRHConceptoCod = '" & gsRHConceptoITOTINGCOD & "' And cPersCod = '" & psPersCod & "' " _
                 & "And cRRHHPeriodo Between '" & Format$(pdFIni, "yyyymm") & "' And " _
                 & "'" & Format$(pdFFin, "yyyymm") & "'"
            Set rsTer = oCon.CargaRecordSet(sqlV)
            If rsTer.EOF And rsTer.BOF Then
                GetUltTercio = 0
            Else
                GetUltTercio = IIf(IsNull(rsTer!Monto), 0, rsTer!Monto)
            End If
        End If
        'oCon.CierraConexion
    'End If
    
    rsTer.Close
    Set rsTer = Nothing
    'Set oCon = Nothing
End Function

Public Function GetUltGrat(ByVal psPersCod As String, pdFecha As Date) As Double
    Dim rsGrat As ADODB.Recordset
    Dim rsBonoCanas As ADODB.Recordset
    Set rsGrat = New ADODB.Recordset
    Set rsBonoCanas = New ADODB.Recordset
    
    'Dim oCon As DConecta
    Dim sqlC As String
    Dim sqlG As String
    'Set oCon = New DConecta
    'ALPA 20101026
    Dim nBonoCanas As Currency
    sqlC = " Set RowCount 1 " _
         & " Select Sum(Abs(nMonto)) nRHConceptoValor from RHPlanillaDetCon Where " _
         & " cPlanillaCod = '" & gsRHPlanillaGratificacion & "' and cRHConceptoCod in ('" & gsRHConceptoCanasNaviCOD & "','" & gsRHConceptoBonoCreEcCOD & "','198','199') And cPersCod = '" & psPersCod & "' " _
         & " And cRRHHPeriodo <= '" & Format$(pdFecha, "yyyymm") & "' And cRRHHPeriodo >= (Select Convert(Varchar(10),dIngreso,112) from RRHH Where cPersCod = '" & psPersCod & "') Group by cRRHHPeriodo  " _
         & " Order by cRRHHPeriodo Desc " _
         & " Set RowCount 0"
    
    
'     sqlC = " select  sum(rc.nRHConceptoValor) nRHConceptoValor "
'     sqlC = sqlC & "from RHConcepto rc "
'     sqlC = sqlC & "where cPersCod='" & psPersCod & "' "
'     sqlC = sqlC & "   and nRHConceptoValor is not null "
'     sqlC = sqlC & "   and rc.cRHConceptoCod in ('" & gsRHConceptoCanasNaviCOD & "','" & gsRHConceptoBonoCreEcCOD & "') "
     
     
    Set rsBonoCanas = oCon.CargaRecordSet(sqlC)
    If rsBonoCanas.EOF And rsBonoCanas.BOF Then
        nBonoCanas = 0
    Else
        nBonoCanas = IIf(IsNull(rsBonoCanas!nRHConceptoValor), 0, rsBonoCanas!nRHConceptoValor)
    End If
     
    sqlG = " Set RowCount 1 "
    sqlG = sqlG & " Select cRRHHPeriodo, Sum(Abs(nMonto)) Monto from RHPlanillaDetCon Where "
    sqlG = sqlG & " cPlanillaCod = '" & gsRHPlanillaGratificacion & "' And cRHConceptoCod = '" & gsRHConceptoITOTINGCOD & "' And cPersCod = '" & psPersCod & "' "
    sqlG = sqlG & " And cRRHHPeriodo <= '" & Format$(pdFecha, "yyyymm") & "' And cRRHHPeriodo >= (Select Convert(Varchar(10),dIngreso,112) from RRHH Where cPersCod = '" & psPersCod & "') Group by cRRHHPeriodo  "
    sqlG = sqlG & " Order by cRRHHPeriodo Desc "
    sqlG = sqlG & " Set RowCount 0"
    '*****
    'If oCon.AbreConexion Then
        Set rsGrat = oCon.CargaRecordSet(sqlG)
        If rsGrat.EOF And rsGrat.BOF Then
            GetUltGrat = 0
        Else
            'ALPA 20101026
            'GetUltGrat = IIf(IsNull(rsGrat!Monto), 0, rsGrat!Monto)
            GetUltGrat = IIf(IsNull(rsGrat!Monto), 0, rsGrat!Monto) - nBonoCanas
        End If
    'End If
    
    rsGrat.Close
    Set rsGrat = Nothing
    'ALPA 20101026
    rsBonoCanas.Close
    Set rsBonoCanas = Nothing
    '************
    'Set oCon = Nothing
End Function

Public Function GetFecIngEmp(psPersCod As String) As Date
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    sqlE = "Select dIngreso from RRHH where cPersCod = '" & psPersCod & "'"
     
    'If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlE)
        If Not RSVacio(rsE) Then
            If Not IsNull(rsE!dIngreso) Then
                GetFecIngEmp = rsE!dIngreso
            End If
        End If
        'oCon.CierraConexion
    'End If
    
    rsE.Close
    Set rsE = Nothing
    'Set oCon = Nothing
End Function

Public Function GetFecCeseEmp(psPersCod As String) As Date
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    sqlE = "Select dCese from RRHH where cPersCod = '" & psPersCod & "'"
     
    'If oCon.AbreConexion Then
        Set rsE = oCon.CargaRecordSet(sqlE)
        If Not RSVacio(rsE) Then
            If Not IsNull(rsE!dCese) Then
                GetFecCeseEmp = rsE!dCese
            Else
                GetFecCeseEmp = CDate("01/01/1900")
            End If
        End If
        'oCon.CierraConexion
    'End If
    
    rsE.Close
    Set rsE = Nothing
    'Set oCon = Nothing
End Function

Public Function GetSumIngAnuEmp(psPersCod As String, pdFecha As Date, psPlanillaActual As String) As Currency
    Dim sqlPD As String
    Dim rsPD As ADODB.Recordset
    Set rsPD = New ADODB.Recordset
    Dim lsNumAnio As String
    Dim lsMes As String

    lsNumAnio = Trim(Str(Year(pdFecha)))
    
    lsMes = Format(pdFecha, "mmdd")
     
     sqlPD = "Select dbo.RHGetSumRet5taAnuEmp('" & psPersCod & "','" & lsNumAnio & "01','" & lsNumAnio & lsMes & "', 0, '" & psPlanillaActual & "') Suma "

    Set rsPD = oCon.CargaRecordSet(sqlPD)
    GetSumIngAnuEmp = IIf(IsNull(rsPD!Suma), "0", rsPD!Suma)
    
    rsPD.Close
    Set rsPD = Nothing
End Function

Public Function GetSumIngAnuEmpRRHH(psPersCod As String, pdFecha As Date) As Currency
    Dim sqlPD As String
    Dim rsPD As ADODB.Recordset
    Set rsPD = New ADODB.Recordset
    Dim lsNumAnio As String

    lsNumAnio = Trim(Str(Year(pdFecha)))
    sqlPD = "Select Sum(nMonto) Suma from RHPlanillaDetCon RH Inner Join RHConceptoTabla RHCT On RH.cRHConceptoCod = RHCT.cRHConceptoCod where cRRHHPeriodo Between '" & lsNumAnio & "01' And '" & lsNumAnio & Trim(Str(Month(pdFecha))) & "' And cPersCod = '" & psPersCod & "' And RH.cRHConceptoCod like '1%' And bRHConceptoImp5ta = 1"
        
    Set rsPD = oCon.CargaRecordSet(sqlPD)
    If IsNull(rsPD!Suma) Then
        GetSumIngAnuEmpRRHH = 0
    Else
        GetSumIngAnuEmpRRHH = rsPD!Suma
    End If
    
    rsPD.Close
    Set rsPD = Nothing
End Function

Public Function GetSumRet5taAnuEmp(psPersCod As String, pdFecha As Date, psPlanillaActual As String) As Currency
    Dim sqlPD As String
    Dim rsPD As ADODB.Recordset
    Set rsPD = New ADODB.Recordset
    Dim lsNumAnio As String

    Dim lsMes As String
    
    lsNumAnio = Trim(Str(Year(pdFecha)))
    If Month(pdFecha) = 12 Then
        lsMes = "12"
    Else
        lsMes = Format(Month(DateAdd("m", 0, pdFecha)), "00")
    End If

    sqlPD = "Select dbo.RHGetSumRet5taAnuEmp('" & psPersCod & "','" & lsNumAnio & "01','" & lsNumAnio & lsMes & Format(pdFecha, "dd") & "',1,'" & psPlanillaActual & "') Suma "
        
    Set rsPD = oCon.CargaRecordSet(sqlPD)
    If IsNull(rsPD!Suma) Then
        GetSumRet5taAnuEmp = 0
    Else
        GetSumRet5taAnuEmp = rsPD!Suma
    End If
    
    rsPD.Close
    Set rsPD = Nothing
End Function

Private Function GetOpeLogico(psExpLogica As String, pnPosOperador As Integer, pnLenOperador As Integer) As String
    Dim lnPos As Integer
    Dim I As Integer
    
    For I = 1 To lnInterprete_OpeTop
        If Interprete_Ope(I).lsOpeTip = "04" Then
            lnPos = InStr(1, psExpLogica, Interprete_Ope(I).lsOpeNom, vbTextCompare)
            If lnPos <> 0 Then
                pnPosOperador = lnPos
                pnLenOperador = Len(Interprete_Ope(I).lsOpeNom)
                GetOpeLogico = Interprete_Ope(I).lsOpeNom
                I = lnInterprete_OpeTop
            End If
        End If
    Next I
End Function

Private Function EvaluaOpeLogica(psOperador As String, psExp1 As String, psExp2 As String) As Boolean
    psOperador = Trim(psOperador)
    If psOperador = "=" Then
        If CCur(psExp1) = CCur(psExp2) Then
            EvaluaOpeLogica = True
        Else
            EvaluaOpeLogica = False
        End If
    ElseIf psOperador = ">" Then
        If CCur(psExp1) > CCur(psExp2) Then
            EvaluaOpeLogica = True
        Else
            EvaluaOpeLogica = False
        End If
    ElseIf psOperador = "<" Then
        If CCur(psExp1) < CCur(psExp2) Then
            EvaluaOpeLogica = True
        Else
            EvaluaOpeLogica = False
        End If
    ElseIf psOperador = "<>" Or psOperador = "><" Then
        If CCur(psExp1) <> CCur(psExp2) Then
            EvaluaOpeLogica = True
        Else
            EvaluaOpeLogica = False
        End If
    ElseIf psOperador = ">=" Then
        If CCur(psExp1) >= CCur(psExp2) Then
            EvaluaOpeLogica = True
        Else
            EvaluaOpeLogica = False
        End If
    ElseIf psOperador = "<=" Then
        If CCur(psExp1) <= CCur(psExp2) Then
            EvaluaOpeLogica = True
        Else
            EvaluaOpeLogica = False
        End If
    End If
End Function

Private Function GetNomTablaCampo(psAlias As String) As String
    Dim sqlTC As String
    Dim rsTC As ADODB.Recordset
    Set rsTC = New ADODB.Recordset
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    Dim lsAux As String
    
    If InStr(1, psAlias, ".") <> 0 Then
        lsAux = Mid(psAlias, InStr(1, psAlias, ".") + 1)
    Else
        lsAux = psAlias
    End If

    sqlTC = "Select cTabAliasTablaColumna From RHTablasAlias Where cAlias = '" & lsAux & "'"
    
    'If oCon.AbreConexion Then
        Set rsTC = oCon.CargaRecordSet(sqlTC)
        If Not RSVacio(rsTC) Then
            GetNomTablaCampo = Trim(rsTC!cTabAliasTablaColumna)
        Else
            GetNomTablaCampo = "X"
        End If
        'oCon.CierraConexion
    'End If
    
    rsTC.Close
    Set rsTC = Nothing
    'Set oCon = Nothing
End Function

Private Function EsOperador(psTipOpe As String, psParametro As String) As Boolean
    Dim I As Integer
    Dim lbBan As Boolean
    
    lbBan = False
    For I = 1 To lnInterprete_OpeTop
        If Interprete_Ope(I).lsOpeTip = psTipOpe Then
            If Interprete_Ope(I).lsOpeNom = psParametro Then lbBan = True
        End If
    Next I
    
    EsOperador = lbBan
End Function

Private Function GetValorParametroFecha(psParametro As String, pdFecIni As Date, pdFecFin As Date) As Date
    'Dim oCon As DConecta
    'Set oCon = New DConecta
    
    Select Case psParametro
        Case "NOW"
            GetValorParametroFecha = oCon.GetFechaHoraServer
        Case "DATE"
            GetValorParametroFecha = Format(oCon.GetFechaHoraServer, gsFormatoFechaView)
        Case "FINI"
            GetValorParametroFecha = pdFecIni
        Case "FFIN"
            GetValorParametroFecha = pdFecFin
        Case "TIME"
            GetValorParametroFecha = oCon.GetHoraServer
    End Select
    
    'Set oCon = Nothing
End Function

Private Function GetMeses(ByVal pdFIni As Date, ByVal pdFFin As Date) As Single
    Dim nMes As Single
    Dim dAux As Date, dTemp As Date
    'Calcular valor en meses
    nMes = 0
    dAux = CDate("01" & "/" & Month(pdFIni) & "/" & Year(pdFIni))
    dAux = DateAdd("m", 1, dAux)
    If Day(pdFIni) > 1 Then
        If dAux < pdFFin Then
            dTemp = CDate("01" & Right(DateAdd("m", 1, pdFIni), 8))
            dTemp = DateAdd("d", -1, pdFIni)
            If Day(dTemp) = 31 Then
                dTemp = DateAdd("d", -1, pdFIni)
            End If
            'nMes = nMes + DateDiff("d", pdFIni, dAux) / 30 'IIf(DateDiff("d", "01" & Right(pdFIni, 8), dAux) = 31, 31, 30)
            nMes = nMes + (30 - Day(pdFIni) + 1) / 30 'IIf(DateDiff("d", "01" & Right(pdFIni, 8), dAux) = 31, 31, 30)
            'If IsDate("30" & Right(pdFIni, 8)) Then
            '    nMes = nMes + (DateDiff("d", pdFIni, "30" & Right(pdFIni, 8)) + 1) / 30  'IIf(DateDiff("d", "01" & Right(pdFIni, 8), dAux) = 31, 31, 30)
            'ElseIf IsDate("29" & Right(pdFIni, 8)) Then
            '    nMes = nMes + (DateDiff("d", pdFIni, "29" & Right(pdFIni, 8)) + 1) / 30  'IIf(DateDiff("d", "01" & Right(pdFIni, 8), dAux) = 31, 31, 30)
            'Else
            '    nMes = nMes + (DateDiff("d", pdFIni, "28" & Right(pdFIni, 8)) + 1) / 30  'IIf(DateDiff("d", "01" & Right(pdFIni, 8), dAux) = 31, 31, 30)
            'End If
        Else
            nMes = nMes + (DateDiff("d", pdFIni, pdFFin) + 1) / 30
            GetMeses = nMes
            Exit Function
        End If
    Else
        nMes = nMes + 1
    End If
    dAux = DateAdd("m", 1, dAux)
    While DateDiff("d", -1, dAux) < pdFFin
        nMes = nMes + 1
        dAux = DateAdd("m", 1, dAux)
    Wend
    If DateAdd("d", -1, dAux) = pdFFin Then
        nMes = nMes + 1
    Else
        dAux = DateAdd("m", -1, dAux)
        nMes = Round(nMes + (DateDiff("d", dAux, pdFFin) + 1) / 30, 2)
    End If
    GetMeses = nMes
End Function

Public Function GetImp5taEmpRRHH(psCodEmp As String, pnMontSueBasNoAfecto As Currency, pdFecIni As Date, pdFecFin As Date, psValor7UIT As String, psValor27UIT As String, psValor54UIT As String, psPorDes5taHasta27 As String, psPorDes5taHasta54 As String, psPorDes5taMasDe54 As String, psIngActual As String, psIngAcum As String, psIngProy As String, psImpuesto As String, psImpuesoAcum As String, psPlanillaCod As String, psPeriodoProc As String) As Currency
    Dim lnNumMes As Integer
    Dim lnFacIng As Integer
    Dim lnFacRet As Integer
    Dim lnMontoImp As Currency
    Dim lnMontoRetAnu As Currency
    Dim lnMontoRetOtrosMes As Currency
    
    Dim lnMontoIngAnu As Currency
    Dim lnMontoOtrosIngMes As Currency
    Dim lnMontoRes As Double
    
    lnNumMes = Month(CDate(pdFecIni))
    
    If lnNumMes < gnMes1 Then
        lnFacIng = gnNumRem - (lnNumMes - 1)
    ElseIf lnNumMes >= gnMes1 And lnNumMes < gnMes2 Then
        lnFacIng = gnNumRem - lnNumMes
    ElseIf lnNumMes = gnMes2 Then
        lnFacIng = 1
    End If
     
    lnFacRet = gnNumRem - 2 - (lnNumMes - 1)
    lnMontoIngAnu = GetSumIngAnuEmp(psCodEmp, pdFecIni, "")
    psIngAcum = Str(lnMontoIngAnu)
    
    lnMontoImp = ((pnMontSueBasNoAfecto * lnFacIng) + lnMontoIngAnu)
    psIngProy = Str(lnMontoImp)
    lnMontoImp = lnMontoImp - CCur(psValor7UIT)
    
    If lnMontoImp <= 0 Then 'EXONERADO DE IMPUESTO
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, pdFecFin, Trim(psPeriodoProc) & Trim(psPlanillaCod))
        psImpuesoAcum = Str(lnMontoRetAnu)
        psImpuesto = "0"
        lnMontoImp = 0
        
    ElseIf lnMontoImp <= CCur(psValor27UIT) Then 'IMPUESTO SOLO DEL 15%
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, pdFecFin, Trim(psPeriodoProc) & Trim(psPlanillaCod))
        psImpuesoAcum = Str(lnMontoRetAnu)
        psImpuesto = Str(((lnMontoImp * CCur(psPorDes5taHasta27)) - lnMontoRetAnu))
        lnMontoRes = ((lnMontoImp * CCur(psPorDes5taHasta27)) - lnMontoRetAnu) / lnFacRet
        
        If lnMontoRes < 0 Then
            GetImp5taEmpRRHH = "0.00"
        Else
            GetImp5taEmpRRHH = Str(Round(lnMontoRes, 2))
        End If
    
    ElseIf lnMontoImp <= CCur(psValor54UIT) Then 'IMPUESTO DEL 15% Y DEL 21%
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, pdFecFin, Trim(psPeriodoProc) & Trim(psPlanillaCod))
        psImpuesoAcum = Str(lnMontoRetAnu)
        psImpuesto = Str((((CCur(psValor27UIT) * CCur(psPorDes5taHasta27)) + (lnMontoImp - CCur(psValor27UIT)) * CCur(psPorDes5taHasta54)) - lnMontoRetAnu))
        lnMontoRes = (((CCur(psValor27UIT) * CCur(psPorDes5taHasta27)) + (lnMontoImp - CCur(psValor27UIT)) * CCur(psPorDes5taHasta54)) - lnMontoRetAnu) / lnFacRet
        
        If lnMontoRes < 0 Then
            GetImp5taEmpRRHH = "0.00"
        Else
            GetImp5taEmpRRHH = Str(Round(lnMontoRes, 2))
        End If
    Else                                         'IMPUESTO DEL 15%, DEL 21% Y DEL 27%
        
        lnMontoRetAnu = GetSumRet5taAnuEmp(psCodEmp, pdFecFin, Trim(psPeriodoProc) & Trim(psPlanillaCod))
        psImpuesoAcum = Str(lnMontoRetAnu)
        psImpuesto = Str((((CCur(psValor27UIT)) * CCur(psPorDes5taHasta27)) + ((CCur(psValor54UIT) - CCur(psValor27UIT)) * CCur(psPorDes5taHasta54)) + ((lnMontoImp - CCur(psValor54UIT)) * CCur(psPorDes5taMasDe54)) - lnMontoRetAnu))
        lnMontoRes = (((CCur(psValor27UIT)) * CCur(psPorDes5taHasta27)) + ((CCur(psValor54UIT) - CCur(psValor27UIT)) * CCur(psPorDes5taHasta54)) + ((lnMontoImp - CCur(psValor54UIT)) * CCur(psPorDes5taMasDe54)) - lnMontoRetAnu) / lnFacRet
        
        If lnMontoRes < 0 Then
            GetImp5taEmpRRHH = "0.00"
        Else
            GetImp5taEmpRRHH = Str(Round(lnMontoRes, 2))
        End If
    End If
End Function

'##ModelId=3AB9030C02CB
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
    Set oCon = Nothing
End Sub

'##ModelId=3AB9030C028E
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    
    Set oCon = New DConecta
    oCon.AbreConexion
    
    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3AB9030C0220
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function GetDiasNoLabUtilEmp(psPersCod As String, pdIni As Date, pdFin As Date) As Currency
    Dim ldFecCese As Date
    Dim sqlC As String
    Dim rsC As New ADODB.Recordset
    Dim lnAcumulador1 As Long
    Dim lnAcumulador2 As Long
    Dim lsFecIni As String
    Dim lsFecFin As String
    Dim lnDiasAntesdeTrabajar As Currency
    Dim lnDiasDespuesdeTrabajar As Currency
    Dim lnDiasDuranteTrabajo As Currency
    Dim ldFecIni As Date
    Dim ldFecFin As Date
    Dim lnFecIni As Integer
    Dim lnFecFin As Integer
    
    lnFecIni = 0
    lnFecFin = 0
    
    ldFecIni = GetFecIngEmp(psPersCod)
    ldFecFin = pdFin
    
    If pdIni > ldFecIni Then
        ldFecIni = pdIni
        lnDiasAntesdeTrabajar = 0
    Else
        lnDiasAntesdeTrabajar = Abs(DateDiff("d", pdIni, ldFecIni)) - GetNumDiasFeriados(pdIni, ldFecIni)
        lnFecIni = GetNumDiasFeriados(ldFecIni, ldFecIni)
        lnFecIni = 0
    End If
    
    
    If rsC.State = 1 Then rsC.Close
    
    sqlC = " Select nRHPeriodoTpo, dEjecutadoInicio = Case When dEjecutadoInicio < '" & Format(ldFecIni, gsFormatoFecha) & "' Then '" & Format(ldFecIni, gsFormatoFecha) & "' Else dEjecutadoInicio End," _
         & "        dEjecutadoFin = Case When dEjecutadoFin > '" & Format(ldFecFin, gsFormatoFecha) & "' Then '" & Format(ldFecFin, gsFormatoFecha) & "' Else dEjecutadoFin End" _
         & " from rhperiodosnolaborados" _
         & " Where (dEjecutadoInicio Between '" & Format(ldFecIni, gsFormatoFecha) & "' And '" & Format(ldFecFin, gsFormatoFecha) & "'" _
         & " Or     dEjecutadoFin Between '" & Format(ldFecIni, gsFormatoFecha) & "' And '" & Format(ldFecFin, gsFormatoFecha) & "'" _
         & " Or    (dEjecutadoInicio > '" & Format(ldFecIni, gsFormatoFecha) & "' And dEjecutadoFin < '" & Format(ldFecFin, gsFormatoFecha) & "'))" _
         & " and cPersCod = '" & psPersCod & "' and nRHPeriodoTpo IN (403,404) And nRHPeriodoEstado = 1 order by dEjecutadoInicio"
         '301 VACACIONES CUSCO
         '403,404,504
    Set rsC = oCon.CargaRecordSet(sqlC)
    
    lnAcumulador1 = 0
    If Not RSVacio(rsC) Then
        While Not rsC.EOF
            If Left(rsC!nRHPeriodoTpo, 1) = "5" Then
                lnAcumulador1 = lnAcumulador1 + (Abs(DateDiff("d", rsC!dEjecutadoInicio, rsC!dEjecutadoFin)) * 60)
            Else
                lnAcumulador1 = lnAcumulador1 + Abs(DateDiff("n", rsC!dEjecutadoInicio, rsC!dEjecutadoFin))
            End If
            rsC.MoveNext
        Wend
    End If
    
    lsFecIni = Format(ldFecIni, gsFormatoFecha)
    lsFecFin = Format(ldFecFin, gsFormatoFecha)
    
    rsC.Close
    
    sqlC = " Select dEjecutadoInicio = Case When dEjecutadoInicio < '" & Format(ldFecIni, gsFormatoFecha) & "' Then '" & Format(ldFecIni, gsFormatoFecha) & "' Else dEjecutadoInicio End," _
         & "        dEjecutadoFin = Case When dEjecutadoFin >  '" & Format(ldFecFin, gsFormatoFecha) & "' Then '" & Format(ldFecFin + 1, gsFormatoFecha) & "' Else dEjecutadoFin End" _
         & " from rhperiodosnolaborados" _
         & " Where (dEjecutadoInicio Between '" & Format(ldFecIni, gsFormatoFecha) & "' And '" & Format(ldFecFin, gsFormatoFecha) & "'" _
         & " Or     dEjecutadoFin Between '" & Format(ldFecIni, gsFormatoFecha) & "' And '" & Format(ldFecFin, gsFormatoFecha) & "'" _
         & " Or    (dEjecutadoInicio > '" & Format(ldFecIni, gsFormatoFecha) & "' And dEjecutadoFin < '" & Format(ldFecFin, gsFormatoFecha) & "'))" _
         & " and cPersCod = '" & psPersCod & "' and nRHPeriodoTpo IN (301,401,501,502,503,601) And nRHPeriodoEstado = 1 order by dEjecutadoInicio"
         '301 CUSCO
         '(401,402,501,502,503,601)
    Set rsC = oCon.CargaRecordSet(sqlC)
    
    lnAcumulador2 = 0
    If Not RSVacio(rsC) Then
        While Not rsC.EOF
            lnAcumulador2 = lnAcumulador2 + (Abs(DateDiff("d", rsC!dEjecutadoInicio, rsC!dEjecutadoFin)) - GetNumDiasFeriados(rsC!dEjecutadoInicio, DateAdd("d", -1, rsC!dEjecutadoFin)))
            rsC.MoveNext
        Wend
    End If
    
    rsC.Close

    ldFecCese = GetFechaCese(psPersCod)
    
    If ldFecCese = CDate("01/01/1900") Or Year(ldFecCese) <> Year(pdFin) Then
        lnDiasDespuesdeTrabajar = 0
    Else
        lnDiasDespuesdeTrabajar = Abs(DateDiff("d", ldFecCese + 1, pdFin + 1)) - GetNumDiasFeriados(ldFecCese + 1, pdFin)
        lnFecFin = GetNumDiasFeriados(ldFecCese, ldFecCese)
        lnFecFin = 0
        If Day(ldFecCese) = 31 And Month(ldFecCese) = 12 Then
            'lnDiasDespuesdeTrabajar = lnDiasDespuesdeTrabajar - 1
            lnFecFin = 0
        End If
    End If

    'If lnDiasDespuesdeTrabajar <> 0 And lnDiasAntesdeTrabajar <> 0 Then
    '    lnDiasDuranteTrabajo = GetNumDiasFeriados(ldFecIni, ldFecCese)
    'Else
        lnDiasDuranteTrabajo = 0
    'End If

    'GetDiasNoLabUtilEmp = (ExprANum(FunEvalua("G_DCPUTIL", psPersCod, pdIni, pdFin, False, "VVVV", ""))) - ((((lnFecFin + lnFecIni + lnDiasAntesdeTrabajar + lnDiasDespuesdeTrabajar + lnAcumulador2 + lnDiasDuranteTrabajo) * 480) + lnAcumulador1) / 480)
    GetDiasNoLabUtilEmp = (ExprANum(FunEvalua("G_DCPUTIL", psPersCod, pdIni, pdFin, False, "VVVV", ""))) - ((((lnFecFin + lnFecIni + lnDiasAntesdeTrabajar + lnDiasDespuesdeTrabajar + lnAcumulador2 + lnDiasDuranteTrabajo) * 480) + lnAcumulador1) / 480)
    
    Dim pers  As String * 13
    Dim nompers  As String * 50
    Dim codemp  As String * 6
    Dim a As String * 6
    Dim b As String * 6
    Dim c As String * 6
    Dim d As String * 6
    Dim e As String * 6
    Dim f As String * 6
    Dim g As String * 6

    pers = psPersCod
    nompers = GetNomPers(psPersCod)
    codemp = GetCodEmp(psPersCod)
    RSet a = lnFecFin
    RSet b = lnFecIni
    RSet c = lnDiasAntesdeTrabajar
    RSet d = lnDiasDespuesdeTrabajar
    RSet e = lnAcumulador2
    RSet f = lnDiasDuranteTrabajo
    RSet g = lnAcumulador1

    cadena = pers & "  " & codemp & "  " & nompers & a & b & c & d & e & f & g
End Function

'TOTAL
Public Function GetTotalDiasNoLabUtilEmp(psPersCod As String, pdIni As Date, pdFin As Date) As Currency
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim lnAcumulador As Currency
    
    Dim ls As String
    sqlC = ""
    'sqlC = sqlC & " Select distinct cPersCod from rhplanilladetcon where cplanillacod in ('" & gsRHPlanillaSueldos & "','E02','E06','E13','E03','E09','E17','E15') and crhconceptocod = '" & gRHConcIngTotIngresos & "'  and crrhhperiodo like '" & Year(pdIni) & "%' And cPersCod <> '1120707158408'"
    'SEGUN CUSCO
    sqlC = sqlC & " Select distinct cPersCod from rhplanilladetcon where cplanillacod in ('" & gsRHPlanillaSueldos & "','E02','E13','E03','E09','E17','E15') and crhconceptocod = '" & gRHConcIngTotIngresos & "'  and crrhhperiodo like '" & Year(pdIni) & "%' And cPersCod <> '1120707158408'"
    sqlC = sqlC & " Union "
    sqlC = sqlC & " Select distinct cPersCod from rhplanilladetcon where cplanillacod in ('E08') and crhconceptocod in ('155','156','157') and crrhhperiodo like '" & Year(pdIni) & "%' And cPersCod <> '1120707158408' And cPersCod Not In (Select cPersCod From RRHH Where dcese < '" & Format(pdIni, gsFormatoFecha) & "')"
    
    
    Set rsC = oCon.CargaRecordSet(sqlC)
    
    ls = ""
    
    lnAcumulador = 0
    If Not RSVacio(rsC) Then
        While Not rsC.EOF
            lnAcumulador = lnAcumulador + (GetDiasNoLabUtilEmp(rsC!cPersCod, pdIni, pdFin))
            rsC.MoveNext

            ls = ls & cadena & Chr(10)
            
        Wend
    End If
    
    
Open "c:\archivo.txt" For Output As #1
   Write #1, ls
Close #1

    
    GetTotalDiasNoLabUtilEmp = lnAcumulador
End Function

Public Function GetTotalMontoUtilEmp(psPersCod As String, pdIni As Date, pdFin As Date) As Currency
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim lnAcumulador As Currency
    
    'sqlC = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon PD " _
         & " Where (( cRHConceptoCod = '" & gRHConcIngTotIngresos & "' And cPlanillaCod in ('" & gsRHPlanillaSueldos & "','E02','E06','E13','E03','E09','E17','E15')) Or (cRHConceptoCod In ('155','156','157') And cPlanillaCod = 'E08') ) " _
         & " And cRRHHPeriodo Between '" & Format(pdIni, gsFormatoMovFecha) & "' And '" & Format(DateAdd("d", 1, pdFin), gsFormatoMovFecha) & "' And cPersCod <> '1120707158408' And cPersCod Not In (Select cPersCod From RRHH where dCese < '" & Format(pdIni, gsFormatoFecha) & "') "
    'SEGUN CUSCO
    sqlC = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon PD " _
         & " Where (( cRHConceptoCod = '" & gRHConcIngTotIngresos & "' And cPlanillaCod in ('" & gsRHPlanillaSueldos & "','E02','E13','E03','E09','E17','E15')) Or (cRHConceptoCod In ('155','156','157') And cPlanillaCod = 'E08') ) " _
         & " And cRRHHPeriodo Between '" & Format(pdIni, gsFormatoMovFecha) & "' And '" & Format(DateAdd("d", 1, pdFin), gsFormatoMovFecha) & "' And cPersCod <> '1120707158408' And cPersCod Not In (Select cPersCod From RRHH where dCese < '" & Format(pdIni, gsFormatoFecha) & "') "
    
    Set rsC = oCon.CargaRecordSet(sqlC)
    
    GetTotalMontoUtilEmp = rsC!Monto
End Function

Public Function GetMontoUtilEmp(psPersCod As String, pdIni As Date, pdFin As Date) As Currency
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim lnAcumulador As Currency
    
    sqlC = " Select IsNull(Sum(nMonto),0) Monto " _
         & " From RHPlanillaDetCon PD " _
         & " Where cPersCod = '" & psPersCod & "' And " _
         & " (( cRHConceptoCod = '" & gRHConcIngTotIngresos & "' And cPlanillaCod in ('" & gsRHPlanillaSueldos & "','E02','E06','E13','E03','E09','E17','E15')) Or (cRHConceptoCod In ('155','156','157') And cPlanillaCod = 'E08') ) " _
         & " And cRRHHPeriodo Between '" & Format(pdIni, gsFormatoMovFecha) & "' And '" & Format(DateAdd("d", 1, pdFin), gsFormatoMovFecha) & "'"
    Set rsC = oCon.CargaRecordSet(sqlC)
    
    GetMontoUtilEmp = rsC!Monto
End Function


Public Function GetNumDiasFeriados(pdFecIni As Date, pdFecFin As Date) As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    
    sql = " Select Count(*) Num from feriado Where dFeriado Between '" & Format(pdFecIni, gsFormatoFecha) & "' And '" & Format(DateAdd("d", 0, pdFecFin), gsFormatoFecha) & "' and dFeriado not in ('2001-12-29 00:00:00.000','2001-10-01 00:00:00.000','2001-07-30 00:00:00.000','2001-01-30 00:00:00.000','2001-11-02 00:00:00.000') "
    Set rs = oCon.CargaRecordSet(sql)
    
    GetNumDiasFeriados = rs!Num
    
End Function


Public Function GetFechaCese(psEmpcCod As String) As Date
    Dim sqlE As String
    Dim rsE As New ADODB.Recordset
    
    sqlE = "Select dCese From RRHH Where cPersCod = '" & psEmpcCod & "'"
    Set rsE = oCon.CargaRecordSet(sqlE)
    
    If Not RSVacio(rsE) Then
        GetFechaCese = IIf(IsNull(rsE!dCese), CDate("01/01/1900"), rsE!dCese)
    End If
    
    rsE.Close
End Function

Public Function GetCodEmp(psPersCod As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    
    sql = " Select crhcod from RRHH Where cperscod = '" & psPersCod & "' "
    Set rs = oCon.CargaRecordSet(sql)
    
    GetCodEmp = rs!cRHCod
End Function

Public Function GetNomPers(psPersCod As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    
    sql = " Select cPersNombre from Persona Where cperscod = '" & psPersCod & "' "
    Set rs = oCon.CargaRecordSet(sql)
    
    GetNomPers = rs!cPersNombre
End Function

Private Function GetDiasVacaciones(psPersCod As String, pdIni As Date, pdFin As Date) As Double
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select Sum(dATEdIFF(Day,dSolicitadoInicio,dSolicitadoFin)) from RHPeriodosNoLaborados Where cPersCod = '" & psPersCod & "' And dSolicitadoInicio >= '" & Format(pdIni, gsFormatoFecha) & "' And dSolicitadoInicio < '" & Format(pdFin, gsFormatoFecha) & "' And nRHPeriodoTpo = 301"
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetDiasVacaciones = 0
    Else
        If IsNull(rs.Fields(0)) Then
            GetDiasVacaciones = 0
        Else
            GetDiasVacaciones = rs.Fields(0)
        End If
    End If
End Function

Private Function GetCopceptoCompleto(psPersCod As String, pnPosicion As Integer) As Double
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " select isnull(rc.nRHConceptoValor,0) nRHConceptoValor "
    sql = sql & "from RHConcepto rc INNER JOIN RHConceptoTabla rt on rc.cRHConceptoCod=rt.cRHConceptoCod "
    sql = sql & "where cPersCod='" & psPersCod & "' and rc.cRHConceptoCod=" & pnPosicion
    sql = sql & "and nRHConceptoValor is not null and cRHPlanillaCod='E01'"
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetCopceptoCompleto = 0
    Else
        If IsNull(rs.Fields(0)) Then
            GetCopceptoCompleto = 0
        Else
            GetCopceptoCompleto = rs.Fields(0)
        End If
    End If
End Function
'EJVG 20110813*********************************************************
Public Function GetUltimaGratificacion(psPersCod As String) As Double
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = "exec stp_sel_ObtenerUltimaGratificacion '" & psPersCod & "'"
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetUltimaGratificacion = 0
    Else
        If IsNull(rs.Fields(0)) Then
            GetUltimaGratificacion = 0
        Else
            GetUltimaGratificacion = rs.Fields(0)
        End If
    End If
End Function
Public Function GetCodAsistenciaMedica(psPersCod As String) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select RHAT.cRHAsistMedPrivCod " _
        & " from rhAsistenciaMedica RHA " _
        & " Inner Join rhasistmedprivtabla RHAT On RHA.cRHAsistMedPrivCod = RHAT.cRHAsistMedPrivCod " _
        & " inner join Persona P on P.cPersCod=RHA.cPersCod " _
        & " where RHA.cPersCod ='" & psPersCod & "'  order by RHA.cUltimaActualizacion Desc" _
        
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetCodAsistenciaMedica = 0
    Else
        If IsNull(rs.Fields(0)) Then
            GetCodAsistenciaMedica = 0
        Else
            GetCodAsistenciaMedica = rs.Fields(0)
        End If
    End If
End Function
Public Function GetObtenerValorConcepto(ByVal psPersCod As String, ByVal psPlanillaCod As String, ByVal psConceptoNemo As String) As Double
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " select nRHConceptoValor " _
        & " from RHConcepto RH " _
        & " inner join RHConceptoTabla RHCT on RH.cRHConceptoCod=RHCT.cRHConceptoCod " _
        & " where RH.cPersCod='" & psPersCod & "' and RH.cRHPlanillaCod='" & psPlanillaCod & "' and RHCT.cRHconceptoMeno='" & psConceptoNemo & "'"

    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetObtenerValorConcepto = 0
    Else
        If IsNull(rs.Fields(0)) Then
            GetObtenerValorConcepto = 0
        Else
            GetObtenerValorConcepto = rs.Fields(0)
        End If
    End If
End Function
'END EJVG****************************************************************
