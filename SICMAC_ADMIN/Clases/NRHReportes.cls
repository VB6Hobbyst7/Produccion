VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NRHReportes"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Actualiza en Base Autorización Fisica
Option Base 0
Option Explicit

Public Event ShowProgress()
Public Event CloseProgress()
Public Event Progress(pnValor As Long, pnTotal As Long)

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB903050003
    Private mlClassDebugID As Long
#End If

Public Function Rep5taRRHH(pgdFecSis As Date, psFecIni As String, psFecFin As String, psEmpresa As String) As String
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsCadena As String
    Dim lnMargen As Integer
    Dim lnPagina As Integer
    Dim lnItem As Long
    Dim lsCadAux1 As String
    Dim lsCadAux4 As String
    
    Dim lsProyeccion As String
    Dim lsIngAcumulado As String
    Dim lsImpProyeccion As String
    Dim lsImpAcumulado As String
    Dim lsImpuesto As String
    
    Dim lsCodigo As String * 10
    Dim lsNombre As String * 35
    Dim lsVProy As String * 18
    Dim lsVIngMes As String * 18
    Dim lsVIngAcum As String * 18
    Dim lsVValUIT As String * 18
    Dim lsVIngAfecto As String * 18
    Dim lsVImpuesto As String * 18
    Dim lsVRetencion As String * 18
    Dim lsVImpuestoMes As String * 18
    
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    
    Dim lsCadUIT7 As String
    Dim lsCadUIT27 As String
    Dim lsCadUIT54 As String
    Dim lsCadPorHasta27 As String
    Dim lsCadPorHasta54 As String
    Dim lsCadPorMas54 As String
    
    Dim lnCorr As Long
    
    Set rsE = oRep.Rep5taRRHH
    
    lsCadena = ""
    If Not (rsE.EOF And rsE.BOF) Then
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena(psEmpresa, 180) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("DETALLE DE RETENCIONES-" & Format(pgdFecSis, gsFormatoFechaView), 180) & oImpresora.gPrnSaltoLinea
        oInterprete.Interprete_InI
        lsCadena = lsCadena & Space(lnMargen) & Encabezado("Item;4; ;2;Código;7; ;3;Apellidos y Nombres;23; ;12;Ingreso mes;16; ;2;Ing. Acumulado;17; ;1;Ing.Anu.Proy;15; ;3;Val.UIT;15; ;3;Ing.Afecto;15; ;3;Impuesto;12; ;6;Impu.Rete;15; ;3;Impu.Mes;15; ;3;", lnItem)
        
        lsCadUIT7 = ExprANum(oInterprete.FunEvalua("V_UIT_7", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", ""))
        lsCadUIT27 = ExprANum(oInterprete.FunEvalua("V_UIT_27", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", ""))
        lsCadUIT54 = ExprANum(oInterprete.FunEvalua("V_UIT_54", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", ""))
        lsCadPorHasta27 = oInterprete.FunEvalua("V_POR_IMP_5TA", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", "")
        lsCadPorHasta54 = oInterprete.FunEvalua("V_POR_5TA_H54", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", "")
        lsCadPorMas54 = oInterprete.FunEvalua("V_POR_5TA_M54", "", CDate(psFecIni), CDate(psFecFin), False, "VVVV", "")
            
        RSet lsVValUIT = Format(lsCadUIT7, "#,##0.00")
            
            
        RaiseEvent ShowProgress
        lnCorr = 0
        
        While Not rsE.EOF
            lnCorr = lnCorr + 1
            lsCodigo = rsE!cRHCod
            lsNombre = PstaNombre(rsE!cPersNombre, False)
            oInterprete.Reinicia
            
            lsCadAux1 = ExprANum(oInterprete.FunEvalua("I_REM_NO_AFEC", rsE!cPersCod, CDate(psFecIni), CDate(psFecFin), False, "", ""))
            lsCadAux4 = Format(oInterprete.GetImp5taEmpRRHH(rsE!cPersCod, CCur(lsCadAux1), CDate(psFecIni), CDate(psFecFin), lsCadUIT7, lsCadUIT27, lsCadUIT54, lsCadPorHasta27, lsCadPorHasta54, lsCadPorMas54, lsCadAux1, lsIngAcumulado, lsProyeccion, lsImpProyeccion, lsImpAcumulado, "", ""), "#0.0000")
            
            RSet lsVIngMes = Format(lsCadAux1, "#,##0.00")
            RSet lsVProy = Format(lsProyeccion, "#,##0.00")
            RSet lsVIngAcum = Format(lsIngAcumulado, "#,##0.00")
            
            If CCur(lsProyeccion) - CCur(lsCadUIT7) < 0 Then
                RSet lsVIngAfecto = Format(0, "#,##0.00")
            Else
                RSet lsVIngAfecto = Format(CCur(lsProyeccion) - CCur(lsCadUIT7), "#,##0.00")
            End If
            lsVImpuesto = FillNum(Format(lsImpProyeccion, "#,##0.00"), 18, " ")
            lsVRetencion = FillNum(Format(lsImpAcumulado, "#,##0.00"), 18, " ")
            lsVImpuestoMes = FillNum(Format(lsCadAux4, "#,##0.00"), 18, " ")
            
            lsCadena = lsCadena & Space(lnMargen) & Format(lnCorr, "0000") & "  " & lsCodigo & lsNombre & lsVIngMes & lsVIngAcum & lsVProy & lsVValUIT & lsVIngAfecto & lsVImpuesto & lsVRetencion & lsVImpuestoMes & oImpresora.gPrnSaltoLinea
            
            lnItem = lnItem + 1
            
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & Space(lnMargen) & Encabezado("Item;4; ;2;Código;7; ;3;Apellidos y Nombres;23; ;12;Ingreso mes;16; ;2;Ing. Acumulado;17; ;1;Ing.Anu.Proy;15; ;3;Val.UIT;15; ;3;Ing.Afecto;15; ;3;Impuesto;12; ;6;Impu.Rete;15; ;3;Impu.Mes;15; ;3;", lnItem)
                lnItem = 0
            End If
            
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        
        RaiseEvent CloseProgress
    End If
    
    rsE.Close
    Set rsE = Nothing
    Set oRep = Nothing
    Set oInterprete = Nothing
    
    Rep5taRRHH = lsCadena
End Function

'Public Function RepContabilidadRRHH(pgdFecSis As Date, psFecIni As String, psFecFin As String, psEmpresa As String) As String
''    Dim sqlE As String
'    Dim rsE As New ADODB.Recordset
''    Dim lsCadena As String
''    Dim lnMargen As Integer
'
''    Dim lbBan As Boolean
''    Dim lsNombre As String * 30
''    Dim lsCodigo As String * 10
''    Dim lsEstFun As String * 15
''    Dim lsEstEmp As String * 15
''    Dim lsConFun As String * 15
''    Dim lsConEmp As String * 15
''    Dim lsTotal As String * 15
''    Dim lsCodAnt As String
''
''    Dim lnAcumuladorT As Integer
''    Dim lnAcumuladorEstFunT As Integer
''    Dim lnAcumuladorEstEmpT As Integer
''    Dim lnAcumuladorConFunT As Integer
''    Dim lnAcumuladorConEmpT As Integer
''    Dim lnAcumuladorPar As Integer
'
'    Dim oRep As DRHReportes
'    Set oRep = New DRHReportes
'
'    Set rsE = oRep.RepContabilidadRRHH(CDate(psFecFin))
'    Set RepContabilidadRRHH = rsE
'
''    lsNombre = ""
''    lsCodigo = ""
''    lsEstFun = ""
''    lsEstEmp = ""
''    lsConFun = ""
''    lsConEmp = ""
''    lsTotal = ""
''
''
''    lsCadena = ""
''    If Not (rsE.EOF And rsE.BOF) Then
''        RaiseEvent ShowProgress
''        lnMargen = 5
''        lsCadena = lsCadena & Space(lnMargen) & "SUPERINTENDENCIA DE BANCA Y SEGUROS" & Space(50) & "ANEXO Nro. 11" & oImpresora.gPrnSaltoLinea
''        lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
''        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("INSTITUCION : " & psEmpresa, 102) & oImpresora.gPrnSaltoLinea
''        lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
''        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("OFICINAS EN FUNCIONAMIENTO Y NUMERO DE PERSONAL", 102) & oImpresora.gPrnSaltoLinea
''
''        lsCadena = lsCadena & Space(lnMargen) & "+" & String(10, "-") & "+" & String(30, "-") & "+" & String(31, "-") & "+" & String(31, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
''        lsCadena = lsCadena & Space(lnMargen) & "¦" & CentrarCadena("CODIGO", 8) & "  ¦" & lsNombre & "¦" & CentrarCadena("ESTABLES", 30) & " ¦" & CentrarCadena("CONTRATADOS", 30) & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
''        lsCadena = lsCadena & Space(lnMargen) & "¦" & CentrarCadena("S.B.S", 8) & " ¦" & CentrarCadena("AGENCIA", 28) & " +" & String(31, "-") & "+" & String(31, "-") & "+" & CentrarCadena("TOTAL", 14) & "¦" & oImpresora.gPrnSaltoLinea
''        lsCadena = lsCadena & Space(lnMargen) & "¦" & lsCodigo & "¦" & lsNombre & "¦" & CentrarCadena("FUNCIONARIOS", 14) & " ¦" & CentrarCadena("EMPLEADOS", 14) & "¦" & CentrarCadena("FUNCIONARIOS", 14) & " ¦" & CentrarCadena("EMPLEADOS", 14) & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
''        lsCadena = lsCadena & Space(lnMargen) & "+" & String(10, "-") & "+" & String(30, "-") & "+" & String(31, "-") & "+" & String(31, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
''
''        lbBan = False
''        lsEstFun = "0"
''        lsEstEmp = "0"
''        lsConFun = "0"
''        lsConEmp = "0"
''
''        While Not rsE.EOF
''            If lsCodAnt <> rsE!cAgeCod And lbBan Then
''                lsTotal = FillNum(Trim(Str(lnAcumuladorPar)), 15, " ")
''                lsEstFun = IIf(CCur(lsEstFun) = 0, "", lsEstFun)
''                lsEstEmp = IIf(CCur(lsEstEmp) = 0, "", lsEstEmp)
''                lsConFun = IIf(CCur(lsConFun) = 0, "", lsConFun)
''                lsConEmp = IIf(CCur(lsConEmp) = 0, "", lsConEmp)
''                lsCadena = lsCadena & Space(lnMargen) & "¦" & lsCodigo & "¦" & lsNombre & "¦" & lsEstFun & "¦" & lsEstEmp & "¦" & lsConFun & "¦" & lsConEmp & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
''                lnAcumuladorPar = 0
''                lsNombre = ""
''                lsCodigo = ""
''                lsEstFun = "0"
''                lsEstEmp = "0"
''                lsConFun = "0"
''                lsConEmp = "0"
''                lsTotal = ""
''            End If
''
''            lsCodigo = rsE!cAgeCod
''            lsNombre = rsE!cAgeDescripcion
''            lnAcumuladorPar = lnAcumuladorPar + rsE!Cantidad
''            If rsE!Condicion = "ESTABLES" Then
''                If rsE!Categoria = "FUNCIONARIO" Then
''                    lsEstFun = FillNum(Trim(Str(CCur(lsEstFun) + rsE!Cantidad)), 15, " ")
''                    lnAcumuladorEstFunT = lnAcumuladorEstFunT + rsE!Cantidad
''                Else
''                    lsEstEmp = FillNum(Trim(Str(CCur(lsEstEmp) + rsE!Cantidad)), 15, " ")
''                    lnAcumuladorEstEmpT = lnAcumuladorEstEmpT + rsE!Cantidad
''                End If
''            Else
''                If rsE!Categoria = "FUNCIONARIO" Then
''                    lsConFun = FillNum(Trim(Str(CCur(lsConFun) + rsE!Cantidad)), 15, " ")
''                    lnAcumuladorConFunT = lnAcumuladorConFunT + rsE!Cantidad
''                Else
''                    lsConEmp = FillNum(Trim(Str(CCur(lsConEmp) + rsE!Cantidad)), 15, " ")
''                    lnAcumuladorConEmpT = lnAcumuladorConEmpT + rsE!Cantidad
''                End If
''            End If
''            lbBan = True
''            lsCodAnt = rsE!cAgeCod
''            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
''            rsE.MoveNext
''            RaiseEvent CloseProgress
''        Wend
''    End If
''
''    lsEstFun = IIf(CCur(lsEstFun) = 0, "", lsEstFun)
''    lsEstEmp = IIf(CCur(lsEstEmp) = 0, "", lsEstEmp)
''    lsConFun = IIf(CCur(lsConFun) = 0, "", lsConFun)
''    lsConEmp = IIf(CCur(lsConEmp) = 0, "", lsConEmp)
''
''    lsTotal = FillNum(Trim(Str(lnAcumuladorPar)), 15, " ")
''    lsCadena = lsCadena & Space(lnMargen) & "¦" & lsCodigo & "¦" & lsNombre & "¦" & lsEstFun & "¦" & lsEstEmp & "¦" & lsConFun & "¦" & lsConEmp & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
''    lnAcumuladorT = lnAcumuladorEstFunT + lnAcumuladorEstEmpT + lnAcumuladorConFunT + lnAcumuladorConEmpT
''
''    lsNombre = ""
''    lsCodigo = ""
''
''    lsEstFun = FillNum(Trim(Str(lnAcumuladorEstFunT)), 15, " ")
''    lsEstEmp = FillNum(Trim(Str(lnAcumuladorEstEmpT)), 15, " ")
''    lsConFun = FillNum(Trim(Str(lnAcumuladorConFunT)), 15, " ")
''    lsConEmp = FillNum(Trim(Str(lnAcumuladorConEmpT)), 15, " ")
''    lsTotal = FillNum(Trim(Str(lnAcumuladorT)), 15, " ")
''
''    lsCadena = lsCadena & Space(lnMargen) & "+" & String(10, "-") & "+" & String(30, "-") & "+" & String(31, "-") & "+" & String(31, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & Space(lnMargen) & "¦" & lsCodigo & "¦" & lsNombre & "¦" & lsEstFun & "¦" & lsEstEmp & "¦" & lsConFun & "¦" & lsConEmp & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & Space(lnMargen) & "+" & String(10, "-") & "+" & String(30, "-") & "+" & String(31, "-") & "+" & String(31, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
''
''    RepContabilidadRRHH = lsCadena
''    rsE.Close
''    Set rsE = Nothing
'End Function

Public Function RepContabilidadRRHH(pgdFecSis As Date, psFecIni As String, psFecFin As String, psEmpresa As String) As String
    Dim sqlE As String
    Dim rsE As New ADODB.Recordset
    Dim lsCadena As String
    Dim lnMargen As Integer
    
    Dim lbBan As Boolean
    Dim lsNombre As String * 30
    Dim lsCodigo As String * 10
    Dim lsEstFun As String * 15
    Dim lsEstEmp As String * 15
    Dim lsConFun As String * 15
    Dim lsConEmp As String * 15
    Dim lsTotal As String * 15
    Dim lsCodAnt As String
    
    Dim lnAcumuladorT As Integer
    Dim lnAcumuladorEstFunT As Integer
    Dim lnAcumuladorEstEmpT As Integer
    Dim lnAcumuladorConFunT As Integer
    Dim lnAcumuladorConEmpT As Integer
    Dim lnAcumuladorPar As Integer
    
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    Set rsE = oRep.RepContabilidadRRHH
    
    lsNombre = ""
    lsCodigo = ""
    lsEstFun = ""
    lsEstEmp = ""
    lsConFun = ""
    lsConEmp = ""
    lsTotal = ""
    
    
    lsCadena = ""
    If Not (rsE.EOF And rsE.BOF) Then
        RaiseEvent ShowProgress
        lnMargen = 5
        lsCadena = lsCadena & Space(lnMargen) & "SUPERINTENDENCIA DE BANCA Y SEGUROS" & Space(50) & "ANEXO Nro. 11" & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("INSTITUCION : " & psEmpresa, 102) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & CentrarCadena("OFICINAS EN FUNCIONAMIENTO Y NUMERO DE PERSONAL", 102) & oImpresora.gPrnSaltoLinea
        
        lsCadena = lsCadena & Space(lnMargen) & "+" & String(10, "-") & "+" & String(30, "-") & "+" & String(31, "-") & "+" & String(31, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & "¦" & CentrarCadena("CODIGO", 8) & "  ¦" & lsNombre & "¦" & CentrarCadena("ESTABLES", 30) & " ¦" & CentrarCadena("CONTRATADOS", 30) & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & "¦" & CentrarCadena("S.B.S", 8) & " ¦" & CentrarCadena("AGENCIA", 28) & " +" & String(31, "-") & "+" & String(31, "-") & "+" & CentrarCadena("TOTAL", 14) & "¦" & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & "¦" & lsCodigo & "¦" & lsNombre & "¦" & CentrarCadena("FUNCIONARIOS", 14) & " ¦" & CentrarCadena("EMPLEADOS", 14) & "¦" & CentrarCadena("FUNCIONARIOS", 14) & " ¦" & CentrarCadena("EMPLEADOS", 14) & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & Space(lnMargen) & "+" & String(10, "-") & "+" & String(30, "-") & "+" & String(31, "-") & "+" & String(31, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
        
        lbBan = False
        lsEstFun = "0"
        lsEstEmp = "0"
        lsConFun = "0"
        lsConEmp = "0"
        
        While Not rsE.EOF
            If lsCodAnt <> rsE!cAgeCod And lbBan Then
                lsTotal = FillNum(Trim(Str(lnAcumuladorPar)), 15, " ")
                lsEstFun = IIf(CCur(lsEstFun) = 0, "", lsEstFun)
                lsEstEmp = IIf(CCur(lsEstEmp) = 0, "", lsEstEmp)
                lsConFun = IIf(CCur(lsConFun) = 0, "", lsConFun)
                lsConEmp = IIf(CCur(lsConEmp) = 0, "", lsConEmp)
                lsCadena = lsCadena & Space(lnMargen) & "¦" & lsCodigo & "¦" & lsNombre & "¦" & lsEstFun & "¦" & lsEstEmp & "¦" & lsConFun & "¦" & lsConEmp & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
                lnAcumuladorPar = 0
                lsNombre = ""
                lsCodigo = ""
                lsEstFun = "0"
                lsEstEmp = "0"
                lsConFun = "0"
                lsConEmp = "0"
                lsTotal = ""
            End If
            
            lsCodigo = rsE!cAgeCod
            lsNombre = rsE!cAgeDescripcion
            lnAcumuladorPar = lnAcumuladorPar + rsE!Cantidad
            If rsE!Condicion = "ESTABLES" Then
                If rsE!Categoria = "FUNCIONARIO" Then
                    lsEstFun = FillNum(Trim(Str(CCur(lsEstFun) + rsE!Cantidad)), 15, " ")
                    lnAcumuladorEstFunT = lnAcumuladorEstFunT + rsE!Cantidad
                Else
                    lsEstEmp = FillNum(Trim(Str(CCur(lsEstEmp) + rsE!Cantidad)), 15, " ")
                    lnAcumuladorEstEmpT = lnAcumuladorEstEmpT + rsE!Cantidad
                End If
            Else
                If rsE!Categoria = "FUNCIONARIO" Then
                    lsConFun = FillNum(Trim(Str(CCur(lsConFun) + rsE!Cantidad)), 15, " ")
                    lnAcumuladorConFunT = lnAcumuladorConFunT + rsE!Cantidad
                Else
                    lsConEmp = FillNum(Trim(Str(CCur(lsConEmp) + rsE!Cantidad)), 15, " ")
                    lnAcumuladorConEmpT = lnAcumuladorConEmpT + rsE!Cantidad
                End If
            End If
            lbBan = True
            lsCodAnt = rsE!cAgeCod
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
            RaiseEvent CloseProgress
        Wend
    End If
    
    lsEstFun = IIf(CCur(lsEstFun) = 0, "", lsEstFun)
    lsEstEmp = IIf(CCur(lsEstEmp) = 0, "", lsEstEmp)
    lsConFun = IIf(CCur(lsConFun) = 0, "", lsConFun)
    lsConEmp = IIf(CCur(lsConEmp) = 0, "", lsConEmp)
    
    lsTotal = FillNum(Trim(Str(lnAcumuladorPar)), 15, " ")
    lsCadena = lsCadena & Space(lnMargen) & "¦" & lsCodigo & "¦" & lsNombre & "¦" & lsEstFun & "¦" & lsEstEmp & "¦" & lsConFun & "¦" & lsConEmp & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
    lnAcumuladorT = lnAcumuladorEstFunT + lnAcumuladorEstEmpT + lnAcumuladorConFunT + lnAcumuladorConEmpT
    
    lsNombre = ""
    lsCodigo = ""
    
    lsEstFun = FillNum(Trim(Str(lnAcumuladorEstFunT)), 15, " ")
    lsEstEmp = FillNum(Trim(Str(lnAcumuladorEstEmpT)), 15, " ")
    lsConFun = FillNum(Trim(Str(lnAcumuladorConFunT)), 15, " ")
    lsConEmp = FillNum(Trim(Str(lnAcumuladorConEmpT)), 15, " ")
    lsTotal = FillNum(Trim(Str(lnAcumuladorT)), 15, " ")
    
    lsCadena = lsCadena & Space(lnMargen) & "+" & String(10, "-") & "+" & String(30, "-") & "+" & String(31, "-") & "+" & String(31, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "¦" & lsCodigo & "¦" & lsNombre & "¦" & lsEstFun & "¦" & lsEstEmp & "¦" & lsConFun & "¦" & lsConEmp & "¦" & lsTotal & "¦" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(lnMargen) & "+" & String(10, "-") & "+" & String(30, "-") & "+" & String(31, "-") & "+" & String(31, "-") & "+" & String(15, "-") & "+" & oImpresora.gPrnSaltoLinea
    
    RepContabilidadRRHH = lsCadena
    rsE.Close
    Set rsE = Nothing
End Function

Public Function GetListaPersonal(psFechaIni As String, psFechaFin As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lnCorr As Long
    Dim lsCodigo As String * 6
    Dim lsNombre As String * 35
    Dim lsFecIng As String * 10
    Dim lsCargo As String * 30
    Dim lsAgencia As String * 23
    Dim lsCondicion As String * 10
    Dim lsSueldo As String * 15
    Set rsE = New ADODB.Recordset
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    Set rsE = oRep.GetListaPersonal()
    
    lsCadena = ""
    
    If Not (rsE.EOF And rsE.BOF) Then
        lsCadena = lsCadena & CabeceraPagina("Lista de Personal - Sueldos", lnPagina, 0, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;12; ;13;Cargo;18; ;10;Agencia;20;FecIng;24;Condicion;11;Sueldo;20;", lnItem)
        RaiseEvent ShowProgress
        lnCorr = 0
        While Not rsE.EOF
            lsCodigo = rsE!cRHCod
            lnCorr = lnCorr + 1
            lsNombre = PstaNombre(rsE!cPersNombre)
            lsFecIng = Format(rsE!dIngreso, gsFormatoFechaView)
            lsCargo = oRH.GetCargo(rsE!cPersCod, Format(CDate(psFechaFin), gsFormatoMovFecha))
            lsAgencia = rsE!cAgeDescripcion
            lsCondicion = IIf(rsE!nRHContratoTpo = 0, "ESTABLE", "CONTRATADO")
            lsSueldo = Format(oRH.GetSueldo(rsE!cPersCod), "@@@@@@@@@@@@.00")
                    
            lsCadena = lsCadena & Format(lnCorr, "0000") & " " & lsCodigo & " " & lsNombre & " " & lsCargo & "  " & lsAgencia & " " & lsFecIng & "   " & lsCondicion & " " & lsSueldo & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem > 52 Then
                lnItem = 0
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Lista de Personal", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;12; ;13;Cargo;18; ;10;Agencia;20;FecIng;24;Condicion;11;Sueldo;20;", lnItem)
            End If
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        RaiseEvent CloseProgress
    End If
    
    rsE.Close
    Set rsE = Nothing
    
    GetListaPersonal = lsCadena
End Function

Public Function GetListaPersonalPorAgencias(pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date)
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    
    Dim lsCodigo As String * 6
    Dim lsNombre As String * 35
    Dim lsFecIng As String * 10
    Dim lsCargo As String * 30
    Dim lsAgencia As String * 23
    Dim lsCondicion As String * 10
    Dim lsSueldo As String * 15
    Dim lnCorr As Long
    Set rsE = New ADODB.Recordset
    
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    Set rsE = oRep.GetListaPersonal()
    lsCadena = ""
    
    lnCorr = 0
    If Not RSVacio(rsE) Then
        lsCadena = lsCadena & CabeceraPagina("Lista de Personal", lnPagina, 0, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;12; ;13;Cargo;18; ;10;Agencia;20;FecIng;24;Condicion;11;Sueldo;20;", lnItem)
        RaiseEvent ShowProgress
        While Not rsE.EOF
            lsCodigo = rsE!cRHCod
            lnCorr = lnCorr + 1
            lsNombre = PstaNombre(rsE!cPersNombre)
            lsFecIng = Format(rsE!dIngreso, gsFormatoFechaView)
            lsCargo = oRH.GetCargo(rsE!cPersCod, Format(pgdFecSis, gsFormatoMovFecha))
            lsAgencia = rsE!cAgeDescripcion
            lsCondicion = IIf(rsE!nRHContratoTpo = 0, "ESTABLE", "CONTRATADO")
                    
            lsCadena = lsCadena & Format(lnCorr, "0000") & " " & lsCodigo & " " & lsNombre & " " & lsCargo & "  " & lsAgencia & " " & lsFecIng & "   " & lsCondicion & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem > 52 Then
                lnItem = 0
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Lista de Personal", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;12; ;13;Cargo;18; ;10;Agencia;20;FecIng;24;Condicion;11;Sueldo;20;", lnItem)
            End If
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        RaiseEvent CloseProgress
    End If
    
    rsE.Close
    Set rsE = Nothing
    
    GetListaPersonalPorAgencias = lsCadena
End Function

Public Function GetContratosVencidos(pdFecIni As Date, pdFecFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Dim lsCadena As String
    Dim lsCodAnt As String
    Dim lnItem As Long
    Dim lnPagina As Long
    Dim lsEmpCod As String * 6
    Dim lsNomPers As String * 38
    Dim lsCargo As String * 38
    Dim lsFechaIni As String * 20
    Dim lsFechaFin As String * 20
    Dim lnCorr As Long
    Set rsC = New ADODB.Recordset
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    Set rsC = oRep.GetContratosVencidos(Format(pdFecIni, "dd/mm/yyyy"), Format(DateAdd("d", 1, pdFecFin), "dd/mm/yyyy"))
    
    lsCadena = ""
    lsCodAnt = ""
    
    If Not (rsC.EOF And rsC.BOF) Then
        RaiseEvent ShowProgress
        lsCadena = lsCadena & CabeceraPagina("Vencimiento de Contrato de Empleados.-CONTRATOS", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;15; ;20;Fec.Ini;14;Fec.Fin;22;Cargo;20; ;22;", lnItem)
        lnCorr = 0
        While Not rsC.EOF
            lsEmpCod = rsC!cRHCod
            lnCorr = lnCorr + 1
            lsNomPers = PstaNombre(rsC!cPersNombre)
            lsFechaIni = Format(rsC!dRHContratoInicio, gsFormatoFechaView)
            lsFechaFin = Format(rsC!dRHContratoFin, gsFormatoFechaView)
            lsCargo = rsC!cRHCargoDescripcion
                
            lsCadena = lsCadena & Format(lnCorr, "0000") & "  " & lsEmpCod & "  " & lsNomPers & "  " & lsFechaIni & "  " & lsFechaFin & "  " & lsCargo & oImpresora.gPrnSaltoLinea
            
            lnItem = lnItem + 1
            
            If lnItem > 54 Then
                lsCadena = lsCadena & CabeceraPagina("Vencimiento de Contrato de Empleados.-CONTRATOS", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;15; ;20;Fec.Ini;14;Fec.Fin;22;Cargo;20; ;22;", lnItem)
            End If
            
            RaiseEvent Progress(rsC.Bookmark, rsC.RecordCount)
            rsC.MoveNext
        Wend
        RaiseEvent CloseProgress
    End If
    
    rsC.Close
    Set rsC = Nothing
    GetContratosVencidos = lsCadena
End Function

'##ModelId=3AB90305014E
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

Public Function GetListaPersonalSoloAgencias(pdFecIni As Date, pdFecFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    Dim rsE As ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    
    Dim lsCodigo As String * 6
    Dim lsNombre As String * 35
    Dim lsFecIng As String * 10
    Dim lsCargo As String * 30
    Dim lsAgencia As String * 23
    Dim lsCondicion As String * 10
    Dim lsSueldo As String * 15
    Dim lnCorr As Long
    
    Set rsE = New ADODB.Recordset
    
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    Set rsE = oRep.GetListaPersonal(False)
    lsCadena = ""
    
    If Not RSVacio(rsE) Then
        lsCadena = lsCadena & CabeceraPagina("Lista de Personal - SOLO AGENCIAS", lnPagina, 0, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;12; ;13;Cargo;18; ;10;Agencia;20;FecIng;24;Condicion;14;Sueldo;20;", lnItem)
        RaiseEvent ShowProgress
        lnCorr = 0
        While Not rsE.EOF
            lsCodigo = rsE!cRHCod
            lnCorr = lnCorr + 1
            lsNombre = PstaNombre(rsE!cPersNombre)
            lsFecIng = Format(rsE!dIngreso, gsFormatoFechaView)
            lsCargo = oRH.GetCargo(rsE!cPersCod, Format(CDate(pdFecFin), gsFormatoMovFecha))
            lsAgencia = rsE!cAgeDescripcion
            lsCondicion = IIf(rsE!nRHContratoTpo = 0, "ESTABLE", "CONTRATADO")
            lsSueldo = Format(oRH.GetSueldo(rsE!cPersCod), "@@@@@@@@@@@@.00")
            
            lsCadena = lsCadena & Format(lnCorr, "0000") & " " & lsCodigo & " " & lsNombre & " " & lsCargo & "  " & lsAgencia & " " & lsFecIng & "   " & lsCondicion & " " & lsSueldo & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem > 52 Then
                lnItem = 0
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Lista de Personal - SOLO AGENCIAS", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;12; ;13;Cargo;18; ;10;Agencia;20;FecIng;24;Condicion;14;Sueldo;20;", lnItem)
            End If
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        RaiseEvent CloseProgress
    End If
    
    rsE.Close
    Set rsE = Nothing
    
    GetListaPersonalSoloAgencias = lsCadena
End Function

Public Function GetListaPersonalSoloSedeAreas(pdFecIni As Date, pdFecFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lnCorr As Long
    
    Dim lsCodigo As String * 6
    Dim lsNombre As String * 35
    Dim lsFecIng As String * 10
    Dim lsCargo As String * 30
    Dim lsAgencia As String * 23
    Dim lsCondicion As String * 10
    Dim lsSueldo As String * 15
    
    Set rsE = New ADODB.Recordset
    
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    Set rsE = oRep.GetListaPersonal(False, True)
    
    lsCadena = ""
    
    If Not RSVacio(rsE) Then
        lsCadena = lsCadena & CabeceraPagina("Lista de Personal - SEDE INSTITUCIONAL POR AREAS", lnPagina, 0, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;12; ;13;Cargo;18; ;10;Agencia;20;FecIng;24;Condicion;14;Sueldo;20;", lnItem)
        RaiseEvent ShowProgress
        lnCorr = 0
        While Not rsE.EOF
            lsCodigo = rsE!cRHCod
            lnCorr = lnCorr + 1
            lsNombre = PstaNombre(rsE!cPersNombre)
            lsFecIng = Format(rsE!dIngreso, gsFormatoFechaView)
            lsCargo = oRH.GetCargo(rsE!cPersCod, Format(CDate(pdFecFin), gsFormatoMovFecha))
            lsAgencia = rsE!CAREADESCRIPCION
            lsCondicion = IIf(rsE!nRHContratoTpo = 0, "ESTABLE", "CONTRATADO")
            lsSueldo = Format(oRH.GetSueldo(rsE!cPersCod), "@@@@@@@@@@@@.00")
            
            lsCadena = lsCadena & Format(lnCorr, "0000") & " " & lsCodigo & " " & lsNombre & " " & lsCargo & "  " & lsAgencia & " " & lsFecIng & "   " & lsCondicion & " " & lsSueldo & oImpresora.gPrnSaltoLinea
            lnItem = lnItem + 1
            If lnItem > 52 Then
                lnItem = 0
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Lista de Personal - SEDE INSTITUCIONAL POR AREAS", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;2;Codigo;6;Nombre;12; ;13;Cargo;18; ;10;Agencia;20;FecIng;24;Condicion;14;Sueldo;20;", lnItem)
            End If
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        RaiseEvent CloseProgress
    End If
    
    rsE.Close
    Set rsE = Nothing
    
    GetListaPersonalSoloSedeAreas = lsCadena
End Function

'Public Function GetTardanzasEmpleados(pdIni As Date, pdFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pgsNombreCompleto As String, pgsRuc As String) As String
'    Dim sqlT As String
'    Dim rsT As ADODB.Recordset
'    Dim lsCadena As String
'    Dim lnPagina As Long
'    Dim lnIndice As Long
'    Dim lsEmpCod As String * 10
'    Dim lsEmpNom As String * 40
'    Dim lsDNI As String * 10
'    Dim lsEmpTar As String * 12
'    Dim lsEmpPerCom As String * 12
'    Dim lsEmpPerMed As String * 12
'    Dim lsEmpPerPer As String * 12
'    Dim lsHE As String * 7
'    Dim lsCodAnt As String
'
'    Set rsT = New ADODB.Recordset
'    Dim lnCorr As Long
'
'    Dim oRep As DRHReportes
'    Set oRep = New DRHReportes
'    Dim oRH As DActualizaDatosRRHH
'    Set oRH = New DActualizaDatosRRHH
'    Dim oAcceso As UAcceso
'    Set oAcceso = New UAcceso
'
'
'    Set rsT = oRep.GetAsistencia(Format(pdIni, gsFormatoFecha), Format(pdFin, gsFormatoFecha))
'
'    lsCadena = ""
'    If Not RSVacio(rsT) Then
'
'        lnCorr = 0
'        While Not rsT.EOF
'            If lsCodAnt <> Left(rsT!cRHCod, 1) Then
'                If lsCadena <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'
'                lsCadena = lsCadena & CabeceraPaginaHE(Space(16) & "REPORTE CONSOLIDADO DEL REGISTRO DE CONTROL DE ASISTENCIA" & _
'                oImpresora.gPrnSaltoLinea & Space(19) & pgsNombreCompleto & _
'                oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
'                oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
'                oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'
'                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;DNI;10;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;HE;6;", lnIndice)
'            End If
'
'            lsEmpCod = rsT!cRHCod
'            lnCorr = lnCorr + 1
'            lsEmpNom = PstaNombre(rsT!Nombre)
'             RSet lsDNI = Left(rsT!DNI, 8)
'             RSet lsEmpTar = rsT!Tardanzas
'             RSet lsEmpPerCom = 0
'             RSet lsEmpPerPer = rsT!Personales
'             RSet lsEmpPerMed = rsT!permitidos
'             RSet lsHE = IIf(IsNull(rsT!HE) Or rsT!HE < 0, Format(0, "#0.000"), Format((rsT!HE / 60), "#0.000"))
'
'            lnIndice = lnIndice + 1
'
'            lsCadena = lsCadena & Format(lnCorr, "0000") & " " & lsEmpCod & lsEmpNom & lsDNI & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & " " & lsHE & oImpresora.gPrnSaltoLinea
'
'            If lnIndice >= 54 Then
'                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'
'                lsCadena = lsCadena & CabeceraPaginaHE(Space(16) & "REPORTE CONSOLIDADO DEL REGISTRO DE CONTROL DE ASISTENCIA" & _
'                oImpresora.gPrnSaltoLinea & Space(19) & pgsNombreCompleto & _
'                oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
'                oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
'                oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'
'                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;DNI;10;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;HE;6;", lnIndice)
'            End If
'
'            lsCodAnt = Left(rsT!cRHCod, 1)
'
'            rsT.MoveNext
'        Wend
'    End If
'
'    rsT.Close
'    Set rsT = Nothing
'
'    GetTardanzasEmpleados = lsCadena
'End Function


Public Function GetTardanzasEmpleados(pdIni As Date, pdFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    Dim sqlT As String
    Dim rsT As ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnIndice As Long
    Dim lsEmpCod As String * 10
    Dim lsEmpNom As String * 40
    Dim lsEmpTar As String * 12
    Dim lsEmpPerCom As String * 12
    Dim lsEmpPerMed As String * 12
    Dim lsEmpPerPer As String * 12
    Dim lsEmpVacaiones As String * 12
    Dim lsCodAnt As String
    
    Set rsT = New ADODB.Recordset
    Dim lnCorr As Long
    
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    'Set rsT = oRep.GetAsistencia(Format(pdIni, gsFormatoFecha), Format(pdFin, gsFormatoFecha))
    Set rsT = oRep.GetAsistencia(Format(pdIni, "mm/dd/yyyy"), Format(pdFin, "mm/dd/yyyy"))
    
    lsCadena = ""
    If Not RSVacio(rsT) Then
        
        lnCorr = 0
        While Not rsT.EOF
            If lsCodAnt <> Left(rsT!cRHCod, 1) Then
                If lsCadena <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMAC ICA ", lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;Dias.Vacac;12;", lnIndice)
            End If
            
            lsEmpCod = rsT!cRHCod
            lnCorr = lnCorr + 1
            lsEmpNom = PstaNombre(rsT!Nombre)
            RSet lsEmpTar = rsT!Tardanzas
            RSet lsEmpPerCom = 0
            RSet lsEmpPerPer = rsT!Personales
            RSet lsEmpPerMed = rsT!permitidos
            
            RSet lsEmpVacaiones = rsT!Dias_Vacaciones
            
            lnIndice = lnIndice + 1
            lsCadena = lsCadena & Format(lnCorr, "0000") & " " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & lsEmpVacaiones & oImpresora.gPrnSaltoLinea
            
            If lnIndice >= 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT.", lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                'lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;", lnIndice)
                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;Dias.Vacac;12;", lnIndice)
            End If
            
            lsCodAnt = Left(rsT!cRHCod, 1)
            
            rsT.MoveNext
        Wend
    End If

    rsT.Close
    Set rsT = Nothing
    
    GetTardanzasEmpleados = lsCadena
End Function

'Public Function GetTardanzasEmpleadosDetalle(pdIni As Date, pdFin As Date, psPersCod As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pgsEmpresaCompleto As String, pgsRuc As String) As String
'    Dim sqlT As String
'    Dim rsT As ADODB.Recordset
'    Set rsT = New ADODB.Recordset
'    Dim lsCadena As String
'    Dim oRep As DRHReportes
'    Set oRep = New DRHReportes
'    Dim oRH As DActualizaDatosRRHH
'    Set oRH = New DActualizaDatosRRHH
'    Dim lsFecha As String * 12
'    Dim lsDia As String * 10
'    Dim lsFechaVig As String * 12
'    Dim lsT1Ini As String * 10
'    Dim lsT1Fin As String * 10
'    Dim lsT2Ini As String * 10
'    Dim lsT2Fin As String * 10
'    Dim lsA1Ini As String * 10
'    Dim lsA1Fin As String * 10
'    Dim lsA2Ini As String * 10
'    Dim lsA2Fin As String * 10
'    Dim lsTar1 As String * 5
'    Dim lsTar2 As String * 5
'    Dim lsHE As String * 7
'    Dim lsTardanzas As String * 18
'    Dim lsPermisos As String * 18
'    Dim lnPagina As Long
'    Dim lnIndice As Long
'    Dim lnTar1 As Long
'    Dim lnTar2 As Long
'    Dim lnHETotal As Double
'    Dim lsTotal As String
'
'    Set rsT = oRep.GetAsistenciaDet(Format(pdIni, gsFormatoFecha), Format(pdFin, gsFormatoFecha), psPersCod)
'
'    lsCadena = ""
'    lnTar1 = 0
'    lnTar2 = 0
'    lnHETotal = 0
'    If Not RSVacio(rsT) Then
'        'lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT." & rsT!cRHCod & " - " & PstaNombre(rsT!cPersNombre), lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'
'         lsCadena = lsCadena & CabeceraPaginaHE(Space(25) & "REGISTRO DE CONTROL DE ASISTENCIA" & _
'         oImpresora.gPrnSaltoLinea & Space(18) & pgsEmpresaCompleto & _
'         oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
'         oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
'         oImpresora.gPrnSaltoLinea & Space(25) & rsT!cRHCod & " - " & PstaNombre(rsT!cPersNombre) & _
'         oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'
'        lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11; ;2;TH1.Ini;8; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;6;T2;5; ;2;HE;6;", lnIndice)
'
'        While Not rsT.EOF
'            LSet lsFecha = rsT!Fecha
'            LSet lsDia = rsT!cDia
'            LSet lsFechaVig = IIf(IsNull(rsT!FechaVig), "", rsT!FechaVig)
'
'            RSet lsT1Ini = rsT!T1Ini
'            RSet lsT1Fin = rsT!T1Fin
'            RSet lsT2Ini = rsT!T2Ini
'            RSet lsT2Fin = rsT!T2Fin
'
'            If IsNull(rsT!Comision) Then
'                If IsNull(rsT!PerNoLabDes) Then
'                    RSet lsA1Ini = rsT!A1Ini
'                    RSet lsA1Fin = rsT!A1Fin
'                    RSet lsA2Ini = rsT!A2Ini
'                    RSet lsA2Fin = rsT!A2Fin
'
'                    RSet lsTar1 = rsT!MT1
'                    RSet lsTar2 = rsT!MT2
'                    RSet lsHE = IIf(IsNull(rsT!HE) Or rsT!HE < 0, Format(0, "#0.000"), Format((rsT!HE / 60), "#0.000"))
'                Else
'                    RSet lsA1Ini = Left(rsT!PerNoLabDes, 8)
'                    RSet lsA1Fin = Left(rsT!PerNoLabDes, 8)
'                    RSet lsA2Ini = Left(rsT!PerNoLabDes, 8)
'                    RSet lsA2Fin = Left(rsT!PerNoLabDes, 8)
'                    RSet lsTar1 = "0"
'                    RSet lsTar2 = "0"
'                    RSet lsHE = "0.000"
'                End If
'            Else
'                RSet lsA1Ini = "COMISION"
'                RSet lsA1Fin = "COMISION"
'                RSet lsA2Ini = "COMISION"
'                RSet lsA2Fin = "COMISION"
'                RSet lsTar1 = "0"
'                RSet lsTar2 = "0"
'            End If
'
'            lnIndice = lnIndice + 1
'            lnHETotal = lnHETotal + IIf(IsNull(rsT!HE) Or rsT!HE < 0, 0, Format((rsT!HE / 60), "#0.000"))
'            lsTotal = CStr(lnHETotal)
'            lsTotal = Format(lsTotal, "#0.000")
'            lsCadena = lsCadena & lsFecha & lsDia & lsFechaVig & lsT1Ini & lsT1Fin & lsT2Ini & lsT2Fin & lsA1Ini & lsA1Fin & lsA2Ini & lsA2Fin & lsTar1 & lsTar2 & " " & lsHE & oImpresora.gPrnSaltoLinea
'
'            If IsNull(rsT!Comision) And IsNull(rsT!PerNoLabDes) Then
'                lnTar1 = lnTar1 + rsT!MT1
'                lnTar2 = lnTar2 + rsT!MT2
'            End If
'
'            If lnIndice >= 54 Then
'                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'                'lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT." & rsT!cRHCod & " - " & rsT!cPersNombre, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'                'lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11; ;2;TH1.Ini;8; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;4; ;2;HE;6;", lnIndice)
'
'                lsCadena = lsCadena & CabeceraPaginaHE(Space(25) & "REGISTRO DE CONTROL DE ASISTENCIA" & _
'                oImpresora.gPrnSaltoLinea & Space(18) & pgsEmpresaCompleto & _
'                oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
'                oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
'                oImpresora.gPrnSaltoLinea & Space(25) & rsT!cRHCod & " - " & PstaNombre(rsT!cPersNombre) & _
'                oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'
'                lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11; ;2;TH1.Ini;8; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;6;T2;5; ;2;HE;6;", lnIndice)
'
'            End If
'            rsT.MoveNext
'        Wend
'
'        lsFecha = ""
'        lsDia = ""
'        lsFechaVig = ""
'        lsT1Ini = ""
'        lsT1Fin = ""
'        lsT2Ini = ""
'        lsT2Fin = ""
'        lsA1Ini = ""
'        lsA1Fin = ""
'        lsA2Ini = ""
'        lsA2Fin = ""
'        RSet lsTar1 = lnTar1
'        RSet lsTar2 = lnTar2
'
'        lsCadena = lsCadena & Encabezado(" ;114;" & lsTar1 & ";5;" & lsTar2 & ";5;" & lsTotal & ";8;", lnIndice)
'        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & "LEYENDA " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " TH1.Ini. Ingreso Turno1 Horario Asignado " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " TH1.Fin. Salida Turno1 Horario Asignado " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " TH2.Ini. Ingreso Turno2 Horario Asignado " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " TH2.Fin. Salida Turno2 Horario Asignado " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " TA1.Ini. Ingreso Turno1 Registro Real " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " TA1.Fin. Salida Turno1 Registro Real " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " TA2.Ini. Ingreso Turno2 Registro Real " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " TA2.Fin. Salida Turno3 Registro Real " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " T1.      Tardanza Turno 1" & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " T2.      Tardanza Turno 2 " & _
'             oImpresora.gPrnSaltoLinea & Space(10) & " HE.      Horas Extras "
'    Else
'        MsgBox "No se encontro información.", vbInformation, "Aviso"
'    End If
'
'    rsT.Close
'    Set rsT = Nothing
'
'    GetTardanzasEmpleadosDetalle = lsCadena
'End Function

Public Function GetTardanzasEmpleadosDetalle(pdIni As Date, pdFin As Date, psPersCod As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    Dim sqlT As String
    Dim rsT As ADODB.Recordset
    Set rsT = New ADODB.Recordset
    Dim lsCadena As String
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    Dim lsFecha As String * 12
    Dim lsDia As String * 10
    Dim lsFechaVig As String * 11
    Dim lsT1Ini As String * 10
    Dim lsT1Fin As String * 10
    Dim lsT2Ini As String * 10
    Dim lsT2Fin As String * 10
    Dim lsA1Ini As String * 10
    Dim lsA1Fin As String * 10
    Dim lsA2Ini As String * 10
    Dim lsA2Fin As String * 10
    Dim lsTar1 As String * 5
    Dim lsTar2 As String * 5
    Dim lsTardanzas As String * 18
    Dim lsPermisos As String * 18
    Dim lnPagina As Long
    Dim lnIndice As Long
    Dim lnTar1 As Integer
    Dim lnTar2 As Integer
    Dim lsAcumMin As String * 5
    Dim lsContTar As String * 4
    Dim lnAcumMin As Integer
    Dim lnContTar As Integer
    
    
    
    'AcumulaMin,ContTardanzas
   ' Set rsT = oRep.GetAsistenciaDet(Format(pdIni, gsFormatoFecha), Format(pdFin, gsFormatoFecha), psPersCod)
    
    Set rsT = oRep.GetAsistenciaDetNew(Format(pdIni, gsFormatoFecha), Format(pdFin, gsFormatoFecha), psPersCod)
    
    lsCadena = ""
    lnTar1 = 0
    lnTar2 = 0
    
    If Not RSVacio(rsT) Then
        lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACI." & rsT!cRHCod & " - " & PstaNombre(rsT!cPersNombre), lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        'lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11; ;2;TH1.Ini;8; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;9;T1;5;T2;5; ;2;", lnIndice)
        lsCadena = lsCadena & Encabezado("Fecha;5; ;5;Dia;5; ;5;Fecha Vig;11; ;3;TH1.Ini;8; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;;1;T1;5;T2;5;NTar;5;MST;4;", lnIndice)
        
        While Not rsT.EOF
            LSet lsFecha = rsT!Fecha
            LSet lsDia = rsT!cDia
            LSet lsFechaVig = IIf(IsNull(rsT!FechaVig), "", rsT!FechaVig)
            RSet lsT1Ini = rsT!T1Ini
            RSet lsT1Fin = rsT!T1Fin
            RSet lsT2Ini = rsT!T2Ini
            RSet lsT2Fin = rsT!T2Fin
            RSet lsA1Ini = rsT!A1Ini
            RSet lsA1Fin = rsT!A1Fin
            RSet lsA2Ini = rsT!A2Ini
            RSet lsA2Fin = rsT!A2Fin
            RSet lsTar1 = rsT!MT1
            RSet lsTar2 = rsT!MT2
            RSet lsAcumMin = rsT!AcumulaMin
            RSet lsContTar = rsT!ContTardanzas
            
            lnIndice = lnIndice + 1
            lsCadena = lsCadena & lsFecha & lsDia & lsFechaVig & lsT1Ini & lsT1Fin & lsT2Ini & lsT2Fin & lsA1Ini & lsA1Fin & lsA2Ini & lsA2Fin & lsTar1 & lsTar2 & lsContTar & lsAcumMin & oImpresora.gPrnSaltoLinea
            
            lnTar1 = lnTar1 + rsT!MT1
            lnTar2 = lnTar2 + rsT!MT2
            lnAcumMin = lnAcumMin + rsT!AcumulaMin
            lnContTar = lnContTar + rsT!ContTardanzas
            
            If lnIndice >= 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT." & rsT!cRHCod & " - " & rsT!cPersNombre, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & lsFecha & lsDia & lsFechaVig & lsT1Ini & lsT1Fin & lsT2Ini & lsT2Fin & lsA1Ini & lsA1Fin & lsA2Ini & lsA2Fin & lsTar1 & lsTar2 & lsContTar & lsAcumMin & oImpresora.gPrnSaltoLinea
            End If
            rsT.MoveNext
        Wend
    
        lsFecha = ""
        lsDia = ""
        lsFechaVig = ""
        lsT1Ini = ""
        lsT1Fin = ""
        lsT2Ini = ""
        lsT2Fin = ""
        lsA1Ini = ""
        lsA1Fin = ""
        lsA2Ini = ""
        lsA2Fin = ""
        RSet lsTar1 = lnTar1
        RSet lsTar2 = lnTar2
        
        RSet lsAcumMin = lnAcumMin
        RSet lsContTar = lnContTar
        
        'lsCadena = lsCadena & Encabezado(" ;114;" & lsTar1 & ";5;" & lsTar2 & ";5;", lnIndice)
        lsCadena = lsCadena & Encabezado(" ;113;" & lsTar1 & ";5;" & lsTar2 & ";5;" & lsContTar & ";4;" & lsAcumMin & ";5;", lnIndice)
    Else
        MsgBox "No se encontro información.", vbInformation, "Aviso"
    End If

    rsT.Close
    Set rsT = Nothing
    
    GetTardanzasEmpleadosDetalle = lsCadena
End Function


Public Function GetRepPlaAFP(pgsRuc As String, pdIni As Date, pdFin As Date) As String
    Dim lsRUC As String * 11
    Dim lsCUSSP As String * 12
    Dim lsAPaterno As String * 20
    Dim lsAMaterno As String * 20
    Dim lsANombres As String * 20
    Dim lsANombresPrimero As String * 10
    Dim lsANombresSegundo As String * 10
    Dim lsRemunera As String * 13
    Dim lsCodAFP As String * 3
    Dim lsTipoMov As String * 1
    Dim lsFecMov As String * 10
    Dim lsCadena As String
    
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    Set rsE = oRep.GetRepAFP(pgsRuc, Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov))
    
    
    If Not RSVacio(rsE) Then
        RaiseEvent ShowProgress
        While Not rsE.EOF
            lsRUC = rsE!RUC
            lsCUSSP = Trim(rsE!CUSSP & "")
            lsAPaterno = GetAPaterno(rsE!cPersNombre)
            lsAMaterno = GetAMaterno(rsE!cPersNombre)
            lsANombres = GetANombres(rsE!cPersNombre)
            lsANombresPrimero = rsE!NPrimero
            lsANombresSegundo = rsE!NSegundo
            lsRemunera = Trim(Str(rsE!Monto))
            lsCodAFP = Trim(rsE!cEmpNumAFP) & ""
            lsTipoMov = ""
            lsFecMov = ""
            'lsCadena = lsCadena & lsRUC & lsCUSSP & lsAPaterno & lsAMaterno & lsANombres & lsRemunera & lsCodAFP & lsTipoMov & lsFecMov & oImpresora.gPrnSaltoLinea
            lsCadena = lsCadena & lsCodAFP & lsCUSSP & lsAPaterno & lsAMaterno & lsANombresPrimero & lsANombresSegundo & lsRemunera & lsTipoMov & lsFecMov & oImpresora.gPrnSaltoLinea
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        RaiseEvent CloseProgress
    Else
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    GetRepPlaAFP = lsCadena
    rsE.Close
    Set rsE = Nothing
End Function

Public Function GetRepPlaAFPValida(pdIni As Date, pdFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    Dim lsItem As String * 4
    Dim lsNombre As String * 30
    Dim lsCalculado As String * 15
    Dim lsCobrado As String * 15
    Dim lsDiferencia As String * 15
    Dim lscRHAFPAbreviatura As String * 3
    Dim lsCadena As String
    Dim lnCorr As Integer
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim lnPagina As Long
    Dim lnItem As Long
    Dim lsAFPAnt As String
    Set rsE = oRep.GetRepAFPValida(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov))
    Dim lnDifAFP As Currency
    Dim lnDifAFPTotal As Currency
    
    If Not RSVacio(rsE) Then
        RaiseEvent ShowProgress
        lnCorr = 0
        lsCadena = ""
        While Not rsE.EOF
            If lsAFPAnt <> rsE!cRHAFPAbreviatura Then
                If lsCadena <> "" Then
                    RSet lsDiferencia = Format(lnDifAFP, "#,##0.00")
                    lnDifAFP = 0
                    lsCadena = lsCadena & Encabezado(" ;4; ;5;TOTAL.AFP;45; ;18;" & lsDiferencia & ";15; ;2; ;17;", lnItem, False)
                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                End If
                
                lsCadena = lsCadena & CabeceraPagina("Lista de Errores en la AFP - PRIMA SEGURO", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;5;Nombre;15; ;15;Calculado;15;Cobrado;18;Diff;15;AFP;5; ;2;Cuenta;7; ;5;", lnItem)
            End If
            
            lnDifAFPTotal = lnDifAFPTotal + rsE!Diferencia
            lnDifAFP = lnDifAFP + rsE!Diferencia
            lnCorr = lnCorr + 1
            lnItem = lnItem + 1
            lsItem = Format(lnCorr, "0000")
            lsNombre = Trim(rsE!cPersNombre)
            RSet lsCalculado = Format(rsE!Calculado, "#,##0.00")
            RSet lsCobrado = Format(rsE!Cobrado, "#,##0.00")
            RSet lsDiferencia = Format(rsE!Diferencia, "#,##0.00")
            lscRHAFPAbreviatura = rsE!cRHAFPAbreviatura
            
            lsCadena = lsCadena & lsItem & Space(2) & lsNombre & Space(2) & lsCalculado & Space(2) & lsCobrado & Space(2) & lsDiferencia & Space(2) & lscRHAFPAbreviatura & Space(2) & rsE!Cuenta & oImpresora.gPrnSaltoLinea
            
            lsAFPAnt = rsE!cRHAFPAbreviatura
            
            If lsAFPAnt <> rsE!cRHAFPAbreviatura Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Lista de Errores en la AFP - PRIMA SEGURO", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;5;Nombre;15; ;15;Calculado;15;Cobrado;18;Diff;15;AFP;5; ;2;Cuenta;7; ;5;", lnItem)
            End If
            
            
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        
        If lsCadena <> "" Then
            RSet lsDiferencia = Format(lnDifAFP, "#,##0.00")
            lnDifAFP = 0
            lsCadena = lsCadena & Encabezado(" ;4; ;5;TOTAL.AFP;45; ;18;" & lsDiferencia & ";15; ;18;", lnItem, False)
            RSet lsDiferencia = Format(lnDifAFPTotal, "#,##0.00")
            lsCadena = lsCadena & Encabezado(" ;4; ;5;TOTAL.GRAL;45; ;18;" & lsDiferencia & ";15; ;18;", lnItem, False)
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
        End If
        
        RaiseEvent CloseProgress
    Else
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    Set rsE = oRep.GetRepAFPValidaComVar(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov))
    
    If Not RSVacio(rsE) Then
        lnDifAFPTotal = 0
        lnDifAFP = 0
        lsAFPAnt = ""
        RaiseEvent ShowProgress
        lnCorr = 0

        While Not rsE.EOF
            If lsAFPAnt <> rsE!cRHAFPAbreviatura Then
                If lsCadena <> "" And lnDifAFP <> 0 Then
                    RSet lsDiferencia = Format(lnDifAFP, "#,##0.00")
                    lnDifAFP = 0
                    lsCadena = lsCadena & Encabezado(" ;4; ;5;TOTAL.AFP;45; ;19;" & lsDiferencia & ";15; ;18;", lnItem, False)
                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                End If
                
                lsCadena = lsCadena & CabeceraPagina("Lista de Errores en la AFP - COMISION VARIABLE", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;5;Nombre;15; ;15;Calculado;15;Cobrado;18;Diff;15;AFP;5; ;2;Cuenta;7; ;5;", lnItem)
            End If
            
            lnDifAFPTotal = lnDifAFPTotal + rsE!Diferencia
            lnDifAFP = lnDifAFP + rsE!Diferencia
            lnCorr = lnCorr + 1
            lnItem = lnItem + 1
            lsItem = Format(lnCorr, "0000")
            lsNombre = Trim(rsE!cPersNombre)
            RSet lsCalculado = Format(rsE!Calculado, "#,##0.00")
            RSet lsCobrado = Format(rsE!Cobrado, "#,##0.00")
            RSet lsDiferencia = Format(rsE!Diferencia, "#,##0.00")
            lscRHAFPAbreviatura = rsE!cRHAFPAbreviatura
            
            lsCadena = lsCadena & lsItem & Space(2) & lsNombre & Space(2) & lsCalculado & Space(2) & lsCobrado & Space(2) & lsDiferencia & Space(2) & lscRHAFPAbreviatura & Space(2) & rsE!Cuenta & oImpresora.gPrnSaltoLinea
            
            lsAFPAnt = rsE!cRHAFPAbreviatura
            
            If lsAFPAnt <> rsE!cRHAFPAbreviatura Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Lista de Errores en la AFP - COMISION VARIABLE", lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;5;Nombre;15; ;15;Calculado;15;Cobrado;18;Diff;15;AFP;5; ;2;Cuenta;7; ;5;", lnItem)
            End If
            
            
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        
        If lsCadena <> "" Then
            RSet lsDiferencia = Format(lnDifAFP, "#,##0.00")
            lnDifAFP = 0
            lsCadena = lsCadena & Encabezado(" ;4; ;5;TOTAL.AFP;45; ;18;" & lsDiferencia & ";15; ;19;", lnItem, False)
            RSet lsDiferencia = Format(lnDifAFPTotal, "#,##0.00")
            lsCadena = lsCadena & Encabezado(" ;4; ;5;TOTAL.GRAL;45; ;18;" & lsDiferencia & ";15; ;19;", lnItem, False)
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
        End If
        
        RaiseEvent CloseProgress
    Else
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    
    GetRepPlaAFPValida = lsCadena
    rsE.Close
    Set rsE = Nothing
End Function


'##ModelId=3AB903050112
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega a una nueva autorizacion fisica
#If DebugMode Then
    '##ModelId=3AB90305005E
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function GetIngresosConceptos(pdIni As Date, pdFin As Date, Optional pnMeses As Integer = 0) As ADODB.Recordset
    Dim rsE As ADODB.Recordset
    Dim lsCodAnt As String
    Set rsE = New ADODB.Recordset
    Dim oRep As DRHReportes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oRep = New DRHReportes
    
    'Set rsE = oRep.GetConceptosPlanillaDetalle(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov), True, True)
    Set rsE = oRep.GetConceptosPlanillaDetalle(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov), True)
    
    rs.Fields.Append "Cod_Emp", adVarChar, Val(adVarChar), adFldMayBeNull
    rs.Fields.Append "Nombre", adVarChar, Val(adVarChar), adFldMayBeNull
    
    If Not (rsE.EOF And rsE.BOF) Then
        While Not rsE.EOF
            rs.Fields.Append rsE!cRHConceptoCod, adVarChar, Val(adVarChar), adFldMayBeNull
            rsE.MoveNext
        Wend
        
        rs.Open
        rsE.MoveFirst
        rs.AddNew
        While Not rsE.EOF
            rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = rsE!cRHConceptoCod
            rsE.MoveNext
        Wend
        
        rsE.MoveFirst
        rs.AddNew
        While Not rsE.EOF
            rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = rsE!cRHConceptoDescripcion
            rsE.MoveNext
        Wend
        
        rsE.MoveFirst
        rs.AddNew
        While Not rsE.EOF
            rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = rsE!cRHConceptoAbreviatura
            rsE.MoveNext
        Wend
        rsE.Close
        
        Set rsE = oRep.GetConceptosPlanillaDetalleDatos(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov), True, 3)
        
        RaiseEvent ShowProgress
        
        While Not rsE.EOF
            If lsCodAnt <> rsE!cRHCod Then
                rs.AddNew
                rs.Fields(0) = rsE!cRHCod
                rs.Fields(1) = PstaNombre(rsE!cPersNombre)
            End If
                
            If IsNull(rs.Fields(Trim(Str(rsE!cRHConceptoCod)))) Then
                rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = rsE!Monto
            Else
                rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = CCur(rs.Fields(Trim(Str(rsE!cRHConceptoCod)))) + rsE!Monto
            End If
            
            lsCodAnt = rsE!cRHCod
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount * 2)
            rsE.MoveNext
        Wend
        Set GetIngresosConceptos = rs
                
'************************************************************************************************************
' PDT
        If pdIni > "31/10/2005" Then
                Dim I As Integer, nTim As Currency, nSubcidio As Currency, n3Porciento As Currency, nNo As Currency
                If Not rs.BOF And Not rs.EOF Then rs.MoveFirst
                Do While Not rs.EOF
                    nTim = 0: I = 1: nSubcidio = 0: nSubcidio = 0: nNo = 0: n3Porciento = 0
                    Do While I < rs.Fields.Count
                        If Not IsNull(rs(I)) Then
                            nTim = nTim + Val(rs(I))
                            Select Case rs.Fields(I).Name
                                Case "129", "151", "150", "154", "155", "156", "157", "159", "161", "163", "165"
                                    nNo = nNo + Val(rs(I))
                                Case "111", "128", "135", "136"
                                    nSubcidio = nSubcidio + Val(rs.Fields(I))
                                Case "110", "139"
                                    n3Porciento = n3Porciento + Val(rs(I))
                            End Select
                        End If
                        I = I + 1
                    Loop
                    If Not IsNull(rs(0)) Then
                        nTim = nTim - nNo
                        GuardarPDT rs(0), Format(pdIni, "yyyymm"), nTim, nTim - nSubcidio, nTim - n3Porciento
        '                If rs(0) = "E00016" Then
        '                    MsgBox nTim
        '                End If
                    End If
        '            MsgBox IIf(IsNull(rs(0)), "", rs(0))
                    RaiseEvent Progress(rs.Bookmark + rs.RecordCount, rs.RecordCount * 2)
                    rs.MoveNext
                Loop
        Else
            RaiseEvent Progress(100, 100)
        End If
''*****************************************************************************************************************
        RaiseEvent CloseProgress
    End If
End Function

Public Function GetDescuentosConceptos(pdIni As Date, pdFin As Date) As ADODB.Recordset
    Dim rsE As ADODB.Recordset
    Dim lsCodAnt As String
    Set rsE = New ADODB.Recordset
    Dim oRep As DRHReportes
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set oRep = New DRHReportes
    
    Set rsE = oRep.GetConceptosPlanillaDetalle(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov), False)
    
    rs.Fields.Append "Cod_Emp", adVarChar, Val(adVarChar), adFldMayBeNull
    rs.Fields.Append "Nombre", adVarChar, Val(adVarChar), adFldMayBeNull
    
    If Not (rsE.EOF And rsE.BOF) Then
        
        While Not rsE.EOF
            rs.Fields.Append rsE!cRHConceptoCod, adVarChar, Val(adVarChar), adFldMayBeNull
            rsE.MoveNext
        Wend
        
        rs.Open
        rsE.MoveFirst
        rs.AddNew
        While Not rsE.EOF
            rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = rsE!cRHConceptoCod
            rsE.MoveNext
        Wend
        
        rsE.MoveFirst
        rs.AddNew
        While Not rsE.EOF
            rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = rsE!cRHConceptoDescripcion
            rsE.MoveNext
        Wend
        
        rsE.MoveFirst
        rs.AddNew
        While Not rsE.EOF
            rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = rsE!cRHConceptoAbreviatura
            rsE.MoveNext
        Wend
        rsE.Close
        
        Set rsE = oRep.GetConceptosPlanillaDetalleDatos(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov), False)
        
        RaiseEvent ShowProgress
        While Not rsE.EOF
            If lsCodAnt <> rsE!cRHCod Then
                rs.AddNew
                rs.Fields(0) = rsE!cRHCod
                rs.Fields(1) = PstaNombre(rsE!cPersNombre)
            End If
                
            If IsNull(rs.Fields(Trim(Str(rsE!cRHConceptoCod)))) Then
                rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = rsE!Monto
            Else
                rs.Fields(Trim(Str(rsE!cRHConceptoCod))) = CCur(rs.Fields(Trim(Str(rsE!cRHConceptoCod)))) + rsE!Monto
            End If
            
            lsCodAnt = rsE!cRHCod
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount * 2)
            rsE.MoveNext
        Wend
        Set GetDescuentosConceptos = rs
        
'************************************************************************************************************
' PDT
        If pdIni > "31/10/2005" Then
                Dim I As Integer, n5ta As Currency, nEsSaludVida As Currency
                If Not rs.BOF And Not rs.EOF Then rs.MoveFirst
                Do While Not rs.EOF
                    n5ta = 0: nEsSaludVida = 0: I = 0
                    Do While I < rs.Fields.Count
'                       If rs!Cod_Emp = "E00064" Then
'                          MsgBox "Cerruti..."
'                       End If
                        If Not IsNull(rs(I)) Then
                            Select Case rs.Fields(I).Name
                                Case "217", "201", "227"
                                    n5ta = n5ta + CDbl(Val(rs.Fields(I)))
                                Case "207"
                                    nEsSaludVida = Val(rs(I))
                            End Select
                        End If
                        I = I + 1
                    Loop
                    If Not IsNull(rs(0)) Then
                        ModificarPDT rs(0), Format(pdIni, "yyyymm"), n5ta, nEsSaludVida
                    End If
        '            MsgBox IIf(IsNull(rs(0)), "", rs(0))
                    RaiseEvent Progress(rs.Bookmark + rs.RecordCount, rs.RecordCount * 2)
                    rs.MoveNext
                Loop
        Else
            RaiseEvent Progress(100, 100)
        End If
'*****************************************************************************************************************
        RaiseEvent CloseProgress
    End If
End Function

Private Function GetAPaterno(psNombre As String) As String
    GetAPaterno = Left(psNombre, InStr(psNombre, "/") - 1)
End Function

Private Function GetAMaterno(psNombre As String) As String
    GetAMaterno = Mid(psNombre, InStr(psNombre, "/") + 1, InStr(psNombre, ",") - InStr(psNombre, "/") - 1)
End Function

Private Function GetANombres(psNombre As String) As String
    GetANombres = Right(psNombre, Len(psNombre) - InStr(psNombre, ","))
End Function

Public Function GetPlanillaConsol(ByVal pcTipoPlanilla As String, ByVal psPeriodo As String, ByVal pnTipo As Integer) As ADODB.Recordset

Dim oRep As DRHReportes
Set oRep = New DRHReportes

Set GetPlanillaConsol = oRep.GetPlanilllaResumen(pcTipoPlanilla, psPeriodo, pnTipo)

End Function
Public Function GetEstadDetRRHHNum(pdIni As Date, pdFin As Date, prI As ADODB.Recordset, prD As ADODB.Recordset) As Boolean
    Dim lsCodAnt As String
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    Set prD = oRep.GetEstadDetRRHH(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov), False, True)
    Set prI = oRep.GetEstadDetRRHH(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov), False, False)
End Function

Public Function GetEstadDetRRHHMonto(pdIni As Date, pdFin As Date, prI As ADODB.Recordset, prD As ADODB.Recordset) As Boolean
    Dim lsCodAnt As String
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    Set prD = oRep.GetEstadDetRRHH(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov), True, True)
    Set prI = oRep.GetEstadDetRRHH(Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov), True, False)
End Function

'Public Function GetTardanzasEmpleadosPorAgencias(pdIni As Date, pdFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pgsNombreCompleto As String, pgsRuc As String) As String
'    Dim sqlT As String
'    Dim rsT As ADODB.Recordset
'    Dim lsCadena As String
'    Dim lnPagina As Long
'    Dim lnIndice As Long
'    Dim lsEmpCod As String * 10
'    Dim lsEmpNom As String * 40
'    Dim lsDNI As String * 10
'    Dim lsEmpTar As String * 12
'    Dim lsEmpPerCom As String * 12
'    Dim lsEmpPerMed As String * 12
'    Dim lsEmpPerPer As String * 12
'
'    Dim lnEmpTar As Currency
'    Dim lnEmpPerCom As Currency
'    Dim lnEmpPerMed As Currency
'    Dim lnEmpPerPer As Currency
'
'    Dim lnEmpTarT As Currency
'    Dim lnEmpPerComT As Currency
'    Dim lnEmpPerMedT As Currency
'    Dim lnEmpPerPerT As Currency
'
'
'    lnEmpTar = 0
'    lnEmpPerCom = 0
'    lnEmpPerMed = 0
'    lnEmpPerPer = 0
'
'    lnEmpTarT = 0
'    lnEmpPerComT = 0
'    lnEmpPerMedT = 0
'    lnEmpPerPerT = 0
'
'    Set rsT = New ADODB.Recordset
'    Dim lnCorr As Long
'    Dim lsCodAnt As String
'
'    Dim oRep As DRHReportes
'    Set oRep = New DRHReportes
'    Dim oRH As DActualizaDatosRRHH
'    Set oRH = New DActualizaDatosRRHH
'
'    Dim lsCodAgeAnt As String
'    lsCodAgeAnt = ""
'
'
'    Set rsT = oRep.GetAsistencia(Format(pdIni, gsFormatoFecha), Format(pdFin, gsFormatoFecha), True)
'
'    lsCadena = ""
'    If Not RSVacio(rsT) Then
'        lnCorr = 0
'        While Not rsT.EOF
'            If lsCodAgeAnt <> rsT!cAgenciaActual Or lsCodAnt <> Left(rsT!cRHCod, 1) Then
'
'                If lsCadena <> "" Then
'                    RSet lsEmpTar = lnEmpTar
'                    RSet lsEmpPerCom = lnEmpPerCom
'                    RSet lsEmpPerPer = lnEmpPerMed
'                    RSet lsEmpPerMed = lnEmpPerPer
'                    lsEmpCod = ""
'                    lsEmpNom = "TOTAL AGENCIA"
'
'                    lnEmpTar = 0
'                    lnEmpPerCom = 0
'                    lnEmpPerMed = 0
'                    lnEmpPerPer = 0
'                    lsCadena = lsCadena & String(115, "-") & oImpresora.gPrnSaltoLinea
'                    lsCadena = lsCadena & "               " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'                    lsCadena = lsCadena & String(115, "-") & oImpresora.gPrnSaltoLinea
'
'                    If lsCodAnt <> Left(rsT!cRHCod, 1) Then
'                        RSet lsEmpTar = lnEmpTarT
'                        RSet lsEmpPerCom = lnEmpPerComT
'                        RSet lsEmpPerPer = lnEmpPerMedT
'                        RSet lsEmpPerMed = lnEmpPerPerT
'                        lsEmpCod = ""
'                        lsEmpNom = "TOTAL GENERAL"
'
'                        lnEmpTarT = 0
'                        lnEmpPerComT = 0
'                        lnEmpPerMedT = 0
'                        lnEmpPerPerT = 0
'
'                        lsCadena = lsCadena & String(115, "-") & oImpresora.gPrnSaltoLinea
'                        lsCadena = lsCadena & "               " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'                        lsCadena = lsCadena & String(115, "-") & oImpresora.gPrnSaltoLinea
'                    End If
'
'                    lnEmpTar = 0
'                    lnEmpPerCom = 0
'                    lnEmpPerMed = 0
'                    lnEmpPerPer = 0
'
'                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'                End If
'
'                'lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT - " & rsT!cAgenciaActual & " - " & rsT!cAgeDescripcion, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'                lsCadena = lsCadena & CabeceraPaginaHE(Space(22) & "REPORTE DEL REGISTRO DE CONTROL DE ASISTENCIA" & _
'                oImpresora.gPrnSaltoLinea & Space(19) & pgsNombreCompleto & _
'                oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
'                oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
'                oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & _
'                oImpresora.gPrnSaltoLinea & Space(30) & rsT!cAgenciaActual & "   " & rsT!cAgeDescripcion & "", lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'
'                'lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;", lnIndice)
'                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;DNI;10;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;", lnIndice)
'            End If
'
'            lnEmpTar = lnEmpTar + rsT!Tardanzas
'            lnEmpPerCom = lnEmpPerCom + 0
'            lnEmpPerMed = lnEmpPerMed + rsT!Personales
'            lnEmpPerPer = lnEmpPerPer + rsT!permitidos
'
'            lnEmpTarT = lnEmpTarT + rsT!Tardanzas
'            lnEmpPerComT = lnEmpPerComT + 0
'            lnEmpPerMedT = lnEmpPerMedT + rsT!Personales
'            lnEmpPerPerT = lnEmpPerPerT + rsT!permitidos
'
'            lsEmpCod = rsT!cRHCod
'            lnCorr = lnCorr + 1
'            lsEmpNom = PstaNombre(rsT!Nombre)
'            RSet lsDNI = Left(rsT!DNI, 8)
'            RSet lsEmpTar = rsT!Tardanzas
'            RSet lsEmpPerCom = 0
'            RSet lsEmpPerPer = rsT!Personales
'            RSet lsEmpPerMed = rsT!permitidos
'
'            lnIndice = lnIndice + 1
'            lsCadena = lsCadena & Format(lnCorr, "0000") & " " & lsEmpCod & lsEmpNom & lsDNI & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'
'            If lnIndice >= 54 Then
'                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'                'lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT - " & rsT!cAgeDescripcion, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'                lsCadena = lsCadena & CabeceraPaginaHE(Space(22) & "REPORTE DEL REGISTRO DE CONTROL DE ASISTENCIA" & _
'                oImpresora.gPrnSaltoLinea & Space(19) & pgsNombreCompleto & _
'                oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
'                oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
'                oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & _
'                oImpresora.gPrnSaltoLinea & Space(30) & rsT!cAgenciaActual & "   " & rsT!cAgeDescripcion & "", lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'
'                'lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;", lnIndice)
'                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;DNI;10;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;", lnIndice)
'            End If
'            lsCodAgeAnt = rsT!cAgenciaActual
'            lsCodAnt = Left(rsT!cRHCod, 1)
'            rsT.MoveNext
'        Wend
'    End If
'
'    rsT.Close
'    Set rsT = Nothing
'
'    If lsCadena <> "" Then
'        RSet lsEmpTar = lnEmpTar
'        RSet lsEmpPerCom = lnEmpPerCom
'        RSet lsEmpPerPer = lnEmpPerMed
'        RSet lsEmpPerMed = lnEmpPerPer
'        lsEmpCod = ""
'        lsEmpNom = "TOTAL AGENCIA"
'
'        lnEmpTar = 0
'        lnEmpPerCom = 0
'        lnEmpPerMed = 0
'        lnEmpPerPer = 0
'        lsCadena = lsCadena & String(115, "-") & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & "               " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & String(115, "-") & oImpresora.gPrnSaltoLinea
'    End If
'
'
'
'    RSet lsEmpTar = lnEmpTarT
'    RSet lsEmpPerCom = lnEmpPerComT
'    RSet lsEmpPerPer = lnEmpPerMedT
'    RSet lsEmpPerMed = lnEmpPerPerT
'    lsEmpCod = ""
'    lsEmpNom = "TOTAL GENERAL"
'
'    lnEmpTar = 0
'    lnEmpPerCom = 0
'    lnEmpPerMed = 0
'    lnEmpPerPer = 0
'
'    lsCadena = lsCadena & String(115, "-") & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & "               " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & String(115, "-") & oImpresora.gPrnSaltoLinea
'
'    lnEmpTar = 0
'    lnEmpPerCom = 0
'    lnEmpPerMed = 0
'    lnEmpPerPer = 0
'
'    GetTardanzasEmpleadosPorAgencias = lsCadena
'End Function

'Public Function GetTardanzasEmpleadosPorAgencias(pdIni As Date, pdFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, psCodagencia As String) As String
'    Dim sqlT As String
'    Dim rsT As ADODB.Recordset
'    Dim lsCadena As String
'    Dim lnPagina As Long
'    Dim lnIndice As Long
'    Dim lsEmpCod As String * 10
'    Dim lsEmpNom As String * 40
'    Dim lsEmpTar As String * 12
'    Dim lsEmpPerCom As String * 12
'    Dim lsEmpPerMed As String * 12
'    Dim lsEmpPerPer As String * 12
'
'    Dim lnEmpTar As Currency
'    Dim lnEmpPerCom As Currency
'    Dim lnEmpPerMed As Currency
'    Dim lnEmpPerPer As Currency
'
'    Dim lnEmpTarT As Currency
'    Dim lnEmpPerComT As Currency
'    Dim lnEmpPerMedT As Currency
'    Dim lnEmpPerPerT As Currency
'
'    lnEmpTar = 0
'    lnEmpPerCom = 0
'    lnEmpPerMed = 0
'    lnEmpPerPer = 0
'
'    lnEmpTarT = 0
'    lnEmpPerComT = 0
'    lnEmpPerMedT = 0
'    lnEmpPerPerT = 0
'
'    Set rsT = New ADODB.Recordset
'    Dim lnCorr As Long
'    Dim lsCodAnt As String
'
'    Dim oRep As DRHReportes
'    Set oRep = New DRHReportes
'    Dim oRH As DActualizaDatosRRHH
'    Set oRH = New DActualizaDatosRRHH
'
'    Dim lsCodAgeAnt As String
'    lsCodAgeAnt = ""
'
'    Set rsT = oRep.GetAsistencia(Format(pdIni, gsFormatoFecha), Format(pdFin, gsFormatoFecha), True, psCodagencia)
'
'    lsCadena = ""
'    If Not RSVacio(rsT) Then
'        lnCorr = 0
'        While Not rsT.EOF
'            If lsCodAgeAnt <> rsT!cAgenciaActual Or lsCodAnt <> Left(rsT!cRHCod, 1) Then
'
'                If lsCadena <> "" Then
'                    RSet lsEmpTar = lnEmpTar
'                    RSet lsEmpPerCom = lnEmpPerCom
'                    RSet lsEmpPerPer = lnEmpPerMed
'                    RSet lsEmpPerMed = lnEmpPerPer
'                    lsEmpCod = ""
'                    lsEmpNom = "TOTAL AGENCIA"
'
'                    lnEmpTar = 0
'                    lnEmpPerCom = 0
'                    lnEmpPerMed = 0
'                    lnEmpPerPer = 0
'                    lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
'                    lsCadena = lsCadena & "     " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'                    lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
'
'                    If lsCodAnt <> Left(rsT!cRHCod, 1) Then
'                        RSet lsEmpTar = lnEmpTarT
'                        RSet lsEmpPerCom = lnEmpPerComT
'                        RSet lsEmpPerPer = lnEmpPerMedT
'                        RSet lsEmpPerMed = lnEmpPerPerT
'                        lsEmpCod = ""
'                        lsEmpNom = "TOTAL GENERAL"
'
'                        lnEmpTarT = 0
'                        lnEmpPerComT = 0
'                        lnEmpPerMedT = 0
'                        lnEmpPerPerT = 0
'
'                        lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
'                        lsCadena = lsCadena & "     " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'                        lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
'                    End If
'
'                    lnEmpTar = 0
'                    lnEmpPerCom = 0
'                    lnEmpPerMed = 0
'                    lnEmpPerPer = 0
'
'                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'                End If
'
'                lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT - " & rsT!cAgenciaActual & " - " & rsT!cAgeDescripcion, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;", lnIndice)
'            End If
'
'            lnEmpTar = lnEmpTar + rsT!Tardanzas
'            lnEmpPerCom = lnEmpPerCom + 0
'            lnEmpPerMed = lnEmpPerMed + rsT!Personales
'            lnEmpPerPer = lnEmpPerPer + rsT!permitidos
'
'            lnEmpTarT = lnEmpTarT + rsT!Tardanzas
'            lnEmpPerComT = lnEmpPerComT + 0
'            lnEmpPerMedT = lnEmpPerMedT + rsT!Personales
'            lnEmpPerPerT = lnEmpPerPerT + rsT!permitidos
'
'            lsEmpCod = rsT!cRHCod
'            lnCorr = lnCorr + 1
'            lsEmpNom = PstaNombre(rsT!Nombre)
'            RSet lsEmpTar = rsT!Tardanzas
'            RSet lsEmpPerCom = 0
'            RSet lsEmpPerPer = rsT!Personales
'            RSet lsEmpPerMed = rsT!permitidos
'
'            lnIndice = lnIndice + 1
'            lsCadena = lsCadena & Format(lnCorr, "0000") & " " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'
'            If lnIndice >= 54 Then
'                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'                lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT - " & rsT!cAgeDescripcion, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
'                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;", lnIndice)
'            End If
'            lsCodAgeAnt = rsT!cAgenciaActual
'            lsCodAnt = Left(rsT!cRHCod, 1)
'            rsT.MoveNext
'        Wend
'    End If
'
'    rsT.Close
'    Set rsT = Nothing
'
'    If lsCadena <> "" Then
'        RSet lsEmpTar = lnEmpTar
'        RSet lsEmpPerCom = lnEmpPerCom
'        RSet lsEmpPerPer = lnEmpPerMed
'        RSet lsEmpPerMed = lnEmpPerPer
'        lsEmpCod = ""
'        lsEmpNom = "TOTAL AGENCIA"
'
'        lnEmpTar = 0
'        lnEmpPerCom = 0
'        lnEmpPerMed = 0
'        lnEmpPerPer = 0
'        lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & "     " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
'    End If
'
'
'
'    RSet lsEmpTar = lnEmpTarT
'    RSet lsEmpPerCom = lnEmpPerComT
'    RSet lsEmpPerPer = lnEmpPerMedT
'    RSet lsEmpPerMed = lnEmpPerPerT
'    lsEmpCod = ""
'    lsEmpNom = "TOTAL GENERAL"
'
'    lnEmpTar = 0
'    lnEmpPerCom = 0
'    lnEmpPerMed = 0
'    lnEmpPerPer = 0
'
'    lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & "     " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
'
'    lnEmpTar = 0
'    lnEmpPerCom = 0
'    lnEmpPerMed = 0
'    lnEmpPerPer = 0
'
'    GetTardanzasEmpleadosPorAgencias = lsCadena
'End Function


Public Function GetTardanzasEmpleadosPorAgencias(pdIni As Date, pdFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, psCodagencia As String) As String
    Dim sqlT As String
    Dim rsT As ADODB.Recordset
    Dim lsCadena As String
    Dim lnPagina As Long
    Dim lnIndice As Long
    Dim lsEmpCod As String * 10
    Dim lsEmpNom As String * 40
    Dim lsEmpTar As String * 12
    Dim lsEmpPerCom As String * 12
    Dim lsEmpPerMed As String * 12
    Dim lsEmpPerPer As String * 12
    
    Dim lnEmpTar As Currency
    Dim lnEmpPerCom As Currency
    Dim lnEmpPerMed As Currency
    Dim lnEmpPerPer As Currency
    
    Dim lnEmpTarT As Currency
    Dim lnEmpPerComT As Currency
    Dim lnEmpPerMedT As Currency
    Dim lnEmpPerPerT As Currency
    
    lnEmpTar = 0
    lnEmpPerCom = 0
    lnEmpPerMed = 0
    lnEmpPerPer = 0
    
    lnEmpTarT = 0
    lnEmpPerComT = 0
    lnEmpPerMedT = 0
    lnEmpPerPerT = 0
    
    Set rsT = New ADODB.Recordset
    Dim lnCorr As Long
    Dim lsCodAnt As String
    
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    Dim lsCodAgeAnt As String
    lsCodAgeAnt = ""
    
    'Set rsT = oRep.GetAsistencia(Format(pdIni, gsFormatoFecha), Format(pdFin, gsFormatoFecha), True, psCodagencia)
    Set rsT = oRep.GetAsistencia(Format(pdIni, "mm/dd/yyyy"), Format(pdFin, "mm/dd/yyyy"), True, psCodagencia)
    
    lsCadena = ""
    If Not RSVacio(rsT) Then
        lnCorr = 0
        While Not rsT.EOF
            If lsCodAgeAnt <> rsT!cAgenciaActual Or lsCodAnt <> Left(rsT!cRHCod, 1) Then
                
                If lsCadena <> "" Then
                    RSet lsEmpTar = lnEmpTar
                    RSet lsEmpPerCom = lnEmpPerCom
                    RSet lsEmpPerPer = lnEmpPerMed
                    RSet lsEmpPerMed = lnEmpPerPer
                    lsEmpCod = ""
                    lsEmpNom = "TOTAL AGENCIA"
                    
                    lnEmpTar = 0
                    lnEmpPerCom = 0
                    lnEmpPerMed = 0
                    lnEmpPerPer = 0
                    lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
                    lsCadena = lsCadena & "     " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
                    lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
                    
                    If lsCodAnt <> Left(rsT!cRHCod, 1) Then
                        RSet lsEmpTar = lnEmpTarT
                        RSet lsEmpPerCom = lnEmpPerComT
                        RSet lsEmpPerPer = lnEmpPerMedT
                        RSet lsEmpPerMed = lnEmpPerPerT
                        lsEmpCod = ""
                        lsEmpNom = "TOTAL GENERAL"
                        
                        lnEmpTarT = 0
                        lnEmpPerComT = 0
                        lnEmpPerMedT = 0
                        lnEmpPerPerT = 0
                        
                        lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
                        lsCadena = lsCadena & "     " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
                        lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
                    End If
                    
                    lnEmpTar = 0
                    lnEmpPerCom = 0
                    lnEmpPerMed = 0
                    lnEmpPerPer = 0
                    
                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                End If
                
                lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT - " & rsT!cAgenciaActual & " - " & rsT!cAgeDescripcion, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;", lnIndice)
            End If
        
            lnEmpTar = lnEmpTar + rsT!Tardanzas
            lnEmpPerCom = lnEmpPerCom + 0
            lnEmpPerMed = lnEmpPerMed + rsT!Personales
            lnEmpPerPer = lnEmpPerPer + rsT!permitidos
            
            lnEmpTarT = lnEmpTarT + rsT!Tardanzas
            lnEmpPerComT = lnEmpPerComT + 0
            lnEmpPerMedT = lnEmpPerMedT + rsT!Personales
            lnEmpPerPerT = lnEmpPerPerT + rsT!permitidos
            
            lsEmpCod = rsT!cRHCod
            lnCorr = lnCorr + 1
            lsEmpNom = PstaNombre(rsT!Nombre)
            RSet lsEmpTar = rsT!Tardanzas
            RSet lsEmpPerCom = 0
            RSet lsEmpPerPer = rsT!Personales
            RSet lsEmpPerMed = rsT!permitidos
            
            lnIndice = lnIndice + 1
            lsCadena = lsCadena & Format(lnCorr, "0000") & " " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
            
            If lnIndice >= 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("Asistencia Personal de la CMACT - " & rsT!cAgeDescripcion, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Item;4; ;2;Cod.Emp.;10;Nombre;10; ;30;Tardanza;12;Per.Comis.;12;Per.Medico;12;Per.Person.;12;", lnIndice)
            End If
            lsCodAgeAnt = rsT!cAgenciaActual
            lsCodAnt = Left(rsT!cRHCod, 1)
            rsT.MoveNext
        Wend
    End If

    rsT.Close
    Set rsT = Nothing
    
    If lsCadena <> "" Then
        RSet lsEmpTar = lnEmpTar
        RSet lsEmpPerCom = lnEmpPerCom
        RSet lsEmpPerPer = lnEmpPerMed
        RSet lsEmpPerMed = lnEmpPerPer
        lsEmpCod = ""
        lsEmpNom = "TOTAL AGENCIA"
        
        lnEmpTar = 0
        lnEmpPerCom = 0
        lnEmpPerMed = 0
        lnEmpPerPer = 0
        lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & "     " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
    End If
    
    
    
    RSet lsEmpTar = lnEmpTarT
    RSet lsEmpPerCom = lnEmpPerComT
    RSet lsEmpPerPer = lnEmpPerMedT
    RSet lsEmpPerMed = lnEmpPerPerT
    lsEmpCod = ""
    lsEmpNom = "TOTAL GENERAL"
    
    lnEmpTar = 0
    lnEmpPerCom = 0
    lnEmpPerMed = 0
    lnEmpPerPer = 0
    
    lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "     " & lsEmpCod & lsEmpNom & lsEmpTar & lsEmpPerCom & lsEmpPerMed & lsEmpPerPer & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & String(105, "-") & oImpresora.gPrnSaltoLinea
    
    lnEmpTar = 0
    lnEmpPerCom = 0
    lnEmpPerMed = 0
    lnEmpPerPer = 0
    
    GetTardanzasEmpleadosPorAgencias = lsCadena
End Function


Public Function GetExportaDataTrabajadores() As String
Dim sql As String
Dim sep As String
Dim rs As New ADODB.Recordset
Dim lsCadena As String
Dim oCon As DConecta
Set oCon = New DConecta
Dim sCond As String

sep = "|"

sql = "select"
sql = sql & " (Select top 1 cPersIDTpo from PersID where cPersCod = RH.cPersCod Order by cPersCod) TipDoc,"
sql = sql & " (Select top 1 cPersIDnro from PersID where cPersCod = RH.cPersCod Order by cPersCod) NroDoc,"
sql = sql & " dbo.GetNombreAPMaterno(cPersNombre) ApMaterno,"
sql = sql & " dbo.GetNombreAPPaterno(cPersNombre) ApPaterno,"
sql = sql & " dbo.GetNombrePrimero(cPersNombre) Nombre, dPersNacCreac dFecNac,"
sql = sql & " (Select cPersNatSexo from PersonaNat where cPersCod = RH.cPersCod ) Sexo,"
sql = sql & " isNull(cPersTelefono,'') telefono,"
sql = sql & " dIngreso dFecIngreso, nRHEstado, 'Empleado' TipTrabajador,"
sql = sql & " dCese dFecBaja, '20100210909' RUC,"
sql = sql & " (Select case when cRHEmplAFPPersCod =  NULL then '2' else '1'  end from rhEmpleado where cPersCod = RH.cPersCod) RegPens"
sql = sql & " from rrhh RH"
sql = sql & " Inner Join Persona P on P.cPersCod = RH.cPErsCod"
sql = sql & " where  not substring(convert(varchar(3), nRHEstado),1,1) in ('8','7')"
lsCadena = ""
If oCon.AbreConexion Then
    Set rs = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    If rs.BOF And rs.EOF Then
    Else
        While Not rs.EOF
            lsCadena = lsCadena & rs!TipDoc & sep & Trim(Mid(rs!NroDoc, 1, 15)) & sep & Trim(Mid(rs!ApMaterno, 1, 20)) & sep
            lsCadena = lsCadena & Trim(Mid(rs!ApPaterno, 1, 20)) & sep & Trim(Mid(rs!Nombre, 1, 20)) & sep
            lsCadena = lsCadena & Format(rs!dFecNac, "DD/MM/YYYY") & sep & rs!Sexo & sep & rs!telefono & sep
            lsCadena = lsCadena & Format(rs!dFecIngreso, "DD/MM/YYYY") & sep
            Select Case rs!d
                Case 101, 201, 301, 302, 401: sCond = 10
                Case 701, 702, 703, 801, 802, 803: sCond = 19
            End Select
            lsCadena = lsCadena & sCond & sep & rs!TipTrabajador & sep & rs!RUC
            
            rs.MoveNext
        Wend
    End If
Else
    GetExportaDataTrabajadores = ""
    MsgBox "No se pudo establecer conexion", vbInformation, "AVISO"
    Exit Function
End If
Set oCon = Nothing



End Function

Public Function GetRepPlaPDT(pdIni As Date, pdFin As Date) As String
    
    Dim lsTipDoc As String
    Dim lsNumDoc As String
    Dim lsNumDiasTrab As String
    Dim lsRemIES As String
    Dim lsRemONP As String
    Dim lsRemSaludAgrario As String
    Dim lsRemArtisas As String
    Dim lsRem5taCat As String
    Dim lsTrib5taCat As String
    
    Dim lsCadena As String
    '**ALPA**08/04/2008************************************************************
    Dim lsCodCon As String
    Dim lsMonDev As String
    Dim lsMonPag As String
    Dim lsFile As String
    Dim fso As Scripting.FileSystemObject
    Dim ts As TextStream
    Dim Impresion As String
    '**End*************************************************************************
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    Set rsE = oRep.GetRepPDT("", Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov))
    
    If Not RSVacio(rsE) Then
        RaiseEvent ShowProgress
        While Not rsE.EOF
            lsTipDoc = rsE!TpoDoc & ""
            lsNumDoc = Trim(rsE!cPersIDnro) & ""
'****ALPA 08/04/2008******************************************
'            lsNumDiasTrab = Format(rsE!DLAB, "00")
'            If rsE!TIES = 0 Then
'                lsRemIES = ""
'            Else
'                lsRemIES = Format(rsE!TIES, "0.00")
'            End If
'            If rsE!TONP = 0 Then
'                lsRemONP = ""
'            Else
'                lsRemONP = Format(rsE!TONP, "0.00")
'            End If
'            If rsE!TES = 0 Then
'                lsRemSaludAgrario = 0
'            Else
'                lsRemSaludAgrario = Format(rsE!TES, "0.00")
'            End If
'            If rsE!TART = 0 Then
'                lsRemArtisas = ""
'            Else
'                lsRemArtisas = Format(rsE!TART, "0.00")
'            End If
'            If rsE!T5ta = 0 Then
'                lsRem5taCat = 0
'            Else
'                lsRem5taCat = Format(rsE!T5ta, "0.00")
'            End If
'            If rsE!I5ta = 0 Then
'                lsTrib5taCat = 0
'            Else
'                lsTrib5taCat = Format(rsE!I5ta, "0.00")
'            End If
            If rsE!cRHConceptoCodSUNAT = "" Then
                lsCodCon = ""
            Else
                lsCodCon = Format(rsE!cRHConceptoCodSUNAT, "0000")
            End If
            If rsE!MontoDeven = 0 Then
                lsMonDev = "0"
            Else
                lsMonDev = Format(rsE!MontoDeven, "0.00")
            End If
            If rsE!MontoPagado = 0 Then
                lsMonPag = "0"
            Else
                lsMonPag = Format(rsE!MontoPagado, "0.00")
            End If
            'lsCadena = lsCadena & lsRUC & lsCUSSP & lsAPaterno & lsAMaterno & lsANombres & lsRemunera & lsCodAFP & lsTipoMov & lsFecMov & oImpresora.gPrnSaltoLinea
            'lsCadena = lsCadena & "01" & "|" & lsNumDoc & "|" & lsNumDiasTrab & "|" & lsRemIES & "|" & lsRemONP & "|" & Format(lsRemSaludAgrario, "0.00") & "|" & lsRemArtisas & "|" & Format(lsRem5taCat, "0.00") & "|" & Format(lsTrib5taCat, "0.00") & "|" & Chr(10)
            lsCadena = lsCadena & lsTipDoc & "|" & lsNumDoc & "|" & lsCodCon & "|" & lsMonDev & "|" & lsMonPag & "|" & Chr(10)
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        RaiseEvent CloseProgress
    Else
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
    '**ALPA*****08/04/2008*********************************************************
    Impresion = ""
    lsFile = App.Path & "\..\Spooler\0601" & Format(CDate(pdIni), "yyyymm") & "20103845328.rem"
        Set fso = New Scripting.FileSystemObject
        If fso.FileExists(lsFile) Then
            If MsgBox("El archivo ya existe, desea reemplazarlo", vbYesNo + vbInformation, "Aviso") = vbNo Then
                Set fso = Nothing
                Impresion = "1"
            End If
        End If
    If Impresion = "" Then
     Set ts = fso.CreateTextFile(lsFile, True)
        ts.Write (lsCadena)
        MsgBox "El archivo se generó satisfactoriamente", vbInformation, "Aviso"
        ts.Close
    Set fso = Nothing
    End If
    '**End*-**************************************************************************
    GetRepPlaPDT = lsCadena
    rsE.Close
    Set rsE = Nothing
End Function
 '**ALPA*****08/04/2008****d*****************************************************
Public Function GetEstabPDT(pdIni As Date, pdFin As Date) As String
    
    Dim lsTipDoc As String
    Dim lsNumDoc As String
    
    Dim lsCadena As String
    
    Dim lsnRUC As String
    Dim lssCodEst As String
    Dim lsnTasa As String
    Dim lsFile As String
    Dim fso As Scripting.FileSystemObject
    Dim ts As TextStream
    Dim Impresion As String
    
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    Set rsE = oRep.GetEstabPDT("", Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov))
    
    If Not RSVacio(rsE) Then
        RaiseEvent ShowProgress
        While Not rsE.EOF
            lsTipDoc = rsE!TpoDoc & ""
            lsNumDoc = Trim(rsE!cPersIDnro) & ""

            If rsE!sRUCEmpresa = "" Then
                lsnRUC = ""
            Else
                lsnRUC = Format(rsE!sRUCEmpresa, "0")
            End If
            If rsE!cAgeCod = 0 Then
                lssCodEst = "00000000000"
            Else
                lssCodEst = Format(rsE!cAgeCod, "0")
            End If
            If rsE!nTasa = 0 Then
                lsnTasa = ""
            Else
                lsnTasa = Format(rsE!nTasa, "0.00")
            End If
            lsCadena = lsCadena & "01" & "|" & lsNumDoc & "|" & lsnRUC & "|" & lssCodEst & "|" & lsnTasa & "|" & Chr(10)
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        RaiseEvent CloseProgress
    Else
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
   
    Impresion = ""
    lsFile = App.Path & "\..\Spooler\0601" & Format(CDate(pdIni), "yyyymm") & "20103845328.tes"
        Set fso = New Scripting.FileSystemObject
        If fso.FileExists(lsFile) Then
            If MsgBox("El archivo ya existe, desea reemplazarlo", vbYesNo + vbInformation, "Aviso") = vbNo Then
                Set fso = Nothing
                Impresion = "1"
            End If
        End If
    If Impresion = "" Then
     Set ts = fso.CreateTextFile(lsFile, True)
        ts.Write (lsCadena)
        MsgBox "El archivo se generó satisfactoriamente", vbInformation, "Aviso"
        ts.Close
    Set fso = Nothing
    End If
    GetEstabPDT = lsCadena
    rsE.Close
    Set rsE = Nothing
End Function
Public Function GetJorLabPDT(pdIni As Date, pdFin As Date) As String
    
    Dim lsTipDoc As String
    Dim lsNumDoc As String
    
    Dim lsCadena As String
    
    Dim lsNdiasTra As String
    Dim lsNHoraTra As String
    Dim lsNHoraTrS As String
    
    Dim lsNMinuTra As String
    Dim lsNMinuTrS As String
    
    Dim lsFile As String
    Dim fso As Scripting.FileSystemObject
    Dim ts As TextStream
    Dim Impresion As String
    
    Dim sqlE As String
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    
    Set rsE = oRep.GetJorLabPDT("", Format(pdIni, gcFormatoMov), Format(pdFin, gcFormatoMov))
    
    If Not RSVacio(rsE) Then
        RaiseEvent ShowProgress
        While Not rsE.EOF
            lsTipDoc = rsE!TpoDoc & ""
            lsNumDoc = Trim(rsE!cPersIDnro) & ""

            If rsE!DiasTraba = "" Then
                lsNdiasTra = ""
            Else
                lsNdiasTra = Format(rsE!DiasTraba, "0")
            End If
            If rsE!HorasTraba = 0 Then
                lsNHoraTra = "0"
            Else
                lsNHoraTra = Format(rsE!HorasTraba, "0")
            End If
            If rsE!HorasExtras = 0 Then
                lsNHoraTrS = "0"
            Else
                lsNHoraTrS = Format(rsE!HorasExtras, "0")
            End If
            'ALPA 20080813****************************************
            If rsE!MinutosTraba = 0 Then
                lsNMinuTra = "0"
            Else
                lsNMinuTra = Format(rsE!MinutosTraba, "0")
            End If
            If rsE!MinutosExtras = 0 Then
                lsNMinuTrS = "0"
            Else
                lsNMinuTrS = Format(rsE!MinutosExtras, "0")
            End If
            '******************************************************
            'lsCadena = lsCadena & "01" & "|" & lsNumDoc & "|" & lsNdiasTra & "|" & lsNHoraTra & "|" & lsNMinuTra & "|" & lsNHoraTrS & "|" & lsNMinuTrS & "|" & Chr(10)
            lsCadena = lsCadena & "01" & "|" & lsNumDoc & "|" & lsNHoraTra & "|" & lsNMinuTra & "|" & lsNHoraTrS & "|" & lsNMinuTrS & "|" & Chr(10)
            RaiseEvent Progress(rsE.Bookmark, rsE.RecordCount)
            rsE.MoveNext
        Wend
        RaiseEvent CloseProgress
    Else
        MsgBox "No existen datos.", vbInformation, "Aviso"
    End If
   
    Impresion = ""
    lsFile = App.Path & "\..\Spooler\0601" & Format(CDate(pdIni), "yyyymm") & "20103845328.jor"
        Set fso = New Scripting.FileSystemObject
        If fso.FileExists(lsFile) Then
            If MsgBox("El archivo ya existe, desea reemplazarlo", vbYesNo + vbInformation, "Aviso") = vbNo Then
                Set fso = Nothing
                Impresion = "1"
            End If
        End If
    If Impresion = "" Then
     Set ts = fso.CreateTextFile(lsFile, True)
        ts.Write (lsCadena)
        MsgBox "El archivo se generó satisfactoriamente", vbInformation, "Aviso"
        ts.Close
    Set fso = Nothing
    End If
    GetJorLabPDT = lsCadena
    rsE.Close
    Set rsE = Nothing
End Function
'**End*************************************************************************
'*****************************************************************************
' Modulo de Impuestos RRHH
'*****************************************************************************
    'tablas pa el reporte de Impuesto de RRHH
    'RHBasePDT
    'RHBasePDTDetalle
    'RHPDT
    'RHVacacionesDevengadas
'*****************************************************************************

Private Function CargarBasePDT(ByVal pcPeriodo As String) As ADODB.Recordset
On Error GoTo CargarBasePDTErr
    Dim oCon As DConecta, sSql As String, rs As ADODB.Recordset, cPeriodoAnterior As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        cPeriodoAnterior = Format(DateAdd("m", -1, pcPeriodo), "yyyymm")
        pcPeriodo = Format(pcPeriodo, "yyyymm")
        sSql = "select b.cRHCod, cPersNombre=replace(p.cPersNombre,'/',' '), cPeriodo=isnull(x.cPeriodo,''), " & _
               " nRemuneracion = (case when x.bsubsidio=0 then isnull(x.nRemuneracion,0) else 0 end), " & _
               " nImpuesto = (case when x.bsubsidio=0 then isnull(x.nImpuesto,0) else 0 end ), " & _
               " Dias = (case when x.bsubsidio=0 then isnull(x.Dias,0) else 0 end ), " & _
               " cTipo=isnull(x.cTipo,'N'), " & _
               " nRemA=b.nRemuneracion, nImpA=b.nImpuesto, nDiaA=b.nDia " & _
               " from RHBasePDT b " & _
               " inner join RRHH r on b.cRHCod = r.cRHCod " & _
               " inner join Persona p on r.cPersCod = p.cPersCod " & _
               " left join RHBasePDTDetalle x on b.cRHCod = x.cRHCod and b.cPeriodo = x.cPeriodo " & _
               " where b.cPeriodo = '" & cPeriodoAnterior & "' order by cPersNombre"
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
    End If
    If Not rs.EOF Then
        Set CargarBasePDT = rs
    End If
    Exit Function
CargarBasePDTErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Function GuardarBasePDT(ByVal pcPeriodo As String, ByVal pnDias As Currency, ByVal pnRemuneracion As Currency, ByVal pnImpuesto As Currency, ByVal pcRHCod As String) As ADODB.Recordset
On Error GoTo GuardarBasePDTErr
    Dim oCon As DConecta, sSql As String, rs As ADODB.Recordset ', cPeriodoAnterior As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        pcPeriodo = Format(pcPeriodo, "yyyymm")
        sSql = " declare @tmp int"
        sSql = sSql & " set @tmp = (select count(*) from RHBasePDT where cPeriodo = '" & pcPeriodo & "' And cRHCod= '" & pcRHCod & "') "
        sSql = sSql & " if @tmp = 0 "
        sSql = sSql & " insert into RHBasePDT(cRHCod,cPeriodo,nRemuneracion,nImpuesto,nDia) "
        sSql = sSql & " values('" & pcRHCod & "','" & pcPeriodo & "'," & pnRemuneracion & "," & pnImpuesto & "," & pnDias & ") "
        sSql = sSql & " else "
        sSql = sSql & " update RHBasePDT set nRemuneracion = " & pnRemuneracion & ",nImpuesto= " & pnImpuesto & ",nDia = " & pnDias & " " & _
                      " where cRHCod = '" & pcRHCod & "' and cPeriodo = '" & pcPeriodo & "'"
        oCon.Ejecutar (sSql)
        oCon.CierraConexion
    End If
    Exit Function
GuardarBasePDTErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Function CargarDiferencia(ByVal pcPeriodo As String) As ADODB.Recordset
    Dim oCon As DConecta, sSql As String, rs As ADODB.Recordset
    On Error GoTo CargarDiferenciaErr
    sSql = "select d.cRHCod, cPersNombre=replace(p.cPersNombre,'/',' '),  " & _
           " nRemuneracionV=d.nRemuneracion, nImpuestoV=d.nImpuesto, DiasV=d.Dias, " & _
           " nRemuneracion=isnull(x.nRemuneracion,0) , nImpuesto=isnull(x.nImpuesto,0), isnull(x.nDia,0) " & _
           " from RHBasePDTDetalle d " & _
           " inner join RRHH r on d.cRHCod = r.cRHCod " & _
           " inner join Persona p on r.cPersCod = p.cPersCod " & _
           " left outer join RHBasePDT x on d.cPeriodo = x.cPeriodo and d.cRHCod = x.cRHCod " & _
           " where d.cPeriodo = '" & pcPeriodo & "' and d.cTipo = 'V' order by d.cRHCod --cPersNombre "
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSql)
        Set CargarDiferencia = rs
        oCon.CierraConexion
    End If
Exit Function
CargarDiferenciaErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Private Sub GuardarVacacionesDevengadas(ByVal pcPeriodo As String, ByVal pcRHCod As String, ByVal pnImpuesto As Currency, ByVal pnRemuneracion As Currency)
On Error GoTo GuardarVacacionesDevengadasErr
    Dim oCon As DConecta, sSql As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSql = " declare @tmp int "
        sSql = sSql & " set @tmp=(select count(*) from RHVacacionesDevengadas where cRHCod = '" & pcRHCod & "' and cPeriodo= '" & Format(pcPeriodo, "yyyymm") & "') "
        sSql = sSql & " if @tmp=0 "
        sSql = sSql & " insert into RHVacacionesDevengadas(cRHCod, cPeriodo, nRemuneracion, nImpuesto) " & _
               " values('" & pcRHCod & "','" & Format(pcPeriodo, "yyyymm") & "'," & pnRemuneracion & "," & pnImpuesto & ")"
        oCon.Ejecutar sSql
        oCon.CierraConexion
    End If
    Exit Sub
GuardarVacacionesDevengadasErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Private Function CargarPDT(ByVal pcPeriodo As String) As ADODB.Recordset
On Error GoTo CargarBasePDTErr
    Dim oCon As DConecta, sSql As String, rs As ADODB.Recordset, cPeriodoAnterior As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSql = "select p.cRHCod, cPersNombre=replace(x.cPersNombre,'/',' '), p.nTIM, p.nEsSalud, p.nRem5ta, p.nEsSaludVida,  p.nImp5ta, " & _
               " nprovRem = d.nRemuneracion, nprovImp = d.nImpuesto, nDevRem = v.nRemuneracion, nDevImp = v.nImpuesto " & _
               " from RHPDT p " & _
               " inner join RRHH r on p.cRHCod = r.cRHCod " & _
               " inner join Persona x on r.cPersCod = x.cPersCod " & _
               " left outer join RHBasePDTDetalle d on p.cRHCod = d.cRHCod and p.cPeriodo = d.cPeriodo and d.cTipo = 'P' " & _
               " left outer join RHVacacionesDevengadas v on p.cRHCod = v.cRHCod and p.cPeriodo = v.cPeriodo " & _
               " where p.cPeriodo = '" & pcPeriodo & "' order by cPersNombre --p.cRHCod "
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
    End If
'    If Not Rs.EOF Then
        Set CargarPDT = rs
'    End If
    Exit Function
CargarBasePDTErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

'*************************************************************************************
'*************************************************************************************

Public Sub ImprimeReporteImpuesto(ByVal pcPeriodo As String, ByVal pdFecha As Date)
On Error GoTo ImprimeReporteImpuesto5taErr
    
    Screen.MousePointer = 11
    ImprimeBasePDT pcPeriodo, pdFecha
    ImprimeDiferenciaEjecutadasDevengadas pcPeriodo, pdFecha
    ImprimePDT pcPeriodo, pdFecha
    Screen.MousePointer = 0
    Exit Sub
ImprimeReporteImpuesto5taErr:
    Screen.MousePointer = 0
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

'*************************************************************************************
'*************************************************************************************

Private Sub ImprimeBasePDT(ByVal psPeriodo As String, ByVal pdFecha As Date)
    On Error GoTo ImprimeBasePDTErr

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1(100) As Excel.Worksheet
    Dim lbExisteHoja  As Boolean
    
    Dim cBSGrupoCod As String
    Dim cBSGrupoDescripcion As String

    Dim lilineas As Integer
    Dim Total As Currency, TotalPos As Currency
    Dim nFactor As Integer
    Dim nCol As Integer, I As Integer, nLetra As Integer
    
    Dim CadTit As String
    Dim CadImp   As String, Ganador As Integer
    Dim rs As New ADODB.Recordset, RsAnterior As ADODB.Recordset
    Dim nNroHoja As Integer, nProSelItem As Integer, nMontoItem As Currency
    Dim lsNomHoja As String, cPeriodoAnterior As String, cRHCod As String

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
        
    RaiseEvent ShowProgress
        
    cPeriodoAnterior = Format(DateAdd("m", -1, psPeriodo), "mmm-yy")
    
    nNroHoja = 1
        
    Set xlHoja1(nNroHoja) = xlLibro.Worksheets.Item(nNroHoja)
    
    xlHoja1(nNroHoja).Name = "Base_PDT_" & Format(psPeriodo, "mmm_yy")
    '**************************************************************************

'         xlAplicacion.Range("A1:A1").ColumnWidth = 5
'         xlAplicacion.Range("B1:B1").ColumnWidth = 5
'         xlAplicacion.Range("C1:C1").ColumnWidth = 20
     
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 1), xlHoja1(nNroHoja).Cells(1, 1)).ColumnWidth = 5
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 2), xlHoja1(nNroHoja).Cells(1, 2)).ColumnWidth = 5
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 3), xlHoja1(nNroHoja).Cells(1, 3)).ColumnWidth = 20
     
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 15)).Font.Bold = True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 4)).Merge True
     'xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(3, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 15)).Merge True
     'xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(5, 2), xlHoja1(nNroHoja).Cells(5, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(5, 15)).Font.Bold = True

'******************************************************
            
    lilineas = 7

    Screen.MousePointer = 11
    xlHoja1(nNroHoja).Cells(2, 2) = "Caja Municipal de Ahorros y Creditos de Trujillo SA"
    
    xlHoja1(nNroHoja).Cells(4, 2) = "Base para el PDT para el Periodo " & Format(psPeriodo, "mmmm - yyyy")
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 15)).Merge True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 14), xlHoja1(nNroHoja).Cells(2, 15)).Merge True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 14), xlHoja1(nNroHoja).Cells(2, 14)).HorizontalAlignment = xlRight
    xlHoja1(nNroHoja).Cells(2, 14) = "Fecha " & pdFecha
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 15), xlHoja1(nNroHoja).Cells(3, 15)).HorizontalAlignment = xlRight
    xlHoja1(nNroHoja).Cells(3, 15) = gsCodUser
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 14), xlHoja1(nNroHoja).Cells(3, 15)).Merge True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 2)).HorizontalAlignment = xlCenter
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas + 1, 15)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas + 1, 15)).HorizontalAlignment = xlCenter
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas + 1, 15)).Cells.Interior.Color = &HC0FFFF
    
    xlHoja1(nNroHoja).Cells(lilineas, 8) = "Provicion " & cPeriodoAnterior
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 4), xlHoja1(nNroHoja).Cells(lilineas, 6)).Merge True
    xlHoja1(nNroHoja).Cells(lilineas, 10) = "Vacacione Ejecutadas " & cPeriodoAnterior
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 7), xlHoja1(nNroHoja).Cells(lilineas, 9)).Merge True
    xlHoja1(nNroHoja).Cells(lilineas, 4) = "Total Anterior A " & cPeriodoAnterior
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 10), xlHoja1(nNroHoja).Cells(lilineas, 12)).Merge True
    xlHoja1(nNroHoja).Cells(lilineas, 14) = "Saldo " & Format(psPeriodo, "mmm-yy")
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 13), xlHoja1(nNroHoja).Cells(lilineas, 15)).Merge True
    lilineas = lilineas + 1
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas - 1, 2), xlHoja1(nNroHoja).Cells(lilineas, 2)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas - 1, 2) = "#"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas - 1, 3), xlHoja1(nNroHoja).Cells(lilineas, 3)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas - 1, 3) = "Nombre"
    xlHoja1(nNroHoja).Cells(lilineas, 4) = "Dias"
    xlHoja1(nNroHoja).Cells(lilineas, 5) = "Remuneracion"
    xlHoja1(nNroHoja).Cells(lilineas, 6) = "Impuesto"
    xlHoja1(nNroHoja).Cells(lilineas, 7) = "Dias"
    xlHoja1(nNroHoja).Cells(lilineas, 8) = "Remuneracion"
    xlHoja1(nNroHoja).Cells(lilineas, 9) = "Impuesto"
    xlHoja1(nNroHoja).Cells(lilineas, 10) = "Dias"
    xlHoja1(nNroHoja).Cells(lilineas, 11) = "Remuneracion"
    xlHoja1(nNroHoja).Cells(lilineas, 12) = "Impuesto"
    xlHoja1(nNroHoja).Cells(lilineas, 13) = "Dias"
    xlHoja1(nNroHoja).Cells(lilineas, 14) = "Remuneracion"
    xlHoja1(nNroHoja).Cells(lilineas, 15) = "Impuesto"
        
    Set rs = CargarBasePDT(psPeriodo)
'    Set RsAnterior = CargarBasePDTAnterior(Format(psPeriodo, "yyyymm"))
    RaiseEvent Progress(rs.AbsolutePosition, rs.RecordCount)
    lilineas = lilineas + 1
    
    Do While Not rs.EOF
        If cRHCod <> rs!cRHCod Then
            cRHCod = rs!cRHCod
            lilineas = lilineas + 1
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 13), xlHoja1(nNroHoja).Cells(lilineas, 13)).Formula = "=" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 4), xlHoja1(nNroHoja).Cells(lilineas, 4)).Address(False, False) & "+" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 7), xlHoja1(nNroHoja).Cells(lilineas, 7)).Address(False, False) & "-" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 10), xlHoja1(nNroHoja).Cells(lilineas, 10)).Address(False, False)
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 14), xlHoja1(nNroHoja).Cells(lilineas, 14)).Formula = "=" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 5), xlHoja1(nNroHoja).Cells(lilineas, 5)).Address(False, False) & "+" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 8), xlHoja1(nNroHoja).Cells(lilineas, 8)).Address(False, False) & "-" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 11), xlHoja1(nNroHoja).Cells(lilineas, 11)).Address(False, False)
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 15), xlHoja1(nNroHoja).Cells(lilineas, 15)).Formula = "=" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 6), xlHoja1(nNroHoja).Cells(lilineas, 6)).Address(False, False) & "+" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 9), xlHoja1(nNroHoja).Cells(lilineas, 9)).Address(False, False) & "-" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 12), xlHoja1(nNroHoja).Cells(lilineas, 12)).Address(False, False)
            xlHoja1(nNroHoja).Cells(lilineas, 2) = rs!cRHCod
            xlHoja1(nNroHoja).Cells(lilineas, 3) = rs!cPersNombre
            xlHoja1(nNroHoja).Cells(lilineas, 4) = rs!nDiaA
            xlHoja1(nNroHoja).Cells(lilineas, 5) = rs!nRemA
            xlHoja1(nNroHoja).Cells(lilineas, 6) = rs!nImpA
        End If
        Select Case rs!cTipo
            Case "P"
                xlHoja1(nNroHoja).Cells(lilineas, 7) = rs!Dias
                xlHoja1(nNroHoja).Cells(lilineas, 8) = rs!nRemuneracion
                xlHoja1(nNroHoja).Cells(lilineas, 9) = rs!nImpuesto
'                lilineas = lilineas + 1
            Case "V"
                xlHoja1(nNroHoja).Cells(lilineas, 10) = rs!Dias
                xlHoja1(nNroHoja).Cells(lilineas, 11) = rs!nRemuneracion
                xlHoja1(nNroHoja).Cells(lilineas, 12) = rs!nImpuesto
'                lilineas = lilineas + 1
'            Case "N"
'                xlHoja1(nNroHoja).Cells(lilineas, 7) = Rs!Dias
'                xlHoja1(nNroHoja).Cells(lilineas, 8) = Rs!nRemuneracion
'                xlHoja1(nNroHoja).Cells(lilineas, 9) = Rs!nImpuesto
'                xlHoja1(nNroHoja).Cells(lilineas, 10) = Rs!Dias
'                xlHoja1(nNroHoja).Cells(lilineas, 11) = Rs!nRemuneracion
'                xlHoja1(nNroHoja).Cells(lilineas, 12) = Rs!nImpuesto
'                lilineas = lilineas + 1
        End Select
        RaiseEvent Progress(rs.AbsolutePosition, rs.RecordCount * 2)
        rs.MoveNext
    Loop
        
    'ExcelCuadro xlHoja1(nNroHoja), 2, 7, 15, lilineas, True, True
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 4), xlHoja1(nNroHoja).Cells(lilineas, 6)).Cells.Interior.Color = &HFCFFE1
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 7), xlHoja1(nNroHoja).Cells(lilineas, 9)).Cells.Interior.Color = &HC0C0C0
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 10), xlHoja1(nNroHoja).Cells(lilineas, 12)).Cells.Interior.Color = &HD8DDC4
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 13), xlHoja1(nNroHoja).Cells(lilineas, 15)).Cells.Interior.Color = &HEAD9E1
            
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 4), xlHoja1(nNroHoja).Cells(lilineas, 4)).NumberFormat = "###,###"
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 5), xlHoja1(nNroHoja).Cells(lilineas, 6)).NumberFormat = "###,###.00"
'
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 7), xlHoja1(nNroHoja).Cells(lilineas, 7)).NumberFormat = "###,###"
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 8), xlHoja1(nNroHoja).Cells(lilineas, 9)).NumberFormat = "###,###.00"
'
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 10), xlHoja1(nNroHoja).Cells(lilineas, 10)).NumberFormat = "###,###"
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 11), xlHoja1(nNroHoja).Cells(lilineas, 12)).NumberFormat = "###,###.00"
'
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 13), xlHoja1(nNroHoja).Cells(lilineas, 13)).NumberFormat = "###,###"
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 14), xlHoja1(nNroHoja).Cells(lilineas, 15)).NumberFormat = "###,###.00"
    
    I = 10
    Do While I <= lilineas
        GuardarBasePDT psPeriodo, xlHoja1(nNroHoja).Cells(I, 13), xlHoja1(nNroHoja).Cells(I, 14), xlHoja1(nNroHoja).Cells(I, 15), xlHoja1(nNroHoja).Cells(I, 2)
        RaiseEvent Progress(I + lilineas, lilineas * 2)
        I = I + 1
    Loop
    
    Screen.MousePointer = 0
'******************************************************
    RaiseEvent CloseProgress
'        xlHoja1(nNroHoja).Cells.Select
    xlHoja1(nNroHoja).Cells.NumberFormat = "###,###,##0.00"
    xlHoja1(nNroHoja).Cells.Font.Size = 8
    xlHoja1(nNroHoja).Cells.EntireColumn.AutoFit

    'Libera los objetos.
    Set xlHoja1(nNroHoja) = Nothing
        
    xlLibro.SaveAs App.Path & "\spooler\BasePDT_" & Format(psPeriodo, "mmmyy")
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
'    xlAplicacion.Windows(1).Visible = True
    Set xlAplicacion = Nothing
    Exit Sub
ImprimeBasePDTErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Err
End Sub

Private Sub ImprimeDiferenciaEjecutadasDevengadas(ByVal psPeriodo As String, ByVal pdFecha As Date)
    On Error GoTo ImprimeBasePDTErr

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1(100) As Excel.Worksheet
    Dim lbExisteHoja  As Boolean
    
    Dim cBSGrupoCod As String
    Dim cBSGrupoDescripcion As String

    Dim lilineas As Integer
    Dim Total As Currency, TotalPos As Currency
    Dim nFactor As Integer
    Dim nCol As Integer, I As Integer, nLetra As Integer
    
    Dim CadTit As String
    Dim CadImp   As String, Ganador As Integer
    Dim rs As New ADODB.Recordset, RsAnterior As ADODB.Recordset
    Dim nNroHoja As Integer, nProSelItem As Integer, nMontoItem As Currency
    Dim lsNomHoja As String, cPeriodoAnterior As String, cRHCod As String
    Dim nDevenRem As Currency, nDevenImp As Currency

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    
    RaiseEvent ShowProgress
    
    cPeriodoAnterior = Format(DateAdd("m", -1, psPeriodo), "mmm-yy")
    
    nNroHoja = 1
        
    Set xlHoja1(nNroHoja) = xlLibro.Worksheets.Item(nNroHoja)
    
    xlHoja1(nNroHoja).Name = "Diferencia_" & Format(psPeriodo, "mmm_yy")
    '**************************************************************************

'         xlAplicacion.Range("A1:A1").ColumnWidth = 5
'         xlAplicacion.Range("B1:B1").ColumnWidth = 5
'         xlAplicacion.Range("C1:C1").ColumnWidth = 20
     
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 1), xlHoja1(nNroHoja).Cells(1, 1)).ColumnWidth = 5
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 2), xlHoja1(nNroHoja).Cells(1, 2)).ColumnWidth = 5
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(1, 3), xlHoja1(nNroHoja).Cells(1, 3)).ColumnWidth = 20
     
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 6)).Font.Bold = True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 4)).Merge True
     'xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(3, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 4)).Merge True
'     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(5, 2), xlHoja1(nNroHoja).Cells(5, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(5, 7)).Font.Bold = True

'******************************************************
            
    lilineas = 6

    Screen.MousePointer = 11
    xlHoja1(nNroHoja).Cells(2, 2) = "Caja Municipal de Ahorros y Creditos de Trujillo SA"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(3, 10)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 9), xlHoja1(nNroHoja).Cells(2, 9)).HorizontalAlignment = xlRight
    xlHoja1(nNroHoja).Cells(2, 10) = "Fecha " & pdFecha
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 10), xlHoja1(nNroHoja).Cells(3, 10)).HorizontalAlignment = xlRight
    xlHoja1(nNroHoja).Cells(3, 10) = gsCodUser
    xlHoja1(nNroHoja).Cells(4, 2) = " DIFERENCIA ENTRE VACACIONES EJECUTADAS Y DEVENGADAS " '& Format(psPeriodo, "mmmm - yyyy")
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 10)).Merge True
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 2)).HorizontalAlignment = xlCenter
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 2), xlHoja1(nNroHoja).Cells(7, 10)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas + 1, 10)).HorizontalAlignment = xlCenter
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas + 1, 10)).Cells.Interior.Color = &HC0FFFF
    
    xlHoja1(nNroHoja).Cells(lilineas, 5) = " Vac. Ejecutadas " ' & cPeriodoAnterior
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 5), xlHoja1(nNroHoja).Cells(lilineas, 6)).Merge True
    xlHoja1(nNroHoja).Cells(lilineas, 7) = " Vac. Devengadas " ' & cPeriodoAnterior
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 7), xlHoja1(nNroHoja).Cells(lilineas, 8)).Merge True
    xlHoja1(nNroHoja).Cells(lilineas, 9) = " Diferencia " ' & Format(psPeriodo, "mmm-yy")
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 9), xlHoja1(nNroHoja).Cells(lilineas, 10)).Merge True
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas + 1, 2)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 2) = "#"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 3), xlHoja1(nNroHoja).Cells(lilineas + 1, 3)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 3) = "Nombre"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 4), xlHoja1(nNroHoja).Cells(lilineas + 1, 4)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 4) = "Dias"
    
    lilineas = lilineas + 1
    
    xlHoja1(nNroHoja).Cells(lilineas, 5) = "Remuneracion"
    xlHoja1(nNroHoja).Cells(lilineas, 6) = "Impuesto"
    xlHoja1(nNroHoja).Cells(lilineas, 7) = "Remuneracion"
    xlHoja1(nNroHoja).Cells(lilineas, 8) = "Impuesto"
    xlHoja1(nNroHoja).Cells(lilineas, 9) = "Remuneracion"
    xlHoja1(nNroHoja).Cells(lilineas, 10) = "Impuesto"
        
'    Set Rs = CargarDiferencia(Format(DateAdd("m", -1, psPeriodo), "yyyymm"))
    Set rs = CargarDiferencia(Format(psPeriodo, "yyyymm"))
    'lilineas = lilineas + 1
    RaiseEvent Progress(rs.AbsolutePosition, rs.RecordCount)
    Do While Not rs.EOF
        If cRHCod <> rs!cRHCod Then
            cRHCod = rs!cRHCod
            lilineas = lilineas + 1
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 9), xlHoja1(nNroHoja).Cells(lilineas, 9)).Formula = "=" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 5), xlHoja1(nNroHoja).Cells(lilineas, 5)).Address(False, False) & "-" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 7), xlHoja1(nNroHoja).Cells(lilineas, 7)).Address(False, False)
            xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 10), xlHoja1(nNroHoja).Cells(lilineas, 10)).Formula = "=" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 6), xlHoja1(nNroHoja).Cells(lilineas, 6)).Address(False, False) & "-" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 8), xlHoja1(nNroHoja).Cells(lilineas, 8)).Address(False, False)
            xlHoja1(nNroHoja).Cells(lilineas, 2) = rs!cRHCod
            xlHoja1(nNroHoja).Cells(lilineas, 3) = rs!cPersNombre
            xlHoja1(nNroHoja).Cells(lilineas, 4) = rs!DiasV
'            xlHoja1(nNroHoja).Cells(lilineas, 5) = "0"
'            xlHoja1(nNroHoja).Cells(lilineas, 6) = "0"
'            xlHoja1(nNroHoja).Cells(lilineas, 7) = "0"
'            xlHoja1(nNroHoja).Cells(lilineas, 8) = "0"
        End If
        'Select Case Rs!cTipo
        '    Case "V"
                xlHoja1(nNroHoja).Cells(lilineas, 5) = rs!nRemuneracionV
                xlHoja1(nNroHoja).Cells(lilineas, 6) = rs!nImpuestoV
            'Case "P"
                xlHoja1(nNroHoja).Cells(lilineas, 7) = IIf(rs!nRemuneracion < 0, 0, rs!nRemuneracion)
                xlHoja1(nNroHoja).Cells(lilineas, 8) = IIf(rs!nImpuesto < 0, 0, rs!nImpuesto)
        'End Select
                                
        If xlHoja1(nNroHoja).Cells(lilineas, 9) >= 0 Then
            nDevenRem = xlHoja1(nNroHoja).Cells(lilineas, 7)
        Else
            nDevenRem = xlHoja1(nNroHoja).Cells(lilineas, 5)
        End If
        
        If xlHoja1(nNroHoja).Cells(lilineas, 10) >= 0 Then
            nDevenImp = xlHoja1(nNroHoja).Cells(lilineas, 8)
        Else
            nDevenImp = xlHoja1(nNroHoja).Cells(lilineas, 6)
        End If
        
        GuardarVacacionesDevengadas psPeriodo, xlHoja1(nNroHoja).Cells(lilineas, 2), nDevenImp, nDevenRem
        RaiseEvent Progress(rs.AbsolutePosition, rs.RecordCount)
        rs.MoveNext
    Loop
        
    'ExcelCuadro xlHoja1(nNroHoja), 2, 6, 10, lilineas, True, True
    
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 4), xlHoja1(nNroHoja).Cells(lilineas, 6)).Cells.Interior.Color = &HFCFFE1
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 7), xlHoja1(nNroHoja).Cells(lilineas, 9)).Cells.Interior.Color = &HC0C0C0
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 10), xlHoja1(nNroHoja).Cells(lilineas, 12)).Cells.Interior.Color = &HD8DDC4
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(10, 13), xlHoja1(nNroHoja).Cells(lilineas, 15)).Cells.Interior.Color = &HEAD9E1
    
    RaiseEvent CloseProgress
    
    Screen.MousePointer = 0
'******************************************************

'    xlHoja1(nNroHoja).Cells.Select
    xlHoja1(nNroHoja).Cells.NumberFormat = "###,###,##0.00"
'    xlHoja1(nNroHoja).Cells.NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(*  - ??_);_(@_)"
    xlHoja1(nNroHoja).Cells.Font.Size = 8
    xlHoja1(nNroHoja).Cells.EntireColumn.AutoFit

    'Libera los objetos.
    Set xlHoja1(nNroHoja) = Nothing
        
    xlLibro.SaveAs App.Path & "\spooler\Diferencia_" & Format(psPeriodo, "mmmyy")
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
'    xlAplicacion.Windows(1).Visible = True
    Set xlAplicacion = Nothing
    Exit Sub
ImprimeBasePDTErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Err
End Sub

Private Sub ImprimePDT(ByVal psPeriodo As String, ByVal pdFecha As Date)
    On Error GoTo ImprimePDTErr

    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1(100) As Excel.Worksheet
    Dim lbExisteHoja  As Boolean
    
    Dim cBSGrupoCod As String
    Dim cBSGrupoDescripcion As String

    Dim lilineas As Integer
    Dim Total As Currency, TotalPos As Currency
    Dim nFactor As Integer
    Dim nCol As Integer, I As Integer, nLetra As Integer
    
    Dim CadTit As String
    Dim CadImp   As String, Ganador As Integer
    Dim rs As New ADODB.Recordset, RsAnterior As ADODB.Recordset
    Dim nNroHoja As Integer, nProSelItem As Integer, nMontoItem As Currency
    Dim lsNomHoja As String, cPeriodoAnterior As String, cRHCod As String

    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Add
    
    RaiseEvent ShowProgress
    
    cPeriodoAnterior = Format(DateAdd("m", -1, psPeriodo), "mmm-yy")
    
    nNroHoja = 1
        
    Set xlHoja1(nNroHoja) = xlLibro.Worksheets.Item(nNroHoja)
    
    xlHoja1(nNroHoja).Name = "PDT" & Format(psPeriodo, "mmm_yy")
    '**************************************************************************
        
         xlAplicacion.Range("A1:A1").ColumnWidth = 3
         xlAplicacion.Range("B1:B1").ColumnWidth = 6
         xlAplicacion.Range("C1:C1").ColumnWidth = 41
         xlAplicacion.Range("D1:D1").ColumnWidth = 10
         
         xlAplicacion.Range("E1:E1").ColumnWidth = 0
         xlAplicacion.Range("F1:F1").ColumnWidth = 0
         xlAplicacion.Range("G1:G1").ColumnWidth = 0
         
         xlAplicacion.Range("H1:H1").ColumnWidth = 10
         
         xlAplicacion.Range("I1:I1").ColumnWidth = 0
         xlAplicacion.Range("J1:J1").ColumnWidth = 0
         xlAplicacion.Range("K1:K1").ColumnWidth = 0
         
         xlAplicacion.Range("L1:L1").ColumnWidth = 10
         xlAplicacion.Range("M1:M1").ColumnWidth = 10
         
         xlAplicacion.Range("N1:N1").ColumnWidth = 0
         xlAplicacion.Range("O1:O1").ColumnWidth = 0
         xlAplicacion.Range("P1:P1").ColumnWidth = 0
         
         xlAplicacion.Range("Q1:Q1").ColumnWidth = 10
     
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 6)).Font.Bold = True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(2, 4)).Merge True
     'xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(3, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 4)).Merge True
'     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(5, 2), xlHoja1(nNroHoja).Cells(5, 4)).Merge True
     xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 2), xlHoja1(nNroHoja).Cells(5, 7)).Font.Bold = True

'******************************************************
            
    lilineas = 6

    Screen.MousePointer = 11
    xlHoja1(nNroHoja).Cells(2, 2) = "Caja Municipal de Ahorros y Creditos de Trujillo SA"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 2), xlHoja1(nNroHoja).Cells(3, 17)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(2, 16), xlHoja1(nNroHoja).Cells(2, 17)).HorizontalAlignment = xlRight
    xlHoja1(nNroHoja).Cells(2, 17) = "Fecha " & pdFecha
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(3, 17), xlHoja1(nNroHoja).Cells(3, 17)).HorizontalAlignment = xlRight
    xlHoja1(nNroHoja).Cells(3, 17) = gsCodUser
    xlHoja1(nNroHoja).Cells(4, 2) = " PDT " & Format(psPeriodo, "mmmm - yyyy")
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 17)).Merge True
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(4, 2), xlHoja1(nNroHoja).Cells(4, 2)).HorizontalAlignment = xlCenter
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(6, 2), xlHoja1(nNroHoja).Cells(7, 17)).Font.Bold = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas + 1, 17)).HorizontalAlignment = xlCenter
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas + 1, 17)).Cells.Interior.Color = &HC0FFFF
    
'    xlHoja1(nNroHoja).Cells(lilineas, 5) = " Vac. Ejecutadas " ' & cPeriodoAnterior
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 5), xlHoja1(nNroHoja).Cells(lilineas, 6)).Merge True
'    xlHoja1(nNroHoja).Cells(lilineas, 7) = " Vac. Devengadas " ' & cPeriodoAnterior
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 7), xlHoja1(nNroHoja).Cells(lilineas, 8)).Merge True
'    xlHoja1(nNroHoja).Cells(lilineas, 9) = " Diferencia " ' & Format(psPeriodo, "mmm-yy")
'    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 9), xlHoja1(nNroHoja).Cells(lilineas, 10)).Merge True
    
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 2)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas + 1, 2)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 2) = "#"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 3)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 3), xlHoja1(nNroHoja).Cells(lilineas + 1, 3)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 3) = "Nombre"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 4)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 4), xlHoja1(nNroHoja).Cells(lilineas + 1, 4)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 4) = "TIM"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 5)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 5), xlHoja1(nNroHoja).Cells(lilineas + 1, 5)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 5) = "Es Salud"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 6)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 6), xlHoja1(nNroHoja).Cells(lilineas + 1, 6)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 6) = "Es Salud Prov"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 7)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 7), xlHoja1(nNroHoja).Cells(lilineas + 1, 7)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 7) = "Es Salud Devengado"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 8)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 8), xlHoja1(nNroHoja).Cells(lilineas + 1, 8)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 8) = "Es Salud OK"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 9)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 9), xlHoja1(nNroHoja).Cells(lilineas + 1, 9)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 9) = "Rem V"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 10)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 10), xlHoja1(nNroHoja).Cells(lilineas + 1, 10)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 10) = "Rem V Provicion"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 11)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 11), xlHoja1(nNroHoja).Cells(lilineas + 1, 11)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 11) = "Rem V Devengado"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 12)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 12), xlHoja1(nNroHoja).Cells(lilineas + 1, 12)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 12) = "Rem V OK"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 13)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 13), xlHoja1(nNroHoja).Cells(lilineas + 1, 13)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 13) = "Es Salud Vida"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 14)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 14), xlHoja1(nNroHoja).Cells(lilineas + 1, 14)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 14) = "Quinta"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 15)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 15), xlHoja1(nNroHoja).Cells(lilineas + 1, 15)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 15) = "Quinta Provicion"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 16)).WrapText = True
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 16), xlHoja1(nNroHoja).Cells(lilineas + 1, 16)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 16) = "Quinta Devengado"
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 17), xlHoja1(nNroHoja).Cells(lilineas + 1, 17)).Merge False
    xlHoja1(nNroHoja).Cells(lilineas, 17) = "Quinta OK"
    lilineas = lilineas + 2
        
'    Set Rs = CargarDiferencia(Format(DateAdd("m", -1, psPeriodo), "yyyymm"))
    Set rs = CargarPDT(Format(psPeriodo, "yyyymm"))
    'lilineas = lilineas + 1
    RaiseEvent Progress(rs.AbsolutePosition, rs.RecordCount)
    Do While Not rs.EOF
        lilineas = lilineas + 1
        xlHoja1(nNroHoja).Cells(lilineas, 2) = rs!cRHCod
        xlHoja1(nNroHoja).Cells(lilineas, 3) = rs!cPersNombre
        xlHoja1(nNroHoja).Cells(lilineas, 4) = rs!nTim
        xlHoja1(nNroHoja).Cells(lilineas, 5) = rs!nEsSalud
        xlHoja1(nNroHoja).Cells(lilineas, 6) = IIf(IsNull(rs!nprovRem), "", rs!nprovRem)
        xlHoja1(nNroHoja).Cells(lilineas, 7) = IIf(IsNull(rs!nDevRem), "", rs!nDevRem)
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 8), xlHoja1(nNroHoja).Cells(lilineas, 8)).Formula = "=" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 5), xlHoja1(nNroHoja).Cells(lilineas, 5)).Address(False, False) & "+" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 6), xlHoja1(nNroHoja).Cells(lilineas, 6)).Address(False, False) & "-" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 7), xlHoja1(nNroHoja).Cells(lilineas, 7)).Address(False, False)
        xlHoja1(nNroHoja).Cells(lilineas, 9) = rs!nRem5ta
        xlHoja1(nNroHoja).Cells(lilineas, 10) = IIf(IsNull(rs!nprovRem), "", rs!nprovRem)
        xlHoja1(nNroHoja).Cells(lilineas, 11) = IIf(IsNull(rs!nDevRem), "", rs!nDevRem)
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 12), xlHoja1(nNroHoja).Cells(lilineas, 12)).Formula = "=" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 9), xlHoja1(nNroHoja).Cells(lilineas, 9)).Address(False, False) & "+" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 10), xlHoja1(nNroHoja).Cells(lilineas, 10)).Address(False, False) & "-" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 11), xlHoja1(nNroHoja).Cells(lilineas, 11)).Address(False, False)
        xlHoja1(nNroHoja).Cells(lilineas, 13) = rs!nEsSaludVida
        xlHoja1(nNroHoja).Cells(lilineas, 14) = rs!nImp5ta
        xlHoja1(nNroHoja).Cells(lilineas, 15) = IIf(IsNull(rs!nprovImp), "", rs!nprovImp)
        xlHoja1(nNroHoja).Cells(lilineas, 16) = IIf(IsNull(rs!nDevImp), "", rs!nDevImp)
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 17), xlHoja1(nNroHoja).Cells(lilineas, 17)).Formula = "=" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 14), xlHoja1(nNroHoja).Cells(lilineas, 14)).Address(False, False) & "+" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 15), xlHoja1(nNroHoja).Cells(lilineas, 15)).Address(False, False) & "-" & xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 16), xlHoja1(nNroHoja).Cells(lilineas, 16)).Address(False, False)
        rs.MoveNext
        RaiseEvent Progress(rs.AbsolutePosition, rs.RecordCount)
    Loop
    lilineas = lilineas + 2
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 17)).Cells.Interior.Color = &HC0FFFF
    xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, 2), xlHoja1(nNroHoja).Cells(lilineas, 17)).Font.Bold = True
    xlHoja1(nNroHoja).Cells(lilineas, 2) = ""
    xlHoja1(nNroHoja).Cells(lilineas, 3) = "TOTAL"
    I = 4
    Do While I < 18
        xlHoja1(nNroHoja).Range(xlHoja1(nNroHoja).Cells(lilineas, I), xlHoja1(nNroHoja).Cells(lilineas, 17)).Formula = "=SUM(" & Chr(64 + I) & "8:" & Chr(64 + I) & lilineas - 1 & ")"
        I = I + 1
    Loop
    
    'ExcelCuadro xlHoja1(nNroHoja), 2, 6, 17, lilineas, True, True
    
    RaiseEvent CloseProgress
    
    Screen.MousePointer = 0
'******************************************************

'    xlHoja1(nNroHoja).Cells.Select
    xlHoja1(nNroHoja).Cells.NumberFormat = "###,###,##0.00"
    xlHoja1(nNroHoja).Cells.Font.Size = 8
'    xlHoja1(nNroHoja).Cells.EntireColumn.AutoFit

    'Libera los objetos.
    Set xlHoja1(nNroHoja) = Nothing
        
    xlLibro.SaveAs App.Path & "\spooler\PDT_" & Format(psPeriodo, "mmmyy")
    Set xlLibro = Nothing
    xlAplicacion.Application.Visible = True
'    xlAplicacion.Windows(1).Visible = True
    Set xlAplicacion = Nothing
    Exit Sub
ImprimePDTErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, Err
End Sub

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

Public Sub GuardarPDT(ByVal pcRHCod As String, ByVal pcPeriodo As String, ByVal pnTIM As Currency, ByVal pnEsSalud As Currency, ByVal pnRem5 As Currency)
On Error GoTo GuardarTIMErr
    Dim oCon As DConecta, sSql As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSql = " declare @tmp int "
        sSql = sSql & " set @tmp=(select count(*) from RHPDT where cRHCod='" & pcRHCod & "' and cPeriodo='" & pcPeriodo & "')"
        sSql = sSql & " if @tmp=0 "
        sSql = sSql & " insert into RHPDT(cRHCod,cPeriodo,nTIM,nEsSalud,nRem5ta) " & _
                " values('" & pcRHCod & "','" & pcPeriodo & "'," & pnTIM & "," & pnEsSalud & "," & pnRem5 & ") "
        oCon.Ejecutar sSql
        oCon.CierraConexion
    End If
    Exit Sub
GuardarTIMErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

Public Sub ModificarPDT(ByVal pcRHCod As String, ByVal pcPeriodo As String, ByVal pnImp5 As Currency, ByVal pnEsSaludVida As Currency)
On Error GoTo GuardarTIMErr
    Dim oCon As DConecta, sSql As String
    Set oCon = New DConecta
    If oCon.AbreConexion Then
        sSql = "update RHPDT set nImp5ta = " & pnImp5 & "," & _
                " nEsSaludVida = " & pnEsSaludVida & " where cRHCod = '" & pcRHCod & "' and cPeriodo = '" & pcPeriodo & "'"
        oCon.Ejecutar sSql
        oCon.CierraConexion
    End If
    Exit Sub
GuardarTIMErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Sub

'**********************************************************************************
'Certificado AFP

Private Function CargarDatosCertificados(ByVal pnAnio As Integer) As ADODB.Recordset
On Error GoTo CargarDatosCertificadosErr
    Dim oCon As DConecta, sSql As String, rs As ADODB.Recordset
    Set oCon = New DConecta
    sSql = "SELECT cPErsCod, cPersNombre, DNI, LTRIM(RTRIM (cNumAFP)) CNUMAFP,AFP, cCUSPP, XXX.cAgenciaAsig, XXX.Agencia, sum(XXX.nMonto) Anual, " & _
           " sum(case when cRHConceptoCod IN ('202','221') then XXX.nMonto else 0 end) Variable, " & _
           " sum(case when cRHConceptoCod IN ('203','220') then XXX.nMonto else 0 end) Fija, " & _
           " sum(case when cRHConceptoCod IN ('204','222') then XXX.nMonto else 0 end) PrimaSeguro " & _
           " FROM ( " & _
           "             Select R.cPErsCod, DBO.GETNOMBRES(cPersNombre) + ' ' + dbo.GetNombreAPPaterno(cPersNombre) + ' ' + dbo.GetNombreAPMaterno(cPersNombre) cPersNombre, " & _
           "             (Select cPersIDnro from PersID PP where  PP.cPErsCod = R.cPErsCod and cPersIDTpo = '1' )  DNI, cNumAFP , " & _
           "             (Select case when cRHAFPAbreviatura  = 'NVU' then 'NUEVA VIDA'  " & _
           "                                      when cRHAFPAbreviatura  = 'HOR' then 'HORIZONTE'  " & _
           "                                      when cRHAFPAbreviatura  = 'PRO' then 'PRO FUTURO'  " & _
           "                                      when cRHAFPAbreviatura  = 'PRI' then 'PRIMA' " & _
           "                                      when cRHAFPAbreviatura  = 'INT' then 'INTEGRA'  end " & _
           "                             from  RHAFP " & _
           "                             where cRHAFPPersCod = R.CRHEMPLAFPPERSCOD) AFP, " & _
           "             R.cCUSPP , Pla.nMonto,cRHConceptoCod,cAgenciaAsig, " & _
           "             (Select cAgeDescripcion from Agencias where cAgeCod =RH.cAgenciaAsig )  Agencia " & _
           "             from RHEmpleado R " & _
           "             Inner Join RRHH RH on RH.cPErsCod = R.cPErsCod " & _
           "             Inner Join Persona P on P.cPErsCod = R.cPErsCod " & _
           "             Inner Join RHPlanillaDetCon PLA on PLA.cPersCod = R.cPersCod and substring(cRRHHPeriodo,1,4)='" & pnAnio & "' " & _
           "             Where cRHConceptoCod in ('202','203','204','220','221','222') " & _
           " ) XXX " & _
           " GROUP BY cPErsCod,cPersNombre,DNI,cNumAFP,AFP,cCUSPP,cAgenciaAsig,Agencia " & _
           " ORDER BY cAgenciaAsig,cPersNombre "
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sSql)
        oCon.CierraConexion
    End If
    Set CargarDatosCertificados = rs
    Exit Function
CargarDatosCertificadosErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Aviso"
End Function

Public Function ImpCertificadoAFP(ByVal pnAnio As Integer, ByVal psFecha As String) As Boolean

On Error GoTo ImpCertificadoAFPErr

Dim lsModelo As String
Dim vCont As Integer
Dim rs As ADODB.Recordset
Dim I As Integer

lsModelo = App.Path & "\spooler\PLANTILLA AFP.doc"
   
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
    
    'Crea Nuevo Documento
    wApp.Documents.Add
    
    Set rs = CargarDatosCertificados(pnAnio)
    
    RaiseEvent ShowProgress
    RaiseEvent Progress(rs.AbsolutePosition, rs.RecordCount)
    Do While Not rs.EOF
        'Lo carga en Memoria
        wAppSource.ActiveDocument.Content.Copy
        
        wApp.Application.Selection.Paste
        
        wApp.Selection.SetRange start:=wApp.Selection.start, End:=wApp.ActiveDocument.Content.End
        wApp.Selection.MoveEnd
                
        With wApp.Selection.Find
            .Text = "<<AFP>>"
            .Replacement.Text = IIf(IsNull(rs!AFP), "", rs!AFP)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<Agencia>>"
            .Replacement.Text = rs!Agencia
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<CodPers>>"
            .Replacement.Text = rs!cPersCod
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
        End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<Año>>"
            .Replacement.Text = pnAnio
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<Nombre>>"
            .Replacement.Text = rs!cPersNombre
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<DNI>>"
            .Replacement.Text = rs!DNI
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<CUSPP>>"
            .Replacement.Text = rs!cCUSPP
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<Total>>"
            .Replacement.Text = FNumero(rs!Anual)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<Variable>>"
            .Replacement.Text = FNumero(rs!Variable)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<Fija>>"
            .Replacement.Text = FNumero(rs!Fija)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<Prima>>"
            .Replacement.Text = FNumero(rs!PrimaSeguro)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        
        With wApp.Selection.Find
            .Text = "<<Fecha>>"
            .Replacement.Text = Format(psFecha, "Trujillo dd - mmmm - yyyy")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
          End With
        wApp.Selection.Find.Execute Replace:=wdReplaceAll
        rs.MoveNext
'        If i = 20 Then
'            Exit Do
'        Else
'            i = i + 1
'        End If
        RaiseEvent Progress(rs.AbsolutePosition, rs.RecordCount)
    Loop
    
    RaiseEvent CloseProgress
            
    'wapp.Selection.Font
    wAppSource.ActiveDocument.Close
    wApp.ActiveDocument.SaveAs App.Path + "\spooler\CertificadoAFP_" & pnAnio & ".doc"
    wApp.ActiveDocument.Close
    Set wAppSource = Nothing
    Set wApp = Nothing
    Exit Function
ImpCertificadoAFPErr:
    MsgBox Err.Number & vbCrLf & Err.Description, vbInformation, "Error"
End Function

Private Function FNumero(vExpNumStr As Variant, Optional vNroDecimales As Integer) As String
Dim nDec As Integer
If Not IsNull(vExpNumStr) Then
   nDec = IIf(vNroDecimales = 0, 2, vNroDecimales)
   If InStr(",", vExpNumStr) = 0 Then
      FNumero = Format(vExpNumStr, "###,###,##0." + String(nDec, "0"))
   Else
      FNumero = Format(Val(vExpNumStr), "###,###,##0." + String(nDec, "0"))
   End If
Else
   FNumero = ""
End If
End Function





'''''''''




Public Function GetHorasExtrasEmpleadosDetalle(pdIni As Date, pdFin As Date, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pgsNombreCompleto As String, pgsRuc As String, pgsTipo As String) As String
    Dim sqlT As String
    Dim rsT As ADODB.Recordset
    Set rsT = New ADODB.Recordset
    Dim lsCadena As String
    Dim oRep As DRHReportes
    Set oRep = New DRHReportes
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    Dim lsFecha As String * 12
    Dim lsDia As String * 10
    Dim lsFechaVig As String * 12
    Dim lsT1Ini As String * 10
    Dim lsT1Fin As String * 10
    Dim lsT2Ini As String * 10
    Dim lsT2Fin As String * 10
    Dim lsA1Ini As String * 10
    Dim lsA1Fin As String * 10
    Dim lsA2Ini As String * 10
    Dim lsA2Fin As String * 10
    Dim lsTar1 As String * 5
    Dim lsTar2 As String * 5
    Dim lsHE As String * 7
    Dim lsPersNom As String * 30
    Dim lsPersCod As String * 15
    Dim lsPersCod2 As String * 15
    Dim lsNomEnc As String * 35
    Dim lnPagina As Long
    Dim lnIndice As Long
    Dim lnTar1 As Long
    Dim lnTar2 As Long
    Dim lnHETotal As Double
    Dim lnTotalGeneral As Double
    Set rsT = oRep.GetHorasExtrasDet(Format(pdIni, gsFormatoFecha), Format(pdFin, gsFormatoFecha))
    
    lsCadena = ""
    lnTar1 = 0
    lnTar2 = 0
    lnTotalGeneral = 0
    If Not RSVacio(rsT) Then
        lsNomEnc = PstaNombre(rsT!cPersNombre)
        lsCadena = lsCadena & CabeceraPaginaHE(Space(21) & "REPORTE DEL REGISTRO DE CONTROL DE ASISTENCIA" & _
                oImpresora.gPrnSaltoLinea & Space(19) & pgsNombreCompleto & _
                oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
                oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
                oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & _
                oImpresora.gPrnSaltoLinea & Space(18) & " Código          Apellidos y Nombres              DNI" & _
                oImpresora.gPrnSaltoLinea & Space(19) & rsT!cRHCod & "      " & lsNomEnc & "  " & rsT!cDNI, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
      
        If pgsTipo = "1" Then
            'lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11;TH1.Ini;11; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;4; ;2;HE;4;", lnIndice)
            lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11;TH1.Ini;11; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;5; ;2;HE;5;", lnIndice)

        Else
            'lsCadena = lsCadena & Encabezado("Fecha;5;;5;Fecha Vig;11;TH1.Ini;13; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;4; ;2;HE;4;", lnIndice)
            'lsCadena = lsCadena & Encabezado("Fecha;5;;5;Fecha Vig;11;TH1.Ini;13; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;5; ;2;HE;5; ;2;REFRIGERIO.", lnIndice)
            lsCadena = lsCadena & Encabezado("Fecha;5;;5;Fecha Vig;11;TH1.Ini;13; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;5; ;2;HE;5;", lnIndice)  'OK
        End If
        
        While Not rsT.EOF
      
        
            LSet lsFecha = rsT!Fecha
            LSet lsDia = rsT!cDia
            LSet lsFechaVig = IIf(IsNull(rsT!FechaVig), "", rsT!FechaVig)
            
            RSet lsT1Ini = rsT!T1Ini
            RSet lsT1Fin = rsT!T1Fin
            RSet lsT2Ini = rsT!T2Ini
            RSet lsT2Fin = rsT!T2Fin
                       
            If IsNull(rsT!Comision) Then
                If IsNull(rsT!PerNoLabDes) Then
                    RSet lsA1Ini = rsT!A1Ini
                    RSet lsA1Fin = rsT!A1Fin
                    RSet lsA2Ini = rsT!A2Ini
                    RSet lsA2Fin = rsT!A2Fin
                    
                    RSet lsTar1 = rsT!MT1
                    RSet lsTar2 = rsT!MT2
                    
                    If pgsTipo = "1" Then
                        RSet lsHE = IIf(IsNull(rsT!HE) Or rsT!HE < 0, Format(0, "#0.000"), Format((rsT!HE / 60), "#0.000"))
                    Else
                        RSet lsHE = IIf(IsNull(rsT!HE) Or rsT!HE < 0, Format(0, "#0.00"), Format((rsT!HE / 60), "#0.00"))
                       
                        If lsHE < 1 Then
                            RSet lsHE = Format(0, "#0.00")
                        End If
                    End If
                    
                    lsPersNom = rsT!cPersNombre
                    
                Else
                    RSet lsA1Ini = Left(rsT!PerNoLabDes, 8)
                    RSet lsA1Fin = Left(rsT!PerNoLabDes, 8)
                    RSet lsA2Ini = Left(rsT!PerNoLabDes, 8)
                    RSet lsA2Fin = Left(rsT!PerNoLabDes, 8)
                    RSet lsTar1 = "0"
                    RSet lsTar2 = "0"
                    RSet lsHE = IIf(pgsTipo = "1", "0.000", "0.00")
                End If
            Else
                RSet lsA1Ini = "COMISION"
                RSet lsA1Fin = "COMISION"
                RSet lsA2Ini = "COMISION"
                RSet lsA2Fin = "COMISION"
                RSet lsTar1 = "0"
                RSet lsTar2 = "0"
            End If
            
            lnIndice = lnIndice + 1
            If pgsTipo = "1" Then
                lsCadena = lsCadena & lsFecha & lsDia & lsFechaVig & lsT1Ini & lsT1Fin & lsT2Ini & lsT2Fin & lsA1Ini & lsA1Fin & lsA2Ini & lsA2Fin & lsTar1 & lsTar2 & lsHE & oImpresora.gPrnSaltoLinea
            Else
                lsCadena = lsCadena & lsFecha & lsFechaVig & lsT1Ini & lsT1Fin & lsT2Ini & lsT2Fin & lsA1Ini & lsA1Fin & lsA2Ini & lsA2Fin & lsTar1 & lsTar2 & lsHE & oImpresora.gPrnSaltoLinea
            End If
            
            If IsNull(rsT!Comision) And IsNull(rsT!PerNoLabDes) Then
                lnTar1 = lnTar1 + rsT!MT1
                lnTar2 = lnTar2 + rsT!MT2
            End If
            
            If lnIndice >= 54 Then
                lnIndice = 1
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPaginaHE(Space(21) & "REPORTE DEL REGISTRO DE CONTROL DE ASISTENCIA" & _
                oImpresora.gPrnSaltoLinea & Space(19) & pgsNombreCompleto & _
                oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
                oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
                oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & _
                oImpresora.gPrnSaltoLinea & Space(18) & " Código          Apellidos y Nombres              DNI" & _
                oImpresora.gPrnSaltoLinea & Space(19) & rsT!cRHCod & "      " & lsNomEnc & "  " & rsT!cDNI, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
      

                If pgsTipo = "1" Then
                    'lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11;TH1.Ini;11; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;4; ;2;HE;4;", lnIndice)
                    lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11;TH1.Ini;11; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;5; ;2;HE;5;", lnIndice)
                Else
                    'lsCadena = lsCadena & Encabezado("Fecha;5;;5;Fecha Vig;11;TH1.Ini;13; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;4; ;2;HE;4;", lnIndice)
                    lsCadena = lsCadena & Encabezado("Fecha;5;;5;Fecha Vig;11;TH1.Ini;13; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;5; ;2;HE;5;", lnIndice)
                End If
            End If
            lnHETotal = Format(lnHETotal + IIf(IsNull(rsT!HE) Or rsT!HE < 0, 0, Format((rsT!HE / 60), "#0.00")), "#0.00")
            lnTotalGeneral = Format(lnTotalGeneral + IIf(IsNull(rsT!HE) Or rsT!HE < 0, 0, Format((rsT!HE / 60), "#0.00")), "#0.00")
            lsPersCod = rsT!cPersCod
            rsT.MoveNext
            If rsT.EOF = True Then
                GoTo Fin
            End If
            lsNomEnc = PstaNombre(rsT!cPersNombre)
            lsPersCod2 = rsT!cPersCod
            
            If lsPersCod <> lsPersCod2 Then
             If pgsTipo = "1" Then
                lsCadena = lsCadena & String(132, "-")
             '   lsCadena = lsCadena & Encabezado("TOTAL HORAS EXTRAS ;124;" & lnHETotal & ";5;", lnIndice, False)
             Else
             '   lsCadena = lsCadena & Encabezado("TOTAL HORAS EXTRAS ;114;" & lnHETotal & ";5;", lnIndice, False)
                lsCadena = lsCadena & String(122, "-")
             End If
             
             lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & "LEYENDA " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TH1.Ini. Ingreso Turno1 Horario Asignado " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TH1.Fin. Salida Turno1 Horario Asignado " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TH2.Ini. Ingreso Turno2 Horario Asignado " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TH2.Fin. Salida Turno2 Horario Asignado " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TA1.Ini. Ingreso Turno1 Registro Real " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TA1.Fin. Salida Turno1 Registro Real " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TA2.Ini. Ingreso Turno2 Registro Real " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TA2.Fin. Salida Turno3 Registro Real " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " T1.      Tardanza Turno 1" & _
             oImpresora.gPrnSaltoLinea & Space(10) & " T2.      Tardanza Turno 2 " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " HE.      Horas Extras "

            
            
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lnIndice = 1
            lsCadena = lsCadena & CabeceraPaginaHE(Space(21) & "REPORTE DEL REGISTRO DE CONTROL DE ASISTENCIA" & _
                oImpresora.gPrnSaltoLinea & Space(19) & pgsNombreCompleto & _
                oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
                oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
                oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & _
                oImpresora.gPrnSaltoLinea & Space(18) & " Código          Apellidos y Nombres              DNI" & _
                oImpresora.gPrnSaltoLinea & Space(19) & rsT!cRHCod & "      " & lsNomEnc & "  " & rsT!cDNI, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
      
   If pgsTipo = "1" Then
                    'lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11;TH1.Ini;11; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;4; ;2;HE;4;", lnIndice)
                    lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11;TH1.Ini;11; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;5; ;2;HE;5;", lnIndice)
                Else
                    'lsCadena = lsCadena & Encabezado("Fecha;5;;5;Fecha Vig;11;TH1.Ini;13; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;4; ;2;HE;4;", lnIndice)
                    lsCadena = lsCadena & Encabezado("Fecha;5;;5;Fecha Vig;11;TH1.Ini;13; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;5; ;2;HE;5;", lnIndice)
                End If

            lnHETotal = 0
                If lnIndice >= 54 Then
                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                    lsCadena = lsCadena & CabeceraPaginaHE(Space(21) & "REPORTE DEL REGISTRO DE CONTROL DE ASISTENCIA" & _
                    oImpresora.gPrnSaltoLinea & Space(19) & pgsNombreCompleto & _
                    oImpresora.gPrnSaltoLinea & Space(34) & " RUC " & pgsRuc & "" & _
                    oImpresora.gPrnSaltoLinea & Space(27) & " DEL " & pdIni & "  AL  " & pdFin & "" & _
                    oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & _
                    oImpresora.gPrnSaltoLinea & Space(18) & " Código          Apellidos y Nombres              DNI" & _
                    oImpresora.gPrnSaltoLinea & Space(19) & rsT!cRHCod & "      " & lsNomEnc & "  " & rsT!cDNI, lnPagina, lnIndice, pgsNomAge, pgsEmpresa, pgdFecSis, "")
      

                    If pgsTipo = "1" Then
                        'lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11;TH1.Ini;11; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;4; ;2;HE;4;", lnIndice)
                        lsCadena = lsCadena & Encabezado("Fecha;5; ;7;Dia;5; ;5;Fecha Vig;11;TH1.Ini;11; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;5; ;2;HE;5;", lnIndice)
                    Else
                        'lsCadena = lsCadena & Encabezado("Fecha;5;;5;Fecha Vig;11;TH1.Ini;13; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;4; ;2;HE;4;", lnIndice)
                        lsCadena = lsCadena & Encabezado("Fecha;5;;5;Fecha Vig;11;TH1.Ini;13; ;2;TH1.Fin;8; ;2;TH2.Ini;8; ;2;TH2.Fin;8; ;2;TA1.Ini;8; ;2;TA1.Fin;8; ;2;TA2.Ini;8; ;2;TA2.Fin;8;T1;5;T2;5; ;2;HE;5;", lnIndice)
                    End If
                End If
            End If
            
            
        Wend
Fin:
        lsFecha = ""
        lsDia = ""
        lsFechaVig = ""
        lsT1Ini = ""
        lsT1Fin = ""
        lsT2Ini = ""
        lsT2Fin = ""
        lsA1Ini = ""
        lsA1Fin = ""
        lsA2Ini = ""
        lsA2Fin = ""
        lsHE = ""
        RSet lsTar1 = lnTar1
        RSet lsTar2 = lnTar2
        If pgsTipo = "1" Then
            lsCadena = lsCadena & String(132, "-")
            '   lsCadena = lsCadena & Encabezado("TOTAL HORAS EXTRAS ;124;" & lnHETotal & ";5;", lnIndice, False)
        Else
            '   lsCadena = lsCadena & Encabezado("TOTAL HORAS EXTRAS ;114;" & lnHETotal & ";5;", lnIndice, False)
            lsCadena = lsCadena & String(122, "-")
        End If
             lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & "LEYENDA " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TH1.Ini. Ingreso Turno1 Horario Asignado " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TH1.Fin. Salida Turno1 Horario Asignado " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TH2.Ini. Ingreso Turno2 Horario Asignado " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TH2.Fin. Salida Turno2 Horario Asignado " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TA1.Ini. Ingreso Turno1 Registro Real " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TA1.Fin. Salida Turno1 Registro Real " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TA2.Ini. Ingreso Turno2 Registro Real " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " TA2.Fin. Salida Turno3 Registro Real " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " T1.      Tardanza Turno 1" & _
             oImpresora.gPrnSaltoLinea & Space(10) & " T2.      Tardanza Turno 2 " & _
             oImpresora.gPrnSaltoLinea & Space(10) & " HE.      Horas Extras "
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        'lsCadena = lsCadena & Encabezado("TOTAL GENERAL HORAS EXTRAS ;124;" & lnTotalGeneral & ";5;", lnIndice, False)

    Else
        MsgBox "No se encontro información.", vbInformation, "Aviso"
    End If

    rsT.Close
    Set rsT = Nothing
    
    GetHorasExtrasEmpleadosDetalle = lsCadena
End Function


