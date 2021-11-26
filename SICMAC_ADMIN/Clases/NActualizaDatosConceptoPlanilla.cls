VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NActualizaDatosConPlanilla"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9B14A40021"
'Actualiza los tipos de planilla y las realciones de conceptos y Planillas
Option Base 0
Option Explicit

Public Event ShowProgress()
Public Event CloseProgress()
Public Event Progress(pnValor As Long, pnTotal As Long)

Private lnTotDias As Double
Private lnNumMonto As Double
'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902DC0235
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902DC02F4
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902DC02E0
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

'Agrega una nueva relación entre concepto y planilla
'##ModelId=3A9B14D600B9
Public Function AgregaConceptoPlanilla(psRHPlanillaCod As String, psRHPlanillaDescripcion As String, psUltimaActualizacion As String, psRHRefPago As String, psRHRefDescripcion As String, psRHRefInicio As String, psRHRefFin As String, pnRHRefParametro As Double, psRHRefPlaSub As String, psRHRefPlaEstado As String, psOpeCodEst As String, psOpeCodCont As String, psOpeAbono As String) As Boolean
    On Error GoTo AgregaConceptoPlanillaErr
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    psRHPlanillaCod = oPla.GetPlanillaCodigo(psRHPlanillaCod)
    oPla.AgregaConceptoPlanilla psRHPlanillaCod, psRHPlanillaDescripcion, psUltimaActualizacion, psRHRefPago, psRHRefDescripcion, psRHRefInicio, psRHRefFin, pnRHRefParametro, psRHRefPlaSub, psRHRefPlaEstado, psOpeCodEst, psOpeCodCont, psOpeAbono

    Exit Function
AgregaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:AgregaConceptoPlanilla Method")
End Function

'Elimina una relacion entre concepto y planilla
'
'##ModelId=3A9B14D600EB
Public Function EliminaConceptoPlanilla(psRHPlanillaCod As String) As Boolean
    On Error GoTo EliminaConceptoPlanillaErr

    'your code goes here...

    Exit Function
EliminaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:EliminaConceptoPlanilla Method")
End Function

'Actualiza una nueva relacion entre concepto y planilla
'##ModelId=3A9B14D6011D
Public Function ModificaConceptoPlanilla(psRHPlanillaCod As String, psRHPlanillaDescripcion As String, psUltimaActualizacion As String, psRHRefPago As String, psRHRefDescripcion As String, psRHRefInicio As String, psRHRefFin As String, pnRHRefParametro As Double, psRHRefPlaSub As String, psRHRefPlaEstado As String, psOpeCodEst As String, psOpeCodCont As String, psOpeAbono As String) As Boolean
    On Error GoTo ModificaConceptoPlanillaErr
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla

    oPla.ModificaConceptoPlanilla psRHPlanillaCod, psRHPlanillaDescripcion, psUltimaActualizacion, psRHRefPago, psRHRefDescripcion, psRHRefInicio, psRHRefFin, pnRHRefParametro, psRHRefPlaSub, psRHRefPlaEstado, psOpeCodEst, psOpeCodCont, psOpeAbono
    
    Exit Function
ModificaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ModificaConceptoPlanilla Method")
End Function

'Actualiza una nueva relacion entre concepto y planilla
'##ModelId=3A9B14D6011D
Public Function ModificaPlanilla(psRHPlanillaPeriodo As String, psRHPlanillaCod As String, prRS As ADODB.Recordset, psUltimaActualizacion As String, psFecImpre As String, psFecIni As String, psFecFin As String, psTipcambio As String, psDescripcion As String, psFecPago As String, psTipcambioFF As String, pnEstado As Integer) As Boolean
    On Error GoTo ModificaConceptoPlanillaErr
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla

    oPla.ModificaPlanilla psRHPlanillaPeriodo, psRHPlanillaCod, prRS, psUltimaActualizacion, psFecImpre, psFecIni, psFecFin, psTipcambio, psDescripcion, psFecPago, psTipcambioFF, pnEstado

    Set oPla = Nothing
    Exit Function
ModificaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ModificaConceptoPlanilla Method")
End Function

'Actualiza una nueva relacion entre concepto y planilla
'##ModelId=3A9B14D6011D
Public Function CierraPlanilla(psRHPlanillaPeriodo As String, psRHPlanillaCod As String) As Boolean
    On Error GoTo ModificaConceptoPlanillaErr
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla

    oPla.CierraPlanilla psRHPlanillaPeriodo, psRHPlanillaCod

    Set oPla = Nothing
    Exit Function
ModificaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ModificaConceptoPlanilla Method")
End Function

'Actualiza una nueva relacion entre concepto y planilla
'##ModelId=3A9B14D6011D
Public Function EliminaPlanilla(psRHPlanillaPeriodo As String, psRHPlanillaCod As String) As Boolean
    On Error GoTo ModificaConceptoPlanillaErr
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla

    oPla.EliminaPlanilla psRHPlanillaPeriodo, psRHPlanillaCod

    Set oPla = Nothing
    Exit Function
ModificaConceptoPlanillaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ModificaConceptoPlanilla Method")
End Function

Public Function EliminaPlanillaEstadosRH(psRHPlanillaCod As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim oCon As DActualizaDatosConPlanilla
    Set oCon = New DActualizaDatosConPlanilla
    
    oCon.EliminaPlanillaEstadosRH psRHPlanillaCod
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

Public Function AgregaPlanillaEstadosRH(psRHPlanillaCod As String, psRHEstado As String) As Recordset
    On Error GoTo GetPlanillasErr
    Dim oCon As DActualizaDatosConPlanilla
    Set oCon = New DActualizaDatosConPlanilla
    
    oCon.AgregaPlanillaEstadosRH psRHPlanillaCod, psRHEstado
    
    Set oCon = Nothing
    Exit Function
GetPlanillasErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConceptoPlanilla:GetPlanillas Method")
End Function

'Agrega un nuevo Informe Social a la Persona
'##ModelId=3A96B0B0022B
Public Function GetReporte(psRHPlanilaCod As String, psConceptoTpo As String, ByVal psTitulo As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    On Error GoTo GetReporteErr
    Dim lsCadena As String
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Dim I As Integer
    Dim lnMax As Integer
    Dim rsAdd As ADODB.Recordset
    Set rsAdd = New ADODB.Recordset
    Dim rsLib As ADODB.Recordset
    Set rsLib = New ADODB.Recordset
    
    Dim lsItem1 As String * 20
    Dim lsItem2 As String * 40
    Dim lsItem3 As String * 20
    Dim lsItem4 As String * 40
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena(Left(psRHPlanilaCod, 50), 120) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena(String(Len(Left(psRHPlanilaCod, 50)), "="), 120) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena(Left(psConceptoTpo, 20), 120) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    Set rsAdd = oPla.GetPlanillaConcepUsados(Left(psRHPlanilaCod, 3), Right(psConceptoTpo, 1))
    Set rsLib = oPla.GetPlanillaConcepLibres(Left(psRHPlanilaCod, 3), Right(psConceptoTpo, 1))
    
    If rsAdd.RecordCount < rsLib.RecordCount Then
        lnMax = rsLib.RecordCount
    Else
        lnMax = rsAdd.RecordCount
    End If
    
    lsItem3 = "CONCEPTOS LIBRES"
    lsItem1 = "   !CONCEPTOS PLANILLA"
    
    lsCadena = lsCadena & "+" & String(125, "-") & "+" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsItem3 & lsItem4 & lsItem1 & lsItem2 & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Encabezado("!Codigo;7; ;8;Nombre;20; ;28;!;1; ;1;Codigo;7; ;5;Nombre;20; ;29;!;1;", 1)
    For I = 0 To lnMax
        lsItem1 = ""
        lsItem2 = ""
        lsItem3 = ""
        lsItem4 = ""
        
        If I < rsAdd.RecordCount Then
            lsItem1 = "   " & rsAdd.Fields(0)
            lsItem2 = rsAdd.Fields(1)
            rsAdd.MoveNext
        End If
        
        If I < rsLib.RecordCount Then
            lsItem3 = "  " & rsLib.Fields(0)
            lsItem4 = rsLib.Fields(1)
            rsLib.MoveNext
        End If
        
        lsCadena = lsCadena & "!" & lsItem3 & " " & lsItem4 & " ! " & lsItem1 & " " & lsItem2 & "!" & oImpresora.gPrnSaltoLinea
    Next I
    lsCadena = lsCadena & "+" & String(125, "-") & "+" & oImpresora.gPrnSaltoLinea
    rsLib.Close
    rsAdd.Close
    Set rsLib = Nothing
    Set rsAdd = Nothing
    
    GetReporte = lsCadena
    Exit Function
GetReporteErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosInformeSocial:AgregaInformeSocial Method")
End Function

Public Function GetBoletas(psPlanillaPeriodo As String, psPlanillaCod As String, psEmpLista As String, psNombrePlanilla As String, psFechaImprePlanilla As String, pgsRuc As String, pgsEmpresa As String, psTextoPlanilla As String) As String
    Dim lsCadena As String
    Dim lsCadAux As String
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    Dim lsEmpDes As String
    
    Set rsC = oPla.GetBoletas(psPlanillaPeriodo, psPlanillaCod, psEmpLista)

    If Not (rsC.BOF And rsC.EOF) Then
        RaiseEvent ShowProgress
        While Not rsC.EOF
            If IIf(oCon.LeeConstSistema(gConstSistBoletaRRHH) = "1", True, False) Then
                lsCadAux = GetBoletaLima(rsC!cPersCod, psPlanillaPeriodo, psPlanillaCod, psNombrePlanilla, psFechaImprePlanilla, pgsRuc, pgsEmpresa, psTextoPlanilla)
                lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & lsCadAux & oImpresora.gPrnSaltoPagina
            Else
                lsCadAux = GetBoleta(rsC!cPersCod, psPlanillaPeriodo, psPlanillaCod, psNombrePlanilla, psFechaImprePlanilla, pgsRuc, pgsEmpresa, psTextoPlanilla)
                lsCadena = lsCadena & lsCadAux & Chr(12)  'oImpresora.gPrnSaltoPagina
            End If
            RaiseEvent Progress(rsC.Bookmark, rsC.RecordCount)
            rsC.MoveNext
        Wend
        RaiseEvent CloseProgress
    End If
    
    rsC.Close
    'Me.Caption = "Generación de Planillas."
    GetBoletas = lsCadena
End Function

'Private Function GetBoleta(psPersCod As String, pdFecha As String, psPlaCod As String, psNombrePlanilla As String, psFechaImprePlanilla As String, pgsRuc As String, pgsEmpresa As String, psTextoPlanilla As String) As String
'    On Error Resume Next
'
'    Dim lsCadena As String
'    Dim rsP As ADODB.Recordset
'    Set rsP = New ADODB.Recordset
'    Dim rsCargo As ADODB.Recordset
'    Set rsCargo = New ADODB.Recordset
'    Dim oPla As DActualizaDatosConPlanilla
'    Set oPla = New DActualizaDatosConPlanilla
'    Dim oRH As DActualizaDatosRRHH
'    Set oRH = New DActualizaDatosRRHH
'
'    Dim lsCod As String * 8
'    Dim lsNombre As String * 36
'    Dim lsCargo As String * 27
'    Dim lsAFP As String * 5
'    Dim lsFecIng As String * 12
'    Dim lsNumC As String * 20
'    Dim lsSueldo As String * 12
'    Dim lsPorMesLabT As String * 13
'    Dim lsPorMesLab As String * 10
'    Dim lsNumCarnetAFP As String * 20
'    Dim lsCondiciaon As String * 16
'    Dim lsCondicion As String * 16
'    Dim lsAgencia As String * 10
'    Dim lsVacaciones As String * 22
'
'    Dim lsConcep(20, 6) As String * 15
'    Dim lsConcepT(1, 6) As String * 15
'    Dim lbNSF As Boolean
'
'    Dim I As Integer
'    Dim lnI As Integer
'    Dim lnD As Integer
'    Dim lnA As Integer
'    Dim lnPosMax As Integer
'    Dim lsCondicionE As String
'    Dim lsNumCE As String
'    Dim ldFecVacIni As Date
'    Dim ldFecVacFin As Date
'
'    Dim lsNegritaOn As String
'    Dim lsNegritaOff As String
'    Dim lnEspacio As Integer
'    Dim lnnI As Integer
'    Dim lnnJ As Integer
'
'    For lnnI = 0 To 20
'        For lnnJ = 0 To 6
'            lsConcep(lnnI, lnnJ) = ""
'        Next lnnJ
'    Next lnnI
'
'
'    lnEspacio = 0
'    lsNegritaOn = oImpresora.gPrnBoldON
'    lsNegritaOff = oImpresora.gPrnBoldOFF
'
'    Set rsP = oPla.GetBoleta(psPersCod, pdFecha, psPlaCod)
'
'    lnI = 0
'    lnD = 0
'    If rsP.EOF And rsP.BOF Then
'        rsP.Close
'        Set rsP = Nothing
'        Exit Function
'    End If
'
'    lsCod = "CODIGO"
'    lsNombre = "NOMBRES Y APELLIDOS"
'   ' lbNSF = rsP!bNoSujFiscalizacion
'    'lsCargo = IIf(Left(psPlaCod, 1) = "E", "CARGO" & IIf(rsP!bConfianza, "-CONFIANZA", "") & IIf(rsP!bNoSujFiscalizacion, "-NSF", "") & IIf(rsP!bPersDireccion, "-PERS DIRECC", ""), "CONVENIO")
'    lsFecIng = "FECHA.ING"
'    lsNumC = IIf(Left(psPlaCod, 1) = "E", "Nª AUTOGENERADO", "")
'
'    lsCadena = lsCadena & lsNegritaOn & Space(lnEspacio) & lsCod & Space(4) & lsNombre & Space(1) & lsCargo & Space(3) & lsFecIng & lsNumC & oImpresora.gPrnSaltoLinea & lsNegritaOff
'
'    lsCod = rsP!cRHCod
'    lsNombre = PstaNombre(rsP!cPersNombre, False)
'    Set rsCargo = oRH.GetCargoRS(psPersCod, pdFecha)
'    lsCargo = rsCargo!Cargo '& IIf(rsCargo!nPDireccion = 1, " Pers Direc", "")
'    lsFecIng = Format(rsP!dIngreso, gsFormatoFechaView)
'    lsNumC = rsP!cNumAFP & ""
'
'    lsCadena = lsCadena & Space(lnEspacio) & lsCod & Space(4) & lsNombre & Space(1) & lsCargo & Space(2) & lsFecIng & lsNumC & oImpresora.gPrnSaltoLinea
'
'    lsSueldo = IIf(Left(psPlaCod, 1) = "E", "SUELDO", IIf(Left(psPlaCod, 1) = "P", "ESTIPENDIO", "SUB.ECONOM"))
'    If rsCargo!bRHCtrAsist Then
'        lsPorMesLabT = "D.LAB/H.LAB"
'    Else
'        lsPorMesLabT = "  D.LAB"
'    End If
'    lsNumC = IIf(Left(psPlaCod, 1) = "E", "CUSPP", "")
'    lsCondicion = "CONDICION"
'    lsAgencia = "AGENCIA"
'    If Left(psPlaCod, 1) = "E" Then
'        If Not IsNull(rsP!cRHAFPAbreviatura) Then
'            lsAFP = "AFP"
'        Else
'            lsAFP = "SNP"
'        End If
'    Else
'        lsAFP = ""
'    End If
'    lsVacaciones = IIf(Left(psPlaCod, 1) = "E", "VACACIONES", "")
'
'    lsCadena = lsCadena & lsNegritaOn & Space(lnEspacio) & lsSueldo & lsPorMesLabT & Space(2) & lsNumC & Space(2) & lsAFP & Space(1) & lsCondicion & lsAgencia & Space(4) & lsVacaciones & oImpresora.gPrnSaltoLinea & lsNegritaOff
'
'    lsSueldo = Format(oRH.GetSueldo(psPersCod), "#,##0.00")
'
'    If Left(psPlaCod, 1) = "E" Then
'        If Not IsNull(rsP!nRHContratoTpo) Then
'            If rsP!nRHContratoTpo = 1 Then
'                lsCondicion = "CONTRATADO"
'            Else
'                lsCondicion = "ESTABLE"
'            End If
'        Else
'            lsCondicion = "CONTRATADO"
'        End If
'    Else
'        lsCondicion = ""
'    End If
'    lsAgencia = rsP!cAgenciaActual
'    lsAFP = rsP!cRHAFPAbreviatura & ""
'    lsNumCarnetAFP = rsP!cCUSPP & ""
'
'    If Left(psPlaCod, 1) = "E" Then
'        If oRH.GetFecVac(psPersCod, ldFecVacIni, ldFecVacFin) Then
'            lsVacaciones = Format(ldFecVacIni, gsFormatoFechaView) & "-" & Format(ldFecVacFin, gsFormatoFechaView)
'        Else
'            lsVacaciones = "__/__/____-__/__/____"
'        End If
'    Else
'        lsVacaciones = ""
'    End If
'
'    lnI = 0
'    lnD = 0
'    lnA = 0
'    rsP.MoveFirst
'    While Not rsP.EOF
'        If rsP!nMonto <> 0 Or Trim(rsP!cRHConceptoCod) = gsRHConceptoITOTING Or Trim(rsP!cRHConceptoCod) = gsRHConceptoDTOTDES Then
'            If Left(rsP!cRHConceptoCod, 2) = "I_" Or Left(rsP!cRHConceptoCod, 2) = "J_" Then
'                lnI = lnI + 1
'                If IsNull(rsP!cRHConceptoAbreviatura) Then
'                    lsConcep(lnI, 1) = Trim(rsP!cRHConceptoCod)
'                Else
'                    lsConcep(lnI, 1) = Trim(rsP!cRHConceptoAbreviatura)
'                End If
'                lsConcep(lnI, 2) = FillNum(Format(rsP!nMonto, "#,##0.00"), 15, " ")
'            ElseIf Left(rsP!cRHConceptoCod, 2) = "D_" Or Left(rsP!cRHConceptoCod, 2) = "E_" Then
'                lnD = lnD + 1
'                If IsNull(rsP!cRHConceptoAbreviatura) Then
'                    lsConcep(lnD, 3) = Trim(rsP!cRHConceptoCod)
'                Else
'                    lsConcep(lnD, 3) = Trim(rsP!cRHConceptoAbreviatura)
'                End If
'
'                lsConcep(lnD, 4) = FillNum(Format(rsP!nMonto, "#,##0.00"), 15, " ")
'            ElseIf Left(rsP!cRHConceptoCod, 2) = "A_" Or Left(rsP!cRHConceptoCod, 2) = "B_" Then
'                lnA = lnA + 1
'                If IsNull(rsP!cRHConceptoAbreviatura) Then
'                    lsConcep(lnA, 5) = Trim(rsP!cRHConceptoCod)
'                Else
'                    lsConcep(lnA, 5) = Trim(rsP!cRHConceptoAbreviatura)
'                End If
'                lsConcep(lnA, 6) = FillNum(Format(rsP!nMonto, "#,##0.00"), 15, " ")
'            Else
'                lsPorMesLab = Trim(Str(rsP!nMonto))
'            End If
'        End If
'        rsP.MoveNext
'    Wend
'
'    If IsNumeric(lsPorMesLab) Then
'
'        If psPlaCod = gsRHPlanillaSueldos Or psPlaCod = "P01" Or psPlaCod = "F01" Then
'            If rsCargo!bRHCtrAsist Then
'                lsPorMesLab = Round(Trim(Str(CCur(lsPorMesLab) * (30 - GetDiasFediados(CDate(Right(pdFecha, 2) & "/" & Mid(pdFecha, 5, 2) & "/" & Left(pdFecha, 4)), CCur(lsPorMesLab))))), 2) & " / " & Trim(Str(GetHorasTrabajadas(CDate(Right(pdFecha, 2) & "/" & Mid(pdFecha, 5, 2) & "/" & Left(pdFecha, 4)), CCur(lsPorMesLab) * 30)))
'            Else
'                lsPorMesLab = Round(Trim(Str(CCur(lsPorMesLab) * (30 - GetDiasFediados(CDate(Right(pdFecha, 2) & "/" & Mid(pdFecha, 5, 2) & "/" & Left(pdFecha, 4)), CCur(lsPorMesLab))))), 2)
'            End If
'        Else
'            lsPorMesLab = Round(Trim(Str(CCur(lsPorMesLab))), 4)
'        End If
'    Else
'        lsPorMesLab = ""
'    End If
'
'    lsCadena = lsCadena & lsSueldo & lsPorMesLab & Space(5) & lsNumCarnetAFP & Space(2) & lsAFP & Space(1) & lsCondicion & Space(3) & lsAgencia & lsVacaciones & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & Space(lnEspacio) & lsSueldo & lsPorMesLab & lsNumCarnetAFP & lsAFP & lsCondicion & lsAgencia & lsVacaciones & oImpresora.gPrnSaltoLinea
'    lsCadena = oImpresora.gPrnSaltoLinea & lsCadena
'    lsCadena = lsNegritaOn & Space(lnEspacio) & Space(lnEspacio) & CentrarCadena("BOLETA DE PAGO " & Trim(psNombrePlanilla) & " " & Format(CDate(psFechaImprePlanilla), "MMMM-YYYY"), 100) & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
'    lsCadena = lsNegritaOn & Space(lnEspacio) & "RUC:" & pgsRuc & CentrarCadena(pgsEmpresa, 65) & psFechaImprePlanilla & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
'    lsCadena = lsNegritaOn & Space(lnEspacio) & Space(lnEspacio) & CentrarCadena(psTextoPlanilla, 100) & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
'
'    'lsCadena = lsCadena & lsNegritaOn & Encabezado("INGRESOS;16; ;16;DESCUENTOS;16; ;16;EMPLEADOR;16; ;16;", 0, lnEspacio) & lsNegritaOff
'
'    If lnI <> 0 Or lnD <> 0 And lnA <> 0 Then
'        lsConcepT(1, 1) = lsConcep(lnI, 1)
'        lsConcepT(1, 2) = lsConcep(lnI, 2)
'        If lnD > 0 Then lsConcepT(1, 3) = lsConcep(lnD, 3)
'        If lnD > 0 Then
'            lsConcepT(1, 4) = lsConcep(lnD, 4)
'        Else
'            lsConcepT(1, 4) = "0.00"
'        End If
'        lsConcepT(1, 5) = "NETO A PAGAR"
'        lsConcepT(1, 6) = FillNum(Format(CCur(lsConcepT(1, 2) - CCur(lsConcepT(1, 4))), "#,##0.00"), 15, " ")
'        lsConcep(lnI, 1) = ""
'        lsConcep(lnI, 2) = ""
'        If lnD > 0 Then lsConcep(lnD, 3) = ""
'        If lnD > 0 Then lsConcep(lnD, 4) = ""
'        If lnA > 0 Then lsConcep(lnA, 5) = ""
'        If lnA > 0 Then lsConcep(lnA, 6) = ""
'        lnI = lnI - 1
'        lnD = lnD - 1
'        lnA = lnA - 1
'    End If
'
'
'    lnPosMax = 1
'    While lnPosMax <= lnI Or lnPosMax <= lnD Or lnPosMax <= lnA
'        lsCadena = lsCadena & Space(lnEspacio) & lsConcep(lnPosMax, 1) & lsConcep(lnPosMax, 2) _
'                 & "  " & lsConcep(lnPosMax, 3) & lsConcep(lnPosMax, 4) _
'                 & "  " & lsConcep(lnPosMax, 5) & lsConcep(lnPosMax, 6) & oImpresora.gPrnSaltoLinea
'        lnPosMax = lnPosMax + 1
'    Wend
'
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsNegritaOn & Space(lnEspacio) & lsConcepT(1, 1) & lsConcepT(1, 2) _
'             & " " & lsConcepT(1, 3) & lsConcepT(1, 4) _
'             & " " & lsConcepT(1, 5) & lsConcepT(1, 6) & lsNegritaOff
'    If lbNSF Then
'        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & "(*) NSF : NO SUJETO A FISCALIZACION" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'    End If
'    Set rsP = Nothing
'    Set oPla = Nothing
'    Set oRH = Nothing
'
'    GetBoleta = lsCadena
'End Function

Private Function GetBoleta(psPersCod As String, pdFecha As String, psPlaCod As String, psNombrePlanilla As String, psFechaImprePlanilla As String, pgsRuc As String, pgsEmpresa As String, psTextoPlanilla As String) As String
    On Error Resume Next

    Dim lsCadena As String
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    Dim rsH1 As ADODB.Recordset
    Set rsH1 = New ADODB.Recordset
    Dim lsCod As String * 8
    
    Dim lsNombre As String * 24
    Dim lsCargo As String * 24
    '********************************************************************
    '** Modificado por ELRO 20110809 10:50:00 según Acta Nº 188-2011/TI-D
    Dim lsPuesto1 As String
    Dim lsPuesto2 As String * 24
    Dim lnPosicion1, lnPosicion2 As Integer
    '********************************************************************
    Dim lsAFP As String * 24
    Dim lsNumCarnetAFP As String * 24
    Dim lsFecIng As String * 24
    Dim lsFecIniVaca As String * 24
    Dim lsFecFinVaca As String * 24
    
    Dim slAFP As String ' nombre de la Entidad Afiliada ( AFP  o SNP)
    Dim slCargo As String 'nombre del Cargo
    
    Dim FechaIni As String
    Dim FechaFin As String
    Dim FechaFinF As String
    Dim Año As String
    Dim Mes As String
    
    Dim RemBasica As Currency
    Dim RemCargo As Currency
    Dim AsigFamiliar As Currency
    Dim SegVida As Currency
    Dim ComSegVida As Currency
    Dim SegSoc As Currency
    Dim Imp5ta As Currency
    Dim TotIng As Currency
    Dim TotDeu As Currency
    Dim TotPag As Currency
    Dim nTotalAporte As Currency
    
    
    Dim lsCodCta As String * 15
    Dim lsNumC As String * 20
    Dim lsSueldo As String * 15
    Dim lsPorMesLabT As String * 10
    Dim lsPorMesLab As String * 4
    
    Dim lsHorasLab As String * 4
    Dim lsHorasExt As String * 4
    Dim lsInasistencias As String * 4
    Dim lsSuspendidos As String * 4
    Dim lsTardanzas As String * 4
    
    Dim lsCondiciaon As String * 16
    Dim lsCondicion As String * 16
    Dim lsAgencia As String * 10
    Dim lsVacaciones As String * 22
    
    Dim lsConcep(20, 6) As String * 15
    Dim lsConcepT(1, 6) As String * 15

    Dim I As Integer
    Dim lnI As Integer
    Dim lnD As Integer
    Dim lnA As Integer
    Dim lnPosMax As Integer
    Dim lsCondicionE As String
    Dim lsNumCE As String
    Dim ldFecVacIni As String
    Dim ldFecVacFin As String
    
    Dim lnEspacio As Integer
    Dim lnnI As Integer
    Dim lnnJ As Integer
    
    Dim lsCodCuenta As String
    
    For lnnI = 0 To 20
        For lnnJ = 0 To 6
            lsConcep(lnnI, lnnJ) = ""
        Next lnnJ
    Next lnnI
    
    
    lnEspacio = 1

    Set rsP = oPla.GetBoleta(psPersCod, pdFecha, psPlaCod)
    
    lnI = 0
    lnD = 0
    If rsP.EOF And rsP.BOF Then
        rsP.Close
        Set rsP = Nothing
        Exit Function
    End If
    
        
    lsCod = rsP!cRHCod 'pendiente para la cabecera
    lsSuspendidos = ""
    
    lsCodCta = oRH.GetCuentaRRHH(psPersCod, gsRHPlanillaSueldos) 'cuenta de ahorros
    lsNombre = PstaNombre(rsP!cPersNombre, False)
    lsCargo = oRH.GetCargo(psPersCod, pdFecha)
    '********************************************************************
    '** Modificado por ELRO 20110809 10:50:00 según Acta Nº 188-2011/TI-D
    lnPosicion1 = InStr(1, lsCargo, "-")
    If lnPosicion1 > 1 Then
        lsCargo = Left(lsCargo, lnPosicion1 - 1)
    End If
    
    lsPuesto1 = oRH.GetCargo(psPersCod, pdFecha)
    lnPosicion2 = InStr(1, lsPuesto1, "-")
    If lnPosicion2 > 1 Then
        lsPuesto2 = Mid(lsPuesto1, lnPosicion2 + 2, 24)
    Else
        lsPuesto2 = ""
    End If
    '********************************************************************
    
    lsAFP = IIf(IsNull(rsP!cRHAFPAbreviatura), "", rsP!cRHAFPAbreviatura)
    lsNumCarnetAFP = rsP!cCUSPP
    lsFecIng = Format(rsP!dIngreso, gsFormatoFechaView)
      
'    If Left(psPlaCod, 3) = "E06" Then
    If oRH.GetFecVac(psPersCod, ldFecVacIni, ldFecVacFin) = True Then
        lsFecIniVaca = Format(ldFecVacIni, gsFormatoFechaView)
        lsFecFinVaca = Format(ldFecVacFin, gsFormatoFechaView)
    Else
        lsFecIniVaca = ""
        lsFecFinVaca = ""
    End If

    'JEOM
    '******************************************
    Dim rsAge As ADODB.Recordset
    Set rsAge = New ADODB.Recordset
    Dim oAge As DRHConcepto
    Set oAge = New DRHConcepto
    
    Dim lsAgenciaPer As String * 24
    Dim lsAreaPer As String * 24
    
    If lsCod <> "" Then
       Set rsAge = oAge.GetDatosPersonaPlanilla(lsCod, pdFecha)

       If Not rsAge.BOF And Not rsAge.EOF Then
          lsAgenciaPer = rsAge!Agencia
          lsAreaPer = rsAge!Area
       End If
    End If
    '******************************************
    
 
     
    Año = Left(pdFecha, 4)
    Mes = Mid(pdFecha, 5, 2)
    FechaIni = "01" & "/" & Mes & "/" & Año
    FechaFin = DateSerial(Año, Mes + 1, 0)
           
    FechaIni = Mes & "/" & "01" & "/" & Año
    FechaFinF = Mid(FechaFin, 4, 2) & "/" & Mid(FechaFin, 1, 2) & "/" & Mid(FechaFin, 7, 4)
    
    Set rsH1 = oRH.GetHorasLaboradas_Extras(FechaIni, FechaFinF, psPersCod)
    
    If Not rsH1.EOF Or Not rsH1.BOF Then
        lsHorasLab = rsH1!horasLaboradas
        lsHorasExt = rsH1!HE
        lsInasistencias = rsH1!Inasistencia
        lsTardanzas = rsH1!Tardanza
    Else
        lsHorasLab = ""
        lsHorasExt = ""
        lsInasistencias = ""
        lsTardanzas = ""
    End If
    lsSueldo = Format(oRH.GetSueldo(psPersCod), "#,##0.00")
    
    lsAgencia = rsP!cAgenciaActual
    slCargo = IIf(Left(psPlaCod, 1) = "E", "CARGO:", "CONVENIO:") 'para titulo Cargo
    If Left(psPlaCod, 1) = "E" Then 'Para titulo AFP
        If Not IsNull(rsP!cRHAFPAbreviatura) Then
            slAFP = "AFP"
        Else
            slAFP = "SNP"
        End If
    Else
        lsAFP = ""
    End If
    
    lsNumC = IIf(Left(psPlaCod, 1) = "E", "", "")
      
    lnI = 0
    lnD = 0
    lnA = 0
    rsP.MoveFirst
    While Not rsP.EOF
        If rsP!nMonto <> 0 Or Trim(rsP!cRHConceptoCod) = gsRHConceptoITOTING Or Trim(rsP!cRHConceptoCod) = gsRHConceptoDTOTDES Then
            If Left(rsP!cRHConceptoCod, 2) = "I_" Or Left(rsP!cRHConceptoCod, 2) = "J_" Then
                lnI = lnI + 1
                If IsNull(rsP!cRHConceptoAbreviatura) Then
                    lsConcep(lnI, 1) = Trim(rsP!cRHConceptoCod)
                Else
                    lsConcep(lnI, 1) = Trim(rsP!cRHConceptoAbreviatura)
                End If
                lsConcep(lnI, 2) = FillNum(Format(rsP!nMonto, "#,##0.00"), 15, " ")
            ElseIf Left(rsP!cRHConceptoCod, 2) = "D_" Or Left(rsP!cRHConceptoCod, 2) = "E_" Then
                lnD = lnD + 1
                If IsNull(rsP!cRHConceptoAbreviatura) Then
                    lsConcep(lnD, 3) = Trim(rsP!cRHConceptoCod)
                Else
                    lsConcep(lnD, 3) = Trim(rsP!cRHConceptoAbreviatura)
                End If
                
                lsConcep(lnD, 4) = FillNum(Format(rsP!nMonto, "#,##0.00"), 15, " ")
            ElseIf Left(rsP!cRHConceptoCod, 2) = "A_" Or Left(rsP!cRHConceptoCod, 2) = "B_" Then
                lnA = lnA + 1
                If IsNull(rsP!cRHConceptoAbreviatura) Then
                    lsConcep(lnA, 5) = Trim(rsP!cRHConceptoCod)
                Else
                    lsConcep(lnA, 5) = Trim(rsP!cRHConceptoAbreviatura)
                End If
                lsConcep(lnA, 6) = FillNum(Format(rsP!nMonto, "#,##0.00"), 15, " ")
            Else
                lsPorMesLab = Trim(Str(rsP!nMonto))
            End If
        End If
        rsP.MoveNext
    Wend
        
    If IsNumeric(lsPorMesLab) Then 'Dias Laborados
        If psPlaCod = gsRHPlanillaSueldos Or psPlaCod = "P01" Or psPlaCod = "F01" Then
            'lsPorMesLab = Round(Trim(Str(CCur(lsPorMesLab) * 30)), 4)
            lsPorMesLab = Round(Trim(Str(CCur(lsPorMesLab))), 4)
        Else
            lsPorMesLab = Round(Trim(Str(CCur(lsPorMesLab))), 4)
        End If
    Else
        lsPorMesLab = ""
    End If
    
        
        If lnI <> 0 Or lnD <> 0 And lnA <> 0 Then
        lsConcepT(1, 1) = lsConcep(lnI, 1)
        lsConcepT(1, 2) = lsConcep(lnI, 2)
        If lnD > 0 Then lsConcepT(1, 3) = lsConcep(lnD, 3)
        If lnD > 0 Then
            lsConcepT(1, 4) = lsConcep(lnD, 4)
        Else
            lsConcepT(1, 4) = "0.00"
        End If
        lsConcepT(1, 5) = "NETO A PAGAR"
        lsConcepT(1, 6) = FillNum(Format(CCur(lsConcepT(1, 2) - CCur(lsConcepT(1, 4))), "#,##0.00"), 15, " ")
        lsConcep(lnI, 1) = ""
        lsConcep(lnI, 2) = ""
        If lnD > 0 Then lsConcep(lnD, 3) = ""
        If lnD > 0 Then lsConcep(lnD, 4) = ""
        If lnA > 0 Then lsConcep(lnA, 5) = ""
        If lnA > 0 Then lsConcep(lnA, 6) = ""
        lnI = lnI - 1
        lnD = lnD - 1
        lnA = lnA - 1
    End If
    
   
 
    'RemBasica = lsConcep(1, 2)
    ' AsigFamiliar= lsConcep(lnPosMax(2), 1) 'Asignacion Familiar
    'AsigFamiliar = lsConcep(2, 2) 'monto asignacion familiar
    'RemCargo = lsConcep(lnPosMax(1), 1) 'titulo Remuneracion al cargo
    'RemCargo = lsConcep(1, 2) 'monto
    'AsigFamiliar = lsConcep(lnPosMax(4), 1) 'titulo Monto Asignacion Familiar
    'AsigFamiliar = lsConcep(4, 2) 'monto Asignacion familiar
    'SegVida = lsConcep(lnPosMax(3), 3) 'AFP titulo
    'SegVida = lsConcep(3, 4) 'AFP monto
    'ComSegVida = lsConcep(4, 4) 'AFP Comision
    'Imp5ta = lsConcep(lnPosMax(2), 3)'titulo impuesto 5 categoria
    'Imp5ta = lsConcep(1, 4) 'monto impuesto 5 categoria
   
    'SegSoc = lsConcep(lnPosMax(1), 5)'titulo seguro social ESSALUD
    'SegSoc = lsConcep(1, 6) 'Monto
        'lsConcep(lnPosMax(4), 3)Cuota Fija titulo   'lsConcep(lnPosMax(4), 4) 'monto
    'lsConcep(lnPosMax(5), 3) & lsConcep(lnPosMax(5), 4) 'Comision  COM. Vari
    'lsConcepT(1, 1) & lsConcepT(1, 2)' Total de Ingresos
    TotIng = lsConcepT(1, 2)
    'lsConcepT(1, 3) & lsConcepT(1, 4)' total DeDu
    TotDeu = lsConcepT(1, 4)
    'lsConcepT(1, 5) & lsConcepT(1, 6) 'Neto a Pagar
    TotPag = lsConcepT(1, 6)
    
    
    
    lsCadena = lsCadena & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
    lsCadena = lsCadena & Space(18) & "BOLETA MENSUAL " & Space(24) & "RUC" & Space(1) & pgsRuc & Space(30) & Space(18) & "BOLETA MENSUAL " & Space(24) & "RUC" & Space(1) & pgsRuc & Chr(10) & Chr(10)
    lsCadena = lsCadena & "DEPOSITADO EN LA CTA. DE AHORROS No." & Space(3) & "" & lsCodCta & Space(4) & "Codigo :" & lsCod & Space(31) & "DEPOSITADO EN LA CTA. DE AHORROS No." & Space(3) & "" & lsCodCta & Space(4) & "Codigo :" & lsCod & Chr(10)
    lsCadena = lsCadena & String(78, "_") & Space(27) & String(78, "_") & Chr(10)
    
    lsCadena = lsCadena & "Apellidos y Nombres :" & Space(2) & lsNombre & Space(4) & "Dias Laborados" & String(2, ".") & ":" & Space(2) & lsPorMesLab & Space(31) & "Apellidos y Nombres :" & Space(2) & lsNombre & Space(4) & "Dias Laborados" & String(2, ".") & ":" & Space(2) & lsPorMesLab & Chr(10)
    lsCadena = lsCadena & slCargo & String(14, ".") & ":" & Space(2) & lsCargo & Space(4) & "Horas Laboradas :" & Space(2) & lsHorasLab & Space(31) & slCargo & String(14, ".") & ":" & Space(2) & lsCargo & Space(4) & "Horas Laboradas :" & Space(2) & lsHorasLab & Chr(10)
    lsCadena = lsCadena & slAFP & String(17, ".") & ":" & Space(2) & lsAFP & Space(4) & "Horas Extras" & String(4, ".") & ":" & Space(2) & lsHorasExt & Space(31) & slAFP & String(17, ".") & ":" & Space(2) & lsAFP & Space(4) & "Horas Extras" & String(4, ".") & ":" & Space(2) & lsHorasExt & Chr(10)
    lsCadena = lsCadena & "C.U.I" & String(15, ".") & ":" & Space(2) & lsNumCarnetAFP & Space(4) & "Inasistencias" & String(2, ".") & " :" & Space(2) & lsInasistencias & Space(31) & "C.U.I" & String(15, ".") & ":" & Space(2) & lsNumCarnetAFP & Space(4) & "Inasistencias" & String(2, ".") & " :" & Space(2) & lsInasistencias & Chr(10)
    lsCadena = lsCadena & "Fecha de Ingreso " & String(3, ".") & ":" & Space(2) & lsFecIng & Space(4) & "Suspendidos " & String(3, ".") & " :" & Space(2) & lsSuspendidos & Space(31) & "Fecha de Ingreso " & String(3, ".") & ":" & Space(2) & lsFecIng & Space(4) & "Suspendidos " & String(3, ".") & " :" & Space(2) & lsSuspendidos & Chr(10)
    lsCadena = lsCadena & "F/Inicio Vacaciones :" & Space(2) & lsFecIniVaca & Space(4) & "Tardanzas " & String(5, ".") & " :" & Space(2) & lsTardanzas & Space(31) & "F/Inicio Vacaciones :" & Space(2) & lsFecIniVaca & Space(4) & "Tardanzas " & String(5, ".") & " :" & Space(2) & lsTardanzas & Chr(10)
    lsCadena = lsCadena & "F/Fin de Vacaciones :" & Space(2) & lsFecFinVaca & Space(58) & "F/Fin de Vacaciones :" & Space(2) & lsFecFinVaca & Chr(10)
    lsCadena = lsCadena & "Agencia             :" & Space(2) & lsAgenciaPer & Space(58) & "Agencia             :" & Space(2) & lsAgenciaPer & Chr(10)
    lsCadena = lsCadena & "Area                :" & Space(2) & lsAreaPer & Space(58) & "Area                :" & Space(2) & lsAreaPer & Chr(10)
    '********************************************************************
    '** Modificado por ELRO 20110809 10:50:00 según Acta Nº 188-2011/TI-D
    lsCadena = lsCadena & lsPuesto2 & Space(81) & lsPuesto2 & Chr(10)
    '********************************************************************
    lsCadena = lsCadena & String(78, "_") & Space(27) & String(78, "_") & Chr(10)
    
    lsCadena = lsCadena & Space(49) & "DESCUENTOS " & Space(5) & "APORTES " & Space(25) & Space(56) & "DESCUENTOS " & Space(5) & "APORTES " & Chr(10)
    lsCadena = lsCadena & Space(2) & "CONCEPTOS " & Space(24) & "INGRESOS " & Space(20) & Space(42) & "CONCEPTOS " & Space(24) & "INGRESOS " & Chr(10)
    lsCadena = lsCadena & Space(48) & "DEL EMPLEADO " & Space(3) & "EMPLEADOR " & Space(25) & Space(54) & "DEL EMPLEADO " & Space(3) & "EMPLEADOR " & Chr(10)
    lsCadena = lsCadena & String(78, "_") & Space(27) & String(78, "_") & Chr(10)
    
    
    lnPosMax = 1
    While lnPosMax <= lnI Or lnPosMax <= lnD Or lnPosMax <= lnA
        If Len(Trim(lsConcep(lnPosMax, 1))) > 0 Then
'        rsP.MoveFirst
'        lsCadena = rsP!concepcod
           lsCadena = lsCadena & Space(lnEspacio) & Space(10) & lsConcep(lnPosMax, 1) & lsConcep(lnPosMax, 2) & _
                      Space(75) & lsConcep(lnPosMax, 1) & lsConcep(lnPosMax, 2) _
                      & Chr(10)
           
        End If
        lnPosMax = lnPosMax + 1
    Wend
    lnPosMax = 1
   While lnPosMax <= lnI Or lnPosMax <= lnD Or lnPosMax <= lnA
        If Len(Trim(lsConcep(lnPosMax, 3))) > 0 Then
           lsCadena = lsCadena & Space(lnEspacio) & Space(10) & lsConcep(lnPosMax, 3) & Space(15) & lsConcep(lnPosMax, 4) & _
           Space(60) & lsConcep(lnPosMax, 3) & Space(15) & lsConcep(lnPosMax, 4) _
                 & Chr(10)
        End If
        lnPosMax = lnPosMax + 1
    Wend
    
    lnPosMax = 1
    While lnPosMax <= lnI Or lnPosMax <= lnD Or lnPosMax <= lnA
        If Len(Trim(lsConcep(lnPosMax, 5))) > 0 Then
            lsCadena = lsCadena & Space(lnEspacio) & Space(10) & lsConcep(lnPosMax, 5) & Space(30) & lsConcep(lnPosMax, 6) & _
            Space(45) & lsConcep(lnPosMax, 5) & Space(30) & lsConcep(lnPosMax, 6) _
            & Chr(10)
            nTotalAporte = nTotalAporte + lsConcep(lnPosMax, 6)
        End If
        lnPosMax = lnPosMax + 1
    Wend
        
    lsCadena = lsCadena & Chr(10)
    lsCadena = lsCadena & "TOTAL INGRESO: " & Trim(lsConcepT(1, 2)) & Space(1) & "TOTAL DESCUENTO: " & Trim(lsConcepT(1, 4)) & Space(1) & "APORTE EMPLEADOR: " & Trim(nTotalAporte) & Space(10 - Len(Trim(nTotalAporte))) & Space(30) & "TOTAL INGRESO: " & Trim(lsConcepT(1, 2)) & Space(1) & "TOTAL DESCUENTO: " & Trim(lsConcepT(1, 4)) & Space(1) & "APORTE EMPLEADOR: " & Trim(nTotalAporte) & Chr(10)
                 
    lsCadena = lsCadena & String(78, "_") & Space(27) & String(78, "_") & Chr(10)
    lsCadena = lsCadena & Space(2) & "**MES" & Space(2) & ArmaFecha(CDate(psFechaImprePlanilla)) & Space(18) & "TOTAL NETO:" & lsConcepT(1, 6) & Space(37) & "**MES" & Space(2) & ArmaFecha(CDate(psFechaImprePlanilla)) & Space(18) & "TOTAL NETO: " & lsConcepT(1, 6) & Chr(10)
    lsCadena = lsCadena & String(78, "_") & Space(27) & String(78, "_") & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
    
    lsCadena = lsCadena & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10) & Chr(10)
    
    'lsCadena = lsCadena & Space(13) & "_ _ _ _ _ _ _ _ _ _ _" & Space(19) & "_ _ _ _ _ _ _ _ _ _ _" & Space(30) & Space(13) & "_ _ _ _ _ _ _ _ _ _ _" & Space(19) & "_ _ _ _ _ _ _ _ _ _ _" & Chr(10)
    'lsCadena = lsCadena & Space(10) & "   Firma del Empleador   " & Space(18) & "Firma del Trabajador " & Space(40) & "   Firma del Empleador   " & Space(18) & "Firma del Trabajador " & Chr(10)
    'lsCadena = lsCadena & Space(10) & "   y/o Jefe de Personal" & Space(59) & Space(25) & "y/o Jefe de Personal" & Chr(10)
    
    'lsCadena = lsCadena & Chr(12)
    
    Set rsP = Nothing
    Set oPla = Nothing
    Set oRH = Nothing
    
    GetBoleta = lsCadena
End Function



'Public Function GetListaPagos(psPlanillaCod As String, psRRHHPeriodo As String, psFecha As String, psTipcambio As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, psTexto As String) As String
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim lsCadena As String
'    Dim lnItem As Long
'    Dim lnPagina As Long
'    Dim lsCodigo As String * 10
'    Dim lsNombre As String * 40
'    Dim lsCuenta As String * 20
'    Dim lsMonto As String * 10
'    Dim lsMontoDolar As String * 10
'
'    Dim lsMonedaAnt As String
'    Dim lsTipoEmpAnt As String
'    Dim lsAgenciaAnt As String
'    Dim lnMontoParte As Currency
'    Dim lnMontocambio As Currency
'    Dim lnMontoTotal As Currency
'    Dim oPla As DActualizaDatosConPlanilla
'    Set oPla = New DActualizaDatosConPlanilla
'
'    Dim lnTipCambio As Double
'
'    Set rs = oPla.GetLista(psRRHHPeriodo, psPlanillaCod)
'
'    lsCadena = ""
'    If Not (rs.EOF And rs.BOF) Then
'        lnTipCambio = rs!nTpoCambio
'        While Not rs.EOF
'            If IIf(IsNull(rs!Moneda), "3", rs!Moneda) <> lsMonedaAnt Or lsTipoEmpAnt <> rs!CONDICION Or lsAgenciaAnt <> rs!cAgenciaActual Then
'
'                If lsCadena <> "" Then
'                    lsCadena = lsCadena & Encabezado(" ;65;" & Format(lnMontoParte, "#,##0.00") & ";10; " & Format(lnMontoParte / lnTipCambio, "#,##0.00") & ";20;", lnItem)
'                    lnMontoParte = 0
'                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'                End If
'
'                If Not IsNull(rs!cCtaCod) Then
'                    lsCadena = lsCadena & CabeceraPagina("LISTA DE PAGOS DE EMPLEADOS - " & rs!cAgenciaActual & " - " & rs!CONDICION & " - " & psPlanillaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, rs!Moneda)
'                    lsCadena = lsCadena & CentrarCadena(psPlanillaCod & " - " & psRRHHPeriodo & " - " & psTexto, 80) & oImpresora.gPrnSaltoLinea
'
'                    If rs!Moneda = "2" Then
'                        lsCadena = lsCadena & Encabezado("Codigo;10;nombre;20; ;20;Cuenta;15;Monto;10; Monto Cambio;20;", lnItem)
'                    Else
'                        lsCadena = lsCadena & Encabezado("Codigo;10;nombre;20; ;20;Cuenta;15;Monto;10; ;45;", lnItem)
'                    End If
'                Else
'                    lsCadena = lsCadena & CabeceraPagina("LISTA DE PAGOS DE EMPLEADOS QUE NO TIENEN CUENTA EN LA CMACT - " & rs!CONDICION & " - " & psPlanillaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "NNN")
'                    lsCadena = lsCadena & Encabezado("Codigo;10;nombre;20; ;20;Cuenta;15;Monto;10; ;45;", lnItem)
'                End If
'            End If
'
'            lnMontoParte = lnMontoParte + rs!nMonto
'            lnMontoTotal = lnMontoTotal + rs!nMonto
'
'            lsCodigo = rs!cRHCod
'            lsNombre = rs!cPersNombre
'            lsCuenta = IIf(IsNull(rs!cCtaCod), "", rs!cCtaCod)
'            RSet lsMonto = Format(rs!nMonto, "#,##0.00")
'
'            If IIf(IsNull(rs!Moneda), "3", rs!Moneda) = "2" Then
'                RSet lsMontoDolar = Format(rs!nMonto / rs!nTpoCambio, "#,##0.00")
'                lsCadena = lsCadena & lsCodigo & lsNombre & lsCuenta & lsMonto & lsMontoDolar & oImpresora.gPrnSaltoLinea
'            Else
'                lsMontoDolar = ""
'                lsCadena = lsCadena & lsCodigo & lsNombre & lsCuenta & lsMonto & oImpresora.gPrnSaltoLinea
'            End If
'
'            lnItem = lnItem + 1
'
'            lsMonedaAnt = IIf(IsNull(rs!Moneda), "3", rs!Moneda)
'            lsTipoEmpAnt = rs!CONDICION
'            lsAgenciaAnt = rs!cAgenciaActual
'
'            If lnItem > 54 Then
'                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'                lsCadena = lsCadena & CabeceraPagina("LISTA DE PAGOS DE EMPLEADOS - " & IIf(rs!CONDICION = "10", "CONTRATADO", "ESTABLE") & " - " & psPlanillaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, rs!Moneda)
'                lsCadena = lsCadena & CentrarCadena(psPlanillaCod & " - " & psRRHHPeriodo & " - " & psTexto, 80) & oImpresora.gPrnSaltoLinea
'                lsCadena = lsCadena & Encabezado("Codigo;10;nombre;20; ;20;Cuenta;15;Monto;15;", lnItem)
'            End If
'
'            rs.MoveNext
'        Wend
'    End If
'
'    If lsCadena <> "" Then
'        lsCadena = lsCadena & Encabezado(" ;65;" & Format(lnMontoParte, "#,##0.00") & ";10; " & Format(lnMontoParte / lnTipCambio, "#,##0.00") & ";20;", lnItem)
'        lsCadena = lsCadena & Encabezado(" ;65;" & Format(lnMontoTotal, "#,##0.00") & ";10; " & Format(lnMontoTotal / lnTipCambio, "#,##0.00") & ";20;", lnItem)
'        lnMontoParte = 0
'    End If
'
'    rs.Close
'    Set rs = Nothing
'    GetListaPagos = lsCadena
'End Function

Public Function GetListaPagos(psPlanillaCod As String, psRRHHPeriodo As String, psFecha As String, psTipcambio As String, pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lsCadena As String
    Dim lnItem As Long
    Dim lnPagina As Long
    Dim lsCodigo As String * 10
    Dim lsNombre As String * 40
    Dim lsCuenta As String * 20
    Dim lsMonto As String * 10
    Dim lsMontoDolar As String * 10
    
    Dim lsMonedaAnt As String
    Dim lsTipoEmpAnt As String
    Dim lsAgenciaAnt As String
    Dim lnMontoParte As Currency
    Dim lnMontocambio As Currency
    Dim lnMontoTotal As Currency
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    Dim lnTipCambio As Double
    
    Set rs = oPla.GetLista(psRRHHPeriodo, psPlanillaCod)
    
    lsCadena = ""
    If Not (rs.EOF And rs.BOF) Then
        lnTipCambio = rs!nTpoCambio
        While Not rs.EOF
            If IIf(IsNull(rs!Moneda), "3", rs!Moneda) <> lsMonedaAnt Or lsTipoEmpAnt <> rs!Condicion Or lsAgenciaAnt <> rs!cAgenciaActual Then
                
                If lsCadena <> "" Then
                    lsCadena = lsCadena & Encabezado(" ;65;" & Format(lnMontoParte, "#,##0.00") & ";10; " & Format(lnMontoParte / lnTipCambio, "#,##0.00") & ";20;", lnItem)
                    lnMontoParte = 0
                    lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                End If
                
                If Not IsNull(rs!cCtaCod) Then
                    lsCadena = lsCadena & CabeceraPagina("LISTA DE PAGOS DE EMPLEADOS - " & rs!cAgenciaActual & " - " & rs!Condicion & " - " & psPlanillaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, rs!Moneda)
                    If rs!Moneda = "2" Then
                        lsCadena = lsCadena & Encabezado("Codigo;10;nombre;20; ;20;Cuenta;15;Monto;10; Monto Cambio;20;", lnItem)
                    Else
                        lsCadena = lsCadena & Encabezado("Codigo;10;nombre;20; ;20;Cuenta;15;Monto;10; ;45;", lnItem)
                    End If
                Else
                    lsCadena = lsCadena & CabeceraPagina("LISTA DE PAGOS DE EMPLEADOS QUE NO TIENEN CUENTA EN LA CMACI - " & rs!Condicion & " - " & psPlanillaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "NNN")
                    lsCadena = lsCadena & Encabezado("Codigo;10;nombre;20; ;20;Cuenta;15;Monto;10; ;45;", lnItem)
                End If
            End If
                
            lnMontoParte = lnMontoParte + rs!nMonto
            lnMontoTotal = lnMontoTotal + rs!nMonto
            
            lsCodigo = rs!cRHCod
            lsNombre = rs!cPersNombre
            lsCuenta = IIf(IsNull(rs!cCtaCod), "", rs!cCtaCod)
            RSet lsMonto = Format(rs!nMonto, "#,##0.00")
            
            If IIf(IsNull(rs!Moneda), "3", rs!Moneda) = "2" Then
                RSet lsMontoDolar = Format(rs!nMonto / rs!nTpoCambio, "#,##0.00")
                lsCadena = lsCadena & lsCodigo & lsNombre & lsCuenta & lsMonto & lsMontoDolar & oImpresora.gPrnSaltoLinea
            Else
                lsMontoDolar = ""
                lsCadena = lsCadena & lsCodigo & lsNombre & lsCuenta & lsMonto & oImpresora.gPrnSaltoLinea
            End If
            
            lnItem = lnItem + 1
            
            lsMonedaAnt = IIf(IsNull(rs!Moneda), "3", rs!Moneda)
            lsTipoEmpAnt = rs!Condicion
            lsAgenciaAnt = rs!cAgenciaActual
            
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                'lsCadena = lsCadena & CabeceraPagina("LISTA DE PAGOS DE EMPLEADOS - " & IIf(rs!Condicion = "10", "CONTRATADO", "ESTABLE") & " - " & psPlanillaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, rs!Moneda)
                lsCadena = lsCadena & CabeceraPagina("LISTA DE PAGOS DE EMPLEADOS - " & IIf(rs!Condicion = "10", "", "") & " - " & psPlanillaCod, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, rs!Moneda)
                lsCadena = lsCadena & Encabezado("Codigo;10;nombre;20; ;20;Cuenta;15;Monto;15;", lnItem)
            End If
            
            rs.MoveNext
        Wend
    End If
    
    If lsCadena <> "" Then
        'lsCadena = lsCadena & Encabezado(" ;65;" & Format(lnMontoParte, "#,##0.00") & ";10; " & Format(lnMontoParte / lnTipCambio, "#,##0.00") & ";20;", lnItem)
        'lsCadena = lsCadena & Encabezado(" ;65;" & Format(lnMontoTotal, "#,##0.00") & ";10; " & Format(lnMontoTotal / lnTipCambio, "#,##0.00") & ";20;", lnItem)
        
        lsCadena = lsCadena & Encabezado(" ;70;" & Format(lnMontoParte, "#,##0.00") & ";10;;20;", lnItem)
        lsCadena = lsCadena & Encabezado(" ;70;" & Format(lnMontoTotal, "#,##0.00") & ";10;;20;", lnItem)
        
        lnMontoParte = 0
    End If
    
    rs.Close
    Set rs = Nothing
    GetListaPagos = lsCadena
End Function
Public Function GetResumenporAgencias(psPlanillaPeriodo As String, psPlanillaCod As String, psFecImpre As String, pgsRuc As String, pgsEmpresa As String, psFecIni As String, psFecSin As String, psNombrePlanilla As String, psTextoPlanilla As String) As String
    Dim lsCadena As String
    Dim I As Integer
    Dim lnItem As Integer
    Dim lnPagina As Integer
    Dim ldFecAsisIni As Date
    Dim ldFecAsisFin As Date
    Dim ldFecAsisFinGra As Date
    Dim ldFecAsisIniGra As Date
    Dim lsEmpDes As String
    Dim sqlC As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim lsConAnt As String
    Dim lsAgeAnt As String
    Dim lsNomAge As String
    Dim oPla As DInterprete
    Set oPla = New DInterprete
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    
    lsCadena = ""
    oPla.ObtFechasAsistencia ldFecAsisIni, ldFecAsisFin, ldFecAsisIniGra, ldFecAsisFinGra, CDate(psFecIni)
    
    Set rsC = oCon.GetAgencias()
    lsEmpDes = ""
    
    lnPagina = 0
    
    If Not RSVacio(rsC) Then
        While Not rsC.EOF
            If lsAgeAnt <> rsC.Fields(1) Then
                lsNomAge = rsC.Fields(0)
                lnPagina = lnPagina + 1
                lnItem = 0
                lsCadena = lsCadena & GetCabeceraPlanilla(lnPagina, 1, "Resumen de Planilla : " & lsNomAge, ldFecAsisIni, ldFecAsisFin, psFecImpre, pgsRuc, pgsEmpresa, psFecIni, psFecSin, psNombrePlanilla, psTextoPlanilla)
                lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True, , 1, "Total Planilla de Personal Estable " & lsNomAge, rsC.Fields(1)) & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True, , 2, "Total Planilla de Personal Contratado " & lsNomAge, rsC.Fields(1)) & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True, , 0, "Total Planilla " & lsNomAge, rsC.Fields(1)) & oImpresora.gPrnSaltoLinea
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            End If
            lsAgeAnt = rsC.Fields(1)
            rsC.MoveNext
        Wend
    End If
    
    lsCadena = lsCadena & GetCabeceraPlanilla(lnPagina, 1, "TOTAL", ldFecAsisIni, ldFecAsisFin, psFecImpre, pgsRuc, pgsEmpresa, psFecIni, psFecSin, psNombrePlanilla, psTextoPlanilla)
    lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True, , 1, "Total Planilla de Personal Estable ") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True, , 2, "Total Planilla de Personal Contratado ") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True, , , "Total Planillas Empleados") & oImpresora.gPrnSaltoLinea
    
    GetResumenporAgencias = lsCadena
End Function


Public Function GetPlanillas(psPlanillaPeriodo As String, psPlanillaCod As String, psEmpLista As String, psNombrePlanilla As String, psFechaImprePlanilla As String, pgsRuc As String, pgsEmpresa As String, psFecIni As String, psFecFin As String, pbParaContabilidad As Boolean, Optional pbOrdenado As Boolean = False, Optional psTextoPlanilla As String = "") As String
    Dim lsCadena As String
    Dim I As Integer
    Dim lnItem As Integer
    Dim lnPagina As Integer
    Dim ldFecAsisIni As Date
    Dim ldFecAsisFin As Date
    Dim ldFecAsisFinGra As Date
    Dim ldFecAsisIniGra As Date
    Dim lsEmpDes As String
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim lsConAnt As String
    Dim lsAgeAnt As String
    Dim lsNomAge As String
    Dim oPla As DInterprete
    Set oPla = New DInterprete
    Dim oPlaRep As DActualizaDatosConPlanilla
    Set oPlaRep = New DActualizaDatosConPlanilla
    
    oPla.ObtFechasAsistencia ldFecAsisIni, ldFecAsisFin, ldFecAsisIniGra, ldFecAsisFinGra, CDate(psFechaImprePlanilla)
    
    lsCadena = ""
        
    Set rsC = oPlaRep.GetPlanillaRep(pbOrdenado, pbParaContabilidad, psPlanillaPeriodo, psPlanillaCod, psEmpLista)
    lnPagina = 0
    
    If Not RSVacio(rsC) Then
        RaiseEvent ShowProgress
        While Not rsC.EOF
            If pbOrdenado Then
                If rsC!TipCon <> lsConAnt Then
                    If pbOrdenado Then
                        lsNomAge = pgsEmpresa
                    Else
                        lsNomAge = rsC!cAgeDescripcion
                    End If
                    lnPagina = lnPagina + 1
                    lnItem = 0
                    If lnPagina <> 1 Then
                        lsCadena = lsCadena & GetItemPlanilla(rsC!cPersCod, psPlanillaPeriodo, psPlanillaCod, "", True, False, IIf(lsConAnt = "ESTABLE", 1, 2))
                        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                    End If
                    lsCadena = lsCadena & GetCabeceraPlanilla(lnPagina, 1, rsC!TipCon & " - " & lsNomAge, ldFecAsisIni, ldFecAsisFin, psFechaImprePlanilla, pgsRuc, pgsEmpresa, psFecIni, psFecFin, psNombrePlanilla, psTextoPlanilla)
                End If
            Else
                If pbParaContabilidad Then
                    If rsC!TipCon <> lsConAnt Or lsAgeAnt <> rsC!cCodAge Then
                        If pbOrdenado Then
                            lsNomAge = pgsEmpresa
                        Else
                            lsNomAge = rsC!cAgeDescripcion
                        End If
                        lnPagina = lnPagina + 1
                        lnItem = 0
                        If lnPagina <> 1 Then
                            lsCadena = lsCadena & GetItemPlanilla(rsC!cPersCod, psPlanillaPeriodo, psPlanillaCod, "", True, False, IIf(lsConAnt = "ESTABLE", 1, 2), , lsAgeAnt)
                            If lsAgeAnt <> rsC!cCodAge Then
                                lsCadena = lsCadena & GetItemPlanilla(rsC!cPersCod, psPlanillaPeriodo, psPlanillaCod, "", True, False, 0, , lsAgeAnt)
                            End If
                            'If rsC!TipCon <> lsConAnt Then
                            '    lsCadena = lsCadena & GetItemPlanilla(rsC!cEmpCod, psPlanillaPeriodo, psPlanillaCod, "", True, False, IIf(lsConAnt = "ESTABLE", 1, 2))
                            'End If
                            
                            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                        End If
                        lsCadena = lsCadena & GetCabeceraPlanilla(lnPagina, 1, rsC!TipCon & " - " & lsNomAge, ldFecAsisIni, ldFecAsisFin, psFechaImprePlanilla, pgsRuc, pgsEmpresa, psFecIni, psFecFin, psNombrePlanilla, psTextoPlanilla)
                    End If
                Else
                    If rsC!TipCon <> lsConAnt Or lsAgeAnt <> rsC!cCodAge Then
                        If pbOrdenado Then
                            lsNomAge = pgsEmpresa
                        Else
                            lsNomAge = rsC!cAgeDescripcion
                        End If
                        lnPagina = lnPagina + 1
                        lnItem = 0
                        If lnPagina <> 1 Then
                            lsCadena = lsCadena & GetItemPlanilla(rsC!cPersCod, psPlanillaPeriodo, psPlanillaCod, "", True, False, IIf(lsConAnt = "ESTABLE", 1, 2), , lsAgeAnt)
                            If lsAgeAnt <> rsC!cCodAge Then
                                lsCadena = lsCadena & GetItemPlanilla(rsC!cPersCod, psPlanillaPeriodo, psPlanillaCod, "", True, False, 0, , lsAgeAnt)
                            End If
                            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                        End If
                        lsCadena = lsCadena & GetCabeceraPlanilla(lnPagina, 1, rsC!TipCon & " - " & lsNomAge, ldFecAsisIni, ldFecAsisFin, psFechaImprePlanilla, pgsRuc, pgsEmpresa, psFecIni, psFecFin, psNombrePlanilla, psTextoPlanilla)
                    End If
                End If
            End If
            
            'Me.Caption = "Generación Planilla " & rsC!cEmpCod
            
            lsCadena = lsCadena & GetItemPlanilla(rsC!cPersCod, psPlanillaPeriodo, psPlanillaCod, "", , True, , IIf(IsNull(rsC!cComentario), "", rsC!cComentario))
            RaiseEvent Progress(rsC.Bookmark, rsC.RecordCount)
            
            lnItem = lnItem + 1
            
            If lnItem = 6 Then
                lnPagina = lnPagina + 1
                lnItem = 0
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & GetCabeceraPlanilla(lnPagina, 1, rsC!TipCon & " - " & lsNomAge, ldFecAsisIni, ldFecAsisFin, psFechaImprePlanilla, pgsRuc, pgsEmpresa, psFecIni, psFecFin, psNombrePlanilla, psTextoPlanilla)
            End If
            lsConAnt = rsC!TipCon
            lsAgeAnt = rsC!cCodAge
            
            rsC.MoveNext
        Wend
        RaiseEvent CloseProgress
    End If
    
    If lnItem = 4 Then
        lnItem = 1
        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
        lsCadena = lsCadena & GetCabeceraPlanilla(lnPagina, 1, lsConAnt & " - " & lsNomAge, ldFecAsisIni, ldFecAsisFin, psFechaImprePlanilla, pgsRuc, pgsEmpresa, psFecIni, psFecFin, psNombrePlanilla, psTextoPlanilla)
    End If
    If Not pbParaContabilidad Then
        lnItem = lnItem + 2
        lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True, False, IIf(lsConAnt = "ESTABLE", 1, 2), , lsAgeAnt)
        lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True, , 0, , lsAgeAnt) & oImpresora.gPrnSaltoLinea
    Else
        lnItem = lnItem + 1
        lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True, , IIf(lsConAnt = "ESTABLE", 1, 2)) & oImpresora.gPrnSaltoLinea
    End If
    If lnItem > 4 Then
        lnItem = 1
        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
        lsCadena = lsCadena & GetCabeceraPlanilla(lnPagina, 1, lsConAnt & " - " & lsNomAge, ldFecAsisIni, ldFecAsisFin, psFechaImprePlanilla, pgsRuc, pgsEmpresa, psFecIni, psFecFin, psNombrePlanilla, psTextoPlanilla)
    End If
    lsCadena = lsCadena & GetItemPlanilla("", psPlanillaPeriodo, psPlanillaCod, "", True) & oImpresora.gPrnSaltoLinea
    
    'Me.Caption = "Generación Planilla "
    

    GetPlanillas = lsCadena
End Function

'0 Total General
'1 Total Empleados
'2 Total Locacion
'Private Function GetItemPlanilla(psPersCod As String, psPlanillaPeriodo As String, psPlanillaCod As String, psProDiaLab As String, Optional pbTotal As Boolean = False, Optional pbConDiasLaborados As Boolean = False, Optional pnSubTotalEmpleados As Integer = 0, Optional psComentario As String = "", Optional psCodAge As String = "") As String
'    On Error Resume Next
'
'    Dim lsCadena As String
'    Dim lsCadAux(6) As String
'    Dim lsCod As String * 7
'    Dim lsNombre As String * 34
'    Dim lsCargo As String * 26
'    Dim lsAFP As String * 4
'    Dim lsNumC As String * 17
'    Dim lsSueldo As String * 15
'    Dim lsPorMesLabT As String * 10
'    Dim lsPorMesLab As String * 10
'    Dim lsCantidadTrab As String * 10
'    Dim lsNumCarnetAFP As String * 17
'    Dim lsCondiciaon As String * 17
'    Dim lsCondicion As String * 17
'    Dim lsAgencia As String * 10
'    Dim lsArea As String * 17
'    Dim lsVacaciones As String * 24
'    Dim lsNivel As String * 6
'    Dim lsComentario As String * 160
'    Dim lbNSF As Boolean
'
'    Dim rsCargo As ADODB.Recordset
'    Set rsCargo = New ADODB.Recordset
'
'    Dim I As Integer
'    Dim lnI As Integer
'    Dim lnD As Integer
'    Dim lnA As Integer
'    Dim lnPosMax As Integer
'    Dim lsCondicionE As String
'    Dim lsNumCE As String
'    Dim ldFecVacIni As Date
'    Dim ldFecVacFin As Date
'
'    Dim lsNegritaOn As String
'    Dim lsNegritaOff As String
'    Dim lnEspacio As Integer
'    Dim lbBan As Boolean
'
'    Dim lnNumConceptos As Integer
'    lnNumConceptos = 16
'    Dim lsConcep(16, 6) As String * 12
'    Dim lsConcepT(1, 6) As String * 12
'
'    lsPorMesLab = ""
'
'    Dim rsP As ADODB.Recordset
'    Set rsP = New ADODB.Recordset
'    Dim oRH As DActualizaDatosRRHH
'    Set oRH = New DActualizaDatosRRHH
'
'    Dim oPla As DActualizaDatosConPlanilla
'    Set oPla = New DActualizaDatosConPlanilla
'
'    Dim lnIn As Integer
'    Dim lnJn As Integer
'
'    For lnIn = 0 To lnNumConceptos
'        For lnJn = 0 To 6
'            lsConcep(lnIn, lnJn) = ""
'        Next lnJn
'    Next lnIn
'
'    For lnJn = 0 To 6
'        lsConcepT(1, lnJn) = ""
'    Next lnJn
'
'
'    lnEspacio = 5
'    lsNegritaOn = oImpresora.gPrnBoldON
'    lsNegritaOff = oImpresora.gPrnBoldOFF
'
'    Set rsP = oPla.GetItemPlanilla(psPersCod, psPlanillaPeriodo, psPlanillaCod, psCodAge, pbTotal, pnSubTotalEmpleados)
'
'    If rsP.EOF And rsP.BOF Then
'        If rsP.State Then rsP.Close
'        Set rsP = oPla.GetItemPlanillaRep(psPersCod)
'        lbBan = False
'    Else
'        lbBan = True
'    End If
'
'    If Not pbTotal Then
'        lsCod = "CODIGO"
'        lsNombre = "NOMBRES Y APELLIDOS"
'        lsCargo = IIf(Left(psPlanillaCod, 1) = "E", "CARGO" & IIf(rsP!bConfianza, "-CONFIANZA", "") & IIf(rsP!bNoSujFiscalizacion, "-NSF", "") & IIf(rsP!bPersDireccion, "-PERS DIRECC", ""), "CONVENIO")
'        lbNSF = rsP!bNoSujFiscalizacion
'        lsNumC = IIf(Left(psPlanillaCod, 1) = "E", "CARNET ESSALUD", "")
'        lsNivel = IIf(Left(psPlanillaCod, 1) = "E", "NIVEL", "")
'        lsArea = "AREA"
'        lsNumCarnetAFP = IIf(Left(psPlanillaCod, 1) = "E", "CUSPP", "")
'        lsAFP = IIf(Left(psPlanillaCod, 1) = "E", "AFP", "")
'        lsVacaciones = IIf(Left(psPlanillaCod, 1) = "E", "VACACIONES", "")
'
'        lsCadena = lsCadena & lsNegritaOn & Space(lnEspacio) & lsCod & lsNombre & " " & lsCargo & lsNumC & lsNivel & lsArea & lsNumCarnetAFP & lsAFP & lsVacaciones & oImpresora.gPrnSaltoLinea & lsNegritaOff
'
'        lsCod = rsP!cRHCod
'        lsNombre = PstaNombre(rsP!cPersNombre, False)
'
'        Set rsCargo = oRH.GetCargoRS(psPersCod, psPlanillaPeriodo)
'        lsCargo = rsCargo!Cargo
'
'        lsNumC = rsP!cCUSPP & ""
'        lsNivel = rsP!cRHEmpNivRem & ""
'        lsArea = Trim(IIf(IsNull(rsP!CAREADESCRIPCION), "", rsP!CAREADESCRIPCION))
'        lsNumCarnetAFP = rsP!cCUSPP & ""
'        lsAFP = rsP!cRHAFPAbreviatura & ""
'        If Left(psPlanillaCod, 1) = "E" Then
'            If oRH.GetFecVac(psPersCod, ldFecVacIni, ldFecVacFin) Then
'                lsVacaciones = Format(ldFecVacIni, gsFormatoFechaView) & " - " & Format(ldFecVacFin, gsFormatoFechaView)
'            Else
'                lsVacaciones = "__/__/____ - __/__/____"
'            End If
'        Else
'            lsVacaciones = ""
'        End If
'
'        lsCadena = lsCadena & Space(lnEspacio) & lsCod & lsNombre & " " & lsCargo & lsNumC & lsNivel & lsArea & lsNumCarnetAFP & lsAFP & lsVacaciones & oImpresora.gPrnSaltoLinea
'
'        If IsNumeric(psProDiaLab) Then lsPorMesLab = Str(CCur(psProDiaLab)) & "d"
'    Else
'        lsCadena = lsCadena & psComentario & oImpresora.gPrnSaltoLinea
'
'        If rsP.EOF And rsP.BOF Then
'            lsCantidadTrab = 0
'        Else
'            If rsP!Num = -1 Then
'                lsCantidadTrab = ""
'            Else
'                 lsCantidadTrab = rsP!Num
'            End If
'        End If
'    End If
'
'    lnI = lnI + 1
'    lnD = lnD + 1
'    lsConcep(lnD, 3) = "DESCUENT :"
'    lsConcep(lnD, 4) = ""
'
'    lnA = lnA + 1
'    lsConcep(lnA, 5) = "APORTES  :"
'    lsConcep(lnA, 6) = ""
'
'    While Not rsP.EOF And lbBan
'        If rsP!nMonto <> 0 Or Trim(rsP!cConcepCod) = gsRHConceptoITOTING Or Trim(rsP!cConcepCod) = "D_TOT_DESC" Then
'            If Left(rsP!cConcepCod, 2) = "I_" Or Left(rsP!cConcepCod, 2) = "J_" Then
'                lnI = lnI + 1
'                lsConcep(lnI, 1) = Trim(rsP!cRHConceptoAbreviatura) & ""
'                RSet lsConcep(lnI, 2) = Format(rsP!nMonto, "#,##0.00")
'            ElseIf Left(rsP!cConcepCod, 2) = "D_" Or Left(rsP!cConcepCod, 2) = "E_" Then
'                lnD = lnD + 1
'                lsConcep(lnD, 3) = Trim(rsP!cRHConceptoAbreviatura)
'                RSet lsConcep(lnD, 4) = Format(rsP!nMonto, "#,##0.00")
'            ElseIf Left(rsP!cConcepCod, 2) = "A_" Or Left(rsP!cConcepCod, 2) = "B_" Then
'                lnA = lnA + 1
'                lsConcep(lnA, 5) = Trim(rsP!cRHConceptoAbreviatura)
'                RSet lsConcep(lnA, 6) = Format(rsP!nMonto, "#,##0.00")
'            Else
'                lsPorMesLab = Trim(Str(rsP!nMonto))
'            End If
'        End If
'        rsP.MoveNext
'    Wend
'
'    lsConcep(1, 1) = "INGRESOS :"
'    lsConcep(1, 2) = ""
'    If psPlanillaCod <> gsRHPlanillaUtilidades Then
'        If pbConDiasLaborados Then
'            If IsNumeric(lsPorMesLab) Then
'                If Not pbTotal Then
'                    If psPlanillaCod = gsRHPlanillaSueldos Or psPlanillaCod = "P01" Or psPlanillaCod = "F01" Then
'                        If rsCargo!bRHCtrAsist Then
'                            lsConcep(1, 2) = Str(CCur(lsPorMesLab) * (30 - GetDiasFediados(CDate(Right(psPlanillaPeriodo, 2) & "/" & Mid(psPlanillaPeriodo, 5, 2) & "/" & Left(psPlanillaPeriodo, 4)), CCur(lsPorMesLab)))) & " / " & Trim(Str(GetHorasTrabajadas(CDate(Right(psPlanillaPeriodo, 2) & "/" & Mid(psPlanillaPeriodo, 5, 2) & "/" & Left(psPlanillaPeriodo, 4)), CCur(lsPorMesLab) * 30)))
'                        Else
'                            lsConcep(1, 2) = Str(CCur(lsPorMesLab) * (30 - GetDiasFediados(CDate(Right(psPlanillaPeriodo, 2) & "/" & Mid(psPlanillaPeriodo, 5, 2) & "/" & Left(psPlanillaPeriodo, 4)), CCur(lsPorMesLab))))
'                        End If
'                    ElseIf psPlanillaCod = gsRHPlanillaLiquidacion Then
'                        lsConcep(1, 2) = ""
'                    Else
'                        lsConcep(1, 2) = Str(CCur(lsPorMesLab)) & " d."
'                    End If
'
'                End If
'            Else
'                If Not pbTotal Then
'                    lsConcep(1, 2) = lsPorMesLab
'                End If
'            End If
'        Else
'            If Not pbTotal Then
'                lsConcep(1, 2) = ""
'            Else
'                lsConcep(1, 2) = lsCantidadTrab
'            End If
'        End If
'    Else
'        lsConcep(1, 2) = ""
'    End If
'
'    lsConcepT(1, 1) = lsConcep(lnI, 1)
'    lsConcepT(1, 2) = lsConcep(lnI, 2)
'    lsConcepT(1, 3) = lsConcep(lnD, 3)
'    lsConcepT(1, 4) = lsConcep(lnD, 4)
'    lsConcepT(1, 5) = "NETO A PAGAR"
'    If Not IsNumeric(lsConcepT(1, 4)) Then RSet lsConcepT(1, 4) = Format(0, "#,##0.00")
'    If Not IsNumeric(lsConcepT(1, 2)) Then RSet lsConcepT(1, 2) = Format(0, "#,##0.00")
'    RSet lsConcepT(1, 6) = Format(CCur(lsConcepT(1, 2) - CCur(lsConcepT(1, 4))), "#,##0.00")
'    lsConcep(lnI, 1) = ""
'    lsConcep(lnI, 2) = ""
'    lsConcep(lnD, 3) = ""
'    lsConcep(lnD, 4) = ""
'    lsConcep(lnA, 5) = ""
'    lsConcep(lnA, 6) = ""
'    lnI = lnI - 1
'    lnD = lnD - 1
'    lnA = lnA - 1
'
'    rsP.Close
'
'    lnPosMax = 1
'    For lnPosMax = 1 To lnNumConceptos
'        lsCadAux(1) = lsCadAux(1) & lsConcep(lnPosMax, 1) & Space(1)
'        lsCadAux(2) = lsCadAux(2) & lsConcep(lnPosMax, 2) & Space(1)
'        lsCadAux(3) = lsCadAux(3) & lsConcep(lnPosMax, 3) & Space(1)
'        lsCadAux(4) = lsCadAux(4) & lsConcep(lnPosMax, 4) & Space(1)
'        lsCadAux(5) = lsCadAux(5) & lsConcep(lnPosMax, 5) & Space(1)
'        lsCadAux(6) = lsCadAux(6) & lsConcep(lnPosMax, 6) & Space(1)
'    Next lnPosMax
'
'    lsCadAux(1) = lsCadAux(1) & lsConcepT(1, 1)
'    lsCadAux(2) = lsCadAux(2) & lsConcepT(1, 2)
'    lsCadAux(3) = lsCadAux(3) & lsConcepT(1, 3)
'    lsCadAux(4) = lsCadAux(4) & lsConcepT(1, 4)
'    lsCadAux(5) = lsCadAux(5) & lsNegritaOn & lsConcepT(1, 5) & lsNegritaOff
'    lsCadAux(6) = lsCadAux(6) & lsNegritaOn & lsConcepT(1, 6) & lsNegritaOff
'
'    For lnPosMax = 1 To 6
'        lsCadena = lsCadena & lsCadAux(lnPosMax) & oImpresora.gPrnSaltoLinea
'    Next lnPosMax
'
'    If (psComentario <> "" And Not pbTotal) Or lbNSF Then
'        If psComentario <> "" Then
'            lsComentario = CentrarCadena("**** " & Trim(Left(psComentario, 160)) & " ****" & IIf(lbNSF, "(*) NSF : NO SUJETO A FISCALIZACION", ""), 180)
'        Else
'            lsComentario = CentrarCadena(IIf(lbNSF, "(*) NSF : NO SUJETO A FISCALIZACION", ""), 180)
'        End If
'        lsCadena = lsCadena & lsComentario & oImpresora.gPrnSaltoLinea
'    End If
'
'    lsCadena = lsCadena & String(226, "-") & oImpresora.gPrnSaltoLinea
'    GetItemPlanilla = lsCadena
'
'    Exit Function
'    Resume
'        MsgBox ""
'End Function

Private Function GetItemPlanilla(psPersCod As String, psPlanillaPeriodo As String, psPlanillaCod As String, psProDiaLab As String, Optional pbTotal As Boolean = False, Optional pbConDiasLaborados As Boolean = False, Optional pnSubTotalEmpleados As Integer = 0, Optional psComentario As String = "", Optional psCodAge As String = "") As String
    On Error Resume Next
    
    Dim lsCadena As String
    Dim lsCadAux(6) As String
    Dim lsCod As String * 7
    Dim lsNombre As String * 34
    Dim lsCargo As String * 26
    Dim lsAFP As String * 4
    Dim lsNumC As String * 17
    Dim lsSueldo As String * 15
    Dim lsPorMesLabT As String * 10
    Dim lsPorMesLab As String * 10
    Dim lsCantidadTrab As String * 10
    Dim lsNumCarnetAFP As String * 17
    Dim lsCondiciaon As String * 17
    Dim lsCondicion As String * 17
    Dim lsAgencia As String * 10
    Dim lsArea As String * 17
    Dim lsVacaciones As String * 24
    Dim lsNivel As String * 6
    Dim lsComentario As String * 160

    Dim I As Integer
    Dim lnI As Integer
    Dim lnD As Integer
    Dim lnA As Integer
    Dim lnPosMax As Integer
    Dim lsCondicionE As String
    Dim lsNumCE As String
    Dim ldFecVacIni As String
    Dim ldFecVacFin As String
    
    Dim lsNegritaOn As String
    Dim lsNegritaOff As String
    Dim lnEspacio As Integer
    Dim lbBan As Boolean
    
    Dim lnNumConceptos As Integer
    lnNumConceptos = 13
    Dim lsConcep(13, 6) As String * 10
    Dim lsConcepT(1, 6) As String * 10
    lsPorMesLab = ""
    
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    Dim lnIn As Integer
    Dim lnJn As Integer
    
    For lnIn = 0 To 13
        For lnJn = 0 To 6
            lsConcep(lnIn, lnJn) = ""
        Next lnJn
    Next lnIn
    
    For lnJn = 0 To 6
        lsConcepT(1, lnJn) = ""
    Next lnJn
    
    
    lnEspacio = 5
    lsNegritaOn = oImpresora.gPrnBoldON
    lsNegritaOff = oImpresora.gPrnBoldOFF

    Set rsP = oPla.GetItemPlanilla(psPersCod, psPlanillaPeriodo, psPlanillaCod, psCodAge, pbTotal, pnSubTotalEmpleados)
    
    If rsP.EOF And rsP.BOF Then
        If rsP.State Then rsP.Close
        Set rsP = oPla.GetItemPlanillaRep(psPersCod)
        lbBan = False
    Else
        lbBan = True
    End If
    
    If Not pbTotal Then
        lsCod = "CODIGO"
        lsNombre = "NOMBRES Y APELLIDOS"
        lsCargo = IIf(Left(psPlanillaCod, 1) = "E", "CARGO", "CONVENIO")
        lsNumC = IIf(Left(psPlanillaCod, 1) = "E", "CARNET ESSALUD", "")
        lsNivel = IIf(Left(psPlanillaCod, 1) = "E", "NIVEL", "")
        lsArea = "AREA"
        lsNumCarnetAFP = IIf(Left(psPlanillaCod, 1) = "E", "CUSPP", "")
        lsAFP = IIf(Left(psPlanillaCod, 1) = "E", "AFP", "")
        lsVacaciones = IIf(Left(psPlanillaCod, 1) = "E", "VACACIONES", "")
        
        lsCadena = lsCadena & lsNegritaOn & Space(lnEspacio) & lsCod & lsNombre & " " & lsCargo & lsNumC & lsNivel & lsArea & lsNumCarnetAFP & lsAFP & lsVacaciones & oImpresora.gPrnSaltoLinea & lsNegritaOff
        
        lsCod = rsP!cRHCod
        lsNombre = PstaNombre(rsP!cPersNombre, False)
        lsCargo = oRH.GetCargo(psPersCod, psPlanillaPeriodo)
        lsNumC = rsP!cCUSPP & ""
        lsNivel = rsP!cRHEmpNivRem & ""
        lsArea = Trim(IIf(IsNull(rsP!CAREADESCRIPCION), "", rsP!CAREADESCRIPCION))
        lsNumCarnetAFP = rsP!cCUSPP & ""
        lsAFP = rsP!cRHAFPAbreviatura & ""
        If Left(psPlanillaCod, 1) = "E" Then
            If oRH.GetFecVac(psPersCod, ldFecVacIni, ldFecVacFin) Then
                lsVacaciones = Format(ldFecVacIni, gsFormatoFechaView) & " - " & Format(ldFecVacFin, gsFormatoFechaView)
            Else
                lsVacaciones = "__/__/____ - __/__/____"
            End If
        Else
            lsVacaciones = ""
        End If
        
        lsCadena = lsCadena & Space(lnEspacio) & lsCod & lsNombre & " " & lsCargo & lsNumC & lsNivel & lsArea & lsNumCarnetAFP & lsAFP & lsVacaciones & oImpresora.gPrnSaltoLinea
        
        If IsNumeric(psProDiaLab) Then lsPorMesLab = Str(CCur(psProDiaLab)) & "d"
    Else
        lsCadena = lsCadena & psComentario & oImpresora.gPrnSaltoLinea
        
        If rsP.EOF And rsP.BOF Then
            lsCantidadTrab = 0
        Else
            If rsP!Num = -1 Then
                lsCantidadTrab = ""
            Else
                 lsCantidadTrab = rsP!Num
            End If
        End If
    End If

    lnI = lnI + 1
    lnD = lnD + 1
    lsConcep(lnD, 3) = "DESCUENT :"
    lsConcep(lnD, 4) = ""

    lnA = lnA + 1
    lsConcep(lnA, 5) = "APORTES  :"
    lsConcep(lnA, 6) = ""
    
    While Not rsP.EOF And lbBan
        If rsP!nMonto <> 0 Or Trim(rsP!cConcepCod) = gsRHConceptoITOTING Or Trim(rsP!cConcepCod) = "D_TOT_DESC" Then
            If Left(rsP!cConcepCod, 2) = "I_" Or Left(rsP!cConcepCod, 2) = "J_" Then
                lnI = lnI + 1
                lsConcep(lnI, 1) = Trim(rsP!cRHConceptoAbreviatura) & ""
                RSet lsConcep(lnI, 2) = Format(rsP!nMonto, "#,##0.00")
            ElseIf Left(rsP!cConcepCod, 2) = "D_" Or Left(rsP!cConcepCod, 2) = "E_" Then
                lnD = lnD + 1
                lsConcep(lnD, 3) = Trim(rsP!cRHConceptoAbreviatura)
                RSet lsConcep(lnD, 4) = Format(rsP!nMonto, "#,##0.00")
            ElseIf Left(rsP!cConcepCod, 2) = "A_" Or Left(rsP!cConcepCod, 2) = "B_" Then
                lnA = lnA + 1
                lsConcep(lnA, 5) = Trim(rsP!cRHConceptoAbreviatura)
                RSet lsConcep(lnA, 6) = Format(rsP!nMonto, "#,##0.00")
            Else
                lsPorMesLab = Trim(Str(rsP!nMonto))
            End If
        End If
        rsP.MoveNext
    Wend
    
    lsConcep(1, 1) = "INGRESOS :"
    lsConcep(1, 2) = ""
    If psPlanillaCod <> gsRHPlanillaUtilidades Then
        If pbConDiasLaborados Then
            If IsNumeric(lsPorMesLab) Then
                If Not pbTotal Then
                    If psPlanillaCod = gsRHPlanillaSueldos Or psPlanillaCod = "P01" Or psPlanillaCod = "F01" Then
                        lsConcep(1, 2) = Str(CCur(lsPorMesLab) * 30) & " d."
                    Else
                        lsConcep(1, 2) = Str(CCur(lsPorMesLab)) & " d."
                    End If
                    
                End If
            Else
                If Not pbTotal Then
                    lsConcep(1, 2) = lsPorMesLab
                End If
            End If
        Else
            If Not pbTotal Then
                lsConcep(1, 2) = ""
            Else
                lsConcep(1, 2) = lsCantidadTrab
            End If
        End If
    Else
        lsConcep(1, 2) = ""
    End If
    
    lsConcepT(1, 1) = lsConcep(lnI, 1)
    lsConcepT(1, 2) = lsConcep(lnI, 2)
    lsConcepT(1, 3) = lsConcep(lnD, 3)
    lsConcepT(1, 4) = lsConcep(lnD, 4)
    lsConcepT(1, 5) = "NETO A PAGAR"
    If Not IsNumeric(lsConcepT(1, 4)) Then RSet lsConcepT(1, 4) = Format(0, "#,##0.00")
    If Not IsNumeric(lsConcepT(1, 2)) Then RSet lsConcepT(1, 2) = Format(0, "#,##0.00")
    RSet lsConcepT(1, 6) = Format(CCur(lsConcepT(1, 2) - CCur(lsConcepT(1, 4))), "#,##0.00")
    lsConcep(lnI, 1) = ""
    lsConcep(lnI, 2) = ""
    lsConcep(lnD, 3) = ""
    lsConcep(lnD, 4) = ""
    lsConcep(lnA, 5) = ""
    lsConcep(lnA, 6) = ""
    lnI = lnI - 1
    lnD = lnD - 1
    lnA = lnA - 1
    
    rsP.Close
    
    lnPosMax = 1
    For lnPosMax = 1 To lnNumConceptos
        lsCadAux(1) = lsCadAux(1) & lsConcep(lnPosMax, 1) & Space(1)
        lsCadAux(2) = lsCadAux(2) & lsConcep(lnPosMax, 2) & Space(1)
        lsCadAux(3) = lsCadAux(3) & lsConcep(lnPosMax, 3) & Space(1)
        lsCadAux(4) = lsCadAux(4) & lsConcep(lnPosMax, 4) & Space(1)
        lsCadAux(5) = lsCadAux(5) & lsConcep(lnPosMax, 5) & Space(1)
        lsCadAux(6) = lsCadAux(6) & lsConcep(lnPosMax, 6) & Space(1)
    Next lnPosMax
    
    lsCadAux(1) = lsCadAux(1) & lsConcepT(1, 1)
    lsCadAux(2) = lsCadAux(2) & lsConcepT(1, 2)
    lsCadAux(3) = lsCadAux(3) & lsConcepT(1, 3)
    lsCadAux(4) = lsCadAux(4) & lsConcepT(1, 4)
    lsCadAux(5) = lsCadAux(5) & lsNegritaOn & lsConcepT(1, 5) & lsNegritaOff
    lsCadAux(6) = lsCadAux(6) & lsNegritaOn & lsConcepT(1, 6) & lsNegritaOff
    
    For lnPosMax = 1 To 6
        lsCadena = lsCadena & lsCadAux(lnPosMax) & oImpresora.gPrnSaltoLinea
    Next lnPosMax
    
    If psComentario <> "" And Not pbTotal Then
        lsComentario = CentrarCadena("**** " & Trim(Left(psComentario, 160)) & " ****", 180)
        lsCadena = lsCadena & lsComentario & oImpresora.gPrnSaltoLinea
    End If
    
    lsCadena = lsCadena & String(160, "-") & oImpresora.gPrnSaltoLinea
    GetItemPlanilla = lsCadena
    
    Exit Function
    Resume
        MsgBox ""
End Function


Private Function GetCabeceraPlanilla(pnPagina As Integer, pnItem As Integer, psTipCon As String, pdFecAsisIni As Date, pdFecAsisFin As Date, psFecImpre As String, pgsRuc As String, pgsEmpresa As String, psFecIni As String, psFecFin As String, psNombrePlanilla As String, psTextoPlanilla As String) As String
    Dim lsCadena As String
    
    lsCadena = ""
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & pgsEmpresa & CentrarCadena(Trim(Left(psNombrePlanilla, 60)) & " " & Format(CDate(psFecImpre), "MMMM - YYYY"), 175) & " FECHA  : " & psFecImpre & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "Periodo: " & Format(CDate(psFecIni), "MMMM") & " - Del " & Format(CDate(psFecIni), gsFormatoFechaView) & " al " & Format(CDate(psFecFin), gsFormatoFechaView) & CentrarCadena("Personal " & psTipCon, 159) & " PAGINA : " & FillNum(Trim(Str(pnPagina)), 3, "0") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & CentrarCadena(Left(Trim(Replace(psTextoPlanilla, oImpresora.gPrnSaltoLinea, " ")), 180), 200) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "RUC: " & pgsRuc & Space(160) & "Control de Asistencia del " & Format(pdFecAsisIni, gsFormatoFechaView) & " al " & Format(pdFecAsisFin, gsFormatoFechaView) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & String(226, "=") & oImpresora.gPrnSaltoLinea
    
    GetCabeceraPlanilla = lsCadena
End Function

Public Function GetCartasCTS(prRS As ADODB.Recordset, pgsRuc As String, pgsEmpresa As String, psFechaIni As String, psFechaFin As String, pgdFecha As Date, pnTipCambio As Currency) As String
    Dim lsCadena As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    lsCadena = ""
    
    If Not (prRS.EOF And prRS.BOF) Then
        RaiseEvent ShowProgress
        While Not prRS.EOF
            If prRS.Fields(0) = "1" Then
                Set rs = oRH.GetRRHHAgenciaActual(prRS.Fields(2))
                lsCadena = lsCadena & GetCartaCTS(prRS, pgsRuc, pgsEmpresa, psFechaIni, psFechaFin, pgdFecha, pnTipCambio, rs)
            End If
            RaiseEvent Progress(prRS.Bookmark, prRS.RecordCount)
            prRS.MoveNext
        Wend
        RaiseEvent CloseProgress
    End If
    GetCartasCTS = lsCadena
End Function

'*********************************
'Private Function GetCartaCTS(prRS As ADODB.Recordset, pgsRuc As String, pgsEmpresa As String, psFechaIni As String, psFechaFin As String, pgdFecha As Date, pnTipCambio As Currency, prRSAGE As ADODB.Recordset) As String
'    Dim lsCadena As String
'    Dim lsMargen As String
'    Dim lsNegritaOn As String
'    Dim lsNegritaOff As String
'    Dim lnEspacio As Integer
'    Dim lsMontoToTal As String * 15
'    Dim lsMontoToTalPar As String * 15
'    Dim lsItem1 As String * 15
'    Dim lsItem2 As String * 15
'    Dim lsCodCta As String
'    Dim oRH As DActualizaDatosRRHH
'    Set oRH = New DActualizaDatosRRHH
'    Dim oInterprete As DInterprete
'    Set oInterprete = New DInterprete
'    Dim ArcSal As Integer
'    Dim lsTexto As String
'
'    ArcSal = FreeFile
'    Open App.Path & "\Txt\TextoCartaCTS.txt" For Input As ArcSal
'
'    Input #ArcSal, lsTexto
'
'
'    lnEspacio = 6
'    lsNegritaOn = oImpresora.gPrnBoldON
'    lsNegritaOff = oImpresora.gPrnBoldOFF
'
'    RSet lsMontoToTal = Format(lnTotDias, "#,##0.00")
'    RSet lsMontoToTalPar = Format(lnNumMonto, "#,##0.00")
'    RSet lsItem1 = prRS!U_CTS_GRATI
'    RSet lsItem2 = prRS!U_CTS_TOT_REM
'
'    lsCodCta = oRH.GetCuentaRRHH(prRS!CodPers, gsRHPlanillaCTS)
'
'    lsCadena = ""
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsNegritaOn & Space(lnEspacio) & Space(lnEspacio) & "    " & CentrarCadena(String(Len("PERIODO " & Format(psFechaIni, "mmmm yyyy") & " - " & Format(psFechaFin, "mmmm yyyy")), "="), 80 - 13) & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
'    lsCadena = lsNegritaOn & Space(lnEspacio) & Space(lnEspacio) & "    " & CentrarCadena("PERIODO " & Format(psFechaIni, "mmmm yyyy") & " - " & Format(psFechaFin, "mmmm yyyy"), 80 - 13) & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
'    'lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsNegritaOn & Space(lnEspacio) & "RUC:" & pgsRuc & CentrarCadena("COMPENSACION  POR  TIEMPO  DE  SERVICIOS", 80 - 24) & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
'    lsCadena = lsNegritaOn & Space(lnEspacio) & "AV. ESPAÑA Nro. 2611-2615 TRUJILLO " & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsNegritaOn & Space(lnEspacio) & pgsEmpresa & CentrarCadena("", 50) & Format(pgdFecha, gsFormatoFechaView) & lsNegritaOff & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsCadena
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "CODIGO    : " & prRS![Cod.Emp.] & " - " & prRSAGE.Fields(0) & " - " & prRSAGE.Fields(1) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "NOMBRE    : " & prRS!Nombre & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "CUENTA    : " & lsCodCta & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    RSet lsItem1 = Format(prRS!U_CTS_MES_TRAB, "#0.00")
'    lsCadena = lsCadena & Space(lnEspacio) & "MESES COMPUTABLES     : " & lsItem1 & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "FECHA INI PERIODO     :      " & psFechaIni & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "FECHA FIN PERIODO     :      " & psFechaFin & oImpresora.gPrnSaltoLinea
'    RSet lsItem1 = Format(prRS!U_CTS_TOT_REM, "#0.00")
'    lsCadena = lsCadena & Space(lnEspacio) & "REMUNERACION          : " & lsItem1 & oImpresora.gPrnSaltoLinea
'
'    RSet lsItem1 = oInterprete.FunEvalua("I_IN_3_AFP", prRS!CodPers, CDate(psFechaIni), CDate(psFechaIni), False, "VVVV", "")
'    lsCadena = lsCadena & Space(lnEspacio) & "INCREMENTO AFP 3.3%   : " & lsItem1 & oImpresora.gPrnSaltoLinea
'    RSet lsItem1 = Format(prRS!U_CTS_GRATI, "#0.00")
'    lsCadena = lsCadena & Space(lnEspacio) & "PROM. GRATIFICA. " & IIf(Month(CDate(psFechaIni) = 11), "NAV", "FFP") & "  : " & lsItem1 & oImpresora.gPrnSaltoLinea
'    RSet lsItem1 = Format(CCur(prRS!U_CTS_GRATI) + CCur(prRS!U_CTS_TOT_REM), "#0.00")
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "REM. INDEMNIZABLE : " & lsItem1 & "    " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & Space(lnEspacio) & "CALCULO :" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "TIEMPO DE SERVICIOS" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "-------------------" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & lsItem1 & " / 12  x " & Format(prRS!U_CTS_MES_TRAB, "#0.00") & " MESES  : " & Format(prRS!I_TOT_ING, "#0.00") & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "Tipo de Cambio : " & Format(pnTipCambio, "#,##0.00") & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & Space(lnEspacio) & "" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsNegritaOn
'    If (Mid(lsCodCta, 6, 1) = Moneda.gMonedaNacional And Len(lsCodCta) = 12) Or (Mid(lsCodCta, 9, 1) = Moneda.gMonedaNacional And Len(lsCodCta) = 18) Then
'        lsCadena = lsCadena & Space(lnEspacio) & "TOTAL DEPOSITADO EN MONEDA NACIONAL" & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & Space(lnEspacio) & "S/. " & Format(prRS!I_TOT_ING, "#0.00") & oImpresora.gPrnSaltoLinea
'    ElseIf (Mid(lsCodCta, 6, 1) = Moneda.gMonedaExtranjera And Len(lsCodCta) = 12) Or (Mid(lsCodCta, 9, 1) = Moneda.gMonedaExtranjera And Len(lsCodCta) = 18) Then
'        lsCadena = lsCadena & Space(lnEspacio) & "TOTAL DEPOSITADO EN MONEDA EXTRANJERA" & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & Space(lnEspacio) & " $  " & Format(CCur(prRS!I_TOT_ING) / pnTipCambio, "#0.00") & oImpresora.gPrnSaltoLinea
'    End If
'    lsCadena = lsCadena & lsNegritaOff
'
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    If lsCodCta <> "" Then
'        lsCadena = lsCadena & Space(lnEspacio) & "ENTIDAD DEPOSITARIA : " & lsNegritaOn & pgsEmpresa & lsNegritaOff
'    Else
'        lsCadena = lsCadena & Space(lnEspacio) & "ENTIDAD DEPOSITARIA : " & "<<NO ESPECIFICADA>>"
'    End If
'
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & JustificaTexto("HE RECIBIDO DE LA " & pgsEmpresa & " " & lsTexto, 80, lnEspacio) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena("---------------------------------", 50) & "   " & CentrarCadena("---------------------------------", 50) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena(Left(prRS!Nombre, 50), 50) & "   " & CentrarCadena("JEFE DE RRHH", 50) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena("(Empleado)", 50) & oImpresora.gPrnSaltoPagina
'
'    Close ArcSal
'    GetCartaCTS = lsCadena & lsCadena
'End Function

'*********************************


Private Function GetCartaCTS(prRS As ADODB.Recordset, pgsRuc As String, pgsEmpresa As String, psFechaIni As String, psFechaFin As String, pgdFecha As Date, pnTipCambio As Currency, prRSAGE As ADODB.Recordset) As String
    Dim lsCadena As String
    Dim lsMargen As String
    Dim lsNegritaOn As String
    Dim lsNegritaOff As String
    Dim lnEspacio As Integer
    Dim lsMontoToTal As String * 15
    Dim lsMontoToTalPar As String * 15
    Dim lsItem1 As String * 15
    Dim lsItem2 As String * 15
    Dim lsItem3 As String * 15
    Dim lsCodCta As String
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    Dim ArcSal As Integer
    Dim lsTexto As String
    Dim nMN As Currency
    Dim nME As Currency
    Dim gMN As Currency
    Dim gME As Currency
    Dim tMN As Currency
    Dim tME As Currency
    Dim nMes As Integer
    Dim sTextMes As String
    
    ArcSal = FreeFile
    Open App.Path & "\Txt\TextoCartaCTS.txt" For Input As ArcSal
    
    Input #ArcSal, lsTexto


    lnEspacio = 6
    lsNegritaOn = oImpresora.gPrnBoldON
    lsNegritaOff = oImpresora.gPrnBoldOFF
    
    RSet lsMontoToTal = Format(lnTotDias, "#,##0.00")
    RSet lsMontoToTalPar = Format(lnNumMonto, "#,##0.00")
    RSet lsItem1 = prRS.Fields(6) 'gratificacion  I_CTS_GRATIF
    RSet lsItem2 = prRS.Fields(7) ' remuneracion  I_CTS_REMUN
    RSet lsItem3 = prRS.Fields(10) 'total
    
    lsCodCta = oRH.GetCuentaRRHH(prRS.Fields(2), gsRHPlanillaCTS)
    
    lsCadena = ""
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    'lsCadena = lsNegritaOn & Space(lnEspacio) & Space(lnEspacio) & "    " & CentrarCadena("PERIODO " & Format(psFechaIni, "mmmm yyyy") & " - " & Format(psFechaFin, "mmmm yyyy"), 80 - 13) & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsNegritaOn & Space(lnEspacio) & CentrarCadena("COMPENSACION  POR  TIEMPO  DE  SERVICIOS", 80 - 20) & lsNegritaOff & oImpresora.gPrnSaltoLinea '& lsCadena
    lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena("T.U.O. del D. LEG. 650. D.S. 01-97-TR", 80 - 20) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsNegritaOn & Space(lnEspacio) & CentrarCadena("LIQUIDACION", 80 - 20) & lsNegritaOff & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea '& lsCadena
     
    lsCadena = lsCadena & Space(3) & "CORRESPONDIENTE A      : " & Space(3) & prRS.Fields(3) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(3) & "PERIODO DE LIQUIDACION : " & Space(3) & "DEL" & Space(3) & psFechaIni & Space(4) & "AL" & Space(3) & psFechaFin & oImpresora.gPrnSaltoLinea
    
    If Year(psFechaIni) = Year(psFechaFin) Then
       nMes = Month(psFechaFin) - Month(psFechaIni) + 1
    Else
       nMes = 12 - Month(psFechaIni) + 1 + Month(psFechaFin)
    End If
    
    If nMes > 1 Then
          sTextMes = "Meses"
       Else
          sTextMes = "Mes"
    End If
  
    lsCadena = lsCadena & Space(3) & "TIEMPO DE SERVICIOS    : " & Space(3) & nMes & " " & sTextMes & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & String(95, "_") & oImpresora.gPrnSaltoLinea
        
    lsCadena = lsCadena & lsNegritaOn & "REMUNERACION COMPUTABLE" & lsNegritaOff & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(70) & lsNegritaOn & "NUEVOS SOLES" & lsNegritaOff & oImpresora.gPrnSaltoLinea
    RSet lsItem1 = Format(prRS.Fields(7), "#0.00")
    lsCadena = lsCadena & Space(8) & "REMUNERACION MENSUAL" & Space(46) & lsItem1 & oImpresora.gPrnSaltoLinea
    RSet lsItem1 = oInterprete.FunEvalua("I_IN_3_AFP", prRS.Fields(2), CDate(psFechaIni), CDate(psFechaIni), False, "VVVV", "")
    lsCadena = lsCadena & Space(8) & "(-)INCREMENTO AFP 10.23%   : " & Space(32) & lsItem1 & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & String(90, "_") & oImpresora.gPrnSaltoLinea
    
    '*************************************************************************************
    RSet lsItem1 = Format(prRS.Fields(7), "#0.00")
    lsCadena = lsCadena & Space(8) & "TOTAL REMUNERACION COMPUTABLE" & Space(10) & psFechaIni & Space(1) & gcPEN_SIMBOLO & Space(13) & lsItem1 & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    
    If (Mid(lsCodCta, 6, 1) = Moneda.gMonedaNacional And Len(lsCodCta) = 12) Or (Mid(lsCodCta, 9, 1) = Moneda.gMonedaNacional And Len(lsCodCta) = 18) Then
        
           nMN = Format(prRS.Fields(7), "#0.00")
           gMN = Format(prRS.Fields(6), "#0.00")
           tMN = Format(prRS.Fields(9), "#0.00")
           nME = 0
           gME = 0
           tME = 0
        
    ElseIf (Mid(lsCodCta, 6, 1) = Moneda.gMonedaExtranjera And Len(lsCodCta) = 12) Or (Mid(lsCodCta, 9, 1) = Moneda.gMonedaExtranjera And Len(lsCodCta) = 18) Then
           nMN = 0
           gMN = 0
           tMN = 0
           nME = Format((prRS.Fields(7) / pnTipCambio), "#0.00")
           gME = Format((prRS.Fields(6) / pnTipCambio), "#0.00")
           tME = nME + gME
    End If
   
        
    lsCadena = lsCadena & Space(6) & "IMPORTE A DEPOSITAR C.T.S. REMUNERAC- EN NUEVOS SOLES S/." & Space(20) & nMN & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(6) & "IMPORTE A DEPOSITAR C.T.S. GRATIFICAC- EN NUEVOS SOLES S/." & Space(19) & gMN & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(12) & "TOTAL IMPORTE A DEPOSITAR C.T.S. EN SOLES S/." & Space(26) & tMN & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & Space(6) & "IMPORTE A DEPOSITAR C.T.S. REMUNERAC- EN U.S. DOLARES $" & Space(22) & nME & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(6) & "IMPORTE A DEPOSITAR C.T.S. GRATIFICAC- EN DOLARES U.S. $" & Space(21) & gME & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(12) & "TOTAL IMPORTE A DEPOSITAR C.T.S. EN DOLARES US $" & Space(23) & tME & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & String(90, "_") & oImpresora.gPrnSaltoLinea
    
    '************************************************************************************
    lsCadena = lsCadena & String(95, "_") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "En la fecha ha sido depositado en la cuenta a plazo fijo C.T.S. N°  " & lsCodCta & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & ", de la CMAC - C, la suma de "
    
    If (Mid(lsCodCta, 6, 1) = Moneda.gMonedaNacional And Len(lsCodCta) = 12) Or (Mid(lsCodCta, 9, 1) = Moneda.gMonedaNacional And Len(lsCodCta) = 18) Then
        lsCadena = lsCadena & "S/.  " & Format(tMN, "0.00") & "  Nuevos Soles." & oImpresora.gPrnSaltoLinea
    ElseIf (Mid(lsCodCta, 6, 1) = Moneda.gMonedaExtranjera And Len(lsCodCta) = 12) Or (Mid(lsCodCta, 9, 1) = Moneda.gMonedaExtranjera And Len(lsCodCta) = 18) Then
        lsCadena = lsCadena & "$/.  " & Format(tME, "0.00") & "  Dolares Americanos.  T/C  " & pnTipCambio & oImpresora.gPrnSaltoLinea
    End If
        
    If Year(psFechaIni) = Year(psFechaFin) Then
       nMes = Month(psFechaFin) - Month(psFechaIni) + 1
    Else
       nMes = 12 - Month(psFechaIni) + 1 + Month(psFechaFin)
    End If
    
    If nMes > 1 Then
          sTextMes = "Meses"
       Else
          sTextMes = "Mes"
    End If
        
    lsCadena = lsCadena & "Por concepto de Compensación Por Tiempo de Servicios prestados a " & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "la CAJA MUNICIPAL DE AHORRO Y CREDITO CUSCO, durante " & nMes & " " & sTextMes & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "En conformmidad de lo antes señalado firmo la presente liquidación y hago constar que está " & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & "regulado por el T.U.O. de la ley de C.T.S. D.S. N° 001-97 y regalmento D.S. 004-97-TR," & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & " DU N° 127-2000 " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & String(95, "_") & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & Space(60) & "Cusco,  " & ArmaFecha(gdFecSis) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(60) & "_____________________________" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(60) & prRS.Fields(3) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(60) & "D.N.I. N° " & oImpresora.gPrnSaltoLinea
    
    'lsCadena = ""
    
    
'    RSet lsItem1 = Format(prRS.Fields(7), "#0.00")
'    lsCadena = lsCadena & Space(lnEspacio) & "MESES COMPUTABLES     : " & lsItem1 & oImpresora.gPrnSaltoLinea

'    lsCadena = lsCadena & Space(lnEspacio) & "FECHA INI PERIODO     :      " & psFechaIni & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "FECHA FIN PERIODO     :      " & psFechaFin & oImpresora.gPrnSaltoLinea

'    RSet lsItem1 = Format(prRS.Fields(5), "#0.00")
'    lsCadena = lsCadena & Space(lnEspacio) & "REMUNERACION          : " & lsItem1 & oImpresora.gPrnSaltoLinea
    
    'RSet lsItem1 = oInterprete.FunEvalua("I_IN_3_AFP", prRS.Fields(2), CDate(psFechaIni), CDate(psFechaIni), False, "VVVV", "")
    'lsCadena = lsCadena & Space(lnEspacio) & "INCREMENTO AFP 3.3%   : " & lsItem1 & oImpresora.gPrnSaltoLinea
'    RSet lsItem1 = Format(prRS.Fields(6), "#0.00")
'    lsCadena = lsCadena & Space(lnEspacio) & "PROM. GRATIFICA. " & IIf(Month(CDate(psFechaIni) = 11), "NAV", "FFP") & "  : " & lsItem1 & oImpresora.gPrnSaltoLinea
'    RSet lsItem1 = Format(CCur(prRS.Fields(6)) + CCur(prRS.Fields(5)), "#0.00")
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "REM. INDEMNIZABLE : " & lsItem1 & "    " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & Space(lnEspacio) & "CALCULO :" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "TIEMPO DE SERVICIOS" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "-------------------" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & lsItem1 & " / 12  x " & Format(prRS.Fields(7), "#0.00") & " MESES  : " & Format(prRS.Fields(11), "#0.00") & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "Tipo de Cambio : " & Format(3.65, "#,##0.00") & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & Space(lnEspacio) & "" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsNegritaOn
'    If (Mid(lsCodCta, 6, 1) = Moneda.gMonedaNacional And Len(lsCodCta) = 12) Or (Mid(lsCodCta, 9, 1) = Moneda.gMonedaNacional And Len(lsCodCta) = 18) Then
'        lsCadena = lsCadena & Space(lnEspacio) & "TOTAL DEPOSITADO EN MONEDA NACIONAL" & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & Space(lnEspacio) & "S/. " & Format(prRS.Fields(11), "#0.00") & oImpresora.gPrnSaltoLinea
'    ElseIf (Mid(lsCodCta, 6, 1) = Moneda.gMonedaExtranjera And Len(lsCodCta) = 12) Or (Mid(lsCodCta, 9, 1) = Moneda.gMonedaExtranjera And Len(lsCodCta) = 18) Then
'        lsCadena = lsCadena & Space(lnEspacio) & "TOTAL DEPOSITADO EN MONEDA EXTRANJERA" & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & Space(lnEspacio) & " $  " & Format(CCur(prRS.Fields(11)) / pnTipCambio, "#0.00") & oImpresora.gPrnSaltoLinea
'    End If
    lsCadena = lsCadena & lsNegritaOff & oImpresora.gPrnSaltoPagina
    
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    If lsCodCta <> "" Then
'        lsCadena = lsCadena & Space(lnEspacio) & "ENTIDAD DEPOSITARIA : " & lsNegritaOn & pgsEmpresa & lsNegritaOff
'    Else
'        lsCadena = lsCadena & Space(lnEspacio) & "ENTIDAD DEPOSITARIA : " & "<<NO ESPECIFICADA>>"
'    End If
    
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & JustificaTexto("HE RECIBIDO DE LA " & pgsEmpresa & " " & lsTexto, 80, lnEspacio) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena("---------------------------------", 80) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena(prRS.Fields(3), 80) & oImpresora.gPrnSaltoPagina
'
    Close ArcSal
    GetCartaCTS = lsCadena & lsCadena
End Function

Public Function GetCartas(prRS As ADODB.Recordset, pgsRuc As String, pgsEmpresa As String, psFechaIni As String, psFechaFin As String, pgdFecha As Date, psFechaImpre As String) As String
    Dim lsCadena As String
    Dim I As Integer
    Dim oInterprete As DInterprete
    Set oInterprete = New DInterprete
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    lsCadena = ""
    
    If Not (prRS.EOF And prRS.BOF) Then
        oInterprete.Interprete_InI
        oInterprete.IniTValor True
        lnTotDias = Abs(oInterprete.FunEvalua("G_TOT_UTIL_REP", "", CDate(psFechaIni), CDate(psFechaFin), True, "E04", ""))
        lnNumMonto = oInterprete.FunEvalua("G_VAL_UTIL", "", CDate(psFechaIni), CDate(psFechaFin), True, "E04", "")
        
        RaiseEvent ShowProgress
        While Not prRS.EOF
            RaiseEvent Progress(prRS.Bookmark, prRS.RecordCount)
            If prRS.Fields(0) = "1" Then
                'Me.Caption = "Generación Cartas " & Str(prRs.Fields(1))
                Set rs = oRH.GetRRHHAgenciaActual(prRS.Fields(2))
                lsCadena = lsCadena & GetCarta(prRS, pgsRuc, pgsEmpresa, psFechaIni, psFechaFin, pgdFecha, psFechaImpre, rs)
                RaiseEvent Progress(prRS.Bookmark, prRS.RecordCount)
            End If
            prRS.MoveNext
        Wend
        RaiseEvent CloseProgress
    End If
    
    GetCartas = lsCadena
End Function

'*********************************
'Private Function GetCarta(prRS As ADODB.Recordset, pgsRuc As String, pgsEmpresa As String, psFechaIni As String, psFechaFin As String, pgdFecha As Date, psFechaImpre As String, prRSAGE As ADODB.Recordset) As String
'    Dim lsCadena As String
'    Dim lsMargen As String
'    Dim lsNegritaOn As String
'    Dim lsNegritaOff As String
'    Dim lnEspacio As Integer
'    Dim lsMontoToTal As String * 15
'    Dim lsMontoToTalPar As String * 15
'    Dim lsTotDiasEmps As String * 15
'    Dim lsTotRemuEmps As String * 15
'    Dim lsItem1 As String * 15
'    Dim lsItem2 As String * 15
'
'    lnEspacio = 6
'    lsNegritaOn = oImpresora.gPrnBoldON
'    lsNegritaOff = oImpresora.gPrnBoldOFF
'
'    RSet lsMontoToTal = Format(lnTotDias, "#,##0.00")
'    RSet lsMontoToTalPar = Format(lnNumMonto, "#,##0.00")
'    RSet lsTotDiasEmps = Format(prRS.Fields(7), "#,##0.00")
'    RSet lsTotRemuEmps = Format(prRS.Fields(8), "#,##0.00")
'    RSet lsItem1 = Format(prRS.Fields(6), "#,##0.00")
'    RSet lsItem2 = Format(prRS.Fields(5), "#,##0.00")
'
'    lsCadena = ""
'    lsCadena = lsNegritaOn & Space(lnEspacio) & Space(lnEspacio) & CentrarCadena("LIQUIDACION DE UTILIDADES EJERCICIO AÑO " & Format(psFechaIni, "YYYY"), 80) & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
'    lsCadena = lsNegritaOn & Space(lnEspacio) & "RUC:" & pgsRuc & CentrarCadena(pgsEmpresa, 65) & psFechaImpre & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
'    lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena("(D.LEG. N 892 Y D.S. N 09-98-TR)", 80) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "CODIGO    : " & prRS.Fields(1) & " &   " & "NOMBRE : " & prRS.Fields(3) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "SITUACION : " & IIf(prRS.Fields(10) <> 0, "ACTIVO - " & prRSAGE.Fields(0) & " - " & prRSAGE.Fields(1), "RETIRADO") & "    " & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "REMUNERACION TOTAL PERCIBIDA AÑO " & Format(psFechaIni, "YYYY") & "    : " & lsItem1 & "    " & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "DIAS LABORADOS                           : " & lsItem2 & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "RENTA ANUAL ANTES DE IMPUESTO            : " & lsMontoToTal & "    " & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "PARTICIPACION LABORAL 5% UTIL            : " & lsMontoToTalPar & "    " & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "TOTAL DIAS LABORADOS TODOS LOS EMPLEADOS : " & lsTotDiasEmps & "    " & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "TOTAL REMUNERACION TODOS LOS EMPLEADOS   : " & lsTotRemuEmps & "    " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & Space(lnEspacio) & "CALCULO :" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "Importe por Dias Laborados + Importe por Ingresos Percibidos" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "(" & prRS.Fields(5) & " * " & Format(CCur(prRS.Fields(9)) / CCur(prRS.Fields(5)), "#,##0.000000000000") & ") + (" & prRS.Fields(6) & " * " & Format(CCur(prRS.Fields(10)) / CCur(prRS.Fields(6)), "#,##0.000000000000") & ")  =  " & oImpresora.gPrnBoldON & prRS.Fields(11) & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "Retenciones :" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "5ta Categoria" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & prRS.Fields(12) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & "Neto a Pagar" & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & oImpresora.gPrnBoldON & Format(CCur(prRS.Fields(11)) - CCur(prRS.Fields(12)), "#,##0.00") & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & JustificaTexto("HE  RECIBIDO DE LA CAJA MUNICIPAL DE  AHORRO Y CREDITO DE  TRUJILLO EL MONTO  DE    S/." & oImpresora.gPrnBoldON & Format(CCur(prRS.Fields(11)) - CCur(prRS.Fields(12)), "#,##0.00") & oImpresora.gPrnBoldOFF & " POR CONCEPTO DE PARTICIPACION LABORAL EN LAS UTILIDADES CORRESPONDIENTES AL EJERCICIO DEL " & Format(psFechaIni, "YYYY") & " EL MISMO QUE CONSIDERO CONFORME A LAS NORMAS LEGALES VIGENTES.", 80, lnEspacio) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena("---------------------------------", 80) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena(prRS.Fields(3), 80) & oImpresora.gPrnSaltoPagina
'
'    GetCarta = lsCadena & lsCadena
'End Function

Private Function GetCarta(prRS As ADODB.Recordset, pgsRuc As String, pgsEmpresa As String, psFechaIni As String, psFechaFin As String, pgdFecha As Date, psFechaImpre As String, prRSAGE As ADODB.Recordset) As String
    Dim lsCadena As String
    Dim lsMargen As String
    Dim lsNegritaOn As String
    Dim lsNegritaOff As String
    Dim lnEspacio As Integer
    Dim lsMontoToTal As String * 15
    Dim lsMontoToTalPar As String * 15
    Dim lsTotDiasEmps As String * 15
    Dim lsTotRemuEmps As String * 15
    Dim lsItem1 As String * 15
    Dim lsItem2 As String * 15
    Dim pardlab As Currency
    Dim paremper As Currency
        
    lnEspacio = 6
    lsNegritaOn = oImpresora.gPrnBoldON
    lsNegritaOff = oImpresora.gPrnBoldOFF
    
    RSet lsMontoToTal = Format(lnTotDias, "#,##0.00")
    RSet lsMontoToTalPar = Format(lnNumMonto, "#,##0.00")
    RSet lsTotDiasEmps = Format(prRS.Fields(7), "#,##0.00")
    RSet lsTotRemuEmps = Format(prRS.Fields(8), "#,##0.00")
    RSet lsItem1 = Format(prRS.Fields(6), "#,##0.00")
    RSet lsItem2 = Format(prRS.Fields(5), "#,##0.00")
    
    lsCadena = ""
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsNegritaOn & Space(lnEspacio) & Space(lnEspacio) & CentrarCadena("LIQUIDACION DE DISTRIBUCION DE UTILIDADES EJERCICIO " & Format(psFechaIni, "YYYY"), 80) & lsNegritaOff & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & lsNegritaOn & "LA CAJA MUNICIPAL DE AHORRO Y CREDITO CUSCO S.A. RUC: N° " & pgsRuc & lsNegritaOff & "Domiciliado en Portal" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & "Espinar N° 146, debidamente representada por el Sr. Carlos Tamayo Caparo, en su calidad de empleador y en" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & "cumplimiento de lo dispuesto por el D.Leg. N°892 y el D.S. N°09-98-TR, deja constancia de la determinación y el pago" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & "de la participación de utilidades del trabajador: " & prRS.Fields(3) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & Space(5) & lsNegritaOn & "CALCULO DEL MONTO DE PARTICIPACION DE UTILIDADES" & oImpresora.gPrnSaltoLinea & lsNegritaOff & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & lsNegritaOn & "1.Utilidad por Distribuir" & lsNegritaOff & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(5) & "- Renta Anual de la Empresa antes de Impuestos" & Space(45) & gcPEN_SIMBOLO & lsMontoToTal & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    lsCadena = lsCadena & Space(5) & Space(5) & "- Porcentaje a Distribuir " & Space(74) & "5 %" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(5) & "- Monto a Distribuir " & Space(70) & gcPEN_SIMBOLO & lsMontoToTalPar & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea 'marg ers044-2016
    
    lsCadena = lsCadena & Space(5) & lsNegritaOn & "2. Cálculo de la Participación" & lsNegritaOff & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(2) & "2.1.  Según días laborados" & lsNegritaOff & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(5) & "- Número de días laborados durante el ejercicio " & Format(psFechaIni, "YYYY") & " por todos los trabajadores" & Space(12) & lsTotDiasEmps & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(7) & "de la empresa con derechos a percibir utilidades" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(5) & "- Número de días laborados durante el ejercicio" & Format(psFechaIni, "YYYY") & "por el trabajador" & Space(21) & lsItem2 & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    pardlab = prRS.Fields(5) * Format(CCur(prRS.Fields(9)) / CCur(prRS.Fields(5)), "#,##0.000000000000")
    lsCadena = lsCadena & Space(5) & Space(5) & "- Participacion del trabajador según días laborados" & Space(49) & pardlab & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & Space(5) & Space(2) & lsNegritaOn & "2.2.  Según Remuneración Percibida" & lsNegritaOff & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(5) & "- Remuneracón computable total pagada durante el ejercicio " & Format(psFechaIni, "YYYY") & " por todos los trabajadrores" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(6) & "de la Empresa con derecho a percibir utilidades" & Space(43) & lsTotRemuEmps & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(5) & "- Remuneración computable percibida durante el ejercicio " & Format(psFechaIni, "YYYY") & " por el trabajador" & Space(12) & lsItem1 & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    paremper = prRS.Fields(6) * Format(CCur(prRS.Fields(10)) / CCur(prRS.Fields(6)), "#,##0.000000000000")
    lsCadena = lsCadena & Space(5) & Space(5) & "- Participación del trabajador según remuneración percibida" & Space(43) & paremper & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & Space(5) & lsNegritaOn & "3. Monto de la participacion a percibir por el trabajador" & lsNegritaOff & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(5) & "- Participación a percibir del trabajador en las utilidades" & Space(32) & Format(CCur(prRS.Fields(11)) - CCur(prRS.Fields(12)), "#,##0.00") & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    gdFecSis = Date
    lsCadena = lsCadena & Space(5) & Space(55) & "Cusco, " & ArmaFecha(gdFecSis) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(55) & "_____________________________" & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(55) & prRS.Fields(3) & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & Space(5) & Space(55) & "D.N.I." & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoPagina
         
     
    'hasta aqui es nuevo
    '*********************************************************************************
    'lsCadena = lsNegritaOn & Space(lnEspacio) & "RUC:" & pgsRuc & CentrarCadena(pgsEmpresa, 65) & psFechaImpre & lsNegritaOff & oImpresora.gPrnSaltoLinea & lsCadena
    'lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena("(D.LEG. N 892 Y D.S. N 09-98-TR)", 80) & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "CODIGO    : " & prRS.Fields(1) & " &   " & "NOMBRE : " & prRS.Fields(3) & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "SITUACION : " & IIf(prRS.Fields(10) <> 0, "ACTIVO - " & prRSAGE.Fields(0) & " - " & prRSAGE.Fields(1), "RETIRADO") & "    " & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "REMUNERACION TOTAL PERCIBIDA AÑO " & Format(psFechaIni, "YYYY") & "    : " & lsItem1 & "    " & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "DIAS LABORADOS                           : " & lsItem2 & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "RENTA ANUAL ANTES DE IMPUESTO            : " & lsMontoToTal & "    " & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "PARTICIPACION LABORAL 5% UTIL            : " & lsMontoToTalPar & "    " & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "TOTAL DIAS LABORADOS TODOS LOS EMPLEADOS : " & lsTotDiasEmps & "    " & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "TOTAL REMUNERACION TODOS LOS EMPLEADOS   : " & lsTotRemuEmps & "    " & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    
    'lsCadena = lsCadena & Space(lnEspacio) & "CALCULO :" & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "Importe por Dias Laborados + Importe por Ingresos Percibidos" & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "(" & prRS.Fields(5) & " * " & Format(CCur(prRS.Fields(9)) / CCur(prRS.Fields(5)), "#,##0.000000000000") & ") + (" & prRS.Fields(6) & " * " & Format(CCur(prRS.Fields(10)) / CCur(prRS.Fields(6)), "#,##0.000000000000") & ")  =  " & oImpresora.gPrnBoldON & prRS.Fields(11) & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "Retenciones :" & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "5ta Categoria" & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & prRS.Fields(12) & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & "Neto a Pagar" & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & oImpresora.gPrnBoldON & Format(CCur(prRS.Fields(11)) - CCur(prRS.Fields(12)), "#,##0.00") & oImpresora.gPrnBoldOFF & oImpresora.gPrnSaltoLinea
    
    'lsCadena = lsCadena & JustificaTexto("HE  RECIBIDO DE LA CAJA MUNICIPAL DE  AHORRO Y CREDITO DE  TRUJILLO EL MONTO  DE    S/." & oImpresora.gPrnBoldON & Format(CCur(prRS.Fields(11)) - CCur(prRS.Fields(12)), "#,##0.00") & oImpresora.gPrnBoldOFF & " POR CONCEPTO DE PARTICIPACION LABORAL EN LAS UTILIDADES CORRESPONDIENTES AL EJERCICIO DEL " & Format(psFechaIni, "YYYY") & " EL MISMO QUE CONSIDERO CONFORME A LAS NORMAS LEGALES VIGENTES.", 80, lnEspacio) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena("---------------------------------", 80) & oImpresora.gPrnSaltoLinea
    'lsCadena = lsCadena & Space(lnEspacio) & CentrarCadena(prRS.Fields(3), 80) & oImpresora.gPrnSaltoPagina
    
    GetCarta = lsCadena & lsCadena
End Function


#If DebugMode Then
    '##ModelId=3AB902DC0249
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Private Function GetBoletaLima(psPersCod As String, pdFecha As String, psPlaCod As String, psNombrePlanilla As String, psFechaImprePlanilla As String, pgsRuc As String, pgsEmpresa As String, psTextoPlanilla As String) As String
    Dim lsCadena As String
    Dim rsP As ADODB.Recordset
    Set rsP = New ADODB.Recordset
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    On Error Resume Next
    
    Dim lsCod As String * 8
    Dim lsNombre As String * 36
    Dim lsCargo As String * 27
    Dim lsAFP As String * 5
    Dim lsFecIng As String * 12
    Dim lsNumC As String * 20
    Dim lsSueldo As String * 15
    Dim lsPorMesLabT As String * 10
    Dim lsPorMesLab As String * 10
    Dim lsNumCarnetAFP As String * 20
    Dim lsCondiciaon As String * 16
    Dim lsCondicion As String * 16
    Dim lsAgencia As String * 10
    Dim lsVacaciones As String * 22
    Dim lsDNI As String * 10
    Dim lsConcep(20, 15) As String * 28
    Dim lsConcepT(1, 15) As String * 28
    Dim lnEspacioSep As Integer
    Dim I As Integer
    Dim j As Integer
    Dim lnI As Integer
    Dim lnD As Integer
    Dim lnA As Integer
    Dim lnPosMax As Integer
    Dim lsCondicionE As String
    Dim lsNumCE As String
    Dim ldFecVacIni As String
    Dim ldFecVacFin As String
    
    Dim lsNegritaOn As String
    Dim lsNegritaOff As String
    Dim lnEspacio As Integer
    Dim lnnI As Integer
    Dim lnnJ As Integer
    
    Dim lsLinea As String * 162
    Dim lnEspacioRep As Integer
    
    For lnnI = 0 To 20
        For lnnJ = 0 To 6
            lsConcep(lnnI, lnnJ) = ""
        Next lnnJ
    Next lnnI
    
    
    lnEspacio = 0
    lsNegritaOn = oImpresora.gPrnBoldON
    lsNegritaOff = oImpresora.gPrnBoldOFF

    Set rsP = oPla.GetBoleta(psPersCod, pdFecha, psPlaCod)
    
    lnI = 0
    lnD = 0
    If rsP.EOF And rsP.BOF Then
        rsP.Close
        Set rsP = Nothing
        Exit Function
    End If
    
    
    lnI = 0
    lnD = 0
    lnA = 0
    rsP.MoveFirst
    While Not rsP.EOF
        If rsP!nMonto <> 0 Or Trim(rsP!cRHConceptoCod) = gsRHConceptoITOTING Or Trim(rsP!cRHConceptoCod) = gsRHConceptoDTOTDES Then
            If Left(rsP!cRHConceptoCod, 2) = "I_" Or Left(rsP!cRHConceptoCod, 2) = "J_" Then
                lnI = lnI + 1
                If IsNull(rsP!cRHConceptoDescripcion) Then
                    lsConcep(lnI, 1) = Trim(rsP!cRHConceptoDescripcion)
                Else
                    lsConcep(lnI, 1) = Trim(rsP!cRHConceptoDescripcion)
                End If
                lsConcep(lnI, 2) = FillNum(Format(rsP!nMonto, "#,##0.00"), 25, " ")
                lsConcep(lnI, 3) = FillNum(rsP!Concepcod, 15, " ")
                
            ElseIf Left(rsP!cRHConceptoCod, 2) = "D_" Or Left(rsP!cRHConceptoCod, 2) = "E_" Then
                lnD = lnD + 1
                If IsNull(rsP!cRHConceptoDescripcion) Then
                    lsConcep(lnD, 4) = Trim(rsP!cRHConceptoDescripcion)
                Else
                    lsConcep(lnD, 4) = Trim(rsP!cRHConceptoDescripcion)
                End If
                lsConcep(lnD, 5) = FillNum(Format(rsP!nMonto, "#,##0.00"), 25, " ")
                lsConcep(lnD, 6) = FillNum(rsP!Concepcod, 15, " ")
            ElseIf Left(rsP!cRHConceptoCod, 2) = "A_" Or Left(rsP!cRHConceptoCod, 2) = "B_" Then
                lnA = lnA + 1
                If IsNull(rsP!cRHConceptoDescripcion) Then
                    lsConcep(lnA, 7) = Trim(rsP!cRHConceptoDescripcion)
                Else
                    lsConcep(lnA, 7) = Trim(rsP!cRHConceptoDescripcion)
                End If
                lsConcep(lnA, 8) = FillNum(Format(rsP!nMonto, "#,##0.00"), 25, " ")
                lsConcep(lnA, 9) = FillNum(rsP!Concepcod, 15, " ")
            Else
                lsPorMesLab = Trim(Str(rsP!nMonto))
            End If
        End If
        rsP.MoveNext
    Wend
    
    If IsNumeric(lsPorMesLab) Then
        If psPlaCod = gsRHPlanillaSueldos Then
            lsPorMesLab = Round(Trim(Str(CCur(lsPorMesLab) * 30)), 4)
        Else
            lsPorMesLab = Round(Trim(Str(CCur(lsPorMesLab))), 4)
        End If
    End If
    
    rsP.MoveFirst
    
    lsCod = rsP!cRHCod
    lsNombre = PstaNombre(rsP!cPersNombre, False)
    lsCargo = oRH.GetCargo(psPersCod, pdFecha)
    lsFecIng = Format(rsP!dIngreso, gsFormatoFechaView)
    lsNumC = rsP!cNumAFP & ""
    lsDNI = rsP!DNI & ""
    
    lsSueldo = Format(oRH.GetSueldo(psPersCod), "#,##0.00")
    
    If Not IsNull(rsP!nRHContratoTpo) Then
        If rsP!nRHContratoTpo = 1 Then
            lsCondicion = "CONTRATADO"
        Else
            lsCondicion = "ESTABLE"
        End If
    Else
        lsCondicion = "CONTRATADO"
    End If
    lsAgencia = rsP!cAgenciaActual
    lsAFP = rsP!cRHAFPAbreviatura & ""
    lsNumCarnetAFP = rsP!cCUSPP & ""
    
    If oRH.GetFecVac(psPersCod, ldFecVacIni, ldFecVacFin) Then
        lsVacaciones = Format(ldFecVacIni, gsFormatoFechaView) & "-" & Format(ldFecVacFin, gsFormatoFechaView)
    Else
        lsVacaciones = "__/__/____-__/__/____"
    End If
    
    lnEspacioSep = 3
    lsVacaciones = ""
    
    lsLinea = Space(lnEspacio) & lsCod & "    " & lsNombre & "     " & lsCargo & Space(lnEspacioSep) & lsCod & "    " & lsNombre & "     " & lsCargo
    lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsLinea = Space(lnEspacio) & lsNumC & lsDNI & "   " & lsPorMesLab & Space(38) & Space(lnEspacioSep) & lsNumC & lsDNI & "   " & lsPorMesLab
    lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsLinea = Space(lnEspacio) & lsVacaciones & "                  " & lsFecIng & "    " & lsSueldo & lsPorMesLab & Space(lnEspacioSep) & lsVacaciones & "                  " & lsFecIng & "    " & lsSueldo & lsPorMesLab
    lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsLinea = Space(63) & lsCondicion & Space(65) & lsCondicion
    lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsCadena = oImpresora.gPrnSaltoLinea & lsCadena
    lsCadena = oImpresora.gPrnSaltoLinea & lsCadena
    
    lsLinea = oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsNegritaOn & Space(lnEspacio) & Space(10) & CentrarCadena(Format(CDate(psFechaImprePlanilla), "MMMM                YYYY"), 80) & lsNegritaOff & Space(5) & lsNegritaOn & Space(lnEspacio) & Space(lnEspacio) & CentrarCadena(Format(CDate(psFechaImprePlanilla), "MMMM                YYYY"), 80) & lsNegritaOff
    lsCadena = lsLinea & oImpresora.gPrnSaltoLinea & lsCadena
    
    If lnI <> 0 Or lnD <> 0 And lnA <> 0 Then
        lsConcepT(1, 1) = lsConcep(lnI, 1)
        lsConcepT(1, 2) = lsConcep(lnI, 2)
        If lnD > 0 Then lsConcepT(1, 4) = lsConcep(lnD, 4)
        If lnD > 0 Then
            lsConcepT(1, 5) = lsConcep(lnD, 5)
        Else
            lsConcepT(1, 5) = "0.00"
        End If
        lsConcepT(1, 7) = "NETO A PAGAR"
        lsConcepT(1, 8) = FillNum(Format(CCur(lsConcepT(1, 2) - CCur(lsConcepT(1, 5))), "#,##0.00"), 15, " ")
        lsConcep(lnI, 1) = ""
        lsConcep(lnI, 2) = ""
        If lnD > 0 Then lsConcep(lnD, 4) = ""
        If lnD > 0 Then lsConcep(lnD, 5) = ""
        If lnA > 0 Then lsConcep(lnA, 7) = ""
        If lnA > 0 Then lsConcep(lnA, 8) = ""
        lnI = lnI - 1
        lnD = lnD - 1
        lnA = lnA - 1
    End If
    
    For j = 1 To lnI
        If IsNumeric(lsConcep(j, 2)) Then
            lsLinea = Space(lnEspacio) & Left(lsConcep(j, 3), 15) & "  " & lsConcep(j, 1) & lsConcep(j, 2) & Space(lnEspacioSep) & Left(lsConcep(j, 3), 15) & "  " & lsConcep(j, 1) & lsConcep(j, 2)
            lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea
        End If
    Next j
    lsConcepT(1, 3) = ""
    
    lsLinea = lsNegritaOn & Space(lnEspacio) & Left(lsConcepT(1, 3), 15) & "  " & lsConcepT(1, 1) & lsConcepT(1, 2) & lsNegritaOff & Space(lnEspacioSep) & lsNegritaOn & Space(lnEspacio) & Left(lsConcepT(1, 3), 15) & "  " & lsConcepT(1, 1) & lsConcepT(1, 2) & lsNegritaOff
    lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    For j = 1 To lnD
        If IsNumeric(lsConcep(j, 5)) Then
            lsLinea = Space(lnEspacio) & Left(lsConcep(j, 6), 15) & "  " & lsConcep(j, 4) & lsConcep(j, 5) & Space(lnEspacioSep) & Left(lsConcep(j, 6), 15) & "  " & lsConcep(j, 4) & lsConcep(j, 5)
            lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea
        End If
    Next j
    lsConcepT(1, 6) = ""
    
    lsLinea = lsNegritaOn & Space(lnEspacio) & Left(lsConcepT(1, 6), 15) & "  " & lsConcepT(1, 4) & lsConcepT(1, 5) & lsNegritaOff & Space(lnEspacioSep) & lsNegritaOn & Space(lnEspacio) & Left(lsConcepT(1, 6), 15) & "  " & lsConcepT(1, 4) & lsConcepT(1, 5) & lsNegritaOff
    lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    For j = 1 To lnA
        If IsNumeric(lsConcep(j, 8)) Then
            lsLinea = Space(lnEspacio) & Left(lsConcep(j, 9), 15) & "  " & lsConcep(j, 7) & lsConcep(j, 8) & Space(lnEspacioSep) & Space(lnEspacio) & Left(lsConcep(j, 9), 15) & "  " & lsConcep(j, 7) & lsConcep(j, 8)
            lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea
        End If
    Next j
    lsConcepT(1, 9) = ""
    
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    
    lsLinea = lsNegritaOn & Space(lnEspacio) & Left(lsConcepT(1, 9), 15) & "  " & lsConcepT(1, 7) & lsConcepT(1, 8) & lsNegritaOff & Space(lnEspacioSep) & lsNegritaOn & Space(lnEspacio) & Left(lsConcepT(1, 9), 15) & "  " & lsConcepT(1, 7) & lsConcepT(1, 8) & lsNegritaOff
    lsCadena = lsCadena & lsLinea & oImpresora.gPrnSaltoLinea
    
    Set rsP = Nothing
    Set oPla = Nothing
    Set oRH = Nothing
    
    GetBoletaLima = lsCadena
    Exit Function
ERROR:
    MsgBox "Error al imprimeir la boleta - " & Err.Description, vbInformation, "Aviso"
End Function

'Public Function GetRep5taCat(psPersCod As String, psDoc As String, psNombre As String, psAnio As String, psRUC As String, pdFecha As Date, pnAddIES As Integer, pnPorDevolver As Integer) As String
'    Dim Sql As String
'    Dim rsIng As ADODB.Recordset
'    Dim rsDes As ADODB.Recordset
'    Dim rsUtil As ADODB.Recordset
'    Dim rsTotales As ADODB.Recordset
'    Dim rsIES As ADODB.Recordset
'    Dim rsDev5ta As ADODB.Recordset
'
'    Dim rsRH As ADODB.Recordset
'    Set rsRH = New ADODB.Recordset
'    Dim sqlRH As String
'
'    Set rsIng = New ADODB.Recordset
'    Set rsDes = New ADODB.Recordset
'    Set rsUtil = New ADODB.Recordset
'    Set rsTotales = New ADODB.Recordset
'    Set rsIES = New ADODB.Recordset
'    Set rsDev5ta = New ADODB.Recordset
'
'    Dim lsIng  As String * 15
'    Dim lsDes  As String * 15
'    Dim lsUtil As String * 15
'    Dim lsTotA As String * 15
'    Dim lsTotB As String * 15
'    Dim lsImp  As String * 15
'    Dim lsIES  As String * 15
'    Dim lsDev5ta As String * 15
'
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    Dim oInt As DInterprete
'    Set oInt = New DInterprete
'
'    Dim lsCadena As String
'    Dim lsMargen As String
'    Dim lnMargen As Integer
'
'    lnMargen = 20
'    lsMargen = Space(lnMargen)
'
'    oCon.AbreConexion
'
'    sqlRH = "Select cRHCod, cAgenciaActual +  ' ' + (Select cAgeDescripcion from agencias where cAgeCod = cAgenciaActual) Agencia   From RRHH Where cPersCod = '" & psPersCod & "' "
'    Set rsRH = oCon.CargaRecordSet(sqlRH)
'
'    oInt.Interprete_InI
'    oInt.IniTValor True
'    If psAnio <= 2001 Then
'        RSet lsImp = Format(21000, "#,##0.00")
'    ElseIf psAnio <= 2003 Then
'        RSet lsImp = Format(21700, "#,##0.00")
'    ElseIf psAnio <= 2004 Then
'        RSet lsImp = Format(22400, "#,##0.00")
'    Else
'        RSet lsImp = Format(ExprANum(oInt.FunEvalua("V_UIT_7", psPersCod, pdFecha, pdFecha, False, "VVVV", "")), "#,##0.00")
'    End If
'
'    Sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '1%' And bRHConceptoImp5ta = 1 And cRHConceptoCod Not In ('147', '148'))"
'    Set rsIng = oCon.CargaRecordSet(Sql)
'
'    Sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' and cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '1%' And bRHConceptoImp5ta = 1 and cRHConceptoCod in ('147', '148'))"
'    Set rsUtil = oCon.CargaRecordSet(Sql)
'
'    Sql = " Select IsNull(Sum(nMonto),0) MontoT , Case When IsNull(Sum(nmonto),0) - " & Format(lsImp, "#.00") & " < 0 Then 0 Else IsNull(Sum(nmonto),0) - " & Format(lsImp, "#.00") & " End  From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod in (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '1%' and bRHConceptoImp5ta = 1)"
'    Set rsTotales = oCon.CargaRecordSet(Sql)
'
'    Sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '2%' And bRHConceptoImp5ta = 1)"
'    Set rsDes = oCon.CargaRecordSet(Sql)
'
'    Sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod In (" & gsRHCAportExtSolid & "," & gsRHCAportLiquidaExtSolid & "," & gsRHCAportExtSolidManual & "))"
'    Set rsIES = oCon.CargaRecordSet(Sql)
'
'    Sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod In (" & gRHConcIngDev5taCategoria & "))"
'    Set rsDev5ta = oCon.CargaRecordSet(Sql)
'
'    lsCadena = ""
'
'
'
'    RSet lsIng = Format(rsIng.Fields(0), "#,##0.00")
'    If pnPorDevolver = 0 Then
'        RSet lsDes = Format(rsDes.Fields(0) - rsDev5ta.Fields(0), "#,##0.00")
'    Else
'        RSet lsDes = Format(rsDes.Fields(0), "#,##0.00")
'    End If
'    RSet lsUtil = Format(rsUtil.Fields(0), "#,##0.00")
'    RSet lsTotA = Format(rsTotales.Fields(0), "#,##0.00")
'    RSet lsTotB = Format(rsTotales.Fields(1), "#,##0.00")
'    RSet lsIES = Format(rsIES.Fields(0), "#,##0.00")
'    RSet lsDev5ta = Format(rsDev5ta.Fields(0), "#,##0.00")
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("CERTIFICADO DE RETENCIONES SOBRE RENTAS DE QUINTA CATEGORIA", 60) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("PERCIBIDAS POR EL TRABAJADOR DURANTE EL EJERCICIO " & psAnio & "", 60) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & JustificaTexto("LA CAJA MUNICIPAL DE AHORRO Y CREDITO DE TRUJILLO CON RUC Nro. " & psRUC & ", Y CON DOMICILIO FISCAL EN LA AV. ESPAÑA 2611 - 2615, TRUJILLO:", 60, lnMargen)
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("CERTIFICA", 60) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena(rsRH.Fields(1) & " - " & rsRH.Fields(0), 60) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & JustificaTexto("Que a don (doña): " & psNombre & ", identificado (a) con LE. " & psDoc & "  se le ha retenido el importe de: S/. " & Format(rsDes.Fields(0) - rsDev5ta.Fields(0), "#,##0.00") & "  como pago a cuenta del impuesto a la renta, correspondiente al ejercicio gravable " & psAnio & " calculado en base a los siguientes ingresos percibidos.", 60, lnMargen)
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "1.  RENTAS BRUTAS DE 5TA. CATEGORÍA" & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & "  _Remuneración percibida en el " & psAnio & "                 " & lsIng & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & "  _Participación en Utilidades ejercicio de " & CInt(psAnio) - 1 & "     " & lsUtil & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "  RENTA BRUTA ANUAL PERCIBIDA                        " & lsTotA & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "2.  DEDUCCIÓN DE LA QUINTA CATEGORÍA (7 UIT)         " & lsImp & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "3.  RENTA IMPONIBLE                                  " & lsTotB & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "4.  IMPUESTO A LA RENTA RETENIDO                     " & lsDes & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    If CCur(lsDev5ta) > 0 And pnPorDevolver = 1 Then
'        lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "5.  IMPUESTO A LA RENTA DEVUELTO                     " & lsDev5ta & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'        lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    End If
'
'    If pnAddIES = 1 Then
'        lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "(*) IMPUESTO EXTRAORDINARIO DE SOLIDARIDAD (" & psAnio & ")    " & lsIES & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    End If
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & Format(pdFecha, "Trujillo, dd mmmm - yyyy") & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("---------------------------------------------", 60) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena(psNombre, 60) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'
'    GetRep5taCat = lsCadena
'End Function


'Public Function GetRep5taCatTot(psAnio As String, psRUC As String, pdFecha As Date, pnAddIES As Integer, pnPorDevolver As Integer) As String
'    Dim Sql As String
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    Dim lsCadena As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    oCon.AbreConexion
'
'    Sql = " Select RH.cRHCod, RH.cPersCod, dbo.PstaNombre(PE.cPersNombre,1) Nombre , IsNull((Select cPersIDnro From PersID PID Where PID.cPersCod = RH.cPersCod And PID.cPersIDTpo = 1),'') DNI from RRHH RH " _
'        & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
'        & " Where RH.cRHCod Like 'E%' And RH.nRHEstado Not Like '[78]%'  And RH.cPersCod In (Select Distinct cPersCod From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPlanillaCod = '" & gsRHPlanillaSueldos & "') Order By RH.cAgenciaActual"
'    Set rs = oCon.CargaRecordSet(Sql)
'
'    lsCadena = ""
'
'    RaiseEvent ShowProgress
'    While Not rs.EOF
'        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina & GetRep5taCat(rs!cPersCod, rs!DNI, rs!Nombre, psAnio, psRUC, pdFecha, pnAddIES, pnPorDevolver)
'        RaiseEvent Progress(rs.Bookmark, rs.RecordCount)
'        rs.MoveNext
'    Wend
'    RaiseEvent CloseProgress
'
'    GetRep5taCatTot = lsCadena
'End Function

Public Function GetHorasTrabajadas(pdIni As Date, pnDiasLab As Currency) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oInt As DInterprete
    Set oInt = New DInterprete
    Dim oconect As NConstSistemas
    Set oconect = New NConstSistemas
    
    
    Dim ldFecAsist1 As Date
    Dim ldFecAsist2 As Date
    Dim ldFecAsist3 As Date
    Dim ldFecAsist4 As Date
    Dim ldFecAsist5 As Date
    
    oCon.AbreConexion
    
    ldFecAsist1 = pdIni
    ldFecAsist2 = pdIni
    ldFecAsist3 = pdIni
    ldFecAsist4 = pdIni
    ldFecAsist5 = pdIni
    
    oInt.ObtFechasAsistencia ldFecAsist1, ldFecAsist2, ldFecAsist3, ldFecAsist4, ldFecAsist5
    
    If pnDiasLab < 23 Then
        ldFecAsist2 = DateAdd("d", pnDiasLab, ldFecAsist1)
    End If
    
    If IIf(oconect.LeeConstSistema(gConstSistBitCentral) = "1", True, False) Then
        sql = "Select count(*) Num from dbcomunes..feriado where dFeriado Between '" & Format(ldFecAsist1, gcFormatoFecha) & "' And '" & Format(ldFecAsist2, gcFormatoFecha) & "'"
    Else
        sql = "Select count(*) Num from dbcomunes..feriado where dFeriado Between '" & Format(ldFecAsist1, gcFormatoFecha) & "' And '" & Format(ldFecAsist2, gcFormatoFecha) & "'"
    End If
    
    Set rs = oCon.CargaRecordSet(sql)
    
    GetHorasTrabajadas = (pnDiasLab - rs!Num) * 7.29167
    rs.Close
    
    Set rs = Nothing
    Set oCon = Nothing

End Function

Public Function GetDiasFediados(pdIni As Date, pnPorcentMesTrab As Currency) As Integer
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim oInt As DInterprete
    Set oInt = New DInterprete
    Dim oconect As NConstSistemas
    Set oconect = New NConstSistemas
    
    
    Dim ldFecAsist1 As Date
    Dim ldFecAsist2 As Date
    Dim ldFecAsist3 As Date
    Dim ldFecAsist4 As Date
    Dim ldFecAsist5 As Date
    
    oCon.AbreConexion
    
    ldFecAsist1 = pdIni
    ldFecAsist2 = pdIni
    ldFecAsist3 = pdIni
    ldFecAsist4 = pdIni
    ldFecAsist5 = pdIni
    
    oInt.ObtFechasAsistencia ldFecAsist1, ldFecAsist2, ldFecAsist3, ldFecAsist4, ldFecAsist5
    
    If IIf(oconect.LeeConstSistema(gConstSistBitCentral) = "1", True, False) Then
        sql = "Select count(*) Num from feriado where dFeriado Between '" & Format(ldFecAsist1, gcFormatoFecha) & "' And '" & Format(ldFecAsist2, gcFormatoFecha) & "'"
    Else
        sql = "Select count(*) Num from feriado where dFeriado Between '" & Format(ldFecAsist1, gcFormatoFecha) & "' And '" & Format(ldFecAsist2, gcFormatoFecha) & "'"
    End If
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If pnPorcentMesTrab < 0.3 Then
        GetDiasFediados = (rs!Num * pnPorcentMesTrab)
    Else
        GetDiasFediados = (rs!Num)
    End If
    
    rs.Close
    
    Set rs = Nothing
    Set oCon = Nothing
End Function



Public Function ListadoBoletas(psRHPeriodo As String, psAgencia As String, pnImprime As Integer) As Recordset
    On Error GoTo ListadoBoletasPlaErr
    Dim oCon As DActualizaDatosConPlanilla
    Set oCon = New DActualizaDatosConPlanilla
    'ListadoBoletas = 1
    Set ListadoBoletas = oCon.GetRHListadoBoletas(psRHPeriodo, psAgencia, pnImprime)
    Set oCon = Nothing
    Exit Function
ListadoBoletasPlaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ListadoBoletas Method")
End Function


Public Function GetRep5taCatTot(psAnio As String, psRUC As String, pdFecha As Date, pnAddIES As Integer) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lsCadena As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    sql = " Select RH.cRHCod, RH.cPersCod, dbo.PstaNombre(PE.cPersNombre,1) Nombre , IsNull((Select cPersIDnro From PersID PID Where PID.cPersCod = RH.cPersCod And PID.cPersIDTpo = 1),'') DNI from RRHH RH " _
        & " Inner Join Persona PE On RH.cPersCod = PE.cPersCod" _
        & " Where RH.cRHCod Like 'E%' And RH.nRHEstado Not Like '[78]%'  And RH.cPersCod In (Select Distinct cPersCod From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPlanillaCod = 'E01') Order By RH.cRHCod"
    Set rs = oCon.CargaRecordSet(sql)
    
    lsCadena = ""
    
    RaiseEvent ShowProgress
    While Not rs.EOF
        lsCadena = lsCadena & oImpresora.gPrnSaltoPagina & GetRep5taCat(rs!cPersCod, rs!DNI, rs!Nombre, psAnio, psRUC, pdFecha, pnAddIES)
        RaiseEvent Progress(rs.Bookmark, rs.RecordCount)
        rs.MoveNext
    Wend
    RaiseEvent CloseProgress
    GetRep5taCatTot = lsCadena
End Function

'Public Function GetRep5taCat(psPersCod As String, psDoc As String, psNombre As String, psAnio As String, psRUC As String, pdFecha As Date, pnAddIES As Integer) As String
'    Dim sql As String
'    Dim rsIng As ADODB.Recordset
'    Dim rsDes As ADODB.Recordset
'    Dim rsUtil As ADODB.Recordset
'    Dim rsTotales As ADODB.Recordset
'    Dim rsIES As ADODB.Recordset
'    Dim rsDev5ta As ADODB.Recordset
'
'    Set rsIng = New ADODB.Recordset
'    Set rsDes = New ADODB.Recordset
'    Set rsUtil = New ADODB.Recordset
'    Set rsTotales = New ADODB.Recordset
'    Set rsIES = New ADODB.Recordset
'    Set rsDev5ta = New ADODB.Recordset
'
'    Dim lsIng  As String * 15
'    Dim lsDes  As String * 15
'    Dim lsUtil As String * 15
'    Dim lsTotA As String * 15
'    Dim lsTotB As String * 15
'    Dim lsImp  As String * 15
'    Dim lsIES  As String * 15
'    Dim lsDev5ta As String * 15
'
'    Dim lsCredConImp As String
'
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    Dim oInt As DInterprete
'    Set oInt = New DInterprete
'
'    Dim lsCadena As String
'    Dim lsMargen As String
'    Dim lnMargen As Integer
'
'    lnMargen = 20
'    lsMargen = Space(lnMargen)
'
'    oCon.AbreConexion
'
'    oInt.Interprete_InI
'    oInt.IniTValor True
'    If psAnio <= 2001 Then
'        RSet lsImp = Format(21000, "#,##0.00")
'    Else
'        RSet lsImp = Format(ExprANum(oInt.FunEvalua("V_UIT_7", psPersCod, pdFecha, pdFecha, False, "VVVV", "")), "#,##0.00")
'    End If
'
'    Select Case psAnio
'        Case 2003
'        Case 2004
'            lsImp = "      " + Trim(Format(22400, "#,##0.00"))
'    End Select
'
'
'
'    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '1%' And bRHConceptoImp5ta = 1 And cRHConceptoCod Not In ('147', '148')) AND cPlanillaCod IN ('E01','E06','E02','E15','E08','E10') "
'    Set rsIng = oCon.CargaRecordSet(sql)
'
'    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' and cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '1%' And bRHConceptoImp5ta = 1 ) AND cPlanillaCod IN ('E04')"
'    Set rsUtil = oCon.CargaRecordSet(sql)
'
'    sql = " Select IsNull(Sum(nMonto),0) MontoT , Case When IsNull(Sum(nmonto),0) - " & Format(lsImp, "#.00") & " < 0 Then 0 Else IsNull(Sum(nmonto),0) - " & Format(lsImp, "#.00") & " End  From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod in (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '1%' and bRHConceptoImp5ta = 1)"
'    Set rsTotales = oCon.CargaRecordSet(sql)
'
'    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '2%' And bRHConceptoImp5ta = 1)"
'    Set rsDes = oCon.CargaRecordSet(sql)
'
'    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod In (" & gsRHCAportExtSolid & "," & gsRHCAportLiquidaExtSolid & "," & gsRHCAportExtSolidManual & "))"
'    Set rsIES = oCon.CargaRecordSet(sql)
'
'    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
'        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod In (" & gRHConcIngDev5taCategoria & "))"
'    Set rsDev5ta = oCon.CargaRecordSet(sql)
'
'    lsCadena = ""
'
'
'
'    RSet lsIng = Format(rsIng.Fields(0), "#,##0.00")
'    RSet lsDes = Format(rsDes.Fields(0) - rsDev5ta.Fields(0), "#,##0.00")
'    RSet lsUtil = Format(rsUtil.Fields(0), "#,##0.00")
'    RSet lsTotA = Format(rsTotales.Fields(0), "#,##0.00")
'    RSet lsTotB = Format(rsTotales.Fields(1), "#,##0.00")
'    RSet lsIES = Format(rsIES.Fields(0), "#,##0.00")
'    RSet lsDev5ta = Format(rsDev5ta.Fields(0), "#,##0.00")
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("CERTIFICADO DE RETENCIONES SOBRE RENTAS DE QUINTA CATEGORIA", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("PERCIBIDAS POR EL TRABAJADOR DURANTE EL EJERCICIO " & psAnio & "", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & JustificaTexto("LA CAJA MUNICIPAL DE AHORRO Y CREDITO DE ICA CON RUC Nro. " & psRUC & ", Y CON DOMICILIO FISCAL EN LA AV. MUNICIPALIDAD Nro.152 , ICA:", 70, lnMargen)
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("CERTIFICA", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & JustificaTexto("Que a don (doña): " & psNombre & ", identificado (a) con LE. " & psDoc & "  se le ha retenido el importe de: S/. " & Format(rsDes.Fields(0), "#,##0.00") & "  por concepto del impuesto a la Renta de quinta Categoria, correspondiente al ejercicio gravable " & psAnio & " calculado en base a los siguientes rentas,deducciones y créditos :", 70, lnMargen)
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "1.  RENTAS BRUTAS DE 5TA. CATEGORÍA" & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & " _ Sueldos,asignaciones,primas,gratificaciones," & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & "   bonificaciones,aguinaldos,comisiones,etc     " & Space(7) & lsIng & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & " _ Participación en Utilidades ejercicio de " & CInt(psAnio) - 1 & Space(7) & lsUtil & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "           REMUNERACION BRUTA ANUAL PERCIBIDA        " & lsTotA & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "2. DEDUCCIÓN (7 UIT)   " & Space(30) & lsImp & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "3. REMUNERACION IMPONIBLE                            " & lsTotB & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "4. IMPUESTO A LA RENTA                        " & lsDes & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
''
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "5. IMPUESTO TOTAL  CALCULADO                        " & Format(rsTotales.Fields(1) * 15 / 100, "#,##0.00") & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "4. IMPUESTO TOTAL   RETENIDO                         " & lsDes & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "5. DEVOLUCION DEL EXCESO/DEFECTO                            " & Format(rsDes.Fields(0) - (rsDev5ta.Fields(0) - (rsDes.Fields(0) - rsDev5ta.Fields(0))), "#,##0.00") & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
'
'
'
'    'If CCur(lsDev5ta) > 0 Then
'    '    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "5.  IMPUESTO A LA RENTA DEVUELTO                     " & lsDev5ta & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    '    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    '    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    'End If
'
'    If pnAddIES = 1 Then
'        lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "(*) IMPUESTO EXTRAORDINARIO DE SOLIDARIDAD (" & psAnio & ") " & lsIES & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    End If
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    'pdFecha
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "Huancayo , " & Format(gdFecSis, "dd mmmm - yyyy") & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
'
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("---------------------------------------------", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("Representante Legal", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'
'    GetRep5taCat = lsCadena
'End Function

Public Function GetRep5taCat(psPersCod As String, psDoc As String, psNombre As String, psAnio As String, psRUC As String, pdFecha As Date, pnAddIES As Integer) As String
    Dim sql As String
    Dim rsIng As ADODB.Recordset
    Dim rsDes As ADODB.Recordset
    Dim rsUtil As ADODB.Recordset
    Dim rsTotales As ADODB.Recordset
    Dim rsIES As ADODB.Recordset
    Dim rsDev5ta As ADODB.Recordset
    
    Set rsIng = New ADODB.Recordset
    Set rsDes = New ADODB.Recordset
    Set rsUtil = New ADODB.Recordset
    Set rsTotales = New ADODB.Recordset
    Set rsIES = New ADODB.Recordset
    Set rsDev5ta = New ADODB.Recordset
    
    Dim lsIng  As String * 15
    Dim lsDes  As String * 15
    Dim lsUtil As String * 15
    Dim lsTotA As String * 15
    Dim lsTotB As String * 15
    Dim lsImp  As String * 15
    Dim lsIES  As String * 15
    Dim lsDev5ta As String * 15
    
    Dim lsCredConImp As String
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    Dim oInt As DInterprete
    Set oInt = New DInterprete
       
    Dim lsCadena As String
    Dim lsMargen As String
    Dim lnMargen As Integer
    
    lnMargen = 20
    lsMargen = Space(lnMargen)
    
    oCon.AbreConexion
           
    oInt.Interprete_InI
    oInt.IniTValor True
    If psAnio <= 2001 Then
        RSet lsImp = Format(21000, "#,##0.00")
    Else
        RSet lsImp = Format(ExprANum(oInt.FunEvalua("V_UIT_7", psPersCod, pdFecha, pdFecha, False, "VVVV", "")), "#,##0.00")
    End If
    
    Select Case psAnio
        Case 2003
        Case 2004
            lsImp = "      " + Trim(Format(22400, "#,##0.00"))
    End Select
    
    
       
    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '1%' And bRHConceptoImp5ta = 1 And cRHConceptoCod Not In ('147', '148')) AND cPlanillaCod IN ('E01','E06','E02','E15','E08','E10') "
    Set rsIng = oCon.CargaRecordSet(sql)

    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' and cPersCod = '" & psPersCod & "'" _
        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '1%' And bRHConceptoImp5ta = 1 ) AND cPlanillaCod IN ('E04')"
    Set rsUtil = oCon.CargaRecordSet(sql)

    sql = " Select IsNull(Sum(nMonto),0) MontoT , Case When IsNull(Sum(nmonto),0) - " & Format(lsImp, "#.00") & " < 0 Then 0 Else IsNull(Sum(nmonto),0) - " & Format(lsImp, "#.00") & " End  From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
        & " And cRHConceptoCod in (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '1%' and bRHConceptoImp5ta = 1)"
    Set rsTotales = oCon.CargaRecordSet(sql)

    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod Like '2%' And bRHConceptoImp5ta = 1)"
    Set rsDes = oCon.CargaRecordSet(sql)
    
    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod In (" & gsRHCAportExtSolid & "," & gsRHCAportLiquidaExtSolid & "," & gsRHCAportExtSolidManual & "))"
    Set rsIES = oCon.CargaRecordSet(sql)
    
    sql = " Select IsNull(Sum(nMonto),0) Monto From RHPlanillaDetCon Where cRRHHPeriodo Like '" & psAnio & "%' And cPersCod = '" & psPersCod & "'" _
        & " And cRHConceptoCod In (Select cRHConceptoCod From RHConceptoTabla Where cRHConceptoCod In (" & gRHConcIngDev5taCategoria & "))"
    Set rsDev5ta = oCon.CargaRecordSet(sql)
    
    lsCadena = ""

    
    
    RSet lsIng = Format(rsIng.Fields(0), "#,##0.00")
    RSet lsDes = Format(rsDes.Fields(0) - rsDev5ta.Fields(0), "#,##0.00")
    RSet lsUtil = Format(rsUtil.Fields(0), "#,##0.00")
    RSet lsTotA = Format(rsTotales.Fields(0), "#,##0.00")
    RSet lsTotB = Format(rsTotales.Fields(1), "#,##0.00")
    RSet lsIES = Format(rsIES.Fields(0), "#,##0.00")
    RSet lsDev5ta = Format(rsDev5ta.Fields(0), "#,##0.00")
    
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("CERTIFICADO DE RETENCIONES SOBRE RENTAS DE QUINTA CATEGORIA", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
       
    
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("PERCIBIDAS POR EL TRABAJADOR DURANTE EL EJERCICIO " & psAnio & "", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & JustificaTexto("LA CAJA MUNICIPAL DE AHORRO Y CREDITO MAYNAS CON RUC Nro. " & psRUC & ", Y CON DOMICILIO FISCAL EN Portal Espinar N° 146, CUSCO:", 70, lnMargen)
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("CERTIFICA", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & JustificaTexto("Que a don (doña): " & psNombre & ", identificado (a) con LE. " & psDoc & "  se le ha retenido el importe de: S/. " & Format(rsDes.Fields(0), "#,##0.00") & "  por concepto del impuesto a la Renta de quinta Categoria, correspondiente al ejercicio gravable " & psAnio & " calculado en base a los siguientes rentas,deducciones y créditos :", 70, lnMargen)
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "1.  RENTAS BRUTAS DE 5TA. CATEGORÍA" & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & lsMargen & " _ Sueldos,asignaciones,primas,gratificaciones," & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & "   bonificaciones,aguinaldos,comisiones,etc     " & Space(7) & lsIng & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & " _ Participación en Utilidades ejercicio de " & CInt(psAnio) - 1 & Space(7) & lsUtil & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "           REMUNERACION BRUTA ANUAL PERCIBIDA        " & lsTotA & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea

    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "2. DEDUCCIÓN (7 UIT)   " & Space(30) & lsImp & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea

    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "3. REMUNERACION IMPONIBLE                            " & lsTotB & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "4. IMPUESTO A LA RENTA                        " & lsDes & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea

    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "5. IMPUESTO TOTAL  CALCULADO                        " & Format(rsTotales.Fields(1) * 15 / 100, "#,##0.00") & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea

''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "4. IMPUESTO TOTAL   RETENIDO                         " & lsDes & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
''
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "5. DEVOLUCION DEL EXCESO/DEFECTO                            " & Format(rsDes.Fields(0) - (rsDev5ta.Fields(0) - (rsDes.Fields(0) - rsDev5ta.Fields(0))), "#,##0.00") & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
''    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    
    
    
    
    'If CCur(lsDev5ta) > 0 Then
    '    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "5.  IMPUESTO A LA RENTA DEVUELTO                     " & lsDev5ta & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    '    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    '    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    'End If
    
    If pnAddIES = 1 Then
        lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "(*) IMPUESTO EXTRAORDINARIO DE SOLIDARIDAD (" & psAnio & ") " & lsIES & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    End If
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    'pdFecha
    gdFecSis = Date
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & "Maynas , " & Format(gdFecSis, "dd mmmm - yyyy") & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsMargen & oImpresora.gPrnSaltoLinea
    
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("---------------------------------------------", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & lsMargen & oImpresora.gPrnNegritaON & CentrarCadena("Representante Legal", 70) & oImpresora.gPrnNegritaOFF & oImpresora.gPrnSaltoLinea

    GetRep5taCat = lsCadena
End Function


Public Function ActualizaConceptosPla(psRHPlanillaCod As String, psRHCodPersona As String, psRHCodConcepto As String, psRHCodMonto As Double, psActualiza As String) As Integer
    On Error GoTo ActualizaConceptosPlaErr
    Dim oCon As DActualizaDatosConPlanilla
    Set oCon = New DActualizaDatosConPlanilla
    oCon.ActualizaRHConceptoPla psRHPlanillaCod, psRHCodPersona, psRHCodConcepto, psRHCodMonto, psActualiza
    Set oCon = Nothing
    Exit Function
ActualizaConceptosPlaErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ActualizaConceptosPla Method")
End Function

'*** PEAC 20131118
Public Function ModificaNumCtaComentario(psCodEmpModiComent As String, psTipoPlanilla As String, psPeriodo As String, psCtaCod As String) As Boolean
    On Error GoTo ModificaNumCtaComentarioErr
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla

    oPla.ModificaNumCtaComentario psCodEmpModiComent, psTipoPlanilla, psPeriodo, psCtaCod

    Set oPla = Nothing
    Exit Function
ModificaNumCtaComentarioErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConceptoPlanilla:ModificaNumCtaComentario Method")
End Function

