VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NRHConcepto"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9701FD02FA"
'Actualiza los Conceptos
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902DB01A8
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902DB027A
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902DB0248
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

'Agrega un nuevo concepto
'##ModelId=3A9AFAF900FF
Public Function AgregaConcepto(psRHConceptoCod As String, psRHConceptoDescripcion As String, psRHConceptoTpo As String, psRHConceptoFormula As String, pnRHConceptoOrden As Integer, psRHConceptoAbreviatura As String, pbRHConceptoImp5ta As String, pbRHConceptoMesTrab As String, psRHConceptoCtaCont As String, psRHNemo As String, psUltimaActualizacion As String, psRHConceptoGrupo As String) As String
    On Error GoTo AgregaConceptoErr
    Dim lsCodigo As String
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    lsCodigo = oCon.GetCodigo(psRHConceptoGrupo)
    oCon.AgregaConcepto lsCodigo, psRHConceptoDescripcion, psRHConceptoTpo, "1", psRHConceptoFormula, pnRHConceptoOrden, psRHConceptoAbreviatura, pbRHConceptoImp5ta, pbRHConceptoMesTrab, psRHConceptoCtaCont, psRHNemo, psUltimaActualizacion
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:AgregaConcepto Method")
End Function



Public Function AgregaConceptoRRHH(psPersCod As String, psRHPlanillaCod As String, psRHConceptoCod As String, pnRHConceptoValor As String, psUltimaActualizacion As String, Optional pbTodoslosGrupos As Boolean = False) As String
    On Error GoTo AgregaConceptoErr
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    oCon.AgregaConceptoRRHH psPersCod, psRHPlanillaCod, psRHConceptoCod, pnRHConceptoValor, psUltimaActualizacion
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:AgregaConcepto Method")
End Function

Public Function ModificaConceptoRRHH(psPersCod As String, psRHPlanillaCod As String, psRHConceptoCod As String, pnRHConceptoValor As String, psUltimaActualizacion As String, Optional pbTodoslosGrupos As Boolean = False) As String
    On Error GoTo AgregaConceptoErr
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    oCon.ModificaConceptoRRHH psPersCod, psRHPlanillaCod, psRHConceptoCod, pnRHConceptoValor, psUltimaActualizacion
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:AgregaConcepto Method")
End Function

Public Function EliminaConceptoRRHH(psPersCod As String, psRHPlanillaCod As String, psRHConceptoCod As String, Optional pbTodoslosGrupos As Boolean = False) As String
    On Error GoTo AgregaConceptoErr
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    oCon.EliminaConceptoRRHH psPersCod, psRHPlanillaCod, psRHConceptoCod
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:AgregaConcepto Method")
End Function


'Elimina  un concepto
'##ModelId=3A9AFAF90127
Public Function EliminaConcepto(psRHConceptoCod As String) As Boolean
    On Error GoTo EliminaConceptoErr
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    oCon.EliminaConcepto psRHConceptoCod

    Exit Function
EliminaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:EliminaConcepto Method")
End Function

'Actualiza un concepto
'##ModelId=3A9AFAF9014F
'Public Function ModificaConcepto(psRHConceptoCod As String, psRHConceptoDescripcion As String, psRHConceptoTpo As String, psRHConceptoFormula As String, pnRHConceptoOrden As Integer, psRHConceptoAbreviatura As String, pbRHConceptoImp5ta As String, pbRHConceptoMesTrab As String, pbRHObligatorio As String, psRHConceptoCtaCont As String, psRHNemo As String, psUltimaActualizacion As String) As Boolean
'    On Error GoTo ModificaConceptoErr
'    Dim OCON As DRHConcepto
'    Set OCON = New DRHConcepto
'
'    OCON.ModificaConcepto psRHConceptoCod, psRHConceptoDescripcion, psRHConceptoTpo, "1", psRHConceptoFormula, pnRHConceptoOrden, psRHConceptoAbreviatura, pbRHConceptoImp5ta, pbRHConceptoMesTrab, pbRHObligatorio, psRHConceptoCtaCont, psRHNemo, psUltimaActualizacion
'
'    Set OCON = Nothing
'    Exit Function
'ModificaConceptoErr:
'    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:ModificaConcepto Method")
'End Function

Public Function ModificaConcepto(psRHConceptoCod As String, psRHConceptoDescripcion As String, psRHConceptoTpo As String, psRHConceptoFormula As String, pnRHConceptoOrden As Integer, psRHConceptoAbreviatura As String, pbRHConceptoImp5ta As String, pbRHConceptoMesTrab As String, psRHConceptoCtaCont As String, psRHNemo As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaConceptoErr
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    oCon.ModificaConcepto psRHConceptoCod, psRHConceptoDescripcion, psRHConceptoTpo, "1", psRHConceptoFormula, pnRHConceptoOrden, psRHConceptoAbreviatura, pbRHConceptoImp5ta, pbRHConceptoMesTrab, psRHConceptoCtaCont, psRHNemo, psUltimaActualizacion
    
    Set oCon = Nothing
    Exit Function
ModificaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:ModificaConcepto Method")
End Function


Public Function GetReporte(pgsNomAge As String, pgsEmpresa As String, pgdFecSis As Date, pgsCodUser As String) As String
    Dim sqlI As String
    Dim rsI As New ADODB.Recordset
    Dim lnP As Long
    Dim lnI As Long
    
    Dim lsCadena As String
    Dim lsFormula As String
    Dim lsEspacio As String * 2
    
    Dim lsItem0 As String * 5
    Dim lsItem1 As String * 20
    Dim lsItem2 As String * 25
    Dim lsItem3 As String * 20
    Dim lsItem4 As String * 3
    Dim lsItem5 As String * 27
    Dim lsItem6 As String * 43
    
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    Set rsI = oCon.GetConceptos(, False, True)
    lsCadena = ""
    lnP = 0
    lnI = 0
    
    If Not RSVacio(rsI) Then
        lsCadena = lsCadena & CabeceraPagina("Relación de Conceptos para el calculo de planilla.-" & pgsCodUser, lnP, lnI, pgsNomAge, pgsEmpresa, pgdFecSis, "NNN")
        lsCadena = lsCadena & Encabezado("Codigo;6;Nemonico;10; ;15;Nombre;12; ;15;Grupo;10; ;7;Est.;5; ;4;Formula;17; ;38;UltMov;7; ;7;", lnI)
    
        While Not rsI.EOF
            
            lsItem0 = rsI!codigo
            lsItem1 = rsI!Nemo
            lsItem2 = rsI!Descrip
            lsItem3 = rsI!DescripcionTpo
            lsItem4 = rsI!Estado
            lsItem5 = rsI!Ult
            lsFormula = rsI!Formula
            lsItem6 = lsFormula
            
            While Trim(lsFormula) <> ""
                lnI = lnI + 1
                lsCadena = lsCadena & lsItem0 & lsItem1 & lsEspacio & lsItem2 & lsEspacio & lsItem3 & lsEspacio _
                         & lsItem4 & lsEspacio & lsEspacio & lsItem6 & lsEspacio & lsItem5 & oImpresora.gPrnSaltoLinea
                lsItem1 = ""
                lsItem2 = ""
                lsItem3 = ""
                lsItem4 = ""
                lsItem5 = ""
                lsFormula = Mid(lsFormula, 44)
                lsItem6 = lsFormula
                
            Wend

            If lnI > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lnI = 0
                lsCadena = lsCadena & CabeceraPagina("Relación de Conceptos para el calculo de planilla.-" & pgsCodUser, lnP, lnI, pgsNomAge, pgsEmpresa, pgdFecSis, "NNN")
                lsCadena = lsCadena & Encabezado("Codigo;6;Nemonico;10; ;15;Nombre;12; ;15;Grupo;10; ;7;Est.;5; ;4;Formula;17; ;38;UltMov;7; ;7;", lnI)
            End If
            
            rsI.MoveNext
        Wend
        
    End If
    rsI.Close
    Set rsI = Nothing
    GetReporte = lsCadena
    Set oCon = Nothing
End Function

Public Function AgregaTablaAlias(psCodigo As String, psNombre As String, psAlias As String, psUltMov As String) As Boolean
    On Error GoTo GetConceptosErr
    Dim oCon As DRHConcepto
    Dim lsCodigo As String
    Set oCon = New DRHConcepto
    
    psCodigo = oCon.GetCodigoTablaAlias(psCodigo)
    oCon.AgregaTablaAlias psCodigo, psNombre, psAlias, psUltMov
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

Public Function EliminaTablaAlias(psCodigo As String) As Boolean
    On Error GoTo GetConceptosErr
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto

    oCon.EliminaTablaAlias psCodigo
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

'Private Function GetRepo(Optional psCodEmp As String = "") As String
'    Dim lsCadena As String
'    Dim lsCabecera As String
'    Dim lbBan As Boolean
'    Dim Item1 As String * 8
'    Dim Item2 As String * 12
'    Dim Item3 As String * 37
'
'    Dim sqlC As String
'    Dim rsC As New ADODB.Recordset
'
'    Dim lnItem  As Integer
'
'    Dim I As Integer
'    Dim j As Integer
'    Dim lnCon As Integer
'
'    Dim Item4 As String * 24
'
'    Dim lsNegritaOn As String
'    Dim lsNegritaOff As String
'
'    Dim slqE As String
'    Dim rsE As New ADODB.Recordset
'
'    lsNegritaOn = oImpresora.gPrnBoldON
'    lsNegritaOff = oImpresora.gPrnBoldOFF
'
'    lsCadena = ""
'    lsCadena = lsCadena & lsNegritaOn & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & CentrarCadena("Conceptos por Persona", 106) & oImpresora.gPrnSaltoLinea
'    lsCadena = lsCadena & CentrarCadena("=========================", 106) & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea
'
'    lnItem = 4
'
'    Item1 = "CodEmp"
'    Item2 = "CodPers"
'    Item3 = "Nombre"
'    Item4 = "MEMONICOS DE CONCEPTOS"
'
'    lsCadena = lsCadena & Item1 & Item2 & Item3 & Space(2) & Item4 & lsNegritaOff & oImpresora.gPrnSaltoLinea
'
'    lsCabecera = lsCadena
'
'    If psCodEmp = "" Then
'        slqE = "Select E.cCodPers, E.cEmpCod, PE.cNomPers From Empleado E Inner Join " & gcCentralPers & "Persona PE On E.cCodPers = PE.cCodPers Where cEmpEst <> '3'"
'    Else
'        slqE = "Select E.cCodPers, E.cEmpCod, PE.cNomPers From Empleado E Inner Join " & gcCentralPers & "Persona PE On E.cCodPers = PE.cCodPers Where cEmpEst <> '3' and E.cEmpCod = '" & psCodEmp & "'"
'    End If
'
'    rsE.Open slqE, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    While Not rsE.EOF
'        lbBan = False
'        Item1 = rsE!cEmpCod
'        Item2 = rsE!cCodPers
'        Item3 = PstaNombre(rsE!cNomPers, False)
'
'        lsCadena = lsCadena & Item1 & Item2 & Item3
'
'        Item1 = ""
'        Item2 = ""
'        Item3 = ""
'        lnCon = -1
'
'        sqlC = " Select cConcepNom from empcon EM" _
'             & " Inner Join Concepto CO On EM.cConcepCod = CO.cConcepCod" _
'             & " Where cEmpCod = '" & rsE!cEmpCod & "' and cCodPers = '" & rsE!cCodPers & "' and cPlaCod = '" & Left(cmbPlaCod, 5) & "' Order By EM.cConcepCod"
'        rsC.Open sqlC, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'        If Not RSVacio(rsC) Then
'            While Not rsC.EOF
'                Item4 = Trim(rsC!cConcepNom)
'                lnCon = lnCon + 1
'                If lnCon Mod 3 = 0 And lnCon <> 0 Then
'                    lsCadena = lsCadena & oImpresora.gPrnSaltoLinea  & Item1 & Item2 & Item3 & Space(2) & Item4
'                Else
'                    lsCadena = lsCadena & Space(2) & Item4
'                End If
'                rsC.MoveNext
'            Wend
'        End If
'
'        rsC.Close
'        Set rsC = Nothing
'
'        lnItem = lnItem + 1
'        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
'
'        If lnItem = 57 Then
'            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
'            lsCadena = lsCadena & lsCabecera
'            lnItem = 4
'        End If
'
'        rsE.MoveNext
'    Wend
'
'    rsE.Close
'    Set rsE = Nothing
'    GetRepo = lsCadena
'End Function



#If DebugMode Then
    '##ModelId=3AB902DB01BC
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If


Public Function AgregaCalculoEPS(psCodCalculoEps As String, psDescripcion As String, psUltimaActualziacion As String, pnEstadoEPS As Integer) As Integer
    On Error GoTo AgregaConceptoErr
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    psCodCalculoEps = Left(psCodCalculoEps, 14) + "EPS"
    AgregaCalculoEPS = 0
    oCon.AgregaRHCalculoEPS psCodCalculoEps, psDescripcion, psUltimaActualziacion, pnEstadoEPS
    AgregaCalculoEPS = 1
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:AgregaConcepto Method")
End Function

Public Function AgregaRHCalculoEPSDet(psCodCalculoEps As String, psCodPersona As String, pnSueldo As Currency, pnSueldo_x_225 As Currency, pnCantPersonas As Integer, pnPlanSinIGV As Currency, pnPromedio As Currency, _
        pnNeto As Currency, pnPagaEmpleado As Currency, pnPagaEmpresa As Currency, pnAdicionalHijos As Currency, pnAdicionalPadres As Currency, pnTotalEmpleado As Currency, _
        pnDesQuincena As Currency, pnDesMensual As Currency, pnDesQuincenaUno As Currency, pnDescQuincenaDos As Currency, pnSaldo As Currency) As String

    On Error GoTo AgregaConceptoErr
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    AgregaRHCalculoEPSDet = 0
    oCon.AgregaRHCalculoEPSDetalle psCodCalculoEps, psCodPersona, pnSueldo, pnSueldo_x_225, pnCantPersonas, pnPlanSinIGV, pnPromedio, _
        pnNeto, pnPagaEmpleado, pnPagaEmpresa, pnAdicionalHijos, pnAdicionalPadres, pnTotalEmpleado, _
        pnDesQuincena, pnDesMensual, pnDesQuincenaUno, pnDescQuincenaDos, pnSaldo
    AgregaRHCalculoEPSDet = 1
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:AgregaConcepto Method")
End Function

Public Function ActualizaRHCalEPSDet(psCodCalculoEps As String) As String

    On Error GoTo AgregaConceptoErr
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    
    ActualizaRHCalEPSDet = 0
    
    oCon.ActualizaRHCalculoEPSDet psCodCalculoEps
    
    ActualizaRHCalEPSDet = 1
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "ActualizaRHCalEPSDet:ActualizaRHCalEPSDet Method")
End Function

