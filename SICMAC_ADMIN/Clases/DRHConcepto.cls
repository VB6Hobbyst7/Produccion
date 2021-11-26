VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRHConcepto"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A9AFE2D0086"
'Actualiza Conceptos
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902FC0097
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902FC01F5
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902FC01AF
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
'##ModelId=3A9AFE7700D2
'Public Function AgregaConcepto(psRHConceptoCod As String, psRHConceptoDescripcion As String, psRHConceptoTpo As String, psRHConceptoEstado As String, psRHConceptoFormula As String, pnRHConceptoOrden As Integer, psRHConceptoAbreviatura As String, pbRHConceptoImp5ta As String, pbRHConceptoMesTrab As String, pbRHObligatorio As String, psRHConceptoCtaCont As String, psRHNemo As String, psUltimaActualizacion As String) As Boolean
'    On Error GoTo AgregaConceptoErr
'    Dim sqlD As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    sqlD = " Insert RHConceptotabla (cRHConceptoCod,cRHConceptoDescripcion,cRHConceptoTpo,cRHConceptoEstado,cRHConceptoFormula,nRHConceptoOrden,cRHConceptoAbreviatura,bRHConceptoImp5ta,bRHConceptoMesTrab, bRHObligatorio,cRHConceptoCtaCont,cRHconceptoMeno,cUltimaActualizacion)" _
'         & " Values('" & psRHConceptoCod & "','" & psRHConceptoDescripcion & "','" & psRHConceptoTpo & "','" & psRHConceptoEstado & "','" & psRHConceptoFormula & "'," & pnRHConceptoOrden & ",'" & psRHConceptoAbreviatura & "'," & pbRHConceptoImp5ta & "," & pbRHConceptoMesTrab & "," & pbRHObligatorio & ",'" & psRHConceptoCtaCont & "','" & psRHNemo & "','" & psUltimaActualizacion & "')"
'
'    If oCon.AbreConexion Then
'        oCon.Ejecutar sqlD
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'AgregaConceptoErr:
'    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:AgregaConcepto Method")
'End Function

Public Function AgregaConcepto(psRHConceptoCod As String, psRHConceptoDescripcion As String, psRHConceptoTpo As String, psRHConceptoEstado As String, psRHConceptoFormula As String, pnRHConceptoOrden As Integer, psRHConceptoAbreviatura As String, pbRHConceptoImp5ta As String, pbRHConceptoMesTrab As String, psRHConceptoCtaCont As String, psRHNemo As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaConceptoErr
    Dim sqlD As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlD = " Insert RHConceptotabla (cRHConceptoCod,cRHConceptoDescripcion,cRHConceptoTpo,cRHConceptoEstado,cRHConceptoFormula,nRHConceptoOrden,cRHConceptoAbreviatura,bRHConceptoImp5ta,bRHConceptoMesTrab,cRHConceptoCtaCont,cRHconceptoMeno,cUltimaActualizacion)" _
         & " Values('" & psRHConceptoCod & "','" & psRHConceptoDescripcion & "','" & psRHConceptoTpo & "','" & psRHConceptoEstado & "','" & psRHConceptoFormula & "'," & pnRHConceptoOrden & ",'" & psRHConceptoAbreviatura & "'," & pbRHConceptoImp5ta & "," & pbRHConceptoMesTrab & ",'" & psRHConceptoCtaCont & "','" & psRHNemo & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlD
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:AgregaConcepto Method")
End Function

Public Function AgregaConceptoRRHH(psPersCod As String, psRHPlanillaCod As String, psRHConceptoCod As String, pnRHConceptoValor As String, psUltimaActualizacion As String, Optional pbTodoslosGrupos As Boolean = False) As String
    On Error GoTo AgregaConceptoErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    If pbTodoslosGrupos Then
        If psRHPlanillaCod = "E06" Then
            sqlC = " Insert RHConcepto (cPersCod,cRHPlanillaCod,cRHConceptoCod,cUltimaActualizacion)" _
                 & " Select cPersCod, '" & psRHPlanillaCod & "', '" & psRHConceptoCod & "','" & psUltimaActualizacion & "' from RRHH RH " _
                 & " Where RH.nRHEstado < 700"
        Else
            sqlC = " Insert RHConcepto (cPersCod,cRHPlanillaCod,cRHConceptoCod,cUltimaActualizacion)" _
                 & " Select cPersCod, cRHPlanillaCod, '" & psRHConceptoCod & "','" & psUltimaActualizacion & "' from RRHH RH " _
                 & " Inner Join  RHPlanillaAplicacion RHPA On RH.nRHEstado = RHPA.nRHEstado" _
                 & " Where cRHPlanillaCod = '" & psRHPlanillaCod & "'"
        End If
    Else
        sqlC = " Insert RHConcepto (cPersCod,cRHPlanillaCod,cRHConceptoCod,nRHConceptoValor,cUltimaActualizacion)" _
             & " Values('" & psPersCod & "','" & psRHPlanillaCod & "','" & psRHConceptoCod & "'," & IIf(pnRHConceptoValor = "", "NULL", pnRHConceptoValor) & ",'" & psUltimaActualizacion & "')"
    End If
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:AgregaConcepto Method")
End Function

Public Function EliminaConceptoRRHH(psPersCod As String, psRHPlanillaCod As String, psRHConceptoCod As String, Optional pbTodoslosGrupos As Boolean = False) As String
    On Error GoTo AgregaConceptoErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim lsPlanilla As String
    
    If pbTodoslosGrupos Then
        sqlC = " Delete RHConcepto " _
             & " Where cRHPlanillaCod in ('" & psRHPlanillaCod & "') And cRHConceptoCod = '" & psRHConceptoCod & "'"
    Else
        sqlC = " Delete RHConcepto " _
             & " Where cPersCod = '" & psPersCod & "' And cRHPlanillaCod = '" & psRHPlanillaCod & "' And cRHConceptoCod = '" & psRHConceptoCod & "'"
    End If
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:AgregaConcepto Method")
End Function

Public Function ModificaConceptoRRHH(psPersCod As String, psRHPlanillaCod As String, psRHConceptoCod As String, pnRHConceptoValor As String, psUltimaActualizacion As String, Optional pbTodoslosGrupos As Boolean = False) As String
    On Error GoTo AgregaConceptoErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    If pbTodoslosGrupos Then
        sqlC = " Update RHConcepto " _
             & " Set nRHConceptoValor = " & pnRHConceptoValor & " " _
             & " Where cRHPlanillaCod = '" & psRHPlanillaCod & "' And cRHConceptoCod = '" & psRHConceptoCod & "'"
    Else
        sqlC = " Update RHConcepto " _
             & " Set nRHConceptoValor = " & pnRHConceptoValor & " " _
             & " Where cPersCod = '" & psPersCod & "' And cRHPlanillaCod = '" & psRHPlanillaCod & "' And cRHConceptoCod = '" & psRHConceptoCod & "'"
    End If
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaConceptoErr:
    Call RaiseError(MyUnhandledError, "NActualizaDatosConcepto:AgregaConcepto Method")
End Function

'Elimina  un concepto
'##ModelId=3A9AFE77010E
Public Function EliminaConcepto(psRHConceptoCod As String) As Boolean
    On Error GoTo EliminaConceptoErr
    Dim sqlD As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlD = " Update RHConceptotabla " _
         & " Set cRHConceptoEstado = '0'" _
         & " Where cRHConceptoCod = '" & psRHConceptoCod & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlD
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaConceptoErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:EliminaConcepto Method")
End Function

'Actualiza un concepto
'##ModelId=3A9AFE77014A
'Public Function ModificaConcepto(psRHConceptoCod As String, psRHConceptoDescripcion As String, psRHConceptoTpo As String, psRHConceptoEstado As String, psRHConceptoFormula As String, pnRHConceptoOrden As Integer, psRHConceptoAbreviatura As String, pbRHConceptoImp5ta As String, pbRHConceptoMesTrab As String, pbRHObligatorio As String, psRHConceptoCtaCont As String, psRHNemo As String, psUltimaActualizacion As String) As Boolean
'    On Error GoTo ModificaConceptoErr
'    Dim sqlD As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    sqlD = " Update RHConceptotabla" _
'         & " Set cRHConceptoDescripcion = '" & psRHConceptoDescripcion & "', cRHConceptoTpo = '" & psRHConceptoTpo & "'," _
'         & " cRHConceptoFormula = '" & psRHConceptoFormula & "', nRHConceptoOrden = " & pnRHConceptoOrden & ", cRHConceptoAbreviatura = '" & psRHConceptoAbreviatura & "', " _
'         & " bRHConceptoImp5ta = " & pbRHConceptoImp5ta & ", bRHConceptoMesTrab = " & pbRHConceptoMesTrab & ", cRHConceptoCtaCont = '" & psRHConceptoCtaCont & "', cRHconceptoMeno = '" & psRHNemo & "', cUltimaActualizacion = '" & psUltimaActualizacion & "', bRHObligatorio = " & pbRHObligatorio & " " _
'         & " Where cRHConceptoCod = '" & psRHConceptoCod & "'"
'
'    If oCon.AbreConexion Then
'        oCon.Ejecutar sqlD
'        oCon.CierraConexion
'    End If
'
'    Set oCon = Nothing
'    Exit Function
'ModificaConceptoErr:
'    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:ModificaConcepto Method")
'End Function

Public Function ModificaConcepto(psRHConceptoCod As String, psRHConceptoDescripcion As String, psRHConceptoTpo As String, psRHConceptoEstado As String, psRHConceptoFormula As String, pnRHConceptoOrden As Integer, psRHConceptoAbreviatura As String, pbRHConceptoImp5ta As String, pbRHConceptoMesTrab As String, psRHConceptoCtaCont As String, psRHNemo As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaConceptoErr
    Dim sqlD As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlD = " Update RHConceptotabla" _
         & " Set cRHConceptoDescripcion = '" & psRHConceptoDescripcion & "', cRHConceptoTpo = '" & psRHConceptoTpo & "'," _
         & " cRHConceptoFormula = '" & psRHConceptoFormula & "', nRHConceptoOrden = " & pnRHConceptoOrden & ", cRHConceptoAbreviatura = '" & psRHConceptoAbreviatura & "', " _
         & " bRHConceptoImp5ta = " & pbRHConceptoImp5ta & ", bRHConceptoMesTrab = " & pbRHConceptoMesTrab & ", cRHConceptoCtaCont = '" & psRHConceptoCtaCont & "', cRHconceptoMeno = '" & psRHNemo & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cRHConceptoCod = '" & psRHConceptoCod & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlD
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaConceptoErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:ModificaConcepto Method")
End Function

'Devuelve los conceptos Activos,
'
'El parametro psTpoConcepto si es vacio dretorna todos los tipos de conceptos
'caso contrario retorna todos los conceptos del tipo indicado
'##ModelId=3A9AFE7A0380
Public Function GetConceptos(Optional psRHConceptoGrupo As String = "", Optional pbSoloActivos As Boolean = True, Optional pbOrdenarporNemonico As Boolean = False) As Recordset
    On Error GoTo GetConceptosErr
    Dim sqlD As String
    Dim sqlDAux As String
    Dim sqlDWhere As String
    Dim sqlDOrder As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlDAux = ""
    If psRHConceptoGrupo <> "" Then
        If pbSoloActivos Then
            sqlDAux = " And cRHConceptoCod like '" & psRHConceptoGrupo & "%'"
        Else
            sqlDAux = " Where cRHConceptoCod like '" & psRHConceptoGrupo & "%'"
        End If
    End If
    If pbSoloActivos Then sqlDWhere = " Where cRHConceptoEstado = 1 "
    
    If pbOrdenarporNemonico Then
        sqlDOrder = "  Order by cRHconceptoMeno"
    Else
        sqlDOrder = "  Order by cRHConceptoCod"
    End If
    
    sqlD = "Select cRHConceptoCod Codigo, cRHConceptoDescripcion Descrip, cRHConceptoTpo Tipo, cRHConceptoFormula Formula, nRHConceptoOrden Orden, cRHConceptoAbreviatura Impre, bRHConceptoImp5ta Imp5, bRHConceptoMesTrab MesTrab, cRHConceptoCtaCont CtaCont, cRHconceptoMeno Nemo, cConsDescripcion DescripcionTpo, cRHConceptoEstado Estado, RHCT.cUltimaActualizacion Ult from rhconceptotabla  RHCT Inner Join Constante CO On RHCT.cRHConceptoTpo = CO.nConsValor And nConsCod  = '" & gRHPlanillaTipoCalculo & "' " & sqlDWhere & sqlDAux & sqlDOrder
    
    If oCon.AbreConexion Then
        Set GetConceptos = oCon.CargaRecordSet(sqlD)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

'Devuelve los conceptos Activos,
'
'El parametro psTpoConcepto si es vacio dretorna todos los tipos de conceptos
'caso contrario retorna todos los conceptos del tipo indicado
'##ModelId=3A9AFE7A0380
Public Function GetConceptosPlanilla(psRHPlanillaCod As String, Optional psRHConceptoGrupo As String = "", Optional pbSoloActivos As Boolean = True, Optional pnTiposConceptos As Integer = -1) As Recordset
    On Error GoTo GetConceptosErr
    Dim sqlD As String
    Dim sqlDAux As String
    Dim sqlDWhere As String
    Dim sqlDOrder As String
    Dim oCon As DConecta
    Dim sqlWhereFijos As String
    Set oCon = New DConecta
    
    If pbSoloActivos Then sqlDWhere = " And cRHConceptoEstado = 1 "
    
    If pnTiposConceptos = -1 Then
        sqlWhereFijos = ""
    ElseIf pnTiposConceptos = 1 Then
        sqlWhereFijos = " And  RHCT.cRHConceptoTpo = '" & RHConceptoTipoCal.RHConceptoTipo_VALOR_PRE_DEFINIDO & "'"
    Else
        sqlWhereFijos = " And  RHCT.cRHConceptoTpo <> '" & RHConceptoTipoCal.RHConceptoTipo_VALOR_PRE_DEFINIDO & "'"
    End If
    
    sqlD = " Select RHPC.cRHConceptoCod Codigo, cRHConceptoDescripcion + ' - ' + STR(RHPC.cRHConceptoCod) Descrip From RHConceptoTabla RHCT" _
         & " Inner Join RHPlanillaConceptoTabla RHPC On RHCT.cRHConceptoCod = RHPC.cRHConceptoCod" _
         & " Where left(RHPC.cRHConceptoCod,1) in ('" & psRHConceptoGrupo & "') And RHPC.cRHPlanillaCod = '" & psRHPlanillaCod & "'" & sqlDWhere & sqlWhereFijos
    
    If oCon.AbreConexion Then
        Set GetConceptosPlanilla = oCon.CargaRecordSet(sqlD)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

'Devuelve los conceptos Activos,
'
'El parametro psTpoConcepto si es vacio dretorna todos los tipos de conceptos
'caso contrario retorna todos los conceptos del tipo indicado
'##ModelId=3A9AFE7A0380
Public Function GetConceptosPlanillaRRHH(psRHPlanillaCod As String, psPersCod As String, Optional psRHConceptoGrupo As String = "", Optional pbSoloActivos As Boolean = True) As Recordset
    On Error GoTo GetConceptosErr
    Dim sqlD As String
    Dim sqlDAux As String
    Dim sqlDWhere As String
    Dim sqlDOrder As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbSoloActivos Then sqlDWhere = " And cRHConceptoEstado = 1 "
          
    If psRHConceptoGrupo = "" Then
       sqlD = " Select RHC.cRHConceptoCod Codigo, nRHConceptoValor Monto From RHConceptoTabla RHCT" _
            & " Inner Join RHConcepto RHC On RHCT.cRHConceptoCod = RHC.cRHConceptoCod" _
            & " Where RHC.cRHPlanillaCod = '" & psRHPlanillaCod & "' And cPersCod = '" & psPersCod & "'" & sqlDWhere
    Else
       sqlD = " Select RHC.cRHConceptoCod Codigo, nRHConceptoValor Monto From RHConceptoTabla RHCT" _
            & " Inner Join RHConcepto RHC On RHCT.cRHConceptoCod = RHC.cRHConceptoCod" _
            & " Where Left(RHC.cRHConceptoCod,1) in ('" & psRHConceptoGrupo & "') And RHPC.cRHPlanillaCod = '" & psRHPlanillaCod & "' And cPersCod = '" & psPersCod & "' " & sqlDWhere
    End If
    
    If oCon.AbreConexion Then
        Set GetConceptosPlanillaRRHH = oCon.CargaRecordSet(sqlD)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function



Public Function GetTablasBase() As Recordset
    On Error GoTo GetConceptosErr
    Dim sqlD As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlD = "select name from sysobjects where xtype in ('V','U') and uid = '1' order by name"

    If oCon.AbreConexion Then
        Set GetTablasBase = oCon.CargaRecordSet(sqlD)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

Public Function GetCamposTabla(psTabla As String) As Recordset
    On Error GoTo GetConceptosErr
    Dim sqlD As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlD = " sp_columns '" & psTabla & "'"

    If oCon.AbreConexion Then
        Set GetCamposTabla = oCon.CargaRecordSet(sqlD)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

'Devulve los conceptos Activos,
'El parametro psTpoConcepto si es vacio dretorna todos los tipos de conceptos
'caso contrario retorna todos los conceptos del tipo indicado
'##ModelId=3A9AFE7A0380
'"" Todos
'"1"  Tablas
'"2"  Estructura Tablas
'"Tabla"  Campos
Public Function GetTablasAlias(Optional psRHGrupo As String = "") As Recordset
    On Error GoTo GetConceptosErr
    Dim sqlD As String
    Dim sqlDAux As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If psRHGrupo = "" Then
        sqlD = "Select Alias = cAlias, SubString(cTabAliasCod,1,5) Codigo from RHTablasAlias Order by cAlias"
    ElseIf psRHGrupo = "1" Then
        sqlD = "Select Alias = cAlias, SubString(cTabAliasCod,1,3) Codigo from RHTablasAlias where cTabAliasCod like '___00' Order by cAlias"
    ElseIf psRHGrupo = "2" Then
        sqlD = "Select cTabAliasCod Codigo, cTabAliasTablaColumna Nombre, cAlias Alias from RHTablasAlias where cTabAliasCod like '___00' Order by cAlias"
    ElseIf psRHGrupo = "3" Then
        sqlD = "Select cTabAliasCod Codigo, cTabAliasTablaColumna Nombre, cAlias Alias from RHTablasAlias where cTabAliasCod like '" & psRHGrupo & "%' Order by cAlias"
    Else
        sqlD = "Select cTabAliasCod Codigo, cTabAliasTablaColumna Nombre, cAlias Alias from RHTablasAlias where cTabAliasCod like '" & psRHGrupo & "__' AND RIGHT(cTabAliasCod,2) <> '00' Order by cAlias"
    End If
    
    If oCon.AbreConexion Then
        Set GetTablasAlias = oCon.CargaRecordSet(sqlD)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function


Public Function ExisteTablaAlias(psTabla As String) As Boolean
    On Error GoTo GetConceptosErr
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rsT As ADODB.Recordset
    Set rsT = New ADODB.Recordset
    Dim sqlD As String
        
   sqlD = "Select cTabAliasCod from RHTablasAlias where cTabAliasTablaColumna = '" & psTabla & "'"
    
    If oCon.AbreConexion Then
        Set rsT = oCon.CargaRecordSet(sqlD)
        If rsT.EOF And rsT.BOF Then
            ExisteTablaAlias = False
        Else
            ExisteTablaAlias = True
        End If
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

Public Function GetCodigoTablaAlias(Optional psCodTab As String = "") As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rsT As ADODB.Recordset
    Set rsT = New ADODB.Recordset
    Dim sqlD As String
    
    If oCon.AbreConexion Then
        If psCodTab = "" Then
            sqlD = "Select cTabAliasCod From RHTablasAlias Where cTabAliasCod like '___00' Order By cTabAliasCod Desc"
            Set rsT = oCon.CargaRecordSet(sqlD)
            If RSVacio(rsT) Then
                GetCodigoTablaAlias = "00100"
            Else
                GetCodigoTablaAlias = FillNum(Trim(Str(CCur(Mid(rsT!cTabAliasCod, 1, 3)) + 1)), 3, "0") & "00"
            End If
        Else
            sqlD = "Select cTabAliasCod From RHTablasAlias Where SubString(cTabAliasCod,1,3) = '" & Left(psCodTab, 3) & "'  Order By cTabAliasCod Desc"
            
            Set rsT = oCon.CargaRecordSet(sqlD)
            If RSVacio(rsT) Then
                GetCodigoTablaAlias = "00100"
            Else
                sqlD = FillNum(Trim(Str(CCur(Mid(rsT!cTabAliasCod, 4, 2)) + 1)), 2, "0")
                GetCodigoTablaAlias = Left(psCodTab, 3) & sqlD
            End If
        End If
    End If
    
    rsT.Close
    Set rsT = Nothing
    Set oCon = Nothing
End Function

Public Function AgregaTablaAlias(psCodigo As String, psNombre As String, psAlias As String, psUltMov As String) As Boolean
    On Error GoTo GetConceptosErr
    Dim oCon As DConecta
    Dim sqlD As String
    Set oCon = New DConecta
        
    sqlD = " Insert RHTablasAlias (cTabAliasCod,cTabAliasTablaColumna,cAlias,cUltimaActualizacion)" _
         & " values('" & psCodigo & "','" & psNombre & "','" & psAlias & "','" & psUltMov & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlD
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

Public Function EliminaTablaAlias(psCodigo As String) As Boolean
    On Error GoTo GetConceptosErr
    Dim oCon As DConecta
    Dim sqlD As String
    Set oCon = New DConecta
        
    sqlD = " Delete RHTablasAlias Where cTabAliasCod like '" & psCodigo & "%'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlD
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function

'Devuelve la formula usada
'##ModelId=3A9AFE980347
Public Function GetFormula(psRHConceptoGrupo As String) As String
    On Error GoTo GetFormulaErr
    Dim sqlD As String
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlD = "Select cRHConceptoFormula Formula From rhconceptotabla Where cRHConceptoEstado = 1 And cRHConceptoCod  = '" & psRHConceptoGrupo & "'"
    
    If oCon.AbreConexion Then
        Set rsD = oCon.CargaRecordSet(sqlD)
        GetFormula = IIf(IsNull(rsD!Formula), "", rsD!Formula)
        oCon.CierraConexion
    End If
    
    rsD.Close
    Set oCon = Nothing
    Exit Function
GetFormulaErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function

'Devuelve ultimo codigo + 1 del tipo del Grupo
'##ModelId=3A9AFE980347
'Public Function GetCodigo(psRHConceptoCodTpo As String) As String
'    On Error GoTo GetFormulaErr
'    Dim sqlD As String
'    Dim rsD As ADODB.Recordset
'    Set rsD = New ADODB.Recordset
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'
'    sqlD = "Select Top 1 cRHConceptoCod Codigo From RHConceptoTabla Where cRHConceptoCod like '" & psRHConceptoCodTpo & "%'  Order by cRHConceptoCod Desc"
'
'    If oCon.AbreConexion Then
'        Set rsD = oCon.CargaRecordSet(sqlD)
'        If rsD.EOF And rsD.BOF Then
'            GetCodigo = psRHConceptoCodTpo & "01"
'        Else
'            GetCodigo = Trim(Str(CInt(rsD!Codigo) + 1))
'        End If
'        oCon.CierraConexion
'    End If
'
'    rsD.Close
'    Set oCon = Nothing
'    Exit Function
'GetFormulaErr:
'    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
'End Function

Public Function GetCodigo(psRHConceptoCodTpo As String) As String
    On Error GoTo GetFormulaErr
    Dim sqlD As String
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlD = "Select Top 1 cRHConceptoCod Codigo From RHConceptoTabla Where cRHConceptoCod like '" & psRHConceptoCodTpo & "%'  Order by cRHConceptoCod Desc"
    
    If oCon.AbreConexion Then
        Set rsD = oCon.CargaRecordSet(sqlD)
        If rsD.EOF And rsD.BOF Then
            GetCodigo = psRHConceptoCodTpo & "01"
        Else
            GetCodigo = Trim(Str(CInt(rsD!codigo) + 1))
        End If
        oCon.CierraConexion
    End If
    
    rsD.Close
    Set oCon = Nothing
    Exit Function
GetFormulaErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function

Public Function GetImpreConcepto(psRHConcepto As String) As String
    On Error GoTo GetFormulaErr
    Dim sqlD As String
    Dim rsD As ADODB.Recordset
    Set rsD = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlD = "Select cRHConceptoAbreviatura Codigo From RHConceptoTabla Where cRHconceptoMeno like '" & psRHConcepto & "'"
    
    If oCon.AbreConexion Then
        Set rsD = oCon.CargaRecordSet(sqlD)
        If rsD.EOF And rsD.BOF Then
            GetImpreConcepto = ""
        Else
            GetImpreConcepto = Trim(rsD!codigo)
        End If
        oCon.CierraConexion
    End If
    
    rsD.Close
    Set oCon = Nothing
    Exit Function
GetFormulaErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetFormula Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902FC00F1
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function AgregaRHCalculoEPS(psCodCalculoEps As String, psDescripcion As String, psUltimaActualziacion As String, pnEstadoEPS As Integer) As String
    On Error GoTo AgregaRHCalculoEPSErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    sqlC = " INSERT INTO RHCalculoEPS " _
          & "  (cCodCalculoEps,cDescripcion,cUltimaActualizacion,nEstadoEPS) " _
          & "  Values ('" & psCodCalculoEps & "','" & psDescripcion & "','" & psUltimaActualziacion & "'," & pnEstadoEPS & ") "
    If oCon.AbreConexion Then
       oCon.Ejecutar sqlC
       oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
AgregaRHCalculoEPSErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConcepto:AgregaRHCalculoEPS Method")
End Function

Public Function AgregaRHCalculoEPSDetalle(psCodCalculoEps As String, psCodPersona As String, pnSueldo As Currency, pnSueldo_x_225 As Currency, pnCantPersonas As Integer, pnPlanSinIGV As Currency, pnPromedio As Currency, _
        pnNeto As Currency, pnPagaEmpleado As Currency, pnPagaEmpresa As Currency, pnAdicionalHijos As Currency, pnAdicionalPadres As Currency, pnTotalEmpleado As Currency, _
        pnDesQuincena As Currency, pnDesMensual As Currency, pnDesQuincenaUno As Currency, pnDescQuincenaDos As Currency, pnSaldo As Currency) As String
    On Error GoTo AgregaRHCalculoEPSDetalleErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    sqlC = " INSERT INTO RHCalculoEPSDET " _
         & " (cCodCalculoEps,cPerscod,nSueldo,nSueldo_x_225,nCantPersonas,nPlanSinIGV,nPromedio, " _
         & " nNeto,nPagaEmpleado,nPagaEmpresa, nAdicionalHijos, nAdicionalPadres, nTotalEmpleado, " _
         & " nDesQuincena,nDesMensual,nDesQuincenaUno,nDescQuincenaDos,nSaldo ) " _
         & " Values ('" & psCodCalculoEps & "','" & psCodPersona & "'," & pnSueldo & "," & pnSueldo_x_225 & " ," & pnCantPersonas & " ," & pnPlanSinIGV & " , " & pnPromedio & " , " _
         & "  " & pnNeto & "," & pnPagaEmpleado & " ," & pnPagaEmpresa & " ," & pnAdicionalHijos & "," & pnAdicionalPadres & "," & pnTotalEmpleado & " , " _
         & "  " & pnDesQuincena & "," & pnDesMensual & " ," & pnDesQuincenaUno & "," & pnDescQuincenaDos & "," & pnSaldo & ")"
    If oCon.AbreConexion Then
       oCon.Ejecutar sqlC
       oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
AgregaRHCalculoEPSDetalleErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConcepto:AgregaRHCalculoEPSDetalle Method")
End Function

Public Function ActualizaRHCalculoEPSDet(psCodCalculoEps) As String
    On Error GoTo ActualizaRHCalculoEPSDetErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    If oCon.AbreConexion Then
       'Actualizar La Agencia
       sqlC = "update rhcalculoepsdet Set CodAgencia = cRHAgenciaCodOficial From  rhcargos rh1,rhcalculoepsdet rhc where rh1.cperscod =rhc.cPerscod and " _
              & " dRHCargoFecha = (Select  max(dRHCargoFecha) from rhcargos rh2 where rh2.cPerscod = rhc.cPerscod ) " _
              & "  and rhc.cCodCalculoEps ='" & psCodCalculoEps & "' "
       oCon.Ejecutar sqlC
       'Plan Base
       sqlC = " update rhcalculoepsdet Set nCodCatBase = nCodCatalogo from  rhpersonaplan rhp ,rhcalculoepsdet rce " _
            & "  where substring(nCodCatalogo,1,2) in ('01','02','03') and rce.cCodCalculoEps ='" & psCodCalculoEps & "' and  rhp.cPersCod =  rce.cPerscod "
       oCon.Ejecutar sqlC
       'Plan Adicional Hijos
       sqlC = " Update rhcalculoepsdet Set nCodCatAdiclHijos = nCodCatalogo from  rhpersonaplan rhp ,rhcalculoepsdet rce where substring(nCodCatalogo,1,2) in ('04','05','06') and " _
                & " rce.cCodCalculoEps ='" & psCodCalculoEps & "' and  rhp.cPersCod =  rce.cPerscod "
       oCon.Ejecutar sqlC
       'Plan  Adicional  Padres
        sqlC = " Update rhcalculoepsdet Set nCodCatAdicPadres = nCodCatalogo from  rhpersonaplan rhp ,rhcalculoepsdet rce  where substring(nCodCatalogo,1,2) in ('07','08') and " _
               & " rce.cCodCalculoEps ='" & psCodCalculoEps & "' and  rhp.cPersCod =  rce.cPerscod "
       oCon.Ejecutar sqlC
        'Actualizar Costo Neto Plan Base
        sqlC = " Update RHCalculoEpsDet Set nCNetoBase = nCostoNeto FROM RHCatalogoPlanEPS RCP, rhcalculoepsdet RCE Where RCE.nCodCatBase = RCP.nCodCatalogo " _
             & " AND  RCE.cCodCalculoEps ='" & psCodCalculoEps & "'  "
        oCon.Ejecutar sqlC
        'Actualizar Costo Neto Adicional Hijos y Cant Hijos
        sqlC = " Update RHCalculoEpsDet set  nCNetoAdicHijos  =nCostoNeto, nCantPerHijos = nCantPer FROM RHCatalogoPlanEPS RCP, rhcalculoepsdet RCE " _
               & "  Where RCE.nCodCatAdiclHijos = RCP.nCodCatalogo AND  RCE.cCodCalculoEps ='" & psCodCalculoEps & "' "
        oCon.Ejecutar sqlC
       
        'Actualizar Costo Neto Adicional Padre Y Cant Padres
        sqlC = " Update RHCalculoEpsDet  set  nCNetoAdicPadres  = nCostoNeto , nCantPerPadres = nCantPer FROM RHCatalogoPlanEPS RCP, rhcalculoepsdet RCE " _
               & " Where RCE.nCodCatAdicPadres = RCP.nCodCatalogo AND  RCE.cCodCalculoEps ='" & psCodCalculoEps & "' "
        oCon.Ejecutar sqlC
       
       oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ActualizaRHCalculoEPSDetErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosConcepto:AgregaRHCalculoEPSDetalle Method")
End Function



Public Function GetDatosPersonaPlanilla(ByVal psRHCodEmp As String, ByVal psFechaFin As String) As ADODB.Recordset
    On Error GoTo GetConceptosErr
    Dim sqlD As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    'ALPA 20110118
    ' Modificado por BRGO 20110707. Se agregó la columna CUSPP
    ' Modificado por EJVG 20110923. Se cambio rh por rhe en left RC
    sqlD = " Select p.cPErsNombre,isnull(ar.cAreaDescripcion,'')Area,isnull(rhct.cRHCargoDescripcion,'')Cargo ,isnull(ag.cAgeDescripcion,'')Agencia,isnull(rh.dIngreso,'')FechaIng,isnull(pid.cPErsIDNRo,'')DNI," & _
            " isnull(RHAF.cRHAFPAbreviatura,'')AFP ,isnull(p.dPersNacCreac,'')FechaNaci,isnull(p.cPersDireccDomicilio,'')Domicilio, " & _
            " isnull(p.cPersTelefono,'')Telefono,rh.cUbicacion cUbicacion,RC.dRHContratoFecha,PN.cPersNatSexo,isnull(ck.cConsDescripcion,'') cConsDescripcionNivel,dRHContratoFin, ISNULL(rhe.cCUSPP,'') cCUSPP " & _
            " from RRHH rh " & _
            " INNER JOIN Persona p ON p.cPersCod=rh.cPersCod " & _
            " INNER JOIN PersonaNat PN On P.cPersCod=PN.cPersCod " & _
            " LEFT JOIN PersID pid ON pid.cPERsCod =p.cPErsCod and cPErsIDTpo=1 " & _
            " LEft JOIN RHCArgos rhc ON rhc.cPersCod=rh.cPersCod and rhc.dRHCargoFecha <= '" & psFechaFin & "' " & _
            " left JOIN RHCargosTabla rhct ON rhct.cRHCargoCod = rhc.cRHCargoCodOficial " & _
            " Left JOIN RHCargosNiveles rctn ON rhct.cRHCargoCod=rctn.cRHCargoCod " & _
            " left join constante CK ON CK.nConsValor=rctn.nRHCargoNivel and CK.nConsCod=6052" & _
            " LEFT JOIN RHEmpleado rhe ON rhe.cPErsCod=RH.cPersCod " & _
            " left JOIN RHAFP RHAf ON RHAF.cRHAFPPerscod = rhe.cRHEmplAFPPerscod " & _
            " INNER JOIN Areas ar ON ar.cAreaCod = rhc.cRHAreaCod   " & _
            " INNER JOIN Agencias ag ON ag.cAgeCod = rhc.cRHAgenciaCod " & _
            " LEFT JOIN (select  cPersCod,MAX(cRHContratoNro) cRhContratoNro, " & _
            "         max(dRHContratoFecha) dRHContratoFecha,max(dRHContratoFin) dRHContratoFin " & _
            " From RHContratoDet " & _
            " group by cPersCod) RC on rh.cPersCod=RC.cPersCod " & _
            " where cRHCod = '" & psRHCodEmp & "'" & _
            " ORDER BY rhc.dRHCargoFecha desc "

    If oCon.AbreConexion Then
        Set GetDatosPersonaPlanilla = oCon.CargaRecordSet(sqlD)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetConceptosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosConcepto:GetConceptos Method")
End Function


