VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizadatosCargo"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A942F310341"
'Actualiza los Cargos y niveles en la base de datos
Option Base 0
Option Explicit

Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3AB902EC03B5
    Private mlClassDebugID As Long
#End If

'##ModelId=3AB902ED010D
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902ED00D1
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

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'Agrega un nuevo cargo o nivel
'##ModelId=3A942F6E024E
Public Function AgregaCargo(psRHCargoCod As String, psRHCargoDescripcion As String, pnRHCargoSueldoMinimo As Double, pnRHCargoSueldoMaximo As Double, psRHCargoCategoria As String, psUltimaActualizacion As String, pnRHCargosOrden As Integer, pnGrado As Integer, pnCtrAsist, pnRHGrupoPlanea As Integer, pnRHGastoDirInd As Integer) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    On Error GoTo AgregaCargoErr

    sqlC = " Insert RHCargosTabla (cRHCargoCod,cRHCargoDescripcion,nRHCargoSueldoMinimo,nRHCargoSueldoMaximo,cRHCargoCategoria,nRHCargoOrden,cUltimaActualizacion,nRHGradoAprobacion,bRHCtrAsist,nRHGrupoPlanea,nRHGastoDirInd)" _
         & " Values('" & psRHCargoCod & "','" & psRHCargoDescripcion & "'," & pnRHCargoSueldoMinimo & "," & pnRHCargoSueldoMaximo & ",'" & psRHCargoCategoria & "'," & pnRHCargosOrden & ",'" & psUltimaActualizacion & "'," & pnGrado & "," & pnCtrAsist & "," & pnRHGrupoPlanea & "," & pnRHGastoDirInd & ")"
    
    If oCon.AbreConexion() Then
        oCon.CargaRecordSet sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosCargo:AgregaCargo Method")
End Function

'Elimina  un nuevo cargo o nivel
'##ModelId=3A942F6E0280
Public Function EliminaCargo(psRHCargoCod As String) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo EliminaCargoErr

    sqlC = "Delete RHCargosTabla where cRHCargoCod like '" & Trim(psRHCargoCod) & "%'"
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
EliminaCargoErr:
    EliminaCargo = False
    Call RaiseError(MyUnhandledError, "DActualizadatosCargo:EliminaCargo Method")
End Function

'Elimina  un nuevo cargo o nivel
'##ModelId=3A942F6E0280
Public Function VerificaUsoCargo(psRHCargoCod As String) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo EliminaCargoErr

    sqlC = "Select cRHCargoCod From AreaCargo where cRHCargoCod like '" & Trim(psRHCargoCod) & "%'"
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        
        If rs.EOF And rs.BOF Then
            VerificaUsoCargo = False
        Else
            VerificaUsoCargo = True
        End If
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
EliminaCargoErr:
    VerificaUsoCargo = False
    Call RaiseError(MyUnhandledError, "DActualizadatosCargo:EliminaCargo Method")
End Function

'Actualiza un nuevo cargo o nivel
'##ModelId=3A942F6E02B2
Public Function ModificaCargo(psRHCargoCod As String, psRHCargoDescripcion As String, pnRHCargoSueldoMinimo As Integer, pnRHCargoSueldoMaximo As Integer, psRHCargoCategoria As String, psUltimaActualizacion As String, pnRHCargosOrden As Integer, pnGrado As Integer, pnCtrAsist As Integer, pnRHGrupoPlanea As Integer, pnRHGastoDirInd As Integer) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo ModificaCargoErr
    
    sqlC = " Update RHCargosTabla" _
         & " Set cRHCargoDescripcion = '" & psRHCargoDescripcion & "',nRHCargoSueldoMinimo = " & pnRHCargoSueldoMinimo & ",nRHCargoSueldoMaximo = " & pnRHCargoSueldoMaximo & ",cRHCargoCategoria = '" & psRHCargoCategoria & "',nRHCargoOrden = " & pnRHCargosOrden & ",cUltimaActualizacion = '" & psUltimaActualizacion & "', nRHGradoAprobacion =  " & pnGrado & ", bRHCtrAsist = " & pnCtrAsist & ", nRHGrupoPlanea = " & pnRHGrupoPlanea & ", nRHGastoDirInd = " & pnRHGastoDirInd & " " _
         & " Where cRHCargoCod = '" & psRHCargoCod & "'"
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        sqlC = " Update RHCargosTabla" _
             & " Set cRHCargoCategoria = '" & psRHCargoCategoria & "'" _
             & " Where cRHCargoCod like '" & Trim(psRHCargoCod) & "___'"
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaCargoErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosCargo:ModificaCargo Method")
End Function

'##ModelId=3A942F730292
Public Function GetCargos(pbNiveles As Boolean, Optional psNivel As String = "", Optional pbTodos As Boolean = False) As Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    On Error GoTo GetCargosErr
    
    If pbNiveles Then
        sqlC = " Select '',cRHCargoCod as Codigo, cRHCargoDescripcion as Descripción, Case nRHCargoSueldoMinimo when 0 then '0.00' else convert(varchar(15),convert(decimal(8,2),nRHCargoSueldoMinimo)) end as SueNiv, Case nRHCargoSueldoMaximo when 0 then '0.00' else convert(varchar(15),convert(decimal(8,2),nRHCargoSueldoMaximo)) end as SueMax, nRHCargoOrden as Orden, CT.cConsDescripcion + space(50) + CA.cRHCargoCategoria as Categoria" _
             & " From " & lsServerAdministracion & "RHCargosTabla CA" _
             & " Inner Join " & lsServerAdministracion & "Constante CT On CA.cRHCargoCategoria = CT.nConsValor And nConsCod like '" & ConstanteCabecera.gViaticosCateg & "'" _
             & " where cRHCargoCod like '___' order by  cRHCargoCod"
    Else
        If Not pbTodos Then
            sqlC = " Select '',cRHCargoCod as Codigo, cRHCargoDescripcion as Descripción, Case nRHCargoSueldoMinimo when 0 then '0.00' else convert(varchar(15),convert(decimal(8,2),nRHCargoSueldoMinimo)) end as SueNiv, nRHGradoAprobacion Grado, Case When bRHCtrAsist = 1 Then 1 Else 0 End CtrAsis, nRHGrupoPlanea, nRHGastoDirInd " _
                 & " From " & lsServerAdministracion & "RHCargosTabla CA" _
                 & " where cRHCargoCod like '" & psNivel & "___%' and cRHCargoCod not like '___' order by cRHCargoCod"
        Else
            sqlC = " Select '',cRHCargoCod as Codigo, cRHCargoDescripcion as Descripción, Case nRHCargoSueldoMinimo when 0 then '0.00' else convert(varchar(15),convert(decimal(8,2),nRHCargoSueldoMinimo)) end as SueNiv, Case nRHCargoSueldoMaximo when 0 then '0.00' else convert(varchar(15),convert(decimal(8,2),nRHCargoSueldoMaximo)) end as SueMax, nRHCargoOrden as Orden, CT.cConsDescripcion + space(50) + CA.cRHCargoCategoria as Categoria" _
                 & " From " & lsServerAdministracion & "RHCargosTabla CA" _
                 & " Left Join " & lsServerAdministracion & "Constante CT On CA.cRHCargoCategoria = CT.nConsValor And nConsCod like '" & ConstanteCabecera.gViaticosCateg & "'" _
                 & " order by cRHCargoCod"
        End If
    End If
    
    If oCon.AbreConexion() Then
        Set GetCargos = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
GetCargosErr:
    Call RaiseError(MyUnhandledError, "DActualizadatosCargo:GetCargos Method")
End Function

Public Function GetAreasCargo(psCargo As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    
    sql = " Select distinct AREA.cAreaCod, AREA.cAreaDescripcion from AreaCargo AC" _
        & " Inner Join Areas AREA On AC.cAreaCod = AREA.cAreaCod" _
        & " Where AC.cRHCargoCod = '" & psCargo & "'"
    
    oCon.AbreConexion
    
    Set GetAreasCargo = oCon.CargaRecordSet(sql)
End Function

Public Function SetAreasCargo(prRS As ADODB.Recordset, psRHCargoCod As String) As Boolean
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    
    oCon.AbreConexion
    
    oCon.BeginTrans
    sql = "Delete AreaCargo Where cRHCargoCod = '" & psRHCargoCod & "'"
    oCon.Ejecutar sql
    
    If prRS Is Nothing Then
        SetAreasCargo = True
        Exit Function
    End If
    If prRS.EOF And prRS.BOF Then
        SetAreasCargo = True
        Exit Function
    End If
    
    prRS.MoveFirst
    
    While Not prRS.EOF
        If prRS.Fields(0) <> "" Then
            sql = " Insert AreaCargo (cAreaCod, cRHCargoCod, nAutPermiso)" _
                & " Values('" & prRS.Fields(0) & "','" & psRHCargoCod & "',0)"
            oCon.Ejecutar sql
        End If
        prRS.MoveNext
    Wend
    oCon.CommitTrans
    SetAreasCargo = True
    Exit Function
End Function

#If DebugMode Then
    '##ModelId=3AB902ED0027
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If
