VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMExecSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'------------------------------------------------------------------------------
'20060324
'modulo de clase para conectar y ejecutar un procedimiento
'almacenado y que retorna el conjunto de datos
'recordset
'los parametros de salida del procedimiento almacenado
'y el valor de retorno del procedimiento almacenado
'------------------------------------------------------------------------------
Option Explicit

'declaracion de variables
Private mConexion As ADODB.Connection
Private mComando As ADODB.Command
Private mRecordSet As ADODB.RecordSet

'arrego de parametros de salida
'Private mOutPuts() As Variant

'valor de retorno
Private mReturnValue As Integer

'declaramos la coleccion de valores de salida
Private mclsOutPut As DCOMOutPut
Private mCOMConecta As COMConecta.DCOMConecta

'20060406
'nuevo código que ejecuta procedmientos almacenados.
Public Property Get ParamOutPut() As DCOMOutPut
    If mclsOutPut Is Nothing Then
        Set mclsOutPut = New DCOMOutPut
    End If

    Set ParamOutPut = mclsOutPut
    
End Property

'Public Property Set DCOMOutPut(vData As DCOMOutPut)
'    Set mvarclsOutPut = vData
'End Property
Private Sub Class_Initialize()
    Set mclsOutPut = New DCOMOutPut
    Set mCOMConecta = New COMConecta.DCOMConecta
End Sub

Private Sub Class_Terminate()
    Set mclsOutPut = Nothing
    Set mCOMConecta = Nothing
End Sub

'declaracion de propiedades
'esta propiedad obtiene el conjunto de datos
'resultado de la ejecución del procedimiento almacenado
Public Property Get RecordSet() As ADODB.RecordSet
    
    If mRecordSet Is Nothing Then
        
        Set mRecordSet = New RecordSet
        
    End If
    
    Set RecordSet = mRecordSet
        
End Property

'propiedad que obtiene el valor de retorno del
'procedimiento almacenado
Public Property Get ReturnValue() As Integer
    
    If mReturnValue = Null Then
        mReturnValue = 0
    End If
    
    ReturnValue = mReturnValue
    
End Property

'3. Mètodos de la clase
Public Function ExecuteStoreProcedure(StoreProcedure As String, _
ParamArray oParametros() As Variant) As ADODB.RecordSet

    Dim oComando As New ADODB.Command
    'Dim oConexion As New ADODB.Connection
    Dim oRecordSet As New ADODB.RecordSet
    Dim nContador, nContador2 As Integer
    Dim oParam As New ADODB.Parameter
    'Dim OutPuts As New DCOMOutPut
    '1. generar u obtener la cadena de conexion
    '....
    
    'nuevas variables
    Dim bConexionAbierta As Boolean
    Dim bTrans As Boolean
    
    On Error GoTo ErrorExecuteProcedure
    
    '2. asignamos la cadena de conexion
    'oConexion.CursorLocation = adUseClient
    'oConexion.ConnectionString = "PROVIDER=SQLOLEDB;driver={SQL Server};server=HYO-SRV-DESA01;uid=sa;pwd=confianza;database=dbcmact;"
    
    '3. cambiamos el codigo de asignacion de cadena de conexion
    'al modelo de ejecucion de la caja
    bConexionAbierta = mCOMConecta.AbreConexion()
    
    If Not bConexionAbierta Then
        'no se pudo establecer la conexion con el servidor
        Exit Function
    Else
        bTrans = True
    End If
    
    '20060406
    'asignamos la conexion activa al comando
    
    'ocomando
    oComando.CommandText = StoreProcedure
    oComando.CommandType = adCmdStoredProc
    
    'oConexion.Open
    oComando.ActiveConnection = mCOMConecta.ConexionActiva
    '3. capturamos los parametros con el metodo refresh
    oComando.Parameters.Refresh
    
    nContador2 = 1
    
    '4. asignamos los valores de los parámetros
    For nContador = 1 To oComando.Parameters.Count - 1
    
        'asigna valores
        If (oComando.Parameters(nContador).Direction = adParamInput) Then
            oComando.Parameters(nContador).Value = oParametros(nContador2 - 1)
            nContador2 = nContador2 + 1
        End If
        
        
    Next
    
    '5.una vez asignado los valores ejecutamos la consulta y asinamos el número
    'de filas afectadas al mReturnValue
    'Set oRecordSet = oComando.Execute(mReturnValue)
    
    oComando.ActiveConnection.BeginTrans
    oRecordSet.Open oComando, , adOpenStatic, adLockReadOnly
    oComando.ActiveConnection.CommitTrans
    
    bTrans = False
    '6. ahora verificamos si el procedimiento almacenado tiene valores
    'de retorno
    For Each oParam In oComando.Parameters
        'verificamos la direccion de los parámetros
        If (oParam.Direction = adParamInputOutput) Or (oParam.Direction = adParamInputOutput) Then
            mclsOutPut.Add (oParam.Value)
        End If
        
        'verificamos si es valor de retorno
        If (oParam.Direction = adParamReturnValue) Then
            'asignamos el retrun value
            mReturnValue = oParam.Value
        End If
        
    Next
    
    'ahora verificamos si hay un conjunto de datos devuelto
    'en la consulta: para eso empleamos el mReturnValue
    If mReturnValue >= 0 Then
        'retornamos el recordset
        If oRecordSet.State = adStateOpen Then
            If oRecordSet.RecordCount > 0 Then
                'Debug.Print oRecordSet.RecordCount
                Set ExecuteStoreProcedure = oRecordSet
            Else
                'si no retornamos null
                Set ExecuteStoreProcedure = Nothing
            End If
        Else
            'si no retornamos null
            Set ExecuteStoreProcedure = Nothing
        End If
    Else
    
        'si no retornamos null
        Set ExecuteStoreProcedure = Nothing
        
    End If
    
    '20060405
    'cerramos la conexion
    'mCOMConecta.CierraConexion
Exit Function
ErrorExecuteProcedure:
    If bTrans Then
        oComando.ActiveConnection.RollbackTrans
    End If
    MsgBox Err.Source + Chr(10) + Chr(13) _
        + Err.Description, vbInformation, "Error en ExecuteProcedure de DCOMExecSP"
End Function

Public Function ExecuteSqlQuery(ByVal psSqlQuery As String)
    On Error GoTo ErrorExecuteQuery
    'nuevas variables
    Dim oConect As ADODB.Connection
    Dim bConexionAbierta As Boolean
    Dim bTrans As Boolean
    
    '3. cambiamos el codigo de asignacion de cadena de conexion
    'al modelo de ejecucion de la caja
    bConexionAbierta = mCOMConecta.AbreConexion()
    
    If Not bConexionAbierta Then
        'no se pudo establecer la conexion con el servidor
        Exit Function
    Else
        bTrans = True
    End If
    
    '20060406
    'asignamos la conexion activa
    
    Set oConect = mCOMConecta.ConexionActiva
    oConect.BeginTrans
    oConect.Execute psSqlQuery
    oConect.CommitTrans
    Set oConect = Nothing
Exit Function
ErrorExecuteQuery:
    If bTrans Then
        oConect.RollbackTrans
    End If
    MsgBox Err.Source + Chr(10) + Chr(13) _
        + Err.Description, vbInformation, "Error en ExecuteSQLQuery de DCOMExecSP"
End Function

