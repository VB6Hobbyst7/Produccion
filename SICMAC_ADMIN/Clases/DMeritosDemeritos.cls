VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMeritosDemeritos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
Option Base 0
Option Explicit


'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    '##ModelId=3B02D59C0115
    Private mlClassDebugID As Long
#End If

'Agrega un nuevo Item de MerDem Tabla
'##ModelId=3B02D4EA01D7
Public Function AgregaMerDemTabla(psRHMerDemTablaCod As String, psRHpcRHMerDemTablaCodDescripcion As String, psMeritoTpo As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaMerDemTablaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Insert RHMerDemTabla (cRHMerDemCod,cRHMerDemDescripcion,cUltimaActualizacion,nRHMErDemTpo)" _
         & " Values('" & psRHMerDemTablaCod & "','" & psRHpcRHMerDemTablaCodDescripcion & "','" & psUltimaActualizacion & "'," & psMeritoTpo & ")"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMerDem:AgregaMerDemTabla Method")
End Function

'Elimina  un Item de MerDem Tabla
'##ModelId=3B02D4EA0213
Public Function EliminaMerDemTabla(psRHMerDemTablaCod As String) As Boolean
    On Error GoTo EliminaMerDemTablaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Delete RHMerDemTabla where cRHMerDemCod = '" & psRHMerDemTablaCod & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMerDem:EliminaMerDemTabla Method")
End Function

'Elimina  un Item de MerDem Tabla
'##ModelId=3B02D4EA0213
Public Function TipoUsado(psRHMerDemTablaCod As String) As Boolean
    On Error GoTo EliminaMerDemTablaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlC = " Select cRHMerDemCod From RHMeritosDemeritos where cRHMerDemCod = '" & psRHMerDemTablaCod & "'"
    
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        
        If rs.EOF And rs.BOF Then
            TipoUsado = False
        Else
            TipoUsado = True
        End If
        
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMerDem:EliminaMerDemTabla Method")
End Function

'Actualiza un Item de MerDem Tabla
'##ModelId=3B02D4EA024F
Public Function ModificaMerDemTabla(psRHMerDemTablaCod As String, psRHpcRHMerDemTablaCodDescripcion As String, psMeritoTpo As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaMerDemTablaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Update RHMerDemTabla" _
         & " Set cRHMerDemDescripcion = '" & psRHpcRHMerDemTablaCodDescripcion & "', cUltimaActualizacion = '" & psUltimaActualizacion & "', nRHMErDemTpo = " & psMeritoTpo & " " _
         & " Where cRHMerDemCod = '" & psRHMerDemTablaCod & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMerDem:ModificaMerDemTabla Method")
End Function

'Agrega un nuevo Item de MerDem
'##ModelId=3B02D4EA028B
Public Function AgregaMerDem(psRHPersCod As String, psRHMerDemCod As String, pnRHItem As Integer, psRHCurLugar As String, psRHCurInicio As String, psRHCurFin As String, pnRHCurUndTiempo As Integer, pnRHCurTiempo As Integer, psComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaMerDemErr
    Dim oCon As DConecta
    Dim sqlC As String
    Dim lsRHCurInicio As String
    Dim lsRHCurFin As String
    Dim lsRHCurUndTiempo As String
    Dim lsRHCurTiempo As String
    
    Set oCon = New DConecta
    
    If psRHCurInicio <> "" Then
        lsRHCurInicio = "'" & psRHCurInicio & "'"
        lsRHCurFin = "'" & psRHCurFin & "'"
        lsRHCurUndTiempo = "NUll"
        lsRHCurTiempo = "NULL"
    Else
        lsRHCurInicio = "NULL"
        lsRHCurFin = "NULL"
        lsRHCurUndTiempo = "'" & pnRHCurUndTiempo & "'"
        lsRHCurTiempo = "" & pnRHCurTiempo & ""
    End If
    
    
    sqlC = " Insert RHMerDem (cRHPersCod,cRHMerDemCod,nRHMerDemItem,cRHMerDemLugar,dRHMerDemInicio,dRHMerDemFin,cRHMerDemUndTiempo,nRHMerDemTiempo,cRHMerDemComentario,cUltimaActualizacion)" _
         & " Values('" & psRHPersCod & "','" & psRHMerDemCod & "'," & pnRHItem & ",'" & psRHCurLugar & "'," & lsRHCurInicio & "," & lsRHCurFin & "," & lsRHCurUndTiempo & "," & lsRHCurTiempo & ",'" & psComentario & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaMerDemErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMerDem:AgregaMerDem Method")
End Function

'Elimina  un Item de MerDem
'##ModelId=3B02D4EA02D1
Public Function EliminaMerDem(psRHMerDemCod As String, psRHPersCod As String, pnRHItem As Integer) As Boolean
    On Error GoTo EliminaMerDemErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Delete RHMerDem" _
         & " Where cRHPersCod = '" & psRHPersCod & "' And cRHMerDemCod = '" & psRHMerDemCod & "' And nRHMerDemItem = " & pnRHItem
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaMerDemErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMerDem:EliminaMerDem Method")
End Function

'Actualiza un Item de MerDem
'##ModelId=3B02D4EA0317
Public Function ModificaMerDem(psPersCod As String, prRS As ADODB.Recordset, psUltimaActualizacion As String, psgsFormatoFecha As String) As Boolean
    On Error GoTo ModificaMerDemErr
    Dim oCon As DConecta
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Set oCon = New DConecta
    Dim lsCadena As String
    Dim sqlC As String
    
    While Not prRS.EOF
        If lsCadena = "" Then
            lsCadena = prRS.Fields(4) & Format(prRS.Fields(5), "yyyymmdd")
        Else
            lsCadena = lsCadena & prRS.Fields(4) & "','" & Format(prRS.Fields(5), "yyyymmdd")
        End If
        prRS.MoveNext
    Wend
    
    prRS.MoveFirst
    sqlC = " Delete rhmeritosdemeritos where cRHMerDemCod + convert(varchar(8),dRHMerDemFecha,112) not in ('" & lsCadena & "') And cRHPersCod = '" & psPersCod & "'"
    
    If oCon.AbreConexion Then
        oCon.BeginTrans
        oCon.Ejecutar sqlC
        
        While Not prRS.EOF
            lsCadena = prRS.Fields(4) & Format(prRS.Fields(5), "yyyymmdd")
            sqlC = " Select cRHMerDemCod From rhmeritosdemeritos where cRHMerDemCod + convert(varchar(8),dRHMerDemFecha,112)  in ('" & lsCadena & "') And cRHPersCod = '" & psPersCod & "'"
            Set rsC = oCon.CargaRecordSet(sqlC)
            
            If rsC.EOF And rsC.BOF Then
                sqlC = " Insert RHMeritosDemeritos (cRHPersCod,cRHMerDemCod,dRHMerDemFecha,cRHMerDemComentario,cUltimaActualizacion)" _
                     & " Values('" & psPersCod & "','" & prRS.Fields(0) & "','" & Format(CDate(prRS.Fields(2)), psgsFormatoFecha) & "','" & prRS.Fields(3) & "','" & psUltimaActualizacion & "')"
            Else
                sqlC = " Update RHMeritosDemeritos" _
                     & " Set cRHMerDemComentario = '" & prRS.Fields(3) & "',cUltimaActualizacion = '" & psUltimaActualizacion & "', dRHMerDemFecha = '" & Format(CDate(prRS.Fields(2)), psgsFormatoFecha) & "', cRHMerDemCod = '" & prRS.Fields(0) & "'" _
                     & " Where cRHPersCod = '" & psPersCod & "' And cRHMerDemCod = '" & prRS.Fields(4) & "' And dRHMerDemFecha = '" & Format(CDate(prRS.Fields(5)), psgsFormatoFecha) & "'"
            End If
            rsC.Close
            oCon.Ejecutar sqlC
            
            prRS.MoveNext
        Wend
        oCon.CommitTrans
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaMerDemErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMerDem:ModificaMerDem Method")
End Function

'Actualiza un Item de MerDem
'##ModelId=3B02D4EA0317
Public Function GetMerDemTabla(Optional pbArbol As Boolean = False) As ADODB.Recordset
    On Error GoTo GetMerDemTablaErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbArbol Then
        sqlC = "Select cRHMerDemCod,cRHMerDemDescripcion,len(cRHMerDemCod) Nivel From RHMerDemTabla"
    Else
        sqlC = "Select '' ITEM, cRHMerDemCod Codigo, cRHMerDemDescripcion Descripcion, ltrim(rtrim(str(CO.nConsValor))) + ' - ' + CO.cConsDescripcion MER_DEM, cUltimaActualizacion Actualizacion from RHMerDemTabla RH Inner Join Constante CO On RH.nRHMErDemTpo = Co.nConsValor And CO.nConsCod = " & gRHDemeritosCab & " Order By cRHMerDemCod"
    End If
    
    If oCon.AbreConexion Then
        Set GetMerDemTabla = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
        
    Set oCon = Nothing
    Exit Function
GetMerDemTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMerDem:ModificaMerDem Method")
End Function

'Actualiza un Item de MerDem
'##ModelId=3B02D4EA0317
Public Function GetMerDems(psPersCod As String) As ADODB.Recordset
    On Error GoTo GetMerDemsErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
        
    sqlC = " Select RHM.cRHMerDemCod Cod, cRHMerDemDescripcion Des, Convert(varchar(10),dRHMerDemFecha,103) Fecha, cRHMerDemComentario Com, RHM.cRHMerDemCod bit1, dRHMerDemFecha Bit2, RHM.cUltimaActualizacion Movimiento" _
         & " From RHMeritosDemeritos RHM" _
         & " Inner Join RHMerDemTabla RHMT On RHM.cRHMerDemCod = RHMT.cRHMerDemCod" _
         & " Where cRHPersCod = '" & psPersCod & "' Order by dRHMerDemFecha"
    
    If oCon.AbreConexion Then
        Set GetMerDems = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
        
    Set oCon = Nothing
    Exit Function
GetMerDemsErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosMerDem:ModificaMerDem Method")
End Function

'##ModelId=3B02D59D0008
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

'##ModelId=3B02D59D00A8
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

#If DebugMode Then
    '##ModelId=3B02D59C0219
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If



