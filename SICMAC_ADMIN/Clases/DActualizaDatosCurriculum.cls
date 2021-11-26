VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaDatosCurriculum"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B02D4B2032B"
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

'Agrega un nuevo Item de Curriculum Tabla
'##ModelId=3B02D4EA01D7
Public Function AgregaCurriculumTabla(psRHCurriculumTablaCod As String, psRHpcRHCurriculumTablaCodDescripcion As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaCurriculumTablaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Insert RHCurriculumTabla (cRHCurriculumCod,cRHCurriculumDescripcion,cUltimaActualizacion)" _
         & " Values('" & psRHCurriculumTablaCod & "','" & psRHpcRHCurriculumTablaCodDescripcion & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:AgregaCurriculumTabla Method")
End Function

'Elimina  un Item de Curriculum Tabla
'##ModelId=3B02D4EA0213
Public Function EliminaCurriculumTabla(psRHCurriculumTablaCod As String) As Boolean
    On Error GoTo EliminaCurriculumTablaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Delete RHCurriculumTabla where cRHCurriculumCod = '" & psRHCurriculumTablaCod & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:EliminaCurriculumTabla Method")
End Function

'Elimina  un Item de Curriculum Tabla
'##ModelId=3B02D4EA0213
Public Function TipoUsado(psRHCurriculumTablaCod As String) As Boolean
    On Error GoTo EliminaCurriculumTablaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlC = " Select cRHCurriculumCod From RHCurriculum where cRHCurriculumCod = '" & psRHCurriculumTablaCod & "'"
    
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
EliminaCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:EliminaCurriculumTabla Method")
End Function

'Actualiza un Item de Curriculum Tabla
'##ModelId=3B02D4EA024F
Public Function ModificaCurriculumTabla(psRHCurriculumTablaCod As String, psRHpcRHCurriculumTablaCodDescripcion As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaCurriculumTablaErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Update RHCurriculumTabla" _
         & " Set cRHCurriculumDescripcion = '" & psRHpcRHCurriculumTablaCodDescripcion & "', cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
         & " Where cRHCurriculumCod = '" & psRHCurriculumTablaCod & "'"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:ModificaCurriculumTabla Method")
End Function

'Agrega un nuevo Item de Curriculum
'##ModelId=3B02D4EA028B
Public Function AgregaCurriculum(psRHPersCod As String, psRHCurriculumCod As String, pnRHItem As Integer, psRHCurLugar As String, psRHCurInicio As String, psRHCurFin As String, pnRHCurUndTiempo As Integer, pnRHCurTiempo As Integer, psComentario As String, psUltimaActualizacion As String) As Boolean
    On Error GoTo AgregaCurriculumErr
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
    
    sqlC = " Insert RHCurriculum (cRHPersCod,cRHCurriculumCod,nRHCurriculumItem,cRHCurriculumLugar,dRHCurriculumInicio,dRHCurriculumFin,cRHCurriculumUndTiempo,nRHCurriculumTiempo,cRHCurriculumComentario,cUltimaActualizacion)" _
         & " Values('" & psRHPersCod & "','" & psRHCurriculumCod & "'," & pnRHItem & ",'" & psRHCurLugar & "'," & lsRHCurInicio & "," & lsRHCurFin & "," & lsRHCurUndTiempo & "," & lsRHCurTiempo & ",'" & psComentario & "','" & psUltimaActualizacion & "')"
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
AgregaCurriculumErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:AgregaCurriculum Method")
End Function

'Elimina  un Item de Curriculum
'##ModelId=3B02D4EA02D1
Public Function EliminaCurriculum(psRHCurriculumCod As String, psRHPersCod As String, pnRHItem As Integer) As Boolean
    On Error GoTo EliminaCurriculumErr
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Delete RHCurriculum" _
         & " Where cRHPersCod = '" & psRHPersCod & "' And cRHCurriculumCod = '" & psRHCurriculumCod & "' And nRHCurriculumItem = " & pnRHItem
    
    If oCon.AbreConexion Then
        oCon.Ejecutar sqlC
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaCurriculumErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:EliminaCurriculum Method")
End Function

'Actualiza un Item de Curriculum
'##ModelId=3B02D4EA0317
Public Function ModificaCurriculum(psPersCod As String, prRS As ADODB.Recordset, prRSExtra As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    On Error GoTo ModificaCurriculumErr
    Dim oCon As DConecta
    Dim sqlC As String
    Dim sqlCExtra As String
    Dim lsRHCurInicio As String
    Dim lsRHCurFin As String
    Dim lsRHCurUndTiempo As String
    Dim lsRHCurTiempo As String
    Set oCon = New DConecta
    Dim rsC As ADODB.Recordset
    Set rsC = New ADODB.Recordset
    Dim lsCadena  As String
    
    lsCadena = ""
    If Not prRS Is Nothing Then
    
      If Not prRS.EOF And Not prRS.BOF Then
        While Not prRS.EOF
            If lsCadena = "" Then
                lsCadena = prRS.Fields(18) & prRS.Fields(19)
            Else
                lsCadena = lsCadena & "','" & prRS.Fields(18) & prRS.Fields(19)
            End If
            prRS.MoveNext
        Wend

        sqlC = " Delete rhCurriculum Where cRHCurriculumCod + RTrim(LTrim(nRHCurriculumItem)) not in ('" & lsCadena & "') And cRHPersCod = '" & psPersCod & "'"
        prRS.MoveFirst
    
      End If
    End If
    'Extra
    lsCadena = ""
    
    If Not prRSExtra Is Nothing Then
    
        If Not prRSExtra.BOF And Not prRSExtra.EOF Then
            While Not prRSExtra.EOF
                If lsCadena = "" Then
                    lsCadena = prRSExtra.Fields(11) & prRSExtra.Fields(12)
                Else
                    lsCadena = lsCadena & "','" & prRSExtra.Fields(11) & prRSExtra.Fields(12)
                End If
                prRSExtra.MoveNext
            Wend
    
            sqlCExtra = " Delete rhCurriculumExtra Where RTrim(LTrim(nRHCurriculumCod)) + RTrim(LTrim(nRHCurriculumItem)) Not in ('" & lsCadena & "') And cRHPersCod = '" & psPersCod & "'"
            prRSExtra.MoveFirst
        End If
    End If
    '------------
    
    If oCon.AbreConexion Then
        If sqlC <> "" Then
            oCon.Ejecutar sqlC
        End If
        If sqlCExtra <> "" Then
            oCon.Ejecutar sqlCExtra
        End If
        While Not prRS.EOF
            sqlC = "Select cRHPersCod From RHCurriculum Where cRHPersCod = '" & psPersCod & "' And nRHCurriculumItem = '" & prRS.Fields(19) & "'"
            Set rsC = oCon.CargaRecordSet(sqlC)
            
            If prRS.Fields(3) = "0" Then
                lsRHCurInicio = "'" & Format(CDate(prRS.Fields(5)), gsFormatoFecha) & "'"
                lsRHCurFin = "'" & Format(CDate(prRS.Fields(6)), gsFormatoFecha) & "'"
                lsRHCurUndTiempo = "NUll"
                lsRHCurTiempo = "NULL"
            ElseIf prRS.Fields(3) = "1" Then
                lsRHCurInicio = "NULL"
                lsRHCurFin = "NULL"
                lsRHCurUndTiempo = "'" & prRS.Fields(7) & "'"
                lsRHCurTiempo = "" & prRS.Fields(9) & ""
            ElseIf prRS.Fields(3) = "2" Then
                lsRHCurInicio = "'" & Format(CDate(prRS.Fields(5)), gsFormatoFecha) & "'"
                lsRHCurFin = "'" & Format(CDate(prRS.Fields(6)), gsFormatoFecha) & "'"
                lsRHCurUndTiempo = "'" & prRS.Fields(7) & "'"
                lsRHCurTiempo = "" & prRS.Fields(9) & ""
            End If
            
            If rsC.EOF And rsC.BOF Then
                rsC.Close
                sqlC = "Select Top 1 nRHCurriculumItem Item From RHCurriculum Where cRHPersCod = '" & psPersCod & "' Order by nRHCurriculumItem Desc"
                Set rsC = oCon.CargaRecordSet(sqlC)
                
                If rsC.EOF And rsC.BOF Then
                    lsCadena = "1"
                Else
                    lsCadena = Trim(Str(CCur(rsC!Item) + 1))
                End If
                
                sqlC = " Insert RHCurriculum (cRHPersCod,cRHCurriculumCod,nRHCurriculumItem,cRHCurriculumLugar,dRHCurriculumInicio,dRHCurriculumFin,cRHCurriculumUndTiempo,nRHCurriculumTiempo,cRHCurriculumComentario,cUltimaActualizacion,nRHProfesionCod, nRHCosto, nRHCurrNivel, bCmactOterogado)" _
                     & " Values('" & psPersCod & "','" & prRS.Fields(0) & "'," & lsCadena & ",'" & prRS.Fields(2) & "'," & lsRHCurInicio & "," & lsRHCurFin & "," & lsRHCurUndTiempo & "," & lsRHCurTiempo & ",'" & prRS.Fields(16) & "','" & psUltimaActualizacion & "'," & prRS.Fields(10) & "," & CCur(prRS.Fields(12)) & "," & prRS.Fields(13) & "," & prRS.Fields(15) & ")"
            Else
                sqlC = " Update RHCurriculum" _
                     & " Set cRHCurriculumCod = '" & prRS.Fields(0) & "', cRHCurriculumLugar = '" & prRS.Fields(2) & "', dRHCurriculumInicio = " & lsRHCurInicio & ", dRHCurriculumFin = " & lsRHCurFin & ", cRHCurriculumUndTiempo = " & lsRHCurUndTiempo & ", nRHCurriculumTiempo = " & lsRHCurTiempo & " , cRHCurriculumComentario = '" & prRS.Fields(16) & "', cUltimaActualizacion = '" & psUltimaActualizacion & "', nRHProfesionCod = " & prRS.Fields(10) & ", nRHCosto = " & CCur(prRS.Fields(12)) & ", nRHCurrNivel = " & prRS.Fields(13) & " , bCmactOterogado = " & prRS.Fields(15) & " " _
                     & " Where cRHPersCod = '" & psPersCod & "' And cRHCurriculumCod = '" & prRS.Fields(18) & "' And nRHCurriculumItem = " & prRS.Fields(19)
            End If
            oCon.Ejecutar sqlC
            prRS.MoveNext
        Wend
        
        'extra
    If Not prRSExtra Is Nothing Then
        While Not prRSExtra.EOF
            sqlC = "Select cRHPersCod From RHCurriculumExtra Where cRHPersCod = '" & psPersCod & "' And nRHCurriculumItem = '" & prRSExtra.Fields(11) & "'"
            Set rsC = oCon.CargaRecordSet(sqlC)
            
            If rsC.EOF And rsC.BOF Then
                rsC.Close
                sqlC = "Select Top 1 nRHCurriculumItem Item From RHCurriculumExtra Where cRHPersCod = '" & psPersCod & "' Order by nRHCurriculumItem Desc"
                Set rsC = oCon.CargaRecordSet(sqlC)
                
                If rsC.EOF And rsC.BOF Then
                    lsCadena = "1"
                Else
                    lsCadena = Trim(Str(CCur(rsC!Item) + 1))
                End If
                
                sqlC = " Insert RHCurriculumExtra (cRHPersCod,nRHCurriculumCod,nRHCurriculumItem,nRHActividadCod,nRHAñosPractica,nCosto,nCodNivel,nOtrogadoCMACT,cRHCurriculumComentario,cUltimaActualizacion)" _
                     & " Values('" & psPersCod & "','" & prRSExtra.Fields(0) & "'," & lsCadena & ",'" & prRSExtra.Fields(2) & "','" & prRSExtra.Fields(4) & "'," & Format(prRSExtra.Fields(5), "#.00") & ",'" & prRSExtra.Fields(6) & "'," & prRSExtra.Fields(8) & ",'" & prRSExtra.Fields(9) & "','" & psUltimaActualizacion & "')"
            Else
                sqlC = " Update RHCurriculumExtra" _
                     & " Set nRHActividadCod = '" & prRSExtra.Fields(2) & "',nRHAñosPractica = '" & prRSExtra.Fields(4) & "',nCosto = " & CCur(prRSExtra.Fields(5)) & ",nCodNivel = '" & prRSExtra.Fields(6) & "',nOtrogadoCMACT='" & prRSExtra.Fields(8) & "',cRHCurriculumComentario='" & prRSExtra.Fields(9) & "',cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
                     & " Where cRHPersCod = '" & psPersCod & "' And nRHCurriculumCod = '" & prRSExtra.Fields(11) & "' And nRHCurriculumItem = " & prRSExtra.Fields(12)
            End If
            oCon.Ejecutar sqlC
            prRSExtra.MoveNext
        Wend
    End If
        '---------------
        
        
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
ModificaCurriculumErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:ModificaCurriculum Method")
End Function

'Actualiza un Item de Curriculum
'##ModelId=3B02D4EA0317
Public Function GetCurriculumTabla(Optional pbArbol As Boolean = False) As ADODB.Recordset
    On Error GoTo GetCurriculumTablaErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    If pbArbol Then
        sqlC = "Select cRHCurriculumCod Codigo, cRHCurriculumDescripcion Descripcion, 2 Nivel from RHCurriculumTabla Order By convert(int,cRHCurriculumCod) "
    Else
        sqlC = "Select '',cRHCurriculumCod Codigo, cRHCurriculumDescripcion Descripcion, cUltimaActualizacion Actualizacion from RHCurriculumTabla Order by convert(int,cRHCurriculumCod)"
    End If
    
    If oCon.AbreConexion Then
        Set GetCurriculumTabla = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
        
    Set oCon = Nothing
    Exit Function
GetCurriculumTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:ModificaCurriculum Method")
End Function


'Actualiza un Item de Curriculum
'##ModelId=3B02D4EA0317
Public Function GetCurriculums(psPersCod As String) As ADODB.Recordset
    On Error GoTo GetCurriculumsErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    
    sqlC = " Select RCT.cRHCurriculumCod, RCT.cRHCurriculumDescripcion Descripcion, cRHCurriculumLugar Lugar," _
         & " Case When dRHCurriculumInicio Is Not Null And cRHCurriculumUndTiempo Is Not Null Then '2' When dRHCurriculumInicio Is Null Then '1' Else '0' End TpoPer, Case When dRHCurriculumInicio Is Not Null And cRHCurriculumUndTiempo Is Not Null Then 'TIEMPO Y PERIODO' When dRHCurriculumInicio Is Null Then 'PERIODO' Else 'TIEMPO' End Per," _
         & " convert(Varchar(10),dRHCurriculumInicio,103) Desde, convert(Varchar(10),dRHCurriculumFin,103) Hasta," _
         & " CO.nConsValor CodUndPer, CO.cConsDescripcion UndPer, nRHCurriculumTiempo NumPeriodos," _
         & " RC.nRHProfesionCod, COP.cConsDescripcion, RC.nRHCosto, nRHCurrNivel, CON.cConsDescripcion, bCmactOterogado CMACTOtrogado, " _
         & " cRHCurriculumComentario Comentario, RCT.cUltimaActualizacion UltMov," _
         & " RCT.cRHCurriculumCod, nRHCurriculumItem Item" _
         & "    From rhcurriculum RC" _
         & " Inner Join RHCurriculumTabla RCT" _
         & "    On RC.cRHCurriculumCod = RCT.cRHCurriculumCod" _
         & " Left Join Constante CO" _
         & "     On CO.nConsValor = RC.cRHCurriculumUndTiempo And CO.nConsCod = '1012'" _
         & " Left Join Constante COP" _
         & "     On COP.nConsValor = RC.nRHProfesionCod And COP.nConsCod = '6044'" _
         & " Left Join Constante CON" _
         & "     On CON.nConsValor = RC.nRHCurrNivel And CON.nConsCod = '6045'" _
         & " Where cRHPersCod = '" & psPersCod & "'"
         
    If oCon.AbreConexion Then
        Set GetCurriculums = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
        
    Set oCon = Nothing
    Exit Function
GetCurriculumsErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:ModificaCurriculum Method")
End Function

Public Function GetCurriculumsExtra(psPersCod As String) As ADODB.Recordset
    On Error GoTo GetCurriculumsErr
    Dim sqlC As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sqlC = " Select nRHCurriculumCod, CTPO.cConsDescripcion, RHE.nRHActividadCod, " _
         & "   Case nRHCurriculumCod" _
         & "        When 0 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6047 And COI.nConsValor = RHE.nRHActividadCod)" _
         & "        When 1 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6048 And COI.nConsValor = RHE.nRHActividadCod)" _
         & "        When 2 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6049 And COI.nConsValor = RHE.nRHActividadCod)" _
         & "        When 3 Then (Select cConsDescripcion From Constante COI Where COI.nConsCod = 6050 And COI.nConsValor = RHE.nRHActividadCod) End Actividad," _
         & "  RHE.nRHAñosPractica, RHE.nCosto, RHE.nCodNivel, CNIV.cConsDescripcion, nOtrogadoCMACT, RHE.cRHCurriculumComentario," _
         & "  RHE.cUltimaActualizacion , RHE.nRHCurriculumCod, RHE.nRHCurriculumItem" _
         & "  From RHCurriculumExtra RHE" _
         & "  Inner Join Constante CTPO On RHE.nRHCurriculumCod = CTPO.nConsValor And CTPO.nConsCod = '6046'" _
         & "  Left  Join Constante CNIV On RHE.nCodNivel = CNIV.nConsValor And CNIV.nConsCod = '6051'" _
         & "  Where RHE.cRHPersCod = '" & psPersCod & "' Order by RHE.nRHCurriculumCod"
    
    If oCon.AbreConexion Then
        Set GetCurriculumsExtra = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
        
    Set oCon = Nothing
    Exit Function
GetCurriculumsErr:
    Call RaiseError(MyUnhandledError, "DActualizaDatosCurriculum:ModificaCurriculum Method")
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

