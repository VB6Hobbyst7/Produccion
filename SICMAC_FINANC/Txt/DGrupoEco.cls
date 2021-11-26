VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGrupoEco"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetGE() As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sql = " Select cGECod,cGENom,cConsDescripcion + Space(50) + STR(nConsValor) Valor From PersGrupoEcon GU " _
        & " Inner Join Constante CO On CO.nConsCod = 4027 And GU.nGETipo = CO.nConsValor Order By cGECod "

    oCon.AbreConexion
    
    Set GetGE = oCon.CargaRecordSet(sql)
    
    oCon.CierraConexion
End Function

Public Function GetGEExiste(psGECod As String) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select cGECod From PersGrupoEcon" _
        & " Where cGECod = '" & psGECod & "'"

    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then
        GetGEExiste = False
    Else
        GetGEExiste = True
    End If
    
    oCon.CierraConexion
End Function

Public Function GetGERelExiste(psGECod As String, psPersCod As String) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select cGECod From PersGERelacion " _
        & " Where cGECod = '" & psGECod & "' And cPersCodRel = '" & psPersCod & "'"

    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then
        GetGERelExiste = False
    Else
        GetGERelExiste = True
    End If
    
    oCon.CierraConexion
End Function

Public Function GetGERelVinExiste(psPersCodEmp As String, psPersCodPers As String) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select cPersCodVinc From PersGEPersVinc " _
        & " Where cPersCodVinc = '" & psPersCodPers & "' And cPersCodRel = '" & psPersCodEmp & "'"

    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(sql)
    
    If Not (rs.EOF And rs.BOF) Then
        GetGERelVinExiste = False
    Else
        GetGERelVinExiste = True
    End If
    
    oCon.CierraConexion
End Function

Public Function ActulizaGE(prRS As ADODB.Recordset) As Boolean
    Dim sql As String
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If prRS Is Nothing Then
        lsCad = ""
        sql = "Delete PersGrupoEcon Where cGECod Not In ('" & lsCad & "')"
        oCon.Ejecutar sql
        Exit Function
    Else
        prRS.MoveFirst
        
        While Not prRS.EOF
            If lsCad = "" Then
                lsCad = prRS.Fields(0)
            Else
                lsCad = lsCad & "','" & prRS.Fields(0)
            End If
            prRS.MoveNext
        Wend
    End If
    
    sql = "Delete PersGrupoEcon Where cGECod Not In ('" & lsCad & "')"
    oCon.Ejecutar sql
        
    prRS.MoveFirst
    
    While Not prRS.EOF
        If GetGEExiste(prRS.Fields(0)) Then
            sql = " Insert PersGrupoEcon (cGECod, cGENom, nGETipo)" _
                & " Values('" & prRS.Fields(0) & "','" & prRS.Fields(1) & "','" & Right(prRS.Fields(2), 2) & "')"
        Else
            sql = " Update PersGrupoEcon" _
                & " Set cGENom = '" & prRS.Fields(1) & "', nGETipo = '" & Right(prRS.Fields(2), 2) & "'" _
                & " Where cGECod = '" & prRS.Fields(0) & "'"
        End If
        oCon.Ejecutar sql
        prRS.MoveNext
    Wend
    
    oCon.CierraConexion
    Set oCon = Nothing
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
End Sub


Public Function GetGEEmp(psGE As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    
    oCon.AbreConexion
    
    sql = " Select  GU.cPersCodRel, PE.cPersNombre, CO.cConsDescripcion + Space(50) + Str(CO.nConsvalor) Valor " _
        & " From PersGERelacion GU" _
        & " Inner Join Persona PE On GU.cPersCodRel = PE.cPersCod" _
        & " Inner Join Constante CO On CO.nConsCod = 4028 And GU.nPrdPersRelac = CO.nConsValor" _
        & " Where cGEcod = '" & psGE & "'"
    
    Set GetGEEmp = oCon.CargaRecordSet(sql)
    
    Set oCon = Nothing
End Function

Public Function GetGEPers(psGE As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    
    oCon.AbreConexion
    
    sql = " Select GE.cPersCodVinc, PE.cPersNombre, CO.cConsDescripcion + Space(50) + Str(CO.nConsvalor) Valor," _
        & " nParticip, cPersCodRel from PersGEPersVinc GE" _
        & " Inner Join Persona PE ON GE.cPersCodVinc = PE.cPersCod" _
        & " Inner Join Constante CO On CO.nConsCod = 4029 And GE.nCargo = CO.nConsValor" _
        & " Where cPersCodRel In (select cPersCodRel From PersGERelacion Where cGEcod = '" & psGE & "')"
    
    Set GetGEPers = oCon.CargaRecordSet(sql)
    
    Set oCon = Nothing
End Function

Public Function ActulizaGERel(prRSEmp As ADODB.Recordset, prRSPers As ADODB.Recordset, psGECod As String) As Boolean
    Dim sql As String
    Dim lsCad As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsCadP As String
    
    oCon.AbreConexion
    
    If prRSEmp Is Nothing Then
        sql = "Delete PersGEPersVinc Where cPersCodRel In (Select cPersCodRel From PersGERelacion Where cGECod = '" & psGECod & "')"
        oCon.Ejecutar sql
        sql = "Delete PersGERelacion Where cGECod = '" & psGECod & "'"
        oCon.Ejecutar sql
        Exit Function
    Else
        prRSEmp.MoveFirst
        
        While Not prRSEmp.EOF
            prRSPers.MoveFirst
            lsCadP = ""
            
            If lsCad = "" Then
                lsCad = prRSEmp.Fields(0)
            Else
                lsCad = lsCad & "','" & prRSEmp.Fields(0)
            End If
            
            If Not prRSPers Is Nothing Then
                While Not prRSPers.EOF
                    If lsCadP = "" Then
                        lsCadP = prRSPers.Fields(0)
                    Else
                        lsCadP = lsCadP & "','" & prRSPers.Fields(0)
                    End If
                    prRSPers.MoveNext
                Wend
            End If
            sql = "Delete PersGEPersVinc Where cPersCodRel = '" & prRSEmp.Fields(0) & "' And cPersCodVinc Not In ('" & lsCadP & "')"
            oCon.Ejecutar sql
            
            prRSEmp.MoveNext
        Wend
    End If
    
    sql = "Delete PersGERelacion Where cGECod = '" & psGECod & "' And cPersCodRel not in ('" & lsCad & "')"
    oCon.Ejecutar sql
        
    prRSEmp.MoveFirst
    While Not prRSEmp.EOF
        If GetGERelExiste(psGECod, prRSEmp.Fields(0)) Then
            sql = " Insert PersGERelacion (cGEcod,cPersCodRel,nPrdPersRelac)" _
                & " Values('" & psGECod & "','" & prRSEmp.Fields(0) & "','" & Right(prRSEmp.Fields(2), 2) & "')"
        Else
            sql = " Update PersGERelacion" _
                & " Set nPrdPersRelac = '" & Right(prRSEmp.Fields(2), 2) & "'" _
                & " Where cGECod = '" & psGECod & "' And cPersCodRel = '" & prRSEmp.Fields(0) & "'"
        End If
        oCon.Ejecutar sql
        prRSEmp.MoveNext
    Wend
    
    If Not prRSPers Is Nothing Then
        prRSPers.MoveFirst
        While Not prRSPers.EOF
            If GetGERelVinExiste(prRSPers.Fields(4), prRSPers.Fields(0)) Then
                sql = " Insert PersGEPersVinc (cPersCodRel,cPersCodVinc,nCargo,nParticip)" _
                    & " Values('" & prRSPers.Fields(4) & "','" & prRSPers.Fields(0) & "','" & Right(prRSPers.Fields(2), 3) & "'," & prRSPers.Fields(3) & ")"
            Else
                sql = " Update PersGEPersVinc" _
                    & " Set nCargo = '" & Right(prRSPers.Fields(2), 2) & "', nParticip = " & prRSPers.Fields(3) & "" _
                    & " Where cPersCodVinc = '" & prRSPers.Fields(0) & "' And cPersCodRel = '" & prRSPers.Fields(4) & "'"
            End If
            oCon.Ejecutar sql
            prRSPers.MoveNext
        Wend
    End If
    oCon.CierraConexion
    Set oCon = Nothing
End Function
