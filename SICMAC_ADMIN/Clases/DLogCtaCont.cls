VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DLogCtaCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetOperacion(psOpGrupo As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select cOpeCod, cOpeDesc, nOpeNiv from opetpo where cOpeCod like '" & psOpGrupo & "%'"

    If oCon.AbreConexion Then
        Set GetOperacion = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function GetCtasCont() As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select cCtaContCod, cCtaContDesc From CtaCont" _
         & " Order By cCtaContCod"
    
    If oCon.AbreConexion Then
        Set GetCtasCont = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function GetConceptoCtaDeb(psOpeCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select A.cObjetoCod, IsNull(BS.cBSDescripcion,(Select cCtaContDesc From CtaCont Where cCtaContCod = A.cObjetoCod)) Descripcion, A.cCtaContCod From CtaBS A" _
         & " Left Join BienesServicios BS On A.cObjetoCod = BS.cBSCod" _
         & " Where A.cOpeCod = '" & psOpeCod & "' Order By A.cObjetoCod"
    
    If oCon.AbreConexion Then
        Set GetConceptoCtaDeb = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function GetConceptoCtaCta(psOpeCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select cCtaContCodD, cCtaContCodH, cCtaContCodOtroD, cCtaContCodOtroH From OpeCtaCta" _
         & " Where cOpeCod = '" & psOpeCod & "' Order by cCtaContCodD"
    
    If oCon.AbreConexion Then
        Set GetConceptoCtaCta = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function GetConceptoObj(psOpeCod As String) As ADODB.Recordset
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    
    sqlC = " Select BSO.cBsCod, BS.cBSDescripcion, BSO.cObj + ' ' + OO.cObjetoDesc Objeto, BSO.cCtaContCod  from BSObj BSO " _
         & " Inner Join BienesServicios BS On BSO.cBSCod = BS.cBSCod" _
         & " Inner Join Objeto OO On BSO.cObj = OO.cObjetoCod" _
         & " Where BSO.cOpeCod = '" & psOpeCod & "' Order by cCtaContCod"
    
    If oCon.AbreConexion Then
        Set GetConceptoObj = oCon.CargaRecordSet(sqlC)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function

Public Function SetCtaBS(psOpeCod As String, prRS As ADODB.Recordset, prRSCta As ADODB.Recordset, prRSObj As ADODB.Recordset, psUltimaActualizacion As String) As Boolean
    Dim oCon As DConecta
    Dim sqlD As String
    Dim sqlH As String
    Dim sqlC As String
    
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If prRS Is Nothing Then
        sqlC = "Delete CtaBS Where cOpeCod = '" & psOpeCod & "'"
        oCon.Ejecutar sqlC
        'Exit Function
    Else
        sqlD = ""
        prRS.MoveFirst
        
        While Not prRS.EOF
            If sqlD = "" Then
                sqlD = prRS.Fields(0)
            Else
                sqlD = sqlD & "','" & prRS.Fields(0)
            End If
            prRS.MoveNext
        Wend
        
        sqlC = " Delete CtaBS Where cOpeCod = '" & psOpeCod & "' And cObjetoCod Not In ('" & sqlD & "') "
        oCon.Ejecutar sqlC
    
        prRS.MoveFirst
        
        While Not prRS.EOF
            If ExisteCtaBS(psOpeCod, prRS.Fields(0)) Then
                sqlD = " Update CtaBS" _
                     & " Set cCtaContCod = '" & prRS.Fields(2) & "'" _
                     & " Where cOpeCod = '" & psOpeCod & "' And cObjetoCod = '" & prRS.Fields(0) & "'"
            Else
                sqlD = " Insert CtaBS (cCtaContCod,nCtaObjOrden,cObjetoCod,cOpeCod,nCtaObjNiv,cCtaObjFiltro,cCtaObjImpre,cCtaObjOpc)" _
                     & " Values ('" & prRS.Fields(2) & "','1','" & prRS.Fields(0) & "','" & psOpeCod & "','2','','','')"
            End If
            oCon.Ejecutar sqlD
            prRS.MoveNext
        Wend
    End If
    
    'CtaCta
    If prRSCta Is Nothing Then
        sqlC = "Delete OpeCtaCta Where cOpeCod = '" & psOpeCod & "'"
        oCon.Ejecutar sqlC
    Else
        sqlC = "Delete OpeCtaCta Where cOpeCod = '" & psOpeCod & "'"
        oCon.Ejecutar sqlC
    
        prRSCta.MoveFirst
        
        While Not prRSCta.EOF
            sqlD = " Insert OpeCtaCta (cOpeCod,cCtaContCodD,cCtaContCodH,cCtaContCodOtroD,cCtaContCodOtroH)" _
                 & " Values ('" & psOpeCod & "','" & prRSCta.Fields(0) & "','" & prRSCta.Fields(1) & "','" & prRSCta.Fields(2) & "','" & prRSCta.Fields(3) & "')"
            oCon.Ejecutar sqlD
            prRSCta.MoveNext
        Wend
    End If
    
    'Objeto
'    If prRSObj Is Nothing Then
'        sqlC = "Delete BSObj Where cOpeCod = '" & psOpeCod & "'"
'        oCon.Ejecutar sqlC
'    Else
'        sqlC = "Delete BSObj Where cOpeCod = '" & psOpeCod & "'"
'        oCon.Ejecutar sqlC
'
'        prRSObj.MoveFirst
'
'        While Not prRSObj.EOF
'            sqlD = " Insert BSObj (cBSCod, cObj, cOpeCod, cCtaContCod, cMovNro)" _
'                 & " Values ('" & prRSObj.Fields(0) & "','" & Left(prRSObj.Fields(2), 5) & "','" & psOpeCod & "','" & prRSObj.Fields(3) & "','" & psUltimaActualizacion & "')"
'            oCon.Ejecutar sqlD
'            prRSObj.MoveNext
'        Wend
'    End If
    
    Set oCon = Nothing
End Function

Public Function ExisteCtaBS(psOpeCod As String, psConcepCod As String) As Boolean
    Dim oCon As DConecta
    Dim sqlC As String
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlC = " Select cOpeCod From CtaBS" _
         & " Where cOpeCod = '" & psOpeCod & "' And cObjetoCod = '" & psConcepCod & "'"
    
    If oCon.AbreConexion Then
        Set rs = oCon.CargaRecordSet(sqlC)
        
        If rs.EOF And rs.BOF Then
            ExisteCtaBS = False
        Else
            ExisteCtaBS = True
        End If
        
        oCon.CierraConexion
    End If
    Set oCon = Nothing
End Function


