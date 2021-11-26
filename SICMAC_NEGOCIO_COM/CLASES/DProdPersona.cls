VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DProdPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function dObtieneCuentasPersona(ByVal psPersCod As String, ByVal psProducto As String, _
       ByVal psRelacion As String, ByVal psEstados As String, Optional ByVal bActivas As Boolean = False) As Recordset
Dim sSql As String
Dim rsCta As Recordset

sSql = "Select PP.cCtaCod, P.cPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
        & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, " _
        & "UPPER(T3.cConsDescripcion) cMoneda  " _
        & "FROM ProductoPersona PP INNER JOIN Producto P INNER " _
        & "JOIN Captaciones C ON P.cCtaCod = C.cCtaCod ON PP.cCtaCod = P.cCtaCod INNER JOIN " _
        & sDBComunes & "Constante T ON P.cPrdEstado = T.cConsValor INNER JOIN " & sDBComunes & "" _
        & "Constante T1 ON PP.nPrdPersRelac = T1.cConsValor INNER JOIN " & sDBComunes & "Constante T2 " _
        & "ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.cConsValor) INNER JOIN " & sDBComunes & "" _
        & "Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.cConsValor) WHERE PP.cPersCod = '" & sPers & "' " _
        & "AND T1.cConsCod = " & gCaptacRelacPersona & " AND T.cConsCod = " & gCaptacEstado & " AND " _
        & "T2.cConsCod = " & gProducto & " AND T3.cConsCod = " & gMoneda

If nProd <> 0 Then
    sSql = sSql & " AND SUBSTRING(PP.cCtaCod,6,3) = '" & nProd & "'"
End If
If bActivas Then
    sSql = sSql & " AND P.cPrdEstado NOT IN (" & gCapEstAnuladaAct & "," & gCapEstAnuladaInac _
        & "," & gCapEstCanceladaAct & "," & gCapEstCanceladaInac & ")"
End If
sSql = sSql & " ORDER BY PP.cCtaCod"
Set rsCta = New Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set dGetCuentasPersona = rsCta
Set rsCta = Nothing
End Function


