VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAjusteDeprecia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset
Dim sConexion As String

Public Function CargaAjusteDeprecia(Optional pnCodigo As Long = -1, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaAjusteDepreciaErr
      psSql = "SELECT nCodigo, cDescrip, cNomArch, cCtaContCodD, cCtaContCodH FROM AjusteDeprecia " & IIf(pnCodigo = -1, "", "WHERE nCodigo = " & pnCodigo)
      Set CargaAjusteDeprecia = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaAjusteDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:CargaAjusteDeprecia Method")
End Function

Public Sub InsertaAjusteDeprecia(ByRef pnCodigo As Long, psDescrip As String, psNomArch As String, psCtaContCodD As String, psCtaContCodH As String)
   On Error GoTo InsertaAjusteDepreciaErr
      psSql = "INSERT AjusteDeprecia (cDescrip, cNomArch, cCtaContCodD, cCtaContCodH ) " _
            & "VALUES ('" & psDescrip & "','" & psNomArch & "','" & psCtaContCodD & "','" & psCtaContCodH & "')"
      dbConec.Ejecutar psSql
      psSql = "SELECT ISNULL(max(nCodigo),0) nMaxCodigo FROM AjusteDeprecia "
      Set prs = dbConec.CargaRecordSet(psSql)
      If Not prs.EOF Then
         pnCodigo = prs!nMaxCodigo
      Else
         pnCodigo = 1
      End If
      prs.Close: Set prs = Nothing
   Exit Sub
InsertaAjusteDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:InsertaAjusteDeprecia Method")
End Sub

Public Sub ActualizaAjusteDeprecia(pnCodigo As Long, psDescrip As String, psNomArch As String, psCtaContCodD As String, psCtaContCodH As String)
   On Error GoTo ActualizaAjusteDepreciaErr
   psSql = "UPDATE AjusteDeprecia SET cDescrip = '" & psDescrip & "', cNomArch = '" & psNomArch & "', cCtaContCodD = '" & psCtaContCodD & "', cCtaContCodH = '" & psCtaContCodH & "' WHERE nCodigo = " & pnCodigo & ""
   dbConec.Ejecutar psSql
   Exit Sub
ActualizaAjusteDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:ActualizaAjusteDeprecia Method")
End Sub

Public Sub EliminaAjusteDeprecia(pnCodigo As Long)
   On Error GoTo EliminaAjusteDepreciaErr
      psSql = "DELETE AjusteDeprecia WHERE nCodigo = " & pnCodigo & ""
      dbConec.Ejecutar psSql
   Exit Sub
EliminaAjusteDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:EliminaAjusteDeprecia Method")
End Sub

Public Function CargaAjusteDepreciaDet(Optional pnCodigo As Long = -1, Optional pnItem As Long = -1, Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
   On Error GoTo CargaAjusteDepreciaErr
      psSql = "SELECT nCodigo, nItem, cBSCod, cUbicacion, cDescrip, dFecAdq, nVidaUtil,nValHis " _
            & "FROM AjusteDepreciaDet " & IIf(pnCodigo = -1, "", "WHERE nCodigo = " & pnCodigo & " ") & IIf(pnItem > -1, IIf(pnCodigo = -1, "WHERE", "AND") & " nItem = " & pnItem, "")
      Set CargaAjusteDepreciaDet = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaAjusteDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:CargaAjusteDeprecia Method")
End Function

Public Sub InsertaAjusteDepreciaDet(pnCodigo As Long, pnItem As Long, psBSCod As String, psUbicacion As String, psDescrip As String, psFecAdq As String, pnVidaUtil As Currency, pnValHist As Currency)
   On Error GoTo InsertaAjusteDepreciaErr
      psSql = "INSERT AjusteDepreciaDet (nCodigo, nItem, cBSCod, cUbicacion, cDescrip, dFecAdq, nVidaUtil, nValHis ) " _
            & "VALUES (" & pnCodigo & "," & pnItem & ",'" & psBSCod & "','" & psUbicacion & "','" & psDescrip & "','" & psFecAdq & "'," & pnVidaUtil & "," & pnValHist & ")"
      dbConec.Ejecutar psSql
   Exit Sub
InsertaAjusteDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:InsertaAjusteDeprecia Method")
End Sub

Public Sub ActualizaAjusteDepreciaDet(pnCodigo As Long, pnItem As Long, psBSCod As String, psUbicacion As String, psDescrip As String, psFecAdq As String, pnVidaUtil As Currency, pnValHist As Currency)
   On Error GoTo ActualizaAjusteDepreciaErr
   psSql = "UPDATE AjusteDepreciaDet SET cBSCod = '" & psBSCod & "', cUbicacion = '" & psUbicacion & "', cDescrip = '" & psDescrip & "', dFecAdq = '" & psFecAdq & "', nVidaUtil = " & pnVidaUtil & ", nValHis = " & pnValHist & " WHERE nCodigo = " & pnCodigo & " and nItem = " & pnItem
   dbConec.Ejecutar psSql
   Exit Sub
ActualizaAjusteDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:ActualizaAjusteDeprecia Method")
End Sub

Public Sub EliminaAjusteDepreciaDet(pnCodigo As Long, pnItem As Long)
   On Error GoTo EliminaAjusteDepreciaErr
      psSql = "DELETE AjusteDepreciaDet WHERE nCodigo = " & pnCodigo & " and nItem = " & pnItem
      dbConec.Ejecutar psSql
   Exit Sub
EliminaAjusteDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:EliminaAjusteDeprecia Method")
End Sub

Public Function GetItemDeprecia(pnCodigo As Long) As Long
On Error GoTo GetItemDepreciaErr
      psSql = "select (ISNULL(MAX(nItem),0)+1) AS nNumMax From AjusteDepreciaDet Where nCodigo = " & pnCodigo
      Set prs = dbConec.CargaRecordSet(psSql)
      If Not prs.EOF Then
         GetItemDeprecia = prs!nNumMax
      Else
         GetItemDeprecia = 1
      End If
   Exit Function
GetItemDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:GetItemDeprecia Method")
End Function

Public Function CargaActivosDeprecia(pnCodigo As Long, psFecAdq As String) As Recordset
   On Error GoTo CargaActivosDepreciaErr
      psSql = "SELECT R.nCodigo, R.cDescrip, R.cNomArch,  RD.cBSCod, RD.cUbicacion, RD.cDescrip, RD.dFecAdq, RD.nVidaUtil, RD.nValHis " _
             & " FROM    AjusteDeprecia R JOIN AjusteDepreciaDet RD On R.nCodigo = RD.nCodigo " _
             & " Where   R.nCodigo = " & pnCodigo & " And RD.dFecAdq <= '" & psFecAdq & "' Order By RD.cDescrip"
      Set CargaActivosDeprecia = dbConec.CargaRecordSet(psSql)
   Exit Function
CargaActivosDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:CargaActivosDeprecia Method")
End Function

Public Sub InsertarBSDeprecia(psCodigo As String, rsDep As ADODB.Recordset)
    On Error GoTo CargaActivosDepreciaErr
    Dim sql As String
   
    rsDep.MoveFirst
    While Not rsDep.EOF
        If rsDep.Fields(0) = "1" Then
            sql = " INSERT AjusteDepreciaDet (nCodigo,cBSCod,cSerie,cUbicacion,cDescrip,dFecAdq,nVidaUtil,nValHis) " _
                & " VALUES (" & psCodigo & ",'" & rsDep.Fields(1) & "','" & rsDep.Fields(2) & "','" & rsDep.Fields(6) & "','" & rsDep.Fields(3) & "','" & Format(CDate(rsDep.Fields(4)), gsFormatoFecha) & "'," & rsDep.Fields(7) & "," & rsDep.Fields(5) & ")"
            dbConec.Ejecutar sql
        End If
        rsDep.MoveNext
    Wend
   
   Exit Sub
CargaActivosDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:CargaActivosDeprecia Method")
End Sub

Public Sub MantenimientoBSDeprecia(psCodigo As String, rsDep As ADODB.Recordset)
    On Error GoTo CargaActivosDepreciaErr
    Dim sql As String
   
    rsDep.MoveFirst
    While Not rsDep.EOF
        If rsDep.Fields(0) = "1" Then
            sql = " UPDATE AjusteDepreciaDet" _
                & " Set cUbicacion = '" & rsDep.Fields(6) & "', cDescrip = '" & rsDep.Fields(3) & "', dFecAdq = '" & Format(CDate(rsDep.Fields(4)), gsFormatoFecha) & "', nVidaUtil = " & rsDep.Fields(7) & ", nValHis = " & rsDep.Fields(5) & " " _
                & " Where nCodigo = " & psCodigo & " And cBSCod = '" & rsDep.Fields(1) & "' And cSerie = '" & rsDep.Fields(2) & "'"
        Else
            sql = " Delete AjusteDepreciaDet" _
                & " Where nCodigo = " & psCodigo & " And cBSCod = '" & rsDep.Fields(1) & "' And cSerie = '" & rsDep.Fields(2) & "'"
        End If
        dbConec.Ejecutar sql
        rsDep.MoveNext
    Wend
    
   Exit Sub
CargaActivosDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:CargaActivosDeprecia Method")
End Sub

Public Function CargaActivosDepreciaDet(pnCodigo As Long, pgsFecha As Date) As Recordset
   On Error GoTo CargaActivosDepreciaErr
       psSql = " Select 1 OK, cBSCod Codigo ,cSerie Serie, cDescrip Descripción, Convert(Varchar(10),dFecAdq,103) Fecha, nValHis  Valor, cUbicacion , nVidaUtil" _
             & " From AjusteDepreciaDet" _
             & " Where nCodigo = " & pnCodigo '& " And Abs(DateDiff(Month,dFecAdq,'" & Format(pgsFecha, gsFormatoFecha) & "')) < nVidaUtil"
      Set CargaActivosDepreciaDet = dbConec.CargaRecordSet(psSql)
   Exit Function
CargaActivosDepreciaErr:
   Call RaiseError(MyUnhandledError, "DAjusteDeprecia:CargaActivosDeprecia Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
Dim oIni As New ClasIni
   Set oIni = Nothing
   Set dbConec = New DConecta
   dbConec.AbreConexion
End Sub

Private Sub Class_Terminate()
   dbConec.CierraConexion
   Set dbConec = Nothing
End Sub
