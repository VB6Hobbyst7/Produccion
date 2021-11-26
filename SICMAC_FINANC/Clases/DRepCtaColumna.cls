VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DRepCtaColumna"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Clase de Datos para Reporte de Cuentas por Columnas"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim sCentralCom As String
Dim dbConec   As DConecta

Dim psSql As String

Public Function CargaRepColumna(Optional psOpeCod As String = "", Optional pnNroCol As Integer = -1, Optional plLike As Boolean = False, Optional LockType As LockTypeEnum = adLockReadOnly, Optional plOpeDesc As Boolean = False, Optional pnBitCentral As Boolean = False) As Recordset
Dim sCond As String
   On Error GoTo CargaRepCtaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      
        'If gbBitCentral = True Then ' Es igual para los dos distribuido y centralizado
          
         sCond = IIf(psOpeCod = "", "", " WHERE cOpeCod " & IIf(plLike, " LIKE '", "= '") & psOpeCod & IIf(plLike, "%'", "'"))
         If pnNroCol >= 0 Then
            sCond = sCond & IIf(sCond = "", " WHERE ", " AND ")
            sCond = sCond & " nNroCol = " & pnNroCol
         End If
         If plOpeDesc Then
            psSql = "SELECT rc.cOpeCod, rc.nNroCol, rc.cDescCol, rc.cTotal, ISNULL(o.cOpeDesc,'') as cOpeDesc FROM RepColumna rc LEFT JOIN " & sCentralCom & "OpeTpo o ON o.cOpeCod = rc.cOpeCod " & Replace(sCond, "cOpecod", "rc.cOpecod") & " ORDER BY rc.cOpeCod, rc.nNroCol"
         Else
            psSql = "SELECT cOpeCod, nNroCol, cDescCol, cTotal FROM RepColumna" & sCond & " ORDER BY cOpeCod, nNroCol"
         End If
         
      Set CargaRepColumna = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaRepCtaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:CargaRepCta Method")
End Function

Public Function CargaRepColumnaCta(psCodOpe As String, pnNroCol As Integer, Optional LockType As LockTypeEnum = adLockReadOnly, Optional plCtaDesc As Boolean = False) As Recordset
   On Error GoTo CargaRepColumnaCtaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      If plCtaDesc Then
         psSql = "SELECT rc.cOpeCod, rc.nNroCol, rc.cCtaContCod, c.cCtaContDesc FROM RepColumnaCta rc JOIN " & sCentralCom & "CtaCont c ON c.cCtaContCod = rc.cCtaContCod " _
               & "WHERE rc.cOpeCod = '" & psCodOpe & "' and rc.nNroCol = " & pnNroCol
      Else
         psSql = "SELECT cOpeCod, nNroCol, cCtaContCod FROM RepColumnaCta WHERE cOpeCod = '" & psCodOpe & "' and nNroCol = " & pnNroCol
      End If
      Set CargaRepColumnaCta = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaRepColumnaCtaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:CargaRepColumnaCta Method")
End Function



Public Function MaxNroColumna(psCodOpe As String) As Integer
Dim prs As ADODB.Recordset
   On Error GoTo MaxNroColumnaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT MAX(nNroCol) nNroCol FROM RepColumna WHERE cOpeCod = '" & psCodOpe & "' "
      Set prs = dbConec.CargaRecordSet(psSql)
      If Not prs.EOF Then
         MaxNroColumna = prs!nNroCol
      Else
         MaxNroColumna = 0
      End If
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
MaxNroColumnaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:MaxNroColumna Method")
End Function

Public Function CargaRepImpresion() As Recordset
   On Error GoTo CargaRepImpresionErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT RC.COPECOD, ISNULL(OT.COPEDESC,'') cOpeDesc, RC.nNroCol, RC.CDESCCOL, RCTA.CCTACONTCOD, CC.CCTACONTDESC " _
            & "FROM RepColumna RC LEFT JOIN RepColumnaCta RCTA ON RCTA.COPECOD=RC.COPECOD AND RCTA.nNroCol=RC.nNroCol " _
            & "     LEFT JOIN " & sCentralCom & "CtaCont CC ON RCTA.CCTACONTCOD=CC.CCTACONTCOD " _
            & "     LEFT JOIN " & sCentralCom & "OpeTpo OT ON OT.COPECOD=RC.COPECOD  ORDER BY RC.COPECOD "
      Set CargaRepImpresion = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaRepImpresionErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:CargaRepImpresion Method")
End Function


Public Sub InsertaRepColumna(psOpeCod As String, pnNroCol As Integer, psDescCol As String, psTotal As String, psMovNro As String)
   On Error GoTo InsertaRepColumnaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT RepColumna (cOpeCod, nNroCol, cDescCol, cTotal, cUltimaActualizacion) " _
            & "VALUES ('" & psOpeCod & "', " & pnNroCol & ", '" & psDescCol & "','" & psTotal & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaRepColumnaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:InsertaRepColumna Method")
End Sub

Public Sub InsertaRepColumnaCta(psOpeCod As String, pnNroCol As Integer, psCtaContCod As String, psMovNro As String)
   On Error GoTo InsertaRepColumnaCtaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT RepColumnaCta (cOpeCod, nNroCol, cCtaContCod, cUltimaActualizacion) VALUES ('" & psOpeCod & "', " & pnNroCol & ", '" & psCtaContCod & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaRepColumnaCtaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:InsertaRepColumnaCta Method")
End Sub

Public Sub EliminaRepColumna(psOpeCod As String, pnNroCol As Integer)
Dim lsError As String
   On Error GoTo EliminaRepColumnaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      dbConec.BeginTrans
      psSql = "DELETE RepColumnaCta WHERE cOpeCod = '" & psOpeCod & "' and nNroCol = " & pnNroCol & ""
      dbConec.Ejecutar psSql
      
      psSql = "DELETE RepColumna WHERE cOpeCod = '" & psOpeCod & "' and nNroCol = " & pnNroCol & ""
      dbConec.Ejecutar psSql
      
      dbConec.CommitTrans
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaRepColumnaErr:
   lsError = Err.Description
   dbConec.RollbackTrans
   Err.Raise Err.Number + MyUnhandledError, "DRepCtaColumna:EliminaRepColumna Method", lsError
End Sub

Public Sub EliminaRepColumnaCta(psOpeCod As String, pnNroCol As Integer, psCtaCod As String)
   On Error GoTo EliminaRepColumnaCtaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE RepColumnaCta  WHERE cOpeCod = '" & psOpeCod & "' and nNroCol = " & pnNroCol & " and cCtaContCod = '" & psCtaCod & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaRepColumnaCtaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:EliminaRepColumnaCta Method")
End Sub

Public Sub ActualizaRepColumna(psOpeCod As String, pnNroCol As Integer, psDescCol As String, pnTotal As Currency, psMovNro As String)
   On Error GoTo ActualizaRepColumnaErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "UPDATE RepColumna Set cDescCol = '" & psDescCol & "', cTotal = " & pnTotal & ", cUltimaActualizacion = '" & psMovNro & "'  WHERE cOpeCod = '" & psOpeCod & "' and nNroCol = " & pnNroCol
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaRepColumnaErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:ActualizaRepColumna Method")
End Sub

      
'Public Function GetRepColumnaCtaSaldo(psOpeCod As String, pnNroCol As Long, pdFecha As Date, Optional pnBitCentral As Boolean = False, Optional gsCodAge As String = "") As ADODB.Recordset
'    On Error GoTo GetRepColumnaCtaSaldoErr
'
'    Set dbConec = New DConecta
'
'    If dbConec.AbreConexion() Then
'
'
'        psSql = "SELECT  SUM(nCtaSaldoImporte) as Total, SUM(nCtaSaldoImporteME) AS TotalME  "
'        psSql = psSql & " FROM    CTASALDO CS JOIN RepColumnaCta OC ON cs.cCtaContCod LIKE oc.cCtaContCod + '%' "
'        psSql = psSql & " inner join ctacont c ON c.cCtaContCod = OC.cCtaContCod "
'        psSql = psSql & " WHERE   OC.nNroCol=" & pnNroCol & " AND OC.cOpeCod = '" & psOpeCod & "' "
'        'psSql = psSql & "    AND CS.NCTASALDOIMPORTE > 0 "
'        psSql = psSql & "    AND CS.dCtaSaldoFecha IN (SELECT MAX(CS1.dCtaSaldoFecha) "
'        psSql = psSql & "                              FROM CTASALDO CS1 "
'        psSql = psSql & "                              WHERE CS1.cCtaContCod = CS.cCtaContCod "
'        psSql = psSql & "                                    AND CS1.dCtaSaldoFecha <='" & Format(pdFecha, gsFormatoFecha) & "') "
'
'        If Len(Trim(gsCodAge)) = 0 Then
'        Else
'            If gsCodAge = "99" Then
'                psSql = psSql & " AND bAgencia=0"
'            Else
'                psSql = psSql & " AND right(c.cCtaContCod, 2)='" & gsCodAge & "' and bAgencia=1"
'            End If
'        End If
'
'        Set GetRepColumnaCtaSaldo = dbConec.CargaRecordSet(psSql)
'
'   End If
'
'   dbConec.CierraConexion
'   Set dbConec = Nothing
'
'   Exit Function
'GetRepColumnaCtaSaldoErr:
'   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetRepColumnaCtaSaldo Method")
'End Function

Public Function GetRepColumnaCtaSaldo(psOpeCod As String, pnNroCol As Long, pdFecha As Date, Optional pnBitCentral As Boolean = False, Optional gsCodAge As String = "") As ADODB.Recordset
    On Error GoTo GetRepColumnaCtaSaldoErr
    
    Set dbConec = New DConecta
   
    If dbConec.AbreConexion() Then
   
      
        psSql = "SELECT  SUM(nCtaSaldoImporte) as Total, SUM(nCtaSaldoImporteME) AS TotalME  "
        psSql = psSql & " FROM    CTASALDO CS JOIN RepColumnaCta OC ON cs.cCtaContCod LIKE oc.cCtaContCod + '%' "
        psSql = psSql & " inner join ctacont c ON c.cCtaContCod = OC.cCtaContCod "
        psSql = psSql & " WHERE   OC.nNroCol=" & pnNroCol & " AND OC.cOpeCod = '" & psOpeCod & "' "
        'psSql = psSql & "    AND CS.NCTASALDOIMPORTE > 0 "
        psSql = psSql & "    AND CS.dCtaSaldoFecha IN (SELECT MAX(CS1.dCtaSaldoFecha) "
        psSql = psSql & "                              FROM CTASALDO CS1 "
        psSql = psSql & "                              WHERE CS1.cCtaContCod = CS.cCtaContCod "
        psSql = psSql & "                                    AND CS1.dCtaSaldoFecha <='" & Format(pdFecha, gsFormatoFecha) & "') "
      
        If Len(Trim(gsCodAge)) = 0 Then
        Else
            If gsCodAge = "99" Then
                psSql = psSql & " AND bAgencia=0"
            Else
                psSql = psSql & " AND right(c.cCtaContCod, 2)='" & gsCodAge & "' and bAgencia=1"
            End If
        End If
      
        Set GetRepColumnaCtaSaldo = dbConec.CargaRecordSet(psSql)
        
   End If
   
   dbConec.CierraConexion
   Set dbConec = Nothing
   
   Exit Function
GetRepColumnaCtaSaldoErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetRepColumnaCtaSaldo Method")
End Function

'ALPA 20080810***************************************************
Public Function GetRepColumnaCtaDesc(psOpeCod As String, pnNroCol As Integer, pdFecha As Date, Optional ByVal pbSinCtaCont As Boolean = False) As ADODB.Recordset
   On Error GoTo GetRepColumnaCtaDescErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
    If pbSinCtaCont Then
        psSql = "SELECT  CC.CCTACONTDESC, RCTA.cDescCol "
        psSql = psSql & "FROM    RepColumnaCta RC JOIN RepColumna RCTA ON rcta.cOpeCod = rc.cOpecod and rcta.nNroCol = rc.nNroCol, "
        psSql = psSql & "        CtaSaldo CS JOIN CtaCont CC ON cc.cCtaContCod = cs.cCtaContCod "
        psSql = psSql & "WHERE   RC.cOpeCod='" & psOpeCod & "' AND RC.nNroCol=" & pnNroCol & " AND "
        psSql = psSql & "        SUBSTRING(CS.cCtaContCod, 1, Len(RTrim(rc.cCtaContCod))) = RTrim(rc.cCtaContCod) "
        'psSql = psSql & "        AND DatedIff(d,'" & Format(pdFecha, gsFormatoFecha) & "',CS.dCtaSaldoFecha) > = -32 "
        psSql = psSql & " AND CS.nCtaSaldoImporte > 0 and datediff(d,CS.dCtaSaldoFecha,(select max(dCtaSaldoFecha) from CtaSaldo where cCtaContCod=cs.cCtaContCod))=0"
        psSql = psSql & " AND CS.nCtaSaldoImporte > 0 "
        psSql = psSql & "GROUP BY cc.cCtaContDesc, RCTA.cDescCol "
    Else
        psSql = "SELECT  cs.cCtaContCod, CC.CCTACONTDESC, RCTA.cDescCol "
        psSql = psSql & "FROM    RepColumnaCta RC JOIN RepColumna RCTA ON rcta.cOpeCod = rc.cOpecod and rcta.nNroCol = rc.nNroCol, "
        psSql = psSql & "        CtaSaldo CS JOIN CtaCont CC ON cc.cCtaContCod = cs.cCtaContCod "
        psSql = psSql & "WHERE   RC.cOpeCod='" & psOpeCod & "' AND RC.nNroCol=" & pnNroCol & " AND "
        psSql = psSql & "        SUBSTRING(CS.cCtaContCod, 1, Len(RTrim(rc.cCtaContCod))) = RTrim(rc.cCtaContCod) "
        'psSql = psSql & "        AND DatedIff(d,'" & Format(pdFecha, gsFormatoFecha) & "',CS.dCtaSaldoFecha) > = -32 "
        psSql = psSql & " AND CS.nCtaSaldoImporte > 0 and datediff(d,CS.dCtaSaldoFecha,(select max(dCtaSaldoFecha) from CtaSaldo where cCtaContCod=cs.cCtaContCod))=0"
        psSql = psSql & "        AND CS.nCtaSaldoImporte > 0 "
        psSql = psSql & "GROUP BY cs.cCtaContCod, cc.cCtaContDesc, RCTA.cDescCol ORDER BY CS.cCtaContCod"
    End If
    
    Set GetRepColumnaCtaDesc = dbConec.CargaRecordSet(psSql)
   End If
   Exit Function
GetRepColumnaCtaDescErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetRepColumnaCtaDesc Method")
End Function

Public Function GetRepColumnaCmacs(pnMoneda As Integer, ByVal pnBitCentral As Boolean) As ADODB.Recordset
   On Error GoTo GetRepColumnaCmacsErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      'psSql = "SELECT Distinct cif.cPersCod, p.cPersNombre, cif.cCtaIFDesc " _
      '      & "FROM   CtaIf cif JOIN Persona p ON p.cPersCod = cif.cPersCod " _
      '      & "                 JOIN ProductoPersona pp ON pp.cPersCod = cif.cPersCod " _
      '      & "WHERE cif.cIFTpo = '" & Format(gTpoIFCmac, "00") & "' and SubString(cif.cCtaIFCod,3,1) = '" & pnMoneda & "' "
      
      If pnBitCentral = True Then
      
            psSql = "Select cPersNombre cObjetoDesc From CtasCmacs Where substring(cctacod,9,1) =" & pnMoneda & " Group by cpersNombre Order by cPersNombre "
      Else
        
            psSql = "Select cNomPers cObjetoDesc From CtasCmacs Where substring(ccodcta,6,1) = '" & pnMoneda & "' Group by cNompers Order by cNomPers "
      
      End If
      
      Set GetRepColumnaCmacs = dbConec.CargaRecordSet(psSql)
   End If
   Exit Function
GetRepColumnaCmacsErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetRepColumnaCmacs Method")
End Function

Public Function GetRepColumnaCmacsCuentas(pnMoneda As Integer, ByVal pnBitCentral As Boolean) As ADODB.Recordset
Dim sSql As String
   On Error GoTo GetRepColumnaCmacsErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
        If pnBitCentral = True Then
            sSql = "Select cPersNombre as cNomPers, cCtaCod  as cCodcta From CtasCmacs Where substring(cCtaCod,9,1) = " & pnMoneda & " Group by cPersNombre,cCtaCod Order by cPersNombre,cCtaCod "
        Else
            sSql = "Select cNomPers, cCodcta From CtasCmacs Where substring(ccodcta,6,1) = '" & pnMoneda & "' Group by cNompers,cCodcta Order by cNomPers,cCodcta "
        End If
      Set GetRepColumnaCmacsCuentas = dbConec.CargaRecordSet(sSql)
   End If
   Exit Function
GetRepColumnaCmacsErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetRepColumnaCmacs Method")
End Function


Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
   Dim oIni As New ClasIni
   sCentralCom = oIni.BaseComunes
   Set oIni = Nothing
End Sub

'MIOL 20130528, SEGUN RQ13286, RQ13287 *********************************************
Public Sub InsertarConfGastoNivel(ByVal pnNivel As Integer, ByVal psConcepto As String, ByVal psFormula As String, ByVal psDescripcion As String, ByVal pnGlosa As Integer, ByVal psMovNro As String, ByVal psAnio As String)
   On Error GoTo InsertarConfGastoNivelErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "Exec stp_ins_ConfGastosNivel " & pnNivel & ",'" & psConcepto & "','" & psFormula & "','" & psDescripcion & "'," & pnGlosa & ",'" & psMovNro & "','" & psAnio & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertarConfGastoNivelErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:InsertarConfGastoNivel Method")
End Sub

Public Sub ActualizarConfGastoNivel(ByVal pnOrden As Integer, ByVal pnNivel As Integer, ByVal psConcepto As String, ByVal psFormula As String, ByVal psDescripcion As String, ByVal pnGlosa As Integer)
   On Error GoTo ActualizarConfGastoNivelErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "Exec stp_upd_ConfGastosNivelxOrden " & pnOrden & "," & pnNivel & ",'" & psConcepto & "','" & psFormula & "','" & psDescripcion & "'," & pnGlosa & ""
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizarConfGastoNivelErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:ActualizarConfGastoNivel Method")
End Sub

Public Sub ActualizaConfGastosNivelOrden(ByVal pnOrdenAct As Integer, ByVal pnNivelAct As Integer, ByVal psConceptoAct As String, ByVal psFormulaAct As String, ByVal psDescripcionAct As String, ByVal pnGlosaAct As Integer, ByVal psMovNroAct As String, _
                                         ByVal pnOrdenReemp As Integer, ByVal pnNivelReemp As Integer, ByVal psConceptoReemp As String, ByVal psFormulaReemp As String, ByVal psDescripcionReemp As String, ByVal pnGlosaReemp As Integer, ByVal psMovNroReemp As String)
   On Error GoTo ActualizarConfGastoNivelErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "Exec stp_upd_ActualizaConfGastosNivelOrden " & pnOrdenAct & "," & pnNivelAct & ",'" & psConceptoAct & "','" & psFormulaAct & "','" & psDescripcionAct & "'," & pnGlosaAct & ",'" & psMovNroAct & "'," & pnOrdenReemp & "," & pnNivelReemp & ",'" & psConceptoReemp & "','" & psFormulaReemp & "','" & psDescripcionReemp & "'," & pnGlosaReemp & ",'" & psMovNroReemp & "'"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizarConfGastoNivelErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:ActualizarConfGastoNivel Method")
End Sub

Public Sub EliminarConfGastoNivel(ByVal pnOrden As Integer)
   On Error GoTo EliminarConfGastoNivelErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "Exec stp_del_ConfGastosNivelxOrden " & pnOrden & ""
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminarConfGastoNivelErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:EliminarConfGastoNivel Method")
End Sub

Public Function ValidarConfGastoNivel(ByVal pnNivel As Integer, ByVal psConcepto As String, ByVal psFormula As String) As ADODB.Recordset
    On Error GoTo ErrValidarConfGastoNivel
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec stp_sel_ConfGastosNivelxValidacion " & pnNivel & ",'" & psConcepto & "','" & psFormula & "'"
        Set ValidarConfGastoNivel = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrValidarConfGastoNivel:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:ValidarConfGastoNivel Method")
End Function

Public Function GetConfGastoNivel(ByVal psAnio As String) As ADODB.Recordset
    On Error GoTo ErrGetConfGastoNivel
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec stp_sel_ConfGastosNivel '" & psAnio & "'"
        Set GetConfGastoNivel = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetConfGastoNivel:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetConfGastoNivel Method")
End Function

Public Function GetConfGastoNivelProyAnual(ByVal psAnio As String) As ADODB.Recordset
    On Error GoTo ErrGetConfGastoNivelProyAnual
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec stp_sel_ConfGastosNivelProyAnualGastoxAnio '" & psAnio & "'"
        Set GetConfGastoNivelProyAnual = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetConfGastoNivelProyAnual:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetConfGastoNivelProyAnual Method")
End Function

Public Function MaxOrdenConfGastosNivel(ByVal psAnio As String) As Integer
Dim prs As ADODB.Recordset
   On Error GoTo MaxOrdenConfGastosNivelErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "stp_sel_ObtenerUltimoOrdenConfGastosNivel '" & psAnio & "'"
      Set prs = dbConec.CargaRecordSet(psSql)
      If Not prs.EOF Then
         MaxOrdenConfGastosNivel = prs!nOrden
      Else
         MaxOrdenConfGastosNivel = 0
      End If
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
MaxOrdenConfGastosNivelErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:MaxOrdenConfGastosNivel Method")
End Function

Public Sub InsertarProyeccionAnual(ByVal pnOrden As Integer, ByVal pnEne As Currency, ByVal pnFeb As Currency, ByVal pnMar As Currency, ByVal pnAbr As Currency, _
                                    ByVal pnMay As Currency, ByVal pnJun As Currency, ByVal pnJul As Currency, ByVal pnAgo As Currency, ByVal pnSet As Currency, _
                                    ByVal pnOct As Currency, ByVal pnNov As Currency, ByVal pnDic As Currency)
   On Error GoTo InsertarProyeccionAnualErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "Exec stp_ins_ProyeccionAnualGasto " & pnOrden & "," & pnEne & "," & pnFeb & "," & pnMar & "," & pnAbr & "," & pnMay & "," & pnJun & "," & pnJul & "," & pnAgo & "," & pnSet & "," & pnOct & "," & pnNov & "," & pnDic & ""
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertarProyeccionAnualErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:InsertarProyeccionAnual Method")
End Sub

Public Sub ActualizarProyeccionAnual(ByVal pnOrden As Integer, ByVal pnEne As Currency, ByVal pnFeb As Currency, ByVal pnMar As Currency, ByVal pnAbr As Currency, _
                                    ByVal pnMay As Currency, ByVal pnJun As Currency, ByVal pnJul As Currency, ByVal pnAgo As Currency, ByVal pnSet As Currency, _
                                    ByVal pnOct As Currency, ByVal pnNov As Currency, ByVal pnDic As Currency)
   On Error GoTo ActualizarProyeccionAnualErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "Exec stp_upd_ProyeccionAnualGastoxOrden " & pnOrden & "," & pnEne & "," & pnFeb & "," & pnMar & "," & pnAbr & "," & pnMay & "," & pnJun & "," & pnJul & "," & pnAgo & "," & pnSet & "," & pnOct & "," & pnNov & "," & pnDic & ""
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizarProyeccionAnualErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:ActualizarProyeccionAnual Method")
End Sub

Public Function GetProyeccionAnualxOrden(ByVal pnOrden As Integer) As ADODB.Recordset
    On Error GoTo ErrGetProyeccionAnualxOrden
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec stp_sel_ProyeccionAnualGastoxOrden '" & pnOrden & "'"
        Set GetProyeccionAnualxOrden = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetProyeccionAnualxOrden:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetProyeccionAnualxOrden Method")
End Function
'END MIOL **********************************************************************************

'MIOL 20130528, SEGUN ERS071 - 2013*********************************************************
Public Function GetColocacionesMorxFec15A(ByVal psFec As String, ByVal pnTpo As Integer) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15A
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        If pnTpo = 1 Then
            psSql = "Exec stp_sel_ColMor15AVigentes '" & psFec & "'"
        ElseIf pnTpo = 2 Then
            psSql = "Exec stp_sel_ColMor15AAtrasados '" & psFec & "'"
        ElseIf pnTpo = 3 Then
            psSql = "Exec stp_sel_ColMor15ARefinanciados '" & psFec & "'"
        ElseIf pnTpo = "0" Then
            psSql = "Exec stp_sel_ColCuadroConsolidadoVigentesRefAtras '" & psFec & "'"
        End If 'NAGL 20170425
        Set GetColocacionesMorxFec15A = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15A:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15A Method")
End Function

Public Function GetColocacionesMorxFec15B(ByVal psFec As String, ByVal pnTpo As Integer) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15B
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        If pnTpo = 1 Then
            psSql = "Exec stp_sel_ColMor15BBruta '" & psFec & "'"
        ElseIf pnTpo = 2 Then
            psSql = "Exec stp_sel_ColMor15BRefinanciado '" & psFec & "'"
        ElseIf pnTpo = 3 Then
            psSql = "Exec stp_sel_ColMor15BAtrasada '" & psFec & "'"
        ElseIf pnTpo = 4 Then
            psSql = "Exec stp_sel_ColMor15BDeudores '" & psFec & "'"
        End If 'NAGL 20170425
        Set GetColocacionesMorxFec15B = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15B:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15B Method")
End Function

Public Function GetColocacionesMorxFec15C(ByVal psFec As String, ByVal pnTpo As Integer) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15C
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        If pnTpo = 1 Then
            psSql = "Exec stp_sel_ColMor15CBruta '" & psFec & "'"
        ElseIf pnTpo = 2 Then
            psSql = "Exec stp_sel_ColMor15CRefinanciada '" & psFec & "'"
        ElseIf pnTpo = 3 Then
            psSql = "Exec stp_sel_ColMor15CAtrasada '" & psFec & "'"
        End If
        Set GetColocacionesMorxFec15C = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15C:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15C Method")
End Function

Public Function GetColocacionesMorxFec15D(ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15D
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_ColMor15DBrutaRefAtrasada '" & psFec & "'"

        Set GetColocacionesMorxFec15D = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15D:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15D Method")
End Function

Public Function GetColocacionesMorxFec15E(ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15E
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_ColMor15EBrutaRefAtrasada '" & psFec & "'"

        Set GetColocacionesMorxFec15E = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15E:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15E Method")
End Function

Public Function GetColocacionesMorxFec15F(ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15F
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_ColMor15FBrutaRefAtrasada'" & psFec & "'"

        Set GetColocacionesMorxFec15F = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15F:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15F Method")
End Function

Public Function GetColocacionesMorxFec15G(ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15G
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_ColMor15GBrutaRefAtrasada'" & psFec & "'"

        Set GetColocacionesMorxFec15G = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15G:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15G Method")
End Function

Public Function GetColocacionesMorxFec15H(ByVal psFec As String, ByVal pnDiasIni As Long, ByVal pnDiasFin As Long) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15H
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_ColMor15HColMor '" & psFec & "'," & pnDiasIni & "," & pnDiasFin & ""

        Set GetColocacionesMorxFec15H = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15H:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15H Method")
End Function

Public Function GetColocacionesMorxFec15I(ByVal psFec As String, ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15I
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_ColMor15IColMor '" & psFec & "','" & psMoneda & "'"

        Set GetColocacionesMorxFec15I = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15I:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15I Method")
End Function

Public Function GetColocacionesMorxFec15J(ByVal psFec As String, ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15J
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
'        If psMoneda = "0" Then
'            psSql = "Exec stp_sel_ColMor15JColMorSum '" & psFec & "'"
'        Else
'            psSql = "Exec stp_sel_ColMor15JColMorProm '" & psFec & "','" & psMoneda & "'"
'        End If
        If psMoneda = "0" Then
            psSql = "Exec stp_sel_ColMor15JColMorSum '" & psFec & "'"
        ElseIf psMoneda = "3" Then
            psSql = "Exec stp_sel_ColMor15JColMorCant '" & psFec & "'"
        Else
            psSql = "Exec stp_sel_ColMor15JColMorProm '" & psFec & "','" & psMoneda & "'"
        End If 'NAGL
        Set GetColocacionesMorxFec15J = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15J:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15J Method")
End Function

Public Function GetColocacionesMorxFec15K(ByVal psFec As String, ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15K
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
'        If psMoneda = "0" Then
'            psSql = "Exec stp_sel_ColMor15KColMorSum '" & psFec & "'"
'        Else
'            psSql = "Exec stp_sel_ColMor15KColMorProm '" & psFec & "','" & psMoneda & "'"
'        End If
        If psMoneda = "0" Then
            psSql = "Exec stp_sel_ColMor15KColMorSum '" & psFec & "'"
        ElseIf psMoneda = "3" Then
            psSql = "Exec stp_sel_ColMor15KClasificarteraReprogxSituCont '" & psFec & "'"
        ElseIf psMoneda = "4" Then
            psSql = "Exec stp_sel_ColMor15KClasificarteraReprogxCatRiesgo '" & psFec & "'"
        Else
            psSql = "Exec stp_sel_ColMor15JColMorProm '" & psFec & "','" & psMoneda & "'"
        End If 'NAGL 20170425
 
        Set GetColocacionesMorxFec15K = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15K:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15K Method")
End Function

Public Function GetColocacionesMorxFec15L(ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetColocacionesMorxFec15L
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_ColMor15LBrutaRefAtrasadaxDepart '" & psFec & "'"

        Set GetColocacionesMorxFec15L = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetColocacionesMorxFec15L:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetColocacionesMorxFec15L Method")
End Function '***NAGL 20170425

Public Function GetCaptacionesMorxFec16A(ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetCaptacionesMorxFec16A
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
            psSql = "Exec stp_sel_ColMor16ACapMor '" & psFec & "'"
        
        Set GetCaptacionesMorxFec16A = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetCaptacionesMorxFec16A:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetCaptacionesMorxFec16A Method")
End Function

Public Function GetCaptacionesMorxFec16BxPF(ByVal psFec As String, ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo ErrGetCaptacionesMorxFec16BxPF
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
            
            psSql = "Exec stp_sel_ColMor16BCapMorPF '" & psFec & "','" & psMoneda & "'"
        
        Set GetCaptacionesMorxFec16BxPF = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetCaptacionesMorxFec16BxPF:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetCaptacionesMorxFec16BxPF Method")
End Function

Public Function GetCaptacionesMorxFec16BxAH(ByVal psFec As String, ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo ErrGetCaptacionesMorxFec16BxAH
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
            
            psSql = "Exec stp_sel_ColMor16BCapMorAH '" & psFec & "','" & psMoneda & "'"
        
        Set GetCaptacionesMorxFec16BxAH = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetCaptacionesMorxFec16BxAH:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetCaptacionesMorxFec16BxAH Method")
End Function

Public Function GetCaptacionesMorxFec16BxCTS(ByVal psFec As String, ByVal psMoneda As String) As ADODB.Recordset
    On Error GoTo ErrGetCaptacionesMorxFec16BxCTS
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
            
            psSql = "Exec stp_sel_ColMor16BCapMorCTS '" & psFec & "','" & psMoneda & "'"
        
        Set GetCaptacionesMorxFec16BxCTS = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetCaptacionesMorxFec16BxCTS:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetCaptacionesMorxFec16BxCTS Method")
End Function

Public Function GetCaptacionesMorxFec16C(ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetCaptacionesMorxFec16C
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
            
            psSql = "Exec stp_sel_ColMor16CCapMor '" & psFec & "'"
        
        Set GetCaptacionesMorxFec16C = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetCaptacionesMorxFec16C:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetCaptacionesMorxFec16C Method")
End Function

'MIOL 20130710, SEGUN RQ13331 *********************************************
Public Function GetAdeudados(ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetAdeudados
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_DetalleAdeudados '" & psFec & "'"

        Set GetAdeudados = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetAdeudados:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetAdeudados Method")
End Function
'END MIOL ******************************************************************
'END MIOL **********************************************************************************

'MIOL 20130712, SEGUN RQ13288 *********************************************
Public Function GetRepProyectados(ByVal pnMes As Integer, ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetRepProyectados
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_ReporteProyvsEjecMes " & pnMes & ",'" & psFec & "'"

        Set GetRepProyectados = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetRepProyectados:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetRepProyectados Method")
End Function

Public Function GetDatosProyEjec(ByVal pnMes As Integer, ByVal psFec As String) As ADODB.Recordset
    On Error GoTo ErrGetDatosProyEjec
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_CargaDatosProyvsEjecMes " & pnMes & ",'" & psFec & "'"

        Set GetDatosProyEjec = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetDatosProyEjec:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetDatosProyEjec Method")
End Function

Public Function GetMovGlosa(ByVal psFecha As String, ByVal psCta As String) As ADODB.Recordset
    On Error GoTo ErrGetMovGlosa
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then

            psSql = "Exec stp_sel_GlosaMovxCta '" & psFecha & "','" & psCta & "'"

        Set GetMovGlosa = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetMovGlosa:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetMovGlosa Method")
End Function
'END MIOL ******************************************************************
'MIOL 20130719, SEGUN ERS088-2013 - SEGUN RQ13387********************************
Public Function GetLimitePosCamb() As ADODB.Recordset
    On Error GoTo ErrGetLimitePosCamb
    Set dbConec = New DConecta
    If dbConec.AbreConexion() Then
        psSql = "Exec stp_sel_LimitePosCamb "
        Set GetLimitePosCamb = dbConec.CargaRecordSet(psSql)
        dbConec.CierraConexion
    End If
    Set dbConec = Nothing
    Exit Function
ErrGetLimitePosCamb:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:GetLimitePosCamb Method")
End Function

Public Sub ActualizarLimitePosCamb(ByVal pnReg As Integer, ByVal pnInt As Integer, ByVal pnTem As Integer, ByVal pnMontoSCReg As Currency, ByVal pnMontoSCInt As Currency, _
                                    ByVal pnMontoSCTem As Currency, ByVal pnMontoSVReg As Currency, ByVal pnMontoSVInt As Currency, ByVal pnMontoSVTem As Currency, ByVal psMovNro As String)
   On Error GoTo ActualizarLimitePosCambErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      dbConec.BeginTrans
      'Regulatorio
      psSql = "stp_upd_LimPosCamb " & pnReg & ",1," & pnMontoSCReg & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql

      psSql = "stp_upd_LimPosCamb " & pnReg & ",2," & pnMontoSVReg & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql
      'Interno
      psSql = "stp_upd_LimPosCamb " & pnInt & ",1," & pnMontoSCInt & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql

      psSql = "stp_upd_LimPosCamb " & pnInt & ",2," & pnMontoSVInt & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql
      'Temprano
      psSql = "stp_upd_LimPosCamb " & pnTem & ",1," & pnMontoSCTem & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql

      psSql = "stp_upd_LimPosCamb " & pnTem & ",2," & pnMontoSVTem & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql

      dbConec.CommitTrans
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizarLimitePosCambErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:ActualizarLimitePosCamb Method")
End Sub

Public Sub InsertarLimitePosCamb(ByVal pnReg As Integer, ByVal pnInt As Integer, ByVal pnTem As Integer, ByVal pnMontoSCReg As Currency, ByVal pnMontoSCInt As Currency, _
                                    ByVal pnMontoSCTem As Currency, ByVal pnMontoSVReg As Currency, ByVal pnMontoSVInt As Currency, ByVal pnMontoSVTem As Currency, ByVal psMovNro As String)
   On Error GoTo InsertarLimitePosCambErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      dbConec.BeginTrans
      'Regulatorio
      psSql = "stp_ins_LimPosCambHis " & pnReg & ",1," & pnMontoSCReg & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql

      psSql = "stp_ins_LimPosCambHis " & pnReg & ",2," & pnMontoSVReg & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql
      'Interno
      psSql = "stp_ins_LimPosCambHis " & pnInt & ",1," & pnMontoSCInt & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql

      psSql = "stp_ins_LimPosCambHis " & pnInt & ",2," & pnMontoSVInt & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql
      'Temprano
      psSql = "stp_ins_LimPosCambHis " & pnTem & ",1," & pnMontoSCTem & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql

      psSql = "stp_ins_LimPosCambHis " & pnTem & ",2," & pnMontoSVTem & ",'" & psMovNro & "'"
      dbConec.Ejecutar psSql

      dbConec.CommitTrans
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertarLimitePosCambErr:
   Call RaiseError(MyUnhandledError, "DRepCtaColumna:InsertarLimitePosCamb Method")
End Sub
'END MIOL ********************************************************
