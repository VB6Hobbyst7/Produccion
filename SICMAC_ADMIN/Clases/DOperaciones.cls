VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOperaciones"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetDocOpe(ByVal psOpeCod As String, Optional pbFormaFlex As Boolean = True) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      If pbFormaFlex Then
        sql = " Select 0, cDocDesc + space(50) + str(OD.nDocTpo), Null,'000','000000' From OpeDoc OD" _
              & " Inner Join Documento DO On OD.nDocTpo = DO.nDocTpo" _
              & " Where cOpeCod = '" & psOpeCod & "' Order By cOpeCod"
      Else
        sql = " Select cDocDesc + space(50) + str(OD.nDocTpo) From OpeDoc OD" _
              & " Inner Join Documento DO On OD.nDocTpo = DO.nDocTpo" _
              & " Where cOpeCod = '" & psOpeCod & "' Order By cOpeCod"
      End If
      Set GetDocOpe = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetOperaciones(ByVal psOpeCodCarb As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select cOpeCod,cOpeDesc, Len(cOpeCod) Nivel from opetpo where copecod like '" & psOpeCodCarb & "%'" _
          & " And cOpeVisible = '" & OpeTpoVisible.gOpeTpoVisibleSi & "'" _
          & " order by cOpeCod"
      Set GetOperaciones = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

'Public Function GetOrdenesCompra(ByVal psOpeCodCarb As String, Optional psMoneda As String = "") As ADODB.Recordset
'   On Error GoTo CargaOpeGruErr
'   Dim oCon As DConecta
'   Set oCon = New DConecta
'   Dim sql As String
'   Dim sqlAdd As String
'
'   Dim oConst As NConstSistemas
'   Set oConst = New NConstSistemas
'
'   Dim lnTope As Integer
'   Dim lsEjercicio As String
'
'   lnTope = oConst.LeeConstSistema(95)
'   lsEjercicio = Trim(oConst.LeeConstSistema(96))
'
'
'   If oCon.AbreConexion() Then
'      If psMoneda <> "" Then
'         sqlAdd = " Substring(cOpeCod,3,1) = '" & psMoneda & "' And "
'      Else
'        sqlAdd = ""
'      End If
'
'      'sql = " Select Distinct MD.cDocNro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + SubString(M.cOpeCod,3,1) , Len(MD.cDocNro) Nivel From Mov M" _
'          & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
'          & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
'          & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
'          & " Inner Join MovGasto MG On M.nMovNro = MG.nMovNro" _
'          & " Inner Join Persona PE On PE.cPersCod = MG.cPersCod" _
'          & " Where " & sqlAdd & " nMovEstado = '" & psOpeCodCarb & "' And nDocTpo = '90' And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "','7')" _
'          & " And NOT EXISTS (SELECT h.nMovNro FROM MovRef h with (nolock) JOIN Mov M1 with (nolock) ON m1.nMovNro = h.nMovNro WHERE m1.nMovEstado = '15' And m1.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "','7') And h.nMovNroRef = M.nMovNro)" _
'          & " Group By MD.cDocNro, M.nMovnro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + SubString(M.cOpeCod,3,1) , Len(MD.cDocNro)" _
'          & " having Sum(MCT.nMovCant) <> (  Select iSNULL(Sum(nMovCant),0) From MovRef MR with (nolock) " _
'          & "                                   Left Join Mov Movi with (nolock) On MR.nMovNro = Movi.nMovNro" _
'          & "                                   Left Join MovCant MCT with (nolock) On MR.nMovNro = MCT.nMovNro" _
'          & "                                   Left Join MovRef MRADD with (nolock) On MRADD.nMovNroRef = MR.nMovNro" _
'          & "                                   Left Join Mov MoviADD with (nolock) On MoviAdd.nMovNro = MRADD.nMovNro And MoviAdd.nMovEstado <> '21'" _
'          & "                               Where MR.nMovNroRef = m.nMovNro And Movi.cOpeCod Like '" & gnAlmaIngXCompras & "'" _
'          & "                                   And Movi.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "'))" _
'          & " Order By MD.cDocNro Desc"
'
'    sql = " Select Distinct Top " & lnTope & " cDocNro, Descrip, Nivel  From ( Select Sum(MCT.nMovCant) Sol, IsNull(AAA.Cant,0) Atendido,  MD.cDocNro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + SubString(M.cOpeCod,3,1) Descrip, Len(MD.cDocNro) Nivel" _
'        & " From Mov M WITH (NOLOCK) " _
'        & " Inner Join MovDoc MD WITH (NOLOCK) On M.nMovNro = MD.nMovNro" _
'        & " Inner Join MovBS MBS WITH (NOLOCK) On M.nMovNro = MBS.nMovNro" _
'        & " Inner Join MovCant MCT WITH (NOLOCK) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
'        & " Inner Join MovGasto MG WITH (NOLOCK) On M.nMovNro = MG.nMovNro" _
'        & " Inner Join Persona PE WITH (NOLOCK) On PE.cPersCod = MG.cPersCod" _
'        & " Left Join (Select MR.nMovNroRef nMovNro, MBS.cBSCod, ISNULL(Sum(nMovCant),0) Cant" _
'        & "            From MovRef MR WITH (NOLOCK) " _
'        & "            Inner Join Mov Movi WITH (NOLOCK) On MR.nMovNro = Movi.nMovNro" _
'        & "            Inner Join MovCant MCT WITH (NOLOCK) On MR.nMovNro = MCT.nMovNro" _
'        & "            Inner Join MovBS MBS WITH (NOLOCK) On MCT.nMovNro = MBS.nMovNro And MCT.nMovItem = MBS.nMovItem" _
'        & "            " _
'        & "            " _
'        & "            Where Movi.cOpeCod = '" & gnAlmaIngXCompras & "' And Movi.nMovFlag IN (" & gMovFlagVigente & "," & gMovFlagModificado & ") " _
'        & "            Group by MR.nMovNroRef, MBS.cBSCod) As AAA On AAA.nMovNro = MBS.nMovNro And AAA.cBSCod = MBS.cBSCod" _
'        & "  Where " & sqlAdd & " nMovEstado = " & psOpeCodCarb & " And nDocTpo = 33 And m.nMovFlag  = 0" _
'        & "  And NOT EXISTS (SELECT h.nMovNro FROM MovRef h WITH (NOLOCK) " _
'        & "                     JOIN Mov M1 WITH (NOLOCK) ON m1.nMovNro = h.nMovNro" _
'        & "                     WHERE m1.nMovEstado = 15 And m1.nMovFlag = " & gMovFlagVigente & " And h.nMovNroRef = M.nMovNro)" _
'        & "     Group By MD.cDocNro, M.nMovnro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + SubString(M.cOpeCod,3,1) , Len(MD.cDocNro), IsNull(AAA.Cant,0) " _
'        & ") As AAA Where Sol > Atendido Order By cDocNro Desc"
'
'        '--Inner Join MovRef MRADD WITH (NOLOCK) On MRADD.nMovNroRef = MR.nMovNro
'        '--Inner Join Mov MoviADD WITH (NOLOCK) On MoviAdd.nMovNro = MRADD.nMovNro And MoviAdd.cOpeCod = '" & gnAlmaIngXComprasConfirma & "'
'        '--And MoviAdd.nMovEstado <> 21
'
'      Set GetOrdenesCompra = oCon.CargaRecordSet(sql)
'      oCon.CierraConexion
'   End If
'   Set oCon = Nothing
'   Exit Function
'CargaOpeGruErr:
'   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
'End Function

'Public Function GetOrdenesCompra(ByVal psOpeCodCarb As String, Optional psMoneda As String = "") As ADODB.Recordset
'   On Error GoTo CargaOpeGruErr
'   Dim oCon As DConecta
'   Set oCon = New DConecta
'   Dim sql As String
'   Dim sqlAdd As String
'
'   If oCon.AbreConexion() Then
'      If psMoneda <> "" Then
'         sqlAdd = " Substring(cOpeCod,3,1) = '" & psMoneda & "' And "
'      Else
'        sqlAdd = ""
'      End If
'
'      sql = " Select Distinct MD.cDocNro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + SubString(M.cOpeCod,3,1) , Len(MD.cDocNro) Nivel From Mov M" _
'          & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
'          & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
'          & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
'          & " Inner Join MovGasto MG On M.nMovNro = MG.nMovNro" _
'          & " Inner Join Persona PE On PE.cPersCod = MG.cPersCod" _
'          & " Where " & sqlAdd & " nMovEstado = '" & psOpeCodCarb & "' And nDocTpo = '90' And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "','7')" _
'          & " And NOT EXISTS (SELECT h.nMovNro FROM MovRef h JOIN Mov M1 ON m1.nMovNro = h.nMovNro WHERE m1.nMovEstado = '15' And m1.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "','7') And h.nMovNroRef = M.nMovNro)" _
'          & " Group By MD.cDocNro, M.nMovnro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + SubString(M.cOpeCod,3,1) , Len(MD.cDocNro)" _
'          & " having Sum(MCT.nMovCant) <> (  Select iSNULL(Sum(nMovCant),0) From MovRef MR " _
'          & "                                   Left Join Mov Movi On MR.nMovNro = Movi.nMovNro" _
'          & "                                   Left Join MovCant MCT On MR.nMovNro = MCT.nMovNro" _
'          & "                                   Left Join MovRef MRADD On MRADD.nMovNroRef = MR.nMovNro" _
'          & "                                   Left Join Mov MoviADD On MoviAdd.nMovNro = MRADD.nMovNro And MoviAdd.nMovEstado <> '21'" _
'          & "                               Where MR.nMovNroRef = m.nMovNro And Movi.cOpeCod Like '" & gnAlmaIngXCompras & "'" _
'          & "                                   And Movi.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "'))" _
'          & " and  MD.cDocNro > '2005-00000103' " _
'          & " Order By MD.cDocNro Desc "
'      Set GetOrdenesCompra = oCon.CargaRecordSet(sql)
'      oCon.CierraConexion
'   End If
'   Set oCon = Nothing
'   Exit Function
'CargaOpeGruErr:
'   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
'End Function

Public Function GetOrdenesCompra(ByVal psOpeCodCarb As String, Optional psMoneda As String = "") As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim sqlAdd As String
   
   If oCon.AbreConexion() Then
      If psMoneda <> "" Then
         sqlAdd = " Substring(cOpeCod,3,1) = '" & psMoneda & "' And "
      Else
        sqlAdd = ""
      End If
             
      sql = " Select Distinct MD.cDocNro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then Str((MC.nMovImporte*-1),10,2) + ' (SOLES)  ' Else Str((MC.nMovImporte*-1),10,2) + ' (DOLARES)  ' End , Len(MD.cDocNro) Nivel From Mov M" _
          & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
          & " Inner Join MovCta MC On M.nMovNro = MC.nMovNro And MC.nMovImporte < 0" _
          & " Left Join MovME ME On MC.nMovNro = ME.nMovNro And ME.nMovMEImporte < 0" _
          & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
          & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
          & " Inner Join MovGasto MG On M.nMovNro = MG.nMovNro" _
          & " Inner Join Persona PE On PE.cPersCod = MG.cPersCod" _
          & " Where " & sqlAdd & " nMovEstado = '" & psOpeCodCarb & "' And nDocTpo = '90' And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "','7') And Left(cCtacontcod,2) = '25'" _
          & " And NOT EXISTS (SELECT h.nMovNro FROM MovRef h JOIN Mov M1 ON m1.nMovNro = h.nMovNro WHERE m1.nMovEstado = '15' And m1.nMovFlag NOT IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "','7') And h.nMovNroRef = M.nMovNro)" _
          & " Group By MD.cDocNro, M.nMovnro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then Str((MC.nMovImporte*-1),10,2) + ' (SOLES)  ' Else Str((MC.nMovImporte*-1),10,2) + ' (DOLARES)  ' End , Len(MD.cDocNro)" _
          & " having Sum(MCT.nMovCant) <> (  Select iSNULL(Sum(nMovCant),0) From MovRef MR " _
          & "                                   Left Join Mov Movi On MR.nMovNro = Movi.nMovNro" _
          & "                                   Left Join MovCant MCT On MR.nMovNro = MCT.nMovNro" _
          & "                                   Left Join MovRef MRADD On MRADD.nMovNroRef = MR.nMovNro" _
          & "                                   Left Join Mov MoviADD On MoviAdd.nMovNro = MRADD.nMovNro And MoviAdd.nMovEstado <> '21'" _
          & "                               Where MR.nMovNroRef = m.nMovNro And Movi.cOpeCod Like '" & gnAlmaIngXCompras & "'" _
          & "                                   And Movi.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "'))" _
          & " and  MD.cDocNro > '2005-00000103' " _
          & " Order By MD.cDocNro Desc "
      Set GetOrdenesCompra = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
'EJVG20111110 ****************************************************
Public Function listarOrdenesCompra(ByVal psOpeCodCarb As Integer, ByVal psAgeCod As String) As ADODB.Recordset
On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String

   If oCon.AbreConexion() Then
      sql = "exec stp_sel_ListaOCxAgencia " & psOpeCodCarb & ",'" & psAgeCod & "'"
      Set listarOrdenesCompra = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
'END *************************************************************
'Public Function GetNotaIngreso(ByVal psMovEstado As String, psOpeCod As String, Optional psDocTpo As String = "42") As ADODB.Recordset
'   On Error GoTo CargaOpeGruErr
'   Dim OCON As DConecta
'   Set OCON = New DConecta
'   Dim sql As String
'
'   Dim oConst As NConstSistemas
'   Set oConst = New NConstSistemas
'
'   Dim lnTope As Integer
'   Dim lsEjercicio As String
'
'   lnTope = oConst.LeeConstSistema(95)
'   lsEjercicio = Trim(oConst.LeeConstSistema(96))
'
'   If OCON.AbreConexion() Then
'      sql = " Select Distinct Top " & lnTope & " MD.cDocNro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + '[' + Ltrim(Str(M.nMovNro)) + ']' + SubString(M.cOpeCod,3,1), Len(MD.cDocNro) Nivel " _
'          & " From Mov M WITH (NOLOCK) " _
'          & " Inner Join MovDoc MD WITH (NOLOCK) On M.nMovNro = MD.nMovNro" _
'          & " " _
'          & " " _
'          & " " _
'          & " " _
'          & " " _
'          & " " _
'          & " Left Join MovGasto MG WITH (NOLOCK) On M.nMovNro = MG.nMovNro" _
'          & " Left Join Persona PE WITH (NOLOCK) On PE.cPersCod = MG.cPersCod" _
'          & " Where MD.cDocNro > '" & lsEjercicio & "' And  SubString(M.cOpeCod,3,1) = '" & Mid(psOpeCod, 3, 1) & "' And nMovEstado = " & psMovEstado & " And nDocTpo = " & psDocTpo & " " _
'          & " And M.nMovFlag =  " & gMovFlagVigente & " Order By MD.cDocNro Desc"
'      ' And MD.dDocFecha >= (Select Max(dSaldo) From BSSaldos)
'      Set GetNotaIngreso = OCON.CargaRecordSet(sql)
'      OCON.CierraConexion
'   End If
'   Set OCON = Nothing
'   Exit Function
'CargaOpeGruErr:
'   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
'End Function
'EJVG20140320 Se agregó pnAgeCod
Public Function GetNotaIngreso(ByVal psMovEstado As String, psOpeCod As String, Optional psDocTpo As String = "42", Optional ByVal pnAgeCod As Integer = 0) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select Distinct MD.cDocNro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + '[' + Ltrim(Str(M.nMovNro)) + ']' + SubString(M.cOpeCod,3,1), Len(MD.cDocNro) Nivel " _
          & " From Mov M" _
          & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
          & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro " & IIf(pnAgeCod > 0, " And MBS.nMovBSOrden = " & pnAgeCod, "") _
          & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
          & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
          & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
          & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
          & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
          & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
          & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
          & " Where SubString(M.cOpeCod,3,1) like '" & Mid(psOpeCod, 3, 1) & "'   And nMovEstado = '" & psMovEstado & "' And nDocTpo = '" & psDocTpo & "'" _
          & " And M.nMovNro In (Select Max(MMD.nMovNro) From MovDoc MMD Where MMD.cDocNro = MD.cDocNro And nDocTpo = '" & psDocTpo & "') " _
          & " And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','5')  Order By MD.cDocNro Desc"
      ' And MD.dDocFecha >= (Select Max(dSaldo) From BSSaldos)
      Set GetNotaIngreso = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function


Public Function GetNotaIngresoReporte(ByVal psMovEstado As String, psOpeCod As String, Optional psDocTpo As String = "42", Optional psDocNro As String = "") As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select Distinct MD.cDocNro, IsNull(PE.cPersNombre,'<SIN PERSONA>') + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + '[' + Ltrim(Str(M.nMovNro)) + ']' + SubString(M.cOpeCod,3,1), Len(MD.cDocNro) Nivel " _
          & " From Mov M" _
          & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
          & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
          & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
          & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
          & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
          & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
          & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
          & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
          & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
          & " Where SubString(M.cOpeCod,3,1) like '" & Mid(psOpeCod, 3, 1) & "' And nMovEstado In ('" & psMovEstado & "') And nDocTpo = '" & psDocTpo & "'" _
          & " And MD.cDocNro = '" & psDocNro & "' And M.nMovNro = "
      sql = sql & "(  Select Max (M.nMovNro) " _
          & " From Mov M" _
          & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
          & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
          & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
          & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
          & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
          & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
          & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
          & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
          & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
          & " Where SubString(M.cOpeCod,3,1) like '" & Mid(psOpeCod, 3, 1) & "' And nMovEstado In ('" & psMovEstado & "') And nDocTpo = '" & psDocTpo & "'" _
          & " And MD.cDocNro = '" & psDocNro & "' )"
          
      Set GetNotaIngresoReporte = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetNotaSalidaReporte(ByVal psMovEstado As String, psOpeCod As String, Optional psDocTpo As String = "42", Optional psDocNro As String = "") As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select Distinct MD.cDocNro, PE.cPersNombre + Space(50) + Case SubString(M.cOpeCod,3,1) When '1' Then '(SOLES)    ' Else '(DOLARES)  ' End + '[' + Ltrim(Str(M.nMovNro)) + ']' + SubString(M.cOpeCod,3,1), Len(MD.cDocNro) Nivel " _
          & " From Mov M with (nolock) " _
          & " Inner Join MovDoc MD with (nolock) On M.nMovNro = MD.nMovNro" _
          & " Inner Join MovBS MBS with (nolock) On M.nMovNro = MBS.nMovNro" _
          & " Left  Join MovBSSerie MBSS with (nolock) On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
          & " Inner Join MovCta MCTA with (nolock) On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
          & " Inner Join MovCant MCT with (nolock) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
          & " Inner Join BienesServicios BS with (nolock) On BS.cBSCod  = MBS.cBSCod" _
          & " Inner Join Constante CO with (nolock) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
          & " Left Join MovGasto MG with (nolock) On M.nMovNro = MG.nMovNro" _
          & " Left Join Persona PE with (nolock) On PE.cPersCod = MG.cPersCod" _
          & " Where SubString(M.cOpeCod,3,1) like '" & Mid(psOpeCod, 3, 1) & "'   And nMovEstado = '" & psMovEstado & "' And nDocTpo = '" & psDocTpo & "'" _
          & " And MD.cDocNro = '" & psDocNro & "' "
      Set GetNotaSalidaReporte = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

'Public Function GetGuiaInterna(ByVal psMovEstado As String, psOpeCod As String, Optional psDocTpo As String = "42") As ADODB.Recordset
'   On Error GoTo CargaOpeGruErr
'   Dim oCon As DConecta
'   Set oCon = New DConecta
'   Dim SQL As String
'
'   Dim oConst As NConstSistemas
'   Set oConst = New NConstSistemas
'
'   Dim lnTope As Integer
'   Dim lsEjercicio As String
'
'   lnTope = oConst.LeeConstSistema(95)
'   lsEjercicio = Trim(oConst.LeeConstSistema(96))
'
'   If oCon.AbreConexion() Then
'        'sql = " Select Distinct MD.cDocNro, IsNull(PE.cPersNombre,'<SIN PERSONA>') + '[' + Ltrim(Str(M.nMovNro)) + ']', Len(MD.cDocNro) Nivel " _
'            & " From Mov M" _
'            & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
'            & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
'            & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
'            & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
'            & " Inner Join MovObjAreaAgencia MOAA On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem" _
'            & " Inner Join Areas ARE On MOAA.cAreaCod = ARE.cAreaCod" _
'            & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
'            & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
'            & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
'            & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
'            & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
'            & " Left Join (Select IsNull(Sum(nMovCant),0) Atendidos, MRMBS.cBSCod, MR.nMovNroRef nMovNro From MovRef MR" _
'            & "             Inner Join MovCant MRMCT On MR.nMovNro = MRMCT.nMovNro" _
'            & "             Inner Join Mov MRM On MRM.nMovNro = MR.nMovNro" _
'            & "             Inner Join MovBS MRMBS On MRMCT.nMovNro = MRMBS.nMovNro And MRMCT.nMovItem = MRMBS.nMovItem" _
'            & "           Where MRM.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') And MRM.cOpeCod like '" & gnAlmaSalXAtencion & "'" _
'            & "           Group By MRMBS.cBSCod,  MR.nMovNroRef ) MREF On MREF.nMovNro = M.nMovNro And MREF.cBSCod = MBS.cBSCod" _
'            & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
'            & " Where nDocTpo = '70' And ((M.nMovEstado = '13' And MREF.Atendidos Is Null) Or (M.nMovEstado = '5' And ISNull(MREF.Atendidos,0) < MCT.nMovCant))  And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') And M.nMovNro Not In ( Select nMovNroRef From MovRef MR Inner Join Mov M On MR.nMovNro = M.nMovNro Where M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')  And cOpeCod in ('" & gnAlmaReqAreaReg & "','" & gnAlmaReqAreaMant & "')) Order by MD.cDocNro Desc "
'
'        SQL = " Select Distinct Top " & lnTope & " MD.cDocNro, IsNull(PE.cPersNombre,'<SIN PERSONA>') + '[' + Ltrim(Str(M.nMovNro)) + ']', Len(MD.cDocNro) Nivel " _
'            & " From Mov M with (nolock) " _
'            & " Inner Join MovDoc MD with (nolock) On M.nMovNro = MD.nMovNro" _
'            & " Inner Join MovBS MBS with (nolock) On M.nMovNro = MBS.nMovNro" _
'            & " Inner Join MovCant MCT with (nolock) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
'            & " Left Join MovGasto MG with (nolock) On M.nMovNro = MG.nMovNro" _
'            & " Left Join Persona PE with (nolock) On PE.cPersCod = MG.cPersCod" _
'            & " Left Join (Select IsNull(Sum(nMovCant),0) Atendidos, MRMBS.cBSCod, MR.nMovNroRef nMovNro From MovRef MR" _
'            & "             Inner Join MovCant MRMCT On MR.nMovNro = MRMCT.nMovNro" _
'            & "             Inner Join Mov MRM On MRM.nMovNro = MR.nMovNro" _
'            & "             Inner Join MovBS MRMBS On MRMCT.nMovNro = MRMBS.nMovNro And MRMCT.nMovItem = MRMBS.nMovItem" _
'            & "           Where  MRM.cMovNro > '" & lsEjercicio & "' And MRM.nMovFlag = " & gMovFlagVigente & " And MRM.cOpeCod = '" & gnAlmaSalXAtencion & "'" _
'            & "           Group By MRMBS.cBSCod,  MR.nMovNroRef ) MREF On MREF.nMovNro = M.nMovNro And MREF.cBSCod = MBS.cBSCod" _
'            & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
'            & " Where MD.cDocNro > '" & lsEjercicio & "'And  nDocTpo = " & TpoDocAlmacenRequerimiento & " And ((M.nMovEstado = " & gMovEstContabNoContable & " And MREF.Atendidos Is Null) Or (M.nMovEstado = " & gMovFlagModificado & " And ISNull(MREF.Atendidos,0) < MCT.nMovCant))  And M.nMovFlag = " & gMovFlagVigente & " And M.nMovNro Not In ( Select nMovNroRef From MovRef MR Inner Join Mov M On MR.nMovNro = M.nMovNro Where M.nMovFlag = " & gMovFlagVigente & " And cOpeCod in ('" & gnAlmaReqAreaReg & "','" & gnAlmaReqAreaMant & "')) Order by MD.cDocNro Desc "
'
'      Set GetGuiaInterna = oCon.CargaRecordSet(SQL)
'      oCon.CierraConexion
'   End If
'   Set oCon = Nothing
'   Exit Function
'CargaOpeGruErr:
'   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
'End Function


Public Function GetGuiaInterna(ByVal psMovEstado As String, psOpeCod As String, psDocTpo As String, pnAlmacen As Integer) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
        sql = " Select Distinct MD.cDocNro, IsNull(PE.cPersNombre,'<SIN PERSONA>') + '[' + Ltrim(Str(M.nMovNro)) + ']', Len(MD.cDocNro) Nivel " _
            & " From Mov M" _
            & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
            & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro " & IIf(pnAlmacen = -1, "", " And MBS.nMovBSOrden = '" & pnAlmacen & "'") _
            & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
            & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
            & " Inner Join MovObjAreaAgencia MOAA On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem" _
            & " Inner Join Areas ARE On MOAA.cAreaCod = ARE.cAreaCod" _
            & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
            & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
            & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
            & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
            & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
            & " Left Join (Select IsNull(Sum(nMovCant),0) Atendidos, MRMBS.cBSCod, MR.nMovNroRef nMovNro From MovRef MR" _
            & "             Inner Join MovCant MRMCT On MR.nMovNro = MRMCT.nMovNro" _
            & "             Inner Join Mov MRM On MRM.nMovNro = MR.nMovNro" _
            & "             Inner Join MovBS MRMBS On MRMCT.nMovNro = MRMBS.nMovNro And MRMCT.nMovItem = MRMBS.nMovItem" _
            & "           Where MRM.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') And MRM.cOpeCod like '" & gnAlmaSalXAtencion & "'" _
            & "           Group By MRMBS.cBSCod,  MR.nMovNroRef ) MREF On MREF.nMovNro = M.nMovNro And MREF.cBSCod = MBS.cBSCod" _
            & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
            & " Where nDocTpo = '70' And ((M.nMovEstado = '13' And MREF.Atendidos Is Null) Or (M.nMovEstado = '5' And ISNull(MREF.Atendidos,0) < MCT.nMovCant))  And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "') And M.nMovNro Not In ( Select nMovNroRef From MovRef MR Inner Join Mov M On MR.nMovNro = M.nMovNro Where M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagModificado & "')  And cOpeCod in ('" & gnAlmaReqAreaReg & "','" & gnAlmaReqAreaMant & "')) Order by MD.cDocNro Desc "
      Set GetGuiaInterna = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function



'Public Function GetGuiaSalida(ByVal psMovEstado As String, psOpeCod As String, Optional psDocTpo As String = "42") As ADODB.Recordset
'   On Error GoTo CargaOpeGruErr
'   Dim OCON As DConecta
'   Set OCON = New DConecta
'   Dim sql As String
'
'   Dim oConst As NConstSistemas
'   Set oConst = New NConstSistemas
'
'   Dim lnTope As Integer
'   Dim lsEjercicio As String
'
'   lnTope = oConst.LeeConstSistema(95) * 2
'   lsEjercicio = Trim(oConst.LeeConstSistema(96))
'
'   If OCON.AbreConexion() Then
'        sql = " Select Distinct Top " & lnTope & " MD.cDocNro, IsNull(PE.cPersNombre,'<SIN PERSONA>') + '[' + Ltrim(Str(M.nMovNro)) + ']', Len(MD.cDocNro) Nivel " _
'            & " From Mov M with (nolock) " _
'            & " Inner Join MovDoc MD with (nolock) On M.nMovNro = MD.nMovNro" _
'            & " Left Join MovGasto MG with (nolock) On M.nMovNro = MG.nMovNro" _
'            & " Left Join Persona PE with (nolock) On PE.cPersCod = MG.cPersCod" _
'            & " " _
'            & " Where nDocTpo = " & TpoDocAlmacenGuiadeRemision & " And M.nMovEstado In (" & gMovEstLogSaleBienAlmacen & "," & gMovEstContabMovContable & ")  And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "')  Order by MD.cDocNro Desc"
'      Set GetGuiaSalida = OCON.CargaRecordSet(sql)
'      OCON.CierraConexion
'   End If
'   Set OCON = Nothing
'   Exit Function
'CargaOpeGruErr:
'   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
'End Function

Public Function GetGuiaSalida(ByVal psMovEstado As String, psOpeCod As String, Optional psDocTpo As String = "42") As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
        sql = " Select Distinct MD.cDocNro, IsNull(PE.cPersNombre,'<SIN PERSONA>') + '[' + Ltrim(Str(M.nMovNro)) + ']', Len(MD.cDocNro) Nivel " _
            & " From Mov M" _
            & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
            & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
            & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
            & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
            & " Inner Join MovObjAreaAgencia MOAA On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem" _
            & " Inner Join Areas ARE On MOAA.cAreaCod = ARE.cAreaCod" _
            & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
            & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
            & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
            & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
            & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
            & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
            & " Where nDocTpo = " & TpoDocAlmacenGuiadeRemision & " And M.nMovEstado In (" & gMovEstLogSaleBienAlmacen & "," & gMovEstContabMovContable & ")  And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "')  Order by MD.cDocNro Desc"
            'And MD.dDocFecha >= (Select Max(dSaldo) From BSSaldos)
      Set GetGuiaSalida = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function


'Public Function GetGuiaSalidaDet(ByVal psOpeCodCarb As String, Optional pbCantidad As Boolean = False, Optional psAlmacen As String = 1) As ADODB.Recordset
'   On Error GoTo CargaOpeGruErr
'   Dim OCON As DConecta
'   Set OCON = New DConecta
'   Dim sql As String
'
'   If OCON.AbreConexion() Then
'        If Not pbCantidad Then
'          sql = " Select M.nMovNro, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion, MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant," _
'              & "        MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, IsNull(MBSS.nIGV,0) nIGV, IsNull(MBSS.nValor,0) nValor , M.cOpeCod, OPE.cOpeDesc, MBS.nMovBSOrden, ARE.cAreaDescripcion, MOAA.cAreaCod, MOAA.cAgeCod, 0 Atendidos, OO.cObjetoCod, OO.cObjetoDesc , dbo.LogAFGetCuentaObj(BS.cBSCod,M.cOpeCod,OO.cObjetoCod) cCtaContCodObj " _
'              & " From Mov M With (Nolock) " _
'              & " Inner Join MovDoc MD with (nolock) On M.nMovNro = MD.nMovNro" _
'              & " Inner Join MovBS MBS with (nolock) On M.nMovNro = MBS.nMovNro" _
'              & " Left  Join MovBSSerie MBSS with (nolock) On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
'              & " Inner Join MovCta MCTA with (nolock) On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
'              & " Inner Join MovObjAreaAgencia MOAA with (nolock) On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem" _
'              & " Inner Join Areas ARE with (nolock) On MOAA.cAreaCod = ARE.cAreaCod" _
'              & " Inner Join MovCant MCT with (nolock) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
'              & " Inner Join BienesServicios BS with (nolock) On BS.cBSCod  = MBS.cBSCod" _
'              & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
'              & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
'              & " Left Join MovObj   MObj On MCTA.nMovNro = MObj.nMovNro And MCTA.nMovItem = MObj.nMovItem And MObj.nMovObjOrden = 2 " _
'              & " Left Join Objeto   OO   On OO.cObjetoCod = MObj.cObjetoCod " _
'              & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
'              & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
'              & " Where nDocTpo = 71 And M.nMovEstado = 10 And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And cDocNro = '" & psOpeCodCarb & "' Order by MCT.nMovItem"
'      Else
'          sql = " Select M.nMovNro, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod," _
'              & " Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' +" _
'              & " Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion," _
'              & " MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant, MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, M.cOpeCod," _
'              & " OPE.cOpeDesc , MBS.nMovBSOrden, ARE.cAreaDescripcion, MOAA.cAreaCod, MOAA.cAgeCod, 0 Atendidos, dbo.LogGetStock( " & psAlmacen & ",BS.cBSCod,(Select nProdTpo from LogOpeProd Where cOpeCod = M.cOpeCod)) Stock, IsNull(dbo.LogGetPrecioPromedio(-1,BS.cBSCod,(Select nProdTpo from LogOpeProd Where cOpeCod = M.cOpeCod)),0) PrePromedio " _
'              & " From Mov M with (nolock) " _
'              & " Inner Join MovDoc MD with (nolock) On M.nMovNro = MD.nMovNro" _
'              & " Inner Join MovBS MBS with (nolock) On M.nMovNro = MBS.nMovNro" _
'              & " Left  Join MovBSSerie MBSS with (nolock) On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
'              & " Inner Join MovCta MCTA with (nolock) On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
'              & " Inner Join MovObjAreaAgencia MOAA with (nolock) On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem" _
'              & " Inner Join Areas ARE with (nolock) On MOAA.cAreaCod = ARE.cAreaCod" _
'              & " Inner Join MovCant MCT with (nolock) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
'              & " Inner Join BienesServicios BS with (nolock) On BS.cBSCod  = MBS.cBSCod" _
'              & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
'              & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
'              & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
'              & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
'              & " Where nDocTpo = '71' And M.nMovEstado = '10'  And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "') And cDocNro = '" & psOpeCodCarb & "' Order By MCT.nMovItem"
'      End If
'
'      Set GetGuiaSalidaDet = OCON.CargaRecordSet(sql)
'      OCON.CierraConexion
'   End If
'   Set OCON = Nothing
'   Exit Function
'CargaOpeGruErr:
'   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
'End Function

Public Function GetGuiaSalidaDet(ByVal psOpeCodCarb As String, Optional pbCantidad As Boolean = False, Optional psAlmacen As String = 1) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
        If Not pbCantidad Then
          sql = " Select M.nMovNro, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion, MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant," _
              & "        MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, IsNull(MBSS.nIGV,0) nIGV, IsNull(MBSS.nValor,0) nValor , M.cOpeCod, OPE.cOpeDesc, MBS.nMovBSOrden, ARE.cAreaDescripcion, MOAA.cAreaCod, MOAA.cAgeCod, 0 Atendidos" _
              & " From Mov M" _
              & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
              & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
              & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
              & " Inner Join MovObjAreaAgencia MOAA On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem" _
              & " Inner Join Areas ARE On MOAA.cAreaCod = ARE.cAreaCod" _
              & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
              & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
              & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
              & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
              & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
              & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
              & " Where nDocTpo = '71' And M.nMovEstado = '10'  And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "') And cDocNro = '" & psOpeCodCarb & "'"
      Else
          sql = " Select M.nMovNro, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod," _
              & " Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' +" _
              & " Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion," _
              & " MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant, MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, M.cOpeCod," _
              & " OPE.cOpeDesc , MBS.nMovBSOrden, ARE.cAreaDescripcion, MOAA.cAreaCod, MOAA.cAgeCod, 0 Atendidos, dbo.LogGetStock( " & psAlmacen & ",BS.cBSCod,(Select nProdTpo from LogOpeProd Where cOpeCod = M.cOpeCod)) Stock, IsNull(dbo.LogGetPrecioPromedio(-1,BS.cBSCod,(Select nProdTpo from LogOpeProd Where cOpeCod = M.cOpeCod)),0) PrePromedio " _
              & " From Mov M" _
              & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
              & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
              & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
              & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
              & " Inner Join MovObjAreaAgencia MOAA On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem" _
              & " Inner Join Areas ARE On MOAA.cAreaCod = ARE.cAreaCod" _
              & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
              & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
              & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
              & " Left Join MovGasto MG On M.nMovNro = MG.nMovNro" _
              & " Left Join Persona PE On PE.cPersCod = MG.cPersCod" _
              & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
              & " Where nDocTpo = '71' And M.nMovEstado = '10'  And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "') And cDocNro = '" & psOpeCodCarb & "'"
      End If
      
      Set GetGuiaSalidaDet = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
Public Function GetAlmacenes() As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select nAlmCod,cAlmDescripcion,3 Nivel From Almacen"
      Set GetAlmacenes = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetAlmacenesUsuarioAlmacen(ByVal psAlmacen As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select nAlmCod,cAlmDescripcion,3 Nivel From Almacen where nAlmCod= '" & psAlmacen & " '"
      Set GetAlmacenesUsuarioAlmacen = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetCtaCntBS(psBSCod As String, psOpeCod As String, psAlmacen As String) As String
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   If oCon.AbreConexion() Then
      'sql = " Select cCtaContCod, Left('110201', nCtaObjNiv*2) from CtaBS Where Left('" & psBSCod & "', nCtaObjNiv*2) Like cObjetoCod"
      sql = "SELECT Replace(cCtaContCod,'AG','" & psAlmacen & "') FROM CTABS WITH (NOLOCK)  WHERE '" & psBSCod & "' LIKE COBJETOCOD + '%' And cOpeCod = '" & psOpeCod & "' order by len(COBJETOCOD) desc "
      Set rs = oCon.CargaRecordSet(sql)
      
      If rs.EOF And rs.BOF Then
        GetCtaCntBS = ""
      Else
        GetCtaCntBS = rs.Fields(0)
      End If
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function VerfBSSerie(psBSCod As String, psSerie As String, Optional psEstado As String = "") As Boolean
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   If oCon.AbreConexion() Then
      If psEstado = "" Then
        sql = " Select cBSCod From BSSerie WITH (NOLOCK) Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
      Else
        sql = " Select cBSCod From BSSerie WITH (NOLOCK) Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "' And nBienEstado = '" & psEstado & "'"
      End If
      Set rs = oCon.CargaRecordSet(sql)
      
      If rs.EOF And rs.BOF Then
        VerfBSSerie = False
      Else
        VerfBSSerie = True
      End If
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function VerfBSSerieMov(psBSCod As String, psSerie As String, pnMovNro As Long) As Boolean
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   If oCon.AbreConexion() Then
      sql = " Select nMovNro from movbsserie WITH (NOLOCK) where nMovNro = '" & pnMovNro & "' And cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
      Set rs = oCon.CargaRecordSet(sql)
      
      If rs.EOF And rs.BOF Then
        VerfBSSerieMov = False
      Else
        VerfBSSerieMov = True
      End If
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetDetOC(ByVal psOrdenCompra As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      sql = " Select M.nMovNro, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod, RTRIM(mcd.cDescrip) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion, MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant nMovCantT ,  MCT.nMovCant - " _
          & " ( Select IsNull(Sum(nMovCant),0) From MovRef MR WITH (NOLOCK) " _
          & "       Left Join Mov Movi WITH (NOLOCK) On MR.nMovNro = Movi.nMovNro" _
          & "       Left Join MovCant MCT WITH (NOLOCK) On MR.nMovNro = MCT.nMovNro" _
          & "       " _
          & "       " _
          & "   Where MR.nMovNroRef = M.nMovNro And MCT.nMovItem = MCTA.nMovItem  And Movi.cOpeCod = '" & gnAlmaIngXCompras & "' And Movi.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "') )  nMovCant ," _
          & "        MD.nDocTpo , MD.cDocNro, MD.dDocFecha" _
          & " From Mov M WITH (NOLOCK) " _
          & " Inner Join MovDoc MD WITH (NOLOCK) On M.nMovNro = MD.nMovNro" _
          & " Inner Join MovBS MBS WITH (NOLOCK) On M.nMovNro = MBS.nMovNro" _
          & " Left  Join MovBSSerie MBSS WITH (NOLOCK) On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
          & " Inner Join MovCta MCTA WITH (NOLOCK) On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
          & " Inner Join MovCant MCT WITH (NOLOCK) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
          & " Inner Join BienesServicios BS WITH (NOLOCK) On BS.cBSCod  = MBS.cBSCod" _
          & " Inner Join Constante CO WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
          & " Inner Join MovGasto MG WITH (NOLOCK) On M.nMovNro = MG.nMovNro" _
          & " Inner Join Persona PE WITH (NOLOCK) On PE.cPersCod = MG.cPersCod" _
          & " inner join MovCotizacDet mcd WITH (NOLOCK) ON mcd.nmovnro=MCTA.nmovnro and mcd.nMovItem = MCTA.nMovItem  " _
          & " Where nMovEstado = " & MovEstado.gMovEstPresupAceptado & " And nMovFlag = " & MovFlag.gMovFlagVigente & " And nDocTpo = " & 90 & " And MD.cDocNro =  '" & psOrdenCompra & "' Order By M.nMovNro, MBS.nMovItem"
       
       '--Left Join MovRef MRADD WITH (NOLOCK) On MRADD.nMovNroRef = MR.nMovNro
       '--Left Join Mov MoviADD WITH (NOLOCK) On MoviAdd.nMovNro = MRADD.nMovNro And MoviAdd.nMovEstado <> '21'
       
      Set GetDetOC = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
'EJVG20111110 ***************************************************
Public Function GetDetalleOCxAgencia(ByVal psOrdenCompra As String, ByVal psAgeCod As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
    
   If oCon.AbreConexion() Then
      sql = "Exec stp_sel_GetDetOCxAgencia '" & psOrdenCompra & "','" & psAgeCod & "'"
      Set GetDetalleOCxAgencia = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
'END**************************************************************
'Public Function GetDetNotaIngreso(ByVal psOpeCodCarb As String) As ADODB.Recordset
'   On Error GoTo CargaOpeGruErr
'   Dim OCON As DConecta
'   Set OCON = New DConecta
'   Dim sql As String
'   Dim oConst As NConstSistemas
'   Set oConst = New NConstSistemas
'
'   If OCON.AbreConexion() Then
'      sql = " Select M.nMovNro, MBS.nMovItem, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion, MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant," _
'          & "        MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, IsNull(MBSS.nIGV,0) nIGV, IsNull(MBSS.nValor,0) nValor, M.cOpeCod, OPE.cOpeDesc, MBS.nMovBSOrden, ALM.cAlmDescripcion, IsNull(MOI.nMovOtroImporte,0) nMovOtroImporte " _
'          & " From Mov M WITH (NOLOCK) " _
'          & " Inner Join MovDoc MD WITH (NOLOCK) On M.nMovNro = MD.nMovNro" _
'          & " Inner Join MovBS MBS WITH (NOLOCK) On M.nMovNro = MBS.nMovNro" _
'          & " Left  Join MovBSSerie MBSS WITH (NOLOCK) On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
'          & " Left  Join MovOtrosItem MOI WITH (NOLOCK) On MBS.nMovNro = MOI.nMovNro And MBS.nMovItem = MOI.nMovItem And cMovOtroVariable = '" & oConst.LeeConstSistema(gConstCtaIGV) & "'" _
'          & " Inner Join MovCta MCTA WITH (NOLOCK) On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
'          & " Inner Join MovCant MCT WITH (NOLOCK) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
'          & " Inner Join BienesServicios BS WITH (NOLOCK) On BS.cBSCod  = MBS.cBSCod" _
'          & " Inner Join Constante CO WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
'          & " Inner Join MovGasto MG WITH (NOLOCK) On M.nMovNro = MG.nMovNro Inner Join Almacen ALM On ALM.nAlmCod = MBS.nMovBSOrden" _
'          & " Inner Join Persona PE WITH (NOLOCK) On PE.cPersCod = MG.cPersCod Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
'          & " Where MD.cDocNro = '" & psOpeCodCarb & "' And nDocTpo = '42' And M.nMovEstado = '" & "13" & "'" _
'          & " And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "') Order by MBS.nMovItem"
'      Set GetDetNotaIngreso = OCON.CargaRecordSet(sql)
'      OCON.CierraConexion
'   End If
'   Set OCON = Nothing
'   Exit Function
'CargaOpeGruErr:
'   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
'End Function

Public Function GetDetNotaIngreso(ByVal psOpeCodCarb As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim oConst As NConstSistemas
   Set oConst = New NConstSistemas
   
   If oCon.AbreConexion() Then
      sql = " Select M.nMovNro, MBS.nMovItem, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod, Rtrim(BS.cBSDescripcion)  + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion, MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant," _
          & "        MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, IsNull(MBSS.nIGV,0) nIGV, IsNull(MBSS.nValor,0) nValor,Isnull(MBSS.cSerieFisica,'') SerieFisica, M.cOpeCod, OPE.cOpeDesc, MBS.nMovBSOrden, ALM.cAlmDescripcion, IsNull(MOI.nMovOtroImporte,0) nMovOtroImporte " _
          & " From Mov M" _
          & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
          & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
          & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
          & " Left  Join MovOtrosItem MOI On MBS.nMovNro = MOI.nMovNro And MBS.nMovItem = MOI.nMovItem And cMovOtroVariable = '" & oConst.LeeConstSistema(gConstCtaIGV) & "'" _
          & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
          & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
          & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
          & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
          & " Inner Join MovGasto MG On M.nMovNro = MG.nMovNro Inner Join Almacen ALM On ALM.nAlmCod = MBS.nMovBSOrden" _
          & " Inner Join Persona PE On PE.cPersCod = MG.cPersCod Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
          & " Where MD.cDocNro = '" & psOpeCodCarb & "' And nDocTpo = '42' And M.nMovEstado = '" & "13" & "'" _
          & " And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "') Order by MBS.nMovItem"
      Set GetDetNotaIngreso = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

'
'Public Function GetDetNotaIngresoReporte(ByVal psOpeCodCarb As String, psMovEstado As String, psDocumento As String) As ADODB.Recordset
'   On Error GoTo CargaOpeGruErr
'   Dim OCON As DConecta
'   Set OCON = New DConecta
'   Dim sql As String
'   Dim oConst As NConstSistemas
'   Set oConst = New NConstSistemas
'
'   If OCON.AbreConexion() Then
'      sql = " Select M.nMovNro, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion, MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant," _
'          & "        MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, IsNull(MBSS.nIGV,0) nIGV, IsNull(MBSS.nValor,0) nValor, M.cOpeCod, OPE.cOpeDesc, MBS.nMovBSOrden, ALM.cAlmDescripcion," _
'          & "        MAGEAREA.cAgeCod, MAGEAREA.cAreaCod, ARE.cAreaDescripcion, ISNull(MOI.nMovOtroImporte,0) nMovOtroImporte" _
'          & " From Mov M WITH (NOLOCK) " _
'          & " Inner Join MovDoc MD WITH (NOLOCK) On M.nMovNro = MD.nMovNro" _
'          & " Inner Join MovBS MBS WITH (NOLOCK) On M.nMovNro = MBS.nMovNro" _
'          & " Left  Join MovBSSerie MBSS WITH (NOLOCK) On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
'          & " Left  Join MovOtrosItem MOI WITH (NOLOCK) On MBS.nMovNro = MOI.nMovNro And MBS.nMovItem = MOI.nMovItem And cMovOtroVariable = '" & oConst.LeeConstSistema(gConstCtaIGV) & "'" _
'          & " Inner Join MovCta MCTA WITH (NOLOCK) On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
'          & " Inner Join MovCant MCT WITH (NOLOCK) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
'          & " Inner Join BienesServicios BS WITH (NOLOCK) On BS.cBSCod  = MBS.cBSCod" _
'          & " Inner Join Constante CO WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
'          & " Left  Join MovGasto MG WITH (NOLOCK) On M.nMovNro = MG.nMovNro Inner Join Almacen ALM On ALM.nAlmCod = MBS.nMovBSOrden" _
'          & " Left  Join MovObjAreaAgencia MAGEAREA WITH (NOLOCK) On MAGEAREA.nMovNro = M.nMovNro" _
'          & " Left  Join Areas ARE WITH (NOLOCK) On MAGEAREA.cAreaCod = ARE.cAreaCod" _
'          & " Left  Join Agencias AGE WITH (NOLOCK) On MAGEAREA.cAgeCod = AGE.cAgeCod" _
'          & " Left  Join Persona PE WITH (NOLOCK) On PE.cPersCod = MG.cPersCod Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
'          & " Where MD.cDocNro = '" & psOpeCodCarb & "' And nDocTpo = '" & psDocumento & "' And M.nMovEstado = '" & psMovEstado & "'" _
'          & " Order By MBS.nMovItem"
'          '"And M.nMovFlag 'Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "')"
'      Set GetDetNotaIngresoReporte = OCON.CargaRecordSet(sql)
'      OCON.CierraConexion
'   End If
'   Set OCON = Nothing
'   Exit Function
'CargaOpeGruErr:
'   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
'End Function

Public Function GetDetNotaIngresoReporte(ByVal psOpeCodCarb As String, psMovEstado As String, psDocumento As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim oConst As NConstSistemas
   Set oConst = New NConstSistemas
   
   If oCon.AbreConexion() Then
      sql = " Select M.nMovNro, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion, MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant," _
          & "        MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, IsNull(MBSS.nIGV,0) nIGV, IsNull(MBSS.nValor,0) nValor,Isnull(MBSS.cSerieFisica,'') SerieFisica, M.cOpeCod, OPE.cOpeDesc, MBS.nMovBSOrden, ALM.cAlmDescripcion," _
          & "        MAGEAREA.cAgeCod, MAGEAREA.cAreaCod, ARE.cAreaDescripcion, ISNull(MOI.nMovOtroImporte,0) nMovOtroImporte" _
          & " From Mov M" _
          & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
          & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
          & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
          & " Left  Join MovOtrosItem MOI On MBS.nMovNro = MOI.nMovNro And MBS.nMovItem = MOI.nMovItem And cMovOtroVariable = '" & oConst.LeeConstSistema(gConstCtaIGV) & "'" _
          & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
          & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
          & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
          & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
          & " Left  Join MovGasto MG On M.nMovNro = MG.nMovNro Inner Join Almacen ALM On ALM.nAlmCod = MBS.nMovBSOrden" _
          & " Left  Join MovObjAreaAgencia MAGEAREA On MAGEAREA.nMovNro = M.nMovNro" _
          & " Left  Join Areas ARE On MAGEAREA.cAreaCod = ARE.cAreaCod" _
          & " Left  Join Agencias AGE On MAGEAREA.cAgeCod = AGE.cAgeCod" _
          & " Left  Join Persona PE On PE.cPersCod = MG.cPersCod Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
          & " Where MD.cDocNro = '" & psOpeCodCarb & "' And nDocTpo = '" & psDocumento & "' And M.nMovEstado = '" & psMovEstado & "'" _
          & " Order By MBS.nMovItem"
          '"And M.nMovFlag 'Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "')"
      Set GetDetNotaIngresoReporte = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function


Public Function GetDetGuiaRemision(ByVal psOpeCodCarb As String, pnProdTpo As Integer, Optional pbCantidad As Boolean = False, Optional psAlmacen As String, Optional pnEstado As String = "13", Optional pbFlag As Boolean = True) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim sqlAdd As String
   Dim sqlFlag As String
   
   If pnEstado = "" Then
        sqlAdd = ""
   Else
        sqlAdd = " And M.nMovEstado = '" & pnEstado & "'"
   End If
   
   If pbFlag Then
        sqlFlag = " And M.nMovFlag = " & MovFlag.gMovFlagVigente & " "
   Else
        sqlFlag = ""
   End If
   
   If oCon.AbreConexion() Then
        If Not pbCantidad Then
          sql = " Select M.nMovNro, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion, MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant," _
              & "        MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, IsNull(MBSS.nIGV,0) nIGV, IsNull(MBSS.nValor,0) nValor , M.cOpeCod, OPE.cOpeDesc, MBS.nMovBSOrden, ARE.cAreaDescripcion, MOAA.cAreaCod, MOAA.cAgeCod, 0 Atendidos" _
              & " From Mov M WITH (NOLOCK) " _
              & " Inner Join MovDoc MD WITH (NOLOCK) On M.nMovNro = MD.nMovNro" _
              & " Inner Join MovBS MBS WITH (NOLOCK) On M.nMovNro = MBS.nMovNro" _
              & " Left  Join MovBSSerie MBSS WITH (NOLOCK) On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
              & " Inner Join MovCta MCTA WITH (NOLOCK) On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
              & " Inner Join MovObjAreaAgencia MOAA WITH (NOLOCK) On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem" _
              & " Inner Join Areas ARE WITH (NOLOCK) On MOAA.cAreaCod = ARE.cAreaCod" _
              & " Inner Join MovCant MCT WITH (NOLOCK) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
              & " Inner Join BienesServicios BS WITH (NOLOCK) On BS.cBSCod  = MBS.cBSCod" _
              & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
              & " Left Join MovGasto MG WITH (NOLOCK) On M.nMovNro = MG.nMovNro" _
              & " Left Join Persona PE WITH (NOLOCK) On PE.cPersCod = MG.cPersCod" _
              & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
              & " Where MD.cDocNro = '" & psOpeCodCarb & "' And nDocTpo = '70' " & sqlAdd & "   " & sqlFlag
      Else
          sql = " Select M.nMovNro, M.cMovDesc, PE.cPersCod, PE.cPersNombre, BS.cBSCod," _
              & " Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' +" _
              & " Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion," _
              & " MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant, MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, M.cOpeCod," _
              & " OPE.cOpeDesc , MBS.nMovBSOrden, ARE.cAreaDescripcion, MOAA.cAreaCod, MOAA.cAgeCod, IsNull(MREF.Atendidos, 0) Atendidos,   dbo.LogGetStock(" & Right(psAlmacen, 2) & ",BS.cBSCod," & pnProdTpo & ") Stock, IsNull(dbo.LogGetPrecioPromedio(-1,BS.cBSCod," & pnProdTpo & "),0) PrePromedio " _
              & " From Mov M WITH (NOLOCK) " _
              & " Inner Join MovDoc MD WITH (NOLOCK) On M.nMovNro = MD.nMovNro" _
              & " Inner Join MovBS MBS WITH (NOLOCK) On M.nMovNro = MBS.nMovNro" _
              & " Left  Join MovBSSerie MBSS WITH (NOLOCK) On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
              & " Inner Join MovCta MCTA WITH (NOLOCK) On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
              & " Inner Join MovObjAreaAgencia MOAA WITH (NOLOCK) On MBS.nMovNro = MOAA.nMovNro And MBS.nMovItem = MOAA.nMovItem" _
              & " Inner Join Areas ARE WITH (NOLOCK) On MOAA.cAreaCod = ARE.cAreaCod" _
              & " Inner Join MovCant MCT WITH (NOLOCK) On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
              & " Inner Join BienesServicios BS WITH (NOLOCK) On BS.cBSCod  = MBS.cBSCod" _
              & " Inner Join Constante CO WITH (NOLOCK) On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
              & " Left Join MovGasto MG WITH (NOLOCK) On M.nMovNro = MG.nMovNro" _
              & " Left Join Persona PE WITH (NOLOCK) On PE.cPersCod = MG.cPersCod" _
              & " Left Join (Select Sum(nMovCant) Atendidos, MRMBS.cBSCod, MR.nMovNroRef nMovNro From MovRef MR" _
              & "             Inner Join MovCant MRMCT On MR.nMovNro = MRMCT.nMovNro" _
              & "             Inner Join Mov MRM On MRM.nMovNro = MR.nMovNro" _
              & "             Inner Join MovBS MRMBS On MRMCT.nMovNro = MRMBS.nMovNro And MRMCT.nMovItem = MRMBS.nMovItem" _
              & "           Where MRM.nMovFlag = " & MovFlag.gMovFlagVigente & " And MRM.cOpeCod = '" & gnAlmaSalXAtencion & "'" _
              & "           Group By MRMBS.cBSCod,  MR.nMovNroRef ) MREF On MREF.nMovNro = M.nMovNro And MREF.cBSCod = MBS.cBSCod" _
              & " Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
              & " Where MD.cDocNro = '" & psOpeCodCarb & "' And nDocTpo = 70 And ((M.nMovEstado = " & pnEstado & " And MREF.Atendidos IS NULL) Or (M.nMovEstado = '5' And ISNull(MREF.Atendidos,0) <= MCT.nMovCant))  And M.nMovFlag = " & MovFlag.gMovFlagVigente & " And M.nMovNro Not In ( Select nMovNroRef From MovRef MR Inner Join Mov M On MR.nMovNro = M.nMovNro Where M.nMovFlag = " & MovFlag.gMovFlagVigente & " And cOpeCod in ('" & gnAlmaReqAreaReg & "','" & gnAlmaReqAreaMant & "'))"
      End If
      
      Set GetDetGuiaRemision = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function


Public Function VerifMovGastoPersonasPorcent(ByVal pnMovNro As Long) As Boolean
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " select top 1 nMovNro from MovGastoPersonasPorcent where  nMovNro = " & pnMovNro
      Set rs = oCon.CargaRecordSet(sql)
      
      If rs.EOF And rs.BOF Then
         VerifMovGastoPersonasPorcent = False
      Else
         VerifMovGastoPersonasPorcent = True
      End If
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

'Ñ - funcion
Public Function GetPersonasMovGastoPorcent(ByVal pnMovNro As Long, ByVal pnMovItem As Integer) As ADODB.Recordset
Dim oCon As DConecta, rs As ADODB.Recordset, cSQL As String
On Error GoTo CargaOpeGruErr

Set oCon = New DConecta
Set rs = New ADODB.Recordset

If oCon.AbreConexion() Then
   cSQL = "select cPersCod,'' as cAgeCod, '' as PrdCod ,nImporte as nMonto from MovGastoPersonasPorcent " & _
          " where  nMovNro = " & pnMovNro & " and nMovItem = " & pnMovItem & " "
   Set rs = oCon.CargaRecordSet(cSQL)
   oCon.CierraConexion
End If

Set GetPersonasMovGastoPorcent = rs
Set rs = Nothing
Set oCon = Nothing
Exit Function

CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetDetNotaIngresoDoc(ByVal pnMovNro As Long, pnTpoDocE As Integer) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      'sql = " Select M.nMovNro, PE.cPersCod, PE.cPersNombre, BS.cBSCod, Rtrim(BS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' cBSDescripcion, MCT.nMovItem, MCTA.nMovImporte, MCTA.cCtaContCod, MCT.nMovCant," _
          & "        MD.nDocTpo , MD.cDocNro, MD.dDocFecha, MBSS.cSerie, M.cOpeCod, OPE.cOpeDesc, MBS.nMovBSOrden, ALM.cAlmDescripcion " _
          & " From Mov M" _
          & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
          & " Inner Join MovBS MBS On M.nMovNro = MBS.nMovNro" _
          & " Left  Join MovBSSerie MBSS On M.nMovNro = MBSS.nMovNro And MBS.nMovItem = MBSS.nMovItem" _
          & " Inner Join MovCta MCTA On M.nMovNro = MCTA.nMovNro And MBS.nMovItem = MCTA.nMovItem" _
          & " Inner Join MovCant MCT On MBS.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
          & " Inner Join BienesServicios BS On BS.cBSCod  = MBS.cBSCod" _
          & " Inner Join Constante CO On CO.nConsValor = BS.nBSUnidad  And CO.nConsCod = '1019'" _
          & " Inner Join MovGasto MG On M.nMovNro = MG.nMovNro Inner Join Almacen ALM On ALM.nAlmCod = MBS.nMovBSOrden" _
          & " Inner Join Persona PE On PE.cPersCod = MG.cPersCod Inner Join OpeTpo OPE On OPE.cOpeCod = M.cOpeCod" _
          & " Where MD.cDocNro = '" & psOpeCodCarb & "' And nDocTpo = '42'" _
          & " And M.nMovFlag Not IN ('" & gMovFlagExtornado & "','" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "')"
      Set GetDetNotaIngresoDoc = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetOpeDoc(ByVal pnMovNro As Long, pnTpoDocE As Integer) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select DO.nDocTpo,cDocNro,dDocFecha, cDocDesc + space(50) + str(DO.nDocTpo) Descrip  From MovDoc MD " _
          & " Inner Join Documento DO On DO.nDocTpo = MD.nDocTpo " _
          & " Where nMovNro = '" & pnMovNro & "' And DO.nDocTpo <> '" & pnTpoDocE & "'"
      Set GetOpeDoc = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetFechaDoc(ByVal psDocNro As String, pnTpoDoc As Integer) As Date
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select dDocFecha from Movdoc where cDocNro = '" & psDocNro & "' And nDocTpo = '" & pnTpoDoc & "'"
      GetFechaDoc = oCon.CargaRecordSet(sql).Fields(0)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetOpeMov(ByVal pnMovNro As Long) As String
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select copecod from Mov where nMovNro = " & pnMovNro & " "
      Set rs = oCon.CargaRecordSet(sql)
      
      If rs.EOF And rs.BOF Then
        GetOpeMov = ""
      Else
        GetOpeMov = rs.Fields(0)
      End If
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Set rs = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
End Sub
Public Function GetStock(ByVal psBSCod As String, pnAlmCod As Integer) As Double
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select nStock from MaestroAlmacen where cBSCod = '" & psBSCod & "' And nAlmCod = " & pnAlmCod & ""
      Set rs = oCon.CargaRecordSet(sql)
      If rs.RecordCount <> 0 Then
        GetStock = rs!nStock
      Else
        GetStock = 0
      End If
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
Public Function GetPrecioPromedio(ByVal psBSCod As String, pnAlmCod As Integer) As Double
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select nPreProm from MaestroAlmacen where cBSCod = '" & psBSCod & "' And nAlmCod = '" & pnAlmCod & "'"
      Set rs = oCon.CargaRecordSet(sql)
      If rs.RecordCount <> 0 Then
        GetPrecioPromedio = rs!nPreProm
      Else
        GetPrecioPromedio = 0
      End If
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
Public Function GetCostoTotal(ByVal psBSCod As String, pnAlmCod As Integer) As Double
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select nCostoTotal from MaestroAlmacen where cBSCod = '" & psBSCod & "' And nAlmCod = '" & pnAlmCod & "'"
      Set rs = oCon.CargaRecordSet(sql)
      If rs.RecordCount <> 0 Then
         GetCostoTotal = rs!nCostoTotal
      Else
         GetCostoTotal = 0
      End If
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
Public Function GetNuevoBienAlmacen(ByVal psBSCod As String, pnAlmCod As Integer) As Boolean
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select cBSCod from MaestroAlmacen where cBSCod = '" & psBSCod & "' And nAlmCod = " & pnAlmCod & ""
      Set rs = oCon.CargaRecordSet(sql)
      If rs.RecordCount = 0 Then
        GetNuevoBienAlmacen = True
       Else
        GetNuevoBienAlmacen = False
      End If
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function VerUltStockMontos(ByVal psBSCod As String, pnAlmCod As Integer) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = "Select nPreUni, nStock, nCosProm From KardexAlmacen " _
          & "Where cBSCod = '" & psBSCod & "' And nAlmCod = " & pnAlmCod & "" _
          & "And dFecha = (Select Max(dFecha) From KardexAlmacen " _
          & "                Where cBSCod = '" & psBSCod & "' and nAlmCod = " & pnAlmCod & ")"
      Set VerUltStockMontos = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetMaestroAlmacen(ByVal pnAlmCod As Integer) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      sql = " Select cBSCod, nPreUni, nStock, nPreProm, nCostoTotal " _
          & " From MaestroAlmacen " _
          & " Where nAlmCod = " & pnAlmCod
      Set GetMaestroAlmacen = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function
Public Function GetPrimerIngreso(ByVal pnAlmCod As Integer) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String

   If oCon.AbreConexion() Then
      sql = " Select cBSCod " _
          & " From KardexAlmacen " _
          & " Where nAlmCod = " & pnAlmCod
      Set GetPrimerIngreso = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

'*** PEAC 20110712
Public Function GetGuiaSalidaPorTransf(ByVal psMoneda As String, ByVal psAge As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   
   If oCon.AbreConexion() Then
      
      sql = "exec stp_sel_GetsalidaPorTransf '" & psMoneda & "','" & psAge & "'"
      
      Set GetGuiaSalidaPorTransf = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:GetGuiaSalidaPorTransf Method")
End Function

'*** PEAC 20110712
Public Function GetDetGuiaRemisionTransf(ByVal psNumDoc As String, ByVal psAgeCod As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim sql As String
   Dim oConst As NConstSistemas
   Set oConst = New NConstSistemas
   
   If oCon.AbreConexion() Then
      
      sql = "exec stp_sel_GetDetGuiaRemisionTransf '" & psNumDoc & "','" & psAgeCod & "'"
      
      Set GetDetGuiaRemisionTransf = oCon.CargaRecordSet(sql)
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:GetDetGuiaRemisionTransf Method")
End Function

