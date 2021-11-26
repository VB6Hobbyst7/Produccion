Attribute VB_Name = "gFunLogistica3"
Option Explicit

Public Function GetAreaCargoAgencia(psPersCod As String) As ADODB.Recordset
Dim oConn As New DConecta, sSQL As String, rc As New ADODB.Recordset
Dim rs As New ADODB.Recordset

rc.Fields.Append "cRHCargoCod", adVarChar, 6, adFldMayBeNull
rc.Fields.Append "cRHAreaCod", adVarChar, 3, adFldMayBeNull
rc.Fields.Append "cRHAgeCod", adVarChar, 2, adFldMayBeNull
rc.Fields.Append "cRHCargo", adVarChar, 40, adFldMayBeNull
rc.Fields.Append "cRHArea", adVarChar, 40, adFldMayBeNull
rc.Fields.Append "cRHAgencia", adVarChar, 40, adFldMayBeNull
rc.Fields.Append "cPersona", adVarChar, 60, adFldMayBeNull
rc.Open

sSQL = "select cRHCargoCodOficial as cRHCargoCod,cRHAreaCodOficial as cRHAreaCod," & _
       "       cRHAgenciaCodOficial as cRHAgeCod,t1.cRHCargoDescripcion as cRHCargo, " & _
       "       cPersona=replace(p.cPersNombre,'/',' '), " & _
       "       t2.cAreaDescripcion as cRHArea, t3.cAgeDescripcion as cRHAgencia" & _
       "  from RHCargos c inner join RHCargosTabla t1 on c.cRHCargoCodOficial = t1.cRHCargoCod " & _
       "                  inner join Areas t2 on c.cRHAreaCodOficial = t2.cAreaCod " & _
       "                  inner join Agencias t3 on c.cRHAgenciaCodOficial = t3.cAgeCod " & _
       "                  inner join Persona p on c.cPersCod = p.cPersCod " & _
       " where c.cPersCod = '" & psPersCod & "' and " & _
       "       dRHCargoFecha = (select max(dRHCargoFecha) from RHCargos where cPersCod = '" & psPersCod & "' and dRHCargoFecha<='200507') "

If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      rc.AddNew
      rc.Fields(0) = rs!cRHCargoCod
      rc.Fields(1) = rs!cRHAreaCod
      rc.Fields(2) = rs!cRHAgeCod
      rc.Fields(3) = Mid(rs!cRHCargo, 1, 40)
      rc.Fields(4) = Mid(rs!cRHArea, 1, 40)
      rc.Fields(5) = Mid(rs!cRHAgencia, 1, 40)
      rc.Fields(6) = Mid(rs!cPersona, 1, 60)
      rc.Update
   End If
   Set GetAreaCargoAgencia = rc
End If
End Function


Public Function CargaComboBoxConstante(ByRef pcComboBox As Control, ByVal pnConsCod As Integer)
Dim sSQL As String, rs As New ADODB.Recordset, oConn As New DConecta

sSQL = "select nConsValor,cConsDescripcion from Constante where nConsCod = " & pnConsCod & " and nConsCod <> nConsValor order by nConsValor"
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   oConn.CierraConexion
End If
If Not rs.EOF Then
   Do While Not rs.EOF
      pcComboBox.AddItem rs(1)
      pcComboBox.ItemData(pcComboBox.ListCount - 1) = rs(0)
      rs.MoveNext
   Loop
   pcComboBox.ListIndex = 0
End If
End Function
