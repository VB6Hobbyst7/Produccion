VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSubasta"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String

Public Function GetCodigo(psPeriodo As String) As String
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "Select Top 1 cSubastaNro From LogSubasta Where cSubastaNro Like '" & psPeriodo & "%' Order By cSubastaNro Desc"
      Set rs = dbConec.CargaRecordSet(psSql)
              
      If rs.EOF And rs.BOF Then
          GetCodigo = psPeriodo & "-00001"
      Else
          GetCodigo = psPeriodo & "-" & Format(CCur(Right(rs!cSubastaNro, 5)) + 1, "00000")
      End If
      
      rs.Close
      Set rs = Nothing
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetUltimaSubasta(Optional psCodigo As String = "") As Long
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = " Select Top 1 nMovNroSub, cSubastaNro From LogSubasta  LS" _
            & " Inner Join Mov M On LS.nMovNroSub = M.nMovNro" _
            & " Where nMovFlag = 0" _
            & " Order By nMovNroSub Desc"
      Set rs = dbConec.CargaRecordSet(psSql)
              
      If rs.EOF And rs.BOF Then
          GetUltimaSubasta = 0
      Else
          psCodigo = rs!cSubastaNro
          GetUltimaSubasta = rs.Fields(0)
      End If
      
      rs.Close
      Set rs = Nothing
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function ValidaLote(pnSubastaActual As Long, pnMovIng As Long) As Boolean
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = " Select nMovNroSub From LogSubastaDetalle  LS " _
            & " Inner Join Mov M On LS.nMovNroIng = M.nMovNro " _
            & " Where M.nMovFlag = 0 " _
            & " And nMovNroIng = " & pnMovIng & " And nMovNroSub = " & pnSubastaActual
      Set rs = dbConec.CargaRecordSet(psSql)
              
      If rs.EOF And rs.BOF Then
          ValidaLote = False
      Else
          ValidaLote = True
      End If
      
      rs.Close
      Set rs = Nothing
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetBSLote(pnSubastaActual As Long, pnMovIng As Long) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = " SELECT BSS.cBSCod, Rtrim(BSS.cBSDescripcion) + '-(' + RTrim(CO.cConsDescripcion)  + '[' + Case bSerie When 0 Then 'N' Else 'S' End +  '])' Descripcion, 4 Nivel" _
            & " FROM LogSubastaDetalle CD" _
            & " LEFT JOIN BienesServicios BSS ON CD.cBSCod = BSS.cBSCod" _
            & " LEFT JOIN Constante CO ON BSS.nBSUnidad = CO.nConsValor And CO.nConsCod = 1019" _
            & " Where nMovNroSub = " & pnSubastaActual & " And nMovNroIng = " & pnMovIng & "" _
            & " ORDER BY CD.cBSCod"
      Set GetBSLote = dbConec.CargaRecordSet(psSql)
              
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function InicioRemate(pnMovNro As Long, psCodigo As String, pdIni As Date, pdFin As Date, prPer As ADODB.Recordset, prBB As ADODB.Recordset) As Boolean
   On Error GoTo CargaOpeGruErr
   
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      dbConec.AbreConexion
        psSql = " Insert LogSubasta(nMovNroSub,cSubastaNro,dInicio,dFin)" _
              & " Values(" & pnMovNro & ",'" & psCodigo & "','" & Format(pdIni, gsFormatoFecha) & "','" & Format(pdFin, gsFormatoFecha) & "')"
        dbConec.Ejecutar psSql
        
        prPer.MoveFirst
        
        While Not prPer.EOF
            psSql = " Insert LogSubastaPers(nMovNroSub,cPersCod,nTpoCargo)" _
                  & " Values(" & pnMovNro & ",'" & prPer.Fields(0) & "','" & Trim(Right(prPer.Fields(2), 5)) & "')"
            dbConec.Ejecutar psSql
            prPer.MoveNext
        Wend
        
        If prBB Is Nothing Then Exit Function
        
        prBB.MoveFirst
        
        While Not prBB.EOF
            psSql = " Insert LogSubastaDetalle (nMovNroSub,nMovNroIng,cBSCod," _
                  & "        cDocNro,dDocFecha,nCantidadIni,nImporteIni," _
                  & "        nImporteDeprecia,nPrecioVentaAnt,nPrecioVenta,nCantidadVendida," _
                  & "        nImporteVendido,nMovFlag, nAlmCod)" _
                  & " Values(" & pnMovNro & ",'" & prBB.Fields(0) & "','" & prBB.Fields(1) & "'" _
                  & "      ,'" & prBB.Fields(3) & "','" & Format(prBB.Fields(4), gsFormatoFecha) & "'," & prBB.Fields(5) & "," & prBB.Fields(6) & "" _
                  & "      ," & prBB.Fields(7) & "," & prBB.Fields(8) & "," & prBB.Fields(9) & "," & prBB.Fields(10) & "" _
                  & "      ," & prBB.Fields(11) & ",0," & prBB.Fields(13) & ")"
            dbConec.Ejecutar psSql
            prBB.MoveNext
        Wend
        
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function CierreRemate(pnMovNro As Long) As Boolean
   On Error GoTo CargaOpeGruErr
   
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      dbConec.AbreConexion
        psSql = " Update LogSubasta " _
              & " Set bCerrada = 1" _
              & " Where  nMovNroSub = " & pnMovNro & ""
        dbConec.Ejecutar psSql
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetNuevaSubasta() As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      psSql = " Select  nMovNroIng, BS.cBSCod, BS.cBSDescripcion, cDocNro, Convert(Varchar(10),dDocFecha,103) Fecha," _
            & " Convert(Decimal(20,2),nCantidadIni) nCantidadIni , Convert(Decimal(20,2),nImporteIni) nImporteIni, Convert(decimal(15,2),(nImporteDeprecia + (nPrecioVenta * (Select (nParValor) /100  From Parametro Where nParProd = 5000 And nParCod = 1006)))) nImporteDeprecia, nPrecioVenta nPrecioVentaAnt, Convert(decimal(15,2),(nPrecioVenta * (Select (100 - nParValor) /100  From Parametro Where nParProd = 5000 And nParCod = 1006))) nPrecioVenta," _
            & " Convert(Decimal(20,2),nCantidadVendida) nCantidadVendida , Convert(Decimal(20,2),nImporteVendido) nImporteVendido, nMovFlag, nAlmCod From LogSubastaDetalle LSD" _
            & " Inner Join BienesServicios BS On BS.cBSCod = LSD.cBSCod Where nMovFlag = 0" _
            & " And nMovNroSub = (Select Top 1 nMovNroSub From LogSubasta Order by nMovNroSub Desc) " _
            & " Union" _
            & " Select  M.nMovNro nMovNroIng, BS.cBSCod cBSCod, BS.cBSDescripcion , MD.cDocNro, Convert(Varchar(10),MD.dDocFecha,103) Fecha, Convert(decimal(20,2),MC.nMovCant) nMovCant," _
            & " Convert(decimal(20,2),MCTA.nMovImporte) nImporteIni, 0.00 nImporteDeprecia, 0.00 nPrecioVentaAnt, Convert(decimal(20,2),MCTA.nMovImporte) nPrecioVenta," _
            & " 0.00 nCantidadVendida, 0.00 nImporteVendido, 0 nMovFlag, MB.nMovBSOrden nAlmCod" _
            & " From Mov M" _
            & " Inner Join MovDoc MD On MD.nMovNro = M.nMovNro" _
            & " Inner Join MovCant MC On M.nMovNro = MC.nMovNro" _
            & " Inner Join MovCta MCTA On MC.nMovNro = MCTA.nMovNro And MC.nMovItem = MCTA.nMovItem" _
            & " Inner Join MovBS MB On MC.nMovNro = MB.nMovNro And MC.nMovItem = MB.nMovItem" _
            & " Inner Join BienesServicios BS On MB.cBSCod = BS.cBSCod" _
            & " where cOpeCod in ('591107','591109') And nMovFlag = 0" _
            & " And nDocTpo = 42 And M.nMovNro Not In (Select nMovNroIng from LogSubastaDetalle)"
      Set GetNuevaSubasta = dbConec.CargaRecordSet(psSql)
      
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetSubasta(pbFiltro As Boolean, pbSoloAbiertos As Boolean) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      If pbFiltro Then
        psSql = " Select cSubastaNro, cMovDesc + '  [' + Ltrim(str(nMovNroSub)) + ']' ,  4 from LogSubasta LS" _
              & " Inner Join Mov M On M.nMovNro = LS.nMovNroSub" _
              & " Where nMovFlag = 0 And bCerrada = " & IIf(pbSoloAbiertos, 0, 1)
      Else
        psSql = " Select cSubastaNro, cMovDesc + '  [' + Ltrim(str(nMovNroSub)) + ']' ,  4 from LogSubasta LS" _
              & " Inner Join Mov M On M.nMovNro = LS.nMovNroSub" _
              & " Where nMovFlag = 0"
      End If
      Set GetSubasta = dbConec.CargaRecordSet(psSql)
      
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetSubastaDet(pnMovNro As Long) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      psSql = " Select nMovNroSub,cSubastaNro,dInicio,dFin,cMovDesc, bCerrada from LogSubasta LS" _
            & " Inner Join Mov M On M.nMovNro = LS.nMovNroSub" _
            & " Where nMovNroSub = " & pnMovNro
      Set GetSubastaDet = dbConec.CargaRecordSet(psSql)
      
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetSubastaPers(pnMovNro As Long) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      psSql = " Select PE.cPersCod, cPersNombre, cConsDescripcion +  space(50) + str(nConsValor)  from LogSubastaPers LS" _
            & " Inner Join Persona PE On Pe.cPersCod = LS.cPersCod" _
            & " Inner join Constante CO On CO.nConsValor = LS.nTpoCargo And CO.nConsCod = '5009'" _
            & " Where nMovNroSub = " & pnMovNro
      Set GetSubastaPers = dbConec.CargaRecordSet(psSql)
      
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetSubastaDetalle(pnMovNro As Long) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      psSql = " Select  nMovNroIng, BS.cBSCod, BS.cBSDescripcion, cDocNro, Convert(Varchar(10),dDocFecha,103) Fecha, " _
            & " nCantidadIni, nImporteIni, nImporteDeprecia, nPrecioVenta nPrecioVentaAnt," _
            & " nPrecioVenta , nCantidadVendida, nImporteVendido, nMovFlag, nALmCod" _
            & " From LogSubastaDetalle LSD" _
            & " Inner Join BienesServicios BS On BS.cBSCod = LSD.cBSCod" _
            & " Where nMovNroSub = " & pnMovNro
      Set GetSubastaDetalle = dbConec.CargaRecordSet(psSql)
      
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetSubastaDetalleResul(pnMovNro As Long) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
   
      psSql = " Select  nMovNroIng, BS.cBSCod, BS.cBSDescripcion," _
            & " dbo.GetSubastaCantVen(" & pnMovNro & ",nMovNroIng,BS.cBSCod) Cantidad," _
            & " dbo.GetSubastaMontoVen(" & pnMovNro & ",nMovNroIng,BS.cBSCod) Monto, nMovFlag" _
            & " From LogSubastaDetalle LSD" _
            & " Inner Join BienesServicios BS On BS.cBSCod = LSD.cBSCod" _
            & " Where nMovNroSub = " & pnMovNro
      Set GetSubastaDetalleResul = dbConec.CargaRecordSet(psSql)
      
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetSubastaSockPrecio(pnSubastaActual As Long, pnMovIng As Long, psBSCod As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      psSql = " Select dbo.GetSubastaStock(" & pnSubastaActual & "," & pnMovIng & ",'" & psBSCod & "') Stock, dbo.GetSubastaPrecio(" & pnSubastaActual & "," & pnMovIng & ",'" & psBSCod & "') , dbo.GetSubastaAlmacen(" & pnSubastaActual & "," & pnMovIng & ",'" & psBSCod & "') Precio"
      Set GetSubastaSockPrecio = dbConec.CargaRecordSet(psSql)
      
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetProvSubasta(pnMovIng As Long, psBSCod As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      psSql = " Select Sum(Monto) MOnto, Anio  from " _
            & " (Select Round(nMovImporte / nMovCant,2) Monto, Left(M.cMovNro,4) Anio from MovBS MBS" _
            & "     Inner Join MovCta MC On MBS.nMovNro = MC.nMovNro And MBS.nMovItem = MC.nMovItem - (Select Count(*) Num from MovBS Where nMovNro = " & pnMovIng & " )" _
            & "     Inner Join MovCant MCT On MCT.nMovNro = MBS.nMovNro And MCT.nMovItem = MBS.nMovItem" _
            & "     Inner Join Mov M On M.nMovNro = MBS.nMovNro" _
            & "     Where MC.nMovNro = " & pnMovIng & "  And cBSCod = '" & psBSCod & "'" _
            & " Union" _
            & " Select ISNULL(Sum(MC.nMovImporte / MCT.nMovCant),0) Monto, Left(M.cMovNro,4) Anio From MovRef MR" _
            & " Inner Join Mov M On MR.nMovNro = M.nMovNro" _
            & "     Inner Join MovBS MBS On MBS.nMovNro = M.nMovNro" _
            & "     Inner Join MovCant MCT On M.nMovNro = MCT.nMovNro And MBS.nMovItem = MCT.nMovItem" _
            & "     Inner Join MovCta MC On M.nMovNro = MC.nMovNro And MBS.nMovItem = MC.nMovItem" _
            & "     Where MR.nMovNroRef = " & pnMovIng & " And cBSCod = '" & psBSCod & "' And cOpeCod = '581101' Group By Left(M.cMovNro,4)" _
            & " ) As B  group by Anio"
      Set GetProvSubasta = dbConec.CargaRecordSet(psSql)
      
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function GetMovSubasta(psUser As String, psOpeCod As String, psCtaContCod As String) As ADODB.Recordset
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      psSql = " Select cOpeDesc, Monto From OpeTpo Ope " _
            & " Inner Join (" _
            & "                 Select Convert(Decimal(19,2),Sum(nMovImporte)) Monto,  cOpeCod from Mov M" _
            & "                         left  Join MovCta MC On M.nMovNro = MC.nMOvnro" _
            & "                         Where cOpeCod In (Select cOpeCodRef from OpeRef where cOpeCod = '" & psOpeCod & "')" _
            & "                         And Right(M.cMovNro,4) = '" & psUser & "' And cCtaContCod like '" & psCtaContCod & "%'" _
            & "                         And nMovEstado = " & gMovEstContabMovContable & " And nMovFlag = " & gMovFlagVigente & "" _
            & "                 Group By cOpeCod" _
            & "         ) As MTO On MTO.cOpeCod = Ope.cOpeCod"

      Set rs = dbConec.CargaRecordSet(psSql)
      Set GetMovSubasta = rs
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaOpeGruErr:
   Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

Public Function VerfDevSubasta(psUser As String, psOpeCod As String, pdFecha As Date) As Boolean
   On Error GoTo CargaOpeGruErr
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      psSql = " Select nMovNro From Mov M" _
            & " Where cOpeCod = '" & psOpeCod & "' And cMovNro like '" & Format(pdFecha, "yyyymmdd") & "%" & psUser & "'" _
            & " And nMovEstado = " & gMovEstContabMovContable & " And nMovFlag = " & gMovFlagVigente
      Set rs = dbConec.CargaRecordSet(psSql)
      
      If rs.EOF And rs.BOF Then
        VerfDevSubasta = False
      Else
        VerfDevSubasta = True
      End If
      
   End If
   Set dbConec = Nothing
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
