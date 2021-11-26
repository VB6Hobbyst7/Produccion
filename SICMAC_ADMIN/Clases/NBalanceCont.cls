VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NBalanceCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim dbConec As DConecta
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String
Dim lsEmprLogo  As String

Public Enum ValidacionBalance
   gValidaCuadreAsiento = 1
   gValidaConvesionME = 2
   gValidaCuentasNoExistentes = 3
   gValidaCuentasNoExistentes2 = 4
   gValidaCuentasAnaliticas = 5
   gValidaCuentasDeOrden = 6
   gValidaSaldosContables = 7
   gValidaCuentasSinPadre = 8
End Enum

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    Set oImp = Nothing
    
    Dim oCons As New NConstSistemas
    lsEmprLogo = oCons.LeeConstSistema(gConstSistNombreAbrevCMAC)
    gbBitCentral = IIf(oCons.LeeConstSistema(gConstSistBitCentral) = "1", True, False)
    gcConvTipo = oCons.LeeConstSistema(gConstSistTipoConverión)
    Set oCons = Nothing
    
Dim oIni As New ClasIni
   sCentralCom = oIni.BaseComunes
Set oIni = Nothing
Set dbConec = New DConecta
If Not dbConec.AbreConexion() Then
   Call RaiseError(MyUnhandledError, "NBalanceCont:Initialize Method. " & Err.Description)
End If

CON = PrnSet("C+")
COFF = PrnSet("C-")
BON = PrnSet("B+")
BOFF = PrnSet("B-")

End Sub

Private Sub Class_Terminate()
dbConec.CierraConexion
Set dbConec = Nothing
End Sub

Public Function CargaBalanceAjusteInflacion(psCtaCod As String, psAnio As String, psMes As String) As Recordset
   On Error GoTo CargaBalanceAjusteInflacionErr
   psSql = "SELECT cBalanceMes, cCtaContCod, SUM(nDebe) as nDebe, SUM(nHaber) as nHaber FROM BalanceEstad " _
                 & "WHERE cCtaContCod like '" & Mid(psCtaCod, 1, 2) & "[12]" & Mid(psCtaCod, 4, 20) & "' and cBalanceAnio = '" & psAnio & "' and cBalanceMes <= '" & psMes & "' and cBalanceCate = '1' and cBalanceTipo IN ('1','2') GROUP BY cBalanceMes, cCtaContCod " _
         & "Union All " _
         & "SELECT '00' cBalanceMes, '" & Mid(psCtaCod, 1, 2) & "1" & Mid(psCtaCod, 4, 20) & "', ISNULL(sum(nCtaSaldoImporte),0) nDebe, 0 nHaber " _
         & "FROM CtaSaldo cs WHERE cs.cCtaContCod like '" & Mid(psCtaCod, 1, 2) & "[126]" & Mid(psCtaCod, 4, 20) & "%' and dCtaSaldoFecha = " _
         & "    (Select max(dCtaSaldoFecha) From CtaSaldo cs1 where cs1.cCtaContCod = cs.cCtaContCod and " _
         & "            cs1.dCtaSaldoFecha <= '" & (psAnio - 1) & "-12-31') " _
         & "ORDER BY cCtaContCod, cBalanceMes "
   Set CargaBalanceAjusteInflacion = dbConec.CargaRecordSet(psSql, adLockReadOnly)
   Exit Function
CargaBalanceAjusteInflacionErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:CargaBalanceAjusteInflacion Method")
End Function

Public Function ValidaBalanceHistoricoAjuste(psAnio As String, psMes As String) As Boolean
Dim nContMN  As Integer, nContME As Integer
Dim sFaltaMN As String
Dim sFaltaME As String
Dim sMesesMN As String
Dim sMesesME As String

   On Error GoTo ValidaBalanceHistoricoAjusteErr
   ValidaBalanceHistoricoAjuste = False
   psSql = "SELECT Distinct cBalanceMes, cBalanceTipo FROM BalanceEstad " _
         & "WHERE cBalanceAnio = '" & psAnio & "' and cBalanceMes <= '" & psMes & "' and cBalanceCate = '1' and cBalanceTipo IN ('1','2') ORDER BY cBalanceMes "
   Set prs = dbConec.CargaRecordSet(psSql, adLockReadOnly)
   If prs.RecordCount = Val(psMes) * 2 Then
      ValidaBalanceHistoricoAjuste = True
   Else
      sMesesMN = "": sMesesME = ""
      Do While Not prs.EOF
         If prs!cBalanceTipo = "1" Then
            sMesesMN = sMesesMN & "'" & Format(prs!cBalanceMes, "00")
         Else
            sMesesME = sMesesME & "'" & Format(prs!cBalanceMes, "00")
         End If
         prs.MoveNext
      Loop
      For nContMN = 1 To Val(psMes)
         If InStr(sMesesMN, Format(nContMN, "00")) = 0 Then
            sFaltaMN = sFaltaMN & Format(nContMN, "00") & ", "
         End If
      Next
      For nContME = 1 To Val(psMes)
         If InStr(sMesesME, Format(nContME, "00")) = 0 Then
            sFaltaME = sFaltaME & Format(nContME, "00") & ", "
         End If
      Next
      If sFaltaMN <> "" Then
         sFaltaMN = Left(sFaltaMN, Len(sFaltaMN) - 2)
      End If
      If sFaltaME <> "" Then
         sFaltaME = Left(sFaltaME, Len(sFaltaME) - 2)
      End If
      If sFaltaME <> "" Or sFaltaMN <> "" Then
         Err.Raise 50001, "NBalanceCont: ValidaBalanceHistoricoAjuste", "Falta Generar los Balances Históricos de los Meses : " & oImpresora.gPrnSaltoLinea & "Moneda Nacional  : " & sFaltaMN & oImpresora.gPrnSaltoLinea & "Moneda Extranjera : " & sFaltaME
      End If
   End If
   RSClose prs
   ValidaBalanceHistoricoAjuste = True
   Exit Function
ValidaBalanceHistoricoAjusteErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ValidaBalanceHistoricoAjuste Method")
End Function


Public Function CargaBalanceGeneral(psTipoRepoFormula As String, psFecha As String, psBalanceCate As String) As Recordset
   On Error GoTo CargaBalanceGeneralErr
   psSql = "SELECT EB.cCodigo,EB.cDescrip,B.nMN,B.nME,B.nTotal,B.nTotAj " _
         & "FROM   RepBaseFormula EB Inner JOIN BalanceGen B ON CONVERT(Int,EB.cCodigo) = B.cCodigo " _
         & "WHERE  B.cFecha = '" & psFecha & "' And EB.cOpeCod='" & psTipoRepoFormula & "' And B.cBalanceCate = '" & psBalanceCate & "' " _
         & "ORDER BY CONVERT(Int,EB.cCodigo)"
   Set CargaBalanceGeneral = dbConec.CargaRecordSet(psSql, adLockReadOnly)
   Exit Function
CargaBalanceGeneralErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:CargaBalanceGeneral Method")
End Function

Public Function GetBalanceSaldoCuenta(psCtaCod As String, psAnio As String, psMes As String) As Currency
   On Error GoTo GetBalanceSaldoCuentaErr
      psSql = "SELECT nSaldoFinImporte FROM BalanceEstad WHERE cCtaContCod = '" & psCtaCod & "' and cBalanceAnio = '" & psAnio & "' and cBalanceMes = '" & psMes & "' and cBalanceCate = '2' and cBalanceTipo = '" & gMonedaNacional & "' "
      Set prs = dbConec.CargaRecordSet(psSql, adLockReadOnly)
      If Not prs.EOF Then
         GetBalanceSaldoCuenta = prs!nSaldoFinImporte
      Else
         GetBalanceSaldoCuenta = 0
      End If
   prs.Close
   Set prs = Nothing
   Exit Function
GetBalanceSaldoCuentaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:GetBalanceSaldoCuenta Method")
End Function

Public Function GetBalanceSaldoCuentaFiltro(psCtaCod As String, psAnio As String, psMes As String, psBalanceTipo As String, Optional pbFiltrarDeudorAcreedor As Boolean = False, Optional pbSaldoAcreedor As Boolean = False, Optional cOpeCod As String = "", Optional cCodigo As String = "") As Currency
   
    Dim lsCtaCaracter As String
    Dim CadSql   As String, CadSql2 As String
    Dim RSalAD   As ADODB.Recordset
    Dim lsCtaDos As String
    Dim oCta     As New DCtaCont
    
    '
    Dim nSaldoTempo As Double
    Dim cCuenta()  As String
    Dim cSigno() As String
    Dim nCantTemp As Integer
    Dim I As Integer
    '
    Dim sOtrasCuentas As String
   
   
    On Error GoTo GetBalanceSaldoCuentaErr
    
    sOtrasCuentas = ""
    
    If pbFiltrarDeudorAcreedor = True Then
        
        'FIltrar las cuentas que estan entre parentesis y dependen de la cuenta principal (SOLO SA Y SD)
        sOtrasCuentas = ""
        sOtrasCuentas = "Select cCtaContCodRef, cSigno From RepBaseFormulaDet "
        sOtrasCuentas = sOtrasCuentas & " WHere cOpeCod='" & cOpeCod & "' and cCodigo='" & cCodigo & "' "
        If Len(psCtaCod) > 3 Then
            sOtrasCuentas = sOtrasCuentas & " And cCtaContCod='" & Left(psCtaCod, 2) & "0" & Mid(psCtaCod, 4, Len(psCtaCod) - 3) & "'  "
        Else
            sOtrasCuentas = sOtrasCuentas & " And cCtaContCod='" & psCtaCod & "'  "
        End If
        
        Set RSalAD = dbConec.CargaRecordSet(sOtrasCuentas)
        sOtrasCuentas = ""
        If Not RSalAD.EOF Then
            Do While Not RSalAD.EOF
                
                ReDim Preserve cCuenta(nCantTemp) As String
                ReDim Preserve cSigno(nCantTemp) As String
                If Len(RSalAD!cCtaContCodRef) > 3 Then
                    cCuenta(nCantTemp) = Left(RSalAD!cCtaContCodRef, 2) & Mid(psCtaCod, 3, 1) & Mid(RSalAD!cCtaContCodRef, 4, Len(RSalAD!cCtaContCodRef) - 3)
                Else
                    cCuenta(nCantTemp) = RSalAD!cCtaContCodRef
                End If
                cSigno(nCantTemp) = RSalAD!cSigno
                nCantTemp = nCantTemp + 1
                RSalAD.MoveNext
            Loop
        End If
        RSalAD.Close
        'Fin de filtro de otras cuentas
         
        'Agrego la cuenta original
        ReDim Preserve cCuenta(nCantTemp) As String
        ReDim Preserve cSigno(nCantTemp) As String
        cCuenta(nCantTemp) = psCtaCod
        cSigno(nCantTemp) = "+"
        'Fin de Cuenta Original
        
        'Sacar el tipo de cuenta registrado
        If Len(Trim(psCtaCod)) > 1 Then
            lsCtaDos = Mid(psCtaCod, 1, 2)
        Else
            lsCtaDos = Mid(psCtaCod, 1, 1)
        End If
       
        Set RSalAD = oCta.CargaCtaContClase(lsCtaDos)
        If Not RSalAD.EOF Then
            lsCtaCaracter = Trim(RSalAD!cCtaCaracter)
        End If
        RSalAD.Close
        CadSql2 = ""
        'Fin Sacar tipo de cuenta
        
        For I = 0 To nCantTemp
        
            CadSql = ""
        
            psSql = "SELECT B.nSaldoFinImporte "
            psSql = psSql & " FROM BalanceEstad B "
            psSql = psSql & " WHERE B.cCtaContCod in('" & Trim(cCuenta(I)) & "') "
            psSql = psSql & " and B.cBalanceAnio = '" & psAnio & "' "
            psSql = psSql & " and B.cBalanceMes = '" & psMes & "' "
            psSql = psSql & " and B.cBalanceCate = '2' "
            psSql = psSql & " and B.cBalanceTipo = '" & psBalanceTipo & "' "
            
            Set prs = dbConec.CargaRecordSet(psSql, adLockReadOnly)
            If Not prs.EOF Then
                If cSigno(I) = "+" Then
                    nSaldoTempo = nSaldoTempo + prs!nSaldoFinImporte
                ElseIf cSigno(I) = "-" Then
                    nSaldoTempo = nSaldoTempo - prs!nSaldoFinImporte
                End If
            End If
            prs.Close
        Next
        
        'Depurando
        If pbSaldoAcreedor Then  'si es saldo acreedor
            If lsCtaCaracter = "A" Then ' si es cuenta Acreedora
                If nSaldoTempo > 0 Then
                    'OK    CadSql2 = " And B.nSaldoFinImporte > 0"
                Else
                    nSaldoTempo = 0
                End If
            Else                    'Si es deudora
                If nSaldoTempo < 0 Then
                    'OK    CadSql2 = " And B.nSaldoFinImporte < 0"
                Else
                    nSaldoTempo = 0
                End If
            End If
        Else                'Si es Saldo deudor
            If lsCtaCaracter = "A" Then ' si es cuenta Acreedora
                If nSaldoTempo < 0 Then
                    'OK    CadSql2 = " And B.nSaldoFinImporte < 0"
                Else
                    nSaldoTempo = 0
                End If
            Else                    'Si es deudora
                If nSaldoTempo > 0 Then
                    'OK    CadSql2 = " And B.nSaldoFinImporte > 0"
                Else
                    nSaldoTempo = 0
                End If
            End If
        End If
        
        GetBalanceSaldoCuentaFiltro = nSaldoTempo
        
    Else
    
        CadSql = ""
    
        psSql = "SELECT B.nSaldoFinImporte "
        psSql = psSql & " FROM BalanceEstad B "
        psSql = psSql & " WHERE B.cCtaContCod IN('" & psCtaCod & "') "
        psSql = psSql & " and B.cBalanceAnio = '" & psAnio & "' "
        psSql = psSql & " and B.cBalanceMes = '" & psMes & "' "
        psSql = psSql & " and B.cBalanceCate = '2' "
        psSql = psSql & " and B.cBalanceTipo = '" & psBalanceTipo & "' "
        
        Set prs = dbConec.CargaRecordSet(psSql, adLockReadOnly)
        If Not prs.EOF Then
            GetBalanceSaldoCuentaFiltro = prs!nSaldoFinImporte
        Else
            GetBalanceSaldoCuentaFiltro = 0
        End If
        prs.Close
    End If
   
    
    Set prs = Nothing
    Exit Function

GetBalanceSaldoCuentaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:GetBalanceSaldoCuenta Method")
End Function


Public Function GetBalanceSaldoCuentaFiltroEX(psCtaCod As String, psAnio As String, psMes As String, psBalanceTipo As String, Optional pbFiltrarDeudorAcreedor As Boolean = False, Optional pbSaldoAcreedor As Boolean = False, Optional cOpeCod As String = "", Optional cCodigo As String = "") As Currency
   
    Dim lsCtaCaracter As String
    Dim CadSql   As String, CadSql2 As String
    Dim RSalAD   As ADODB.Recordset
    Dim lsCtaDos As String
    Dim oCta     As New DCtaCont
    
    '
    Dim sOtrasCuentas As String
   
   
    On Error GoTo GetBalanceSaldoCuentaErr
    
    sOtrasCuentas = ""
    
    If pbFiltrarDeudorAcreedor = True Then
        
        'FIltrar las cuentas que estan entre parentesis y dependen de la cuenta principal (SOLO SA Y SD)
        sOtrasCuentas = ""
        sOtrasCuentas = "Select cCtaContCodRef From RepBaseFormulaDet "
        sOtrasCuentas = sOtrasCuentas & " WHere cOpeCod='" & cOpeCod & "' and cCodigo='" & cCodigo & "' "
        sOtrasCuentas = sOtrasCuentas & " And cCtaContCod='" & psCtaCod & "'  "
        
        Set RSalAD = dbConec.CargaRecordSet(sOtrasCuentas)
        sOtrasCuentas = ""
        If Not RSalAD.EOF Then
            Do While Not RSalAD.EOF
                If Len(Trim(sOtrasCuentas)) = 0 Then
                    sOtrasCuentas = "'" & RSalAD!cCtaContCodRef & "'"
                Else
                    sOtrasCuentas = sOtrasCuentas & ", '" & RSalAD!cCtaContCodRef & "'"
                End If
                RSalAD.MoveNext
            Loop
        End If
        RSalAD.Close
        'Fin de filtro de otras cuentas
        
        'Sacar el tipo de cuenta registrado
        If Len(Trim(psCtaCod)) > 1 Then
            lsCtaDos = Mid(psCtaCod, 1, 2)
        Else
            lsCtaDos = Mid(psCtaCod, 1, 1)
        End If
       
        Set RSalAD = oCta.CargaCtaContClase(lsCtaDos)
        If Not RSalAD.EOF Then
            lsCtaCaracter = Trim(RSalAD!cCtaCaracter)
        End If
        RSalAD.Close
        CadSql2 = ""
    
        If Len(Trim(sOtrasCuentas)) = 0 Then
            If pbSaldoAcreedor Then  'si es saldo acreedor
                If lsCtaCaracter = "A" Then ' si es cuenta Acreedora
                    CadSql2 = " And B.nSaldoFinImporte > 0"
                Else                    'Si es deudora
                    CadSql2 = " And B.nSaldoFinImporte < 0"
                End If
            Else                'Si es Saldo deudor
               If lsCtaCaracter = "A" Then ' si es cuenta Acreedora
                   CadSql2 = " And B.nSaldoFinImporte < 0"
               Else                    'Si es deudora
                   CadSql2 = " And B.nSaldoFinImporte > 0"
               End If
            End If
        Else
            If pbSaldoAcreedor Then  'si es saldo acreedor
                If lsCtaCaracter = "A" Then ' si es cuenta Acreedora
                    CadSql2 = " HAVING SUM(B.nSaldoFinImporte) > 0"
                Else                    'Si es deudora
                    CadSql2 = " HAVING SUM(B.nSaldoFinImporte) < 0"
                End If
            Else                'Si es Saldo deudor
               If lsCtaCaracter = "A" Then ' si es cuenta Acreedora
                   CadSql2 = " HAVING SUM(B.nSaldoFinImporte) < 0"
               Else                    'Si es deudora
                   CadSql2 = " HAVING SUM(B.nSaldoFinImporte) > 0"
               End If
            End If
        End If
        
        
    End If
     
    CadSql = ""
    
    If Len(Trim(sOtrasCuentas)) = 0 Then
        psSql = "SELECT B.nSaldoFinImporte "
    Else
        psSql = "SELECT ISNULL(SUM(B.nSaldoFinImporte),0) as nSaldoFinImporte "
    End If
    
    psSql = psSql & " FROM BalanceEstad B "
    
    If Len(Trim(sOtrasCuentas)) = 0 Then
        psSql = psSql & " WHERE B.cCtaContCod = '" & psCtaCod & "' "
    Else
        psSql = psSql & " WHERE B.cCtaContCod IN('" & psCtaCod & "', " & sOtrasCuentas & ") "
    End If
    
    psSql = psSql & " and B.cBalanceAnio = '" & psAnio & "' "
    psSql = psSql & " and B.cBalanceMes = '" & psMes & "' "
    psSql = psSql & " and B.cBalanceCate = '2' "
    psSql = psSql & " and B.cBalanceTipo = '" & psBalanceTipo & "' "
    psSql = psSql & CadSql2
    

    
    Set prs = dbConec.CargaRecordSet(psSql, adLockReadOnly)
    If Not prs.EOF Then
        GetBalanceSaldoCuentaFiltroEX = prs!nSaldoFinImporte
    Else
        GetBalanceSaldoCuentaFiltroEX = 0
    End If
   
    prs.Close
    Set prs = Nothing
    Exit Function

GetBalanceSaldoCuentaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:GetBalanceSaldoCuenta Method")
End Function

Public Function ImprimeValidaBalance(psFechaIni As String, psFechaFin As String, pnLinPage As Integer, pnTipoValida As ValidacionBalance, Optional psEmprLog As String = "", Optional pbProcCierre As Boolean = False, Optional pnTipoCambio As Currency = 0) As String
Dim nLin   As Integer
Dim P      As Integer
Dim sTexto As String
Dim sTitulo As String
Dim prs     As ADODB.Recordset
Dim lsCabecera As String
Dim lsCabTipo  As String
   On Error GoTo ImprimeValidaBalanceErr
   nLin = nLin + 1
   sTitulo = "V A L I D A C I O N   D E   A S I E N T O S"
   Linea lsCabecera, Cabecera(sTitulo, P, "", pnLinPage, , , lsEmprLogo)
   nLin = 7
   lsCabTipo = ""
   Select Case pnTipoValida
      Case gValidaCuadreAsiento
            psSql = "SELECT m.cMovNro, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMEImporte),0) as nImporteME " _
               & "FROM   mov m JOIN movcta mc on m.nmovnro = mc.nmovnro " _
               & "        LEFT JOIN movme me on me.nmovnro = mc.nmovnro and me.nmovitem = mc.nmovitem " _
               & "WHERE  m.nMovEstado = " & gMovEstContabMovContable & " and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and substring(cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & "' and " _
               & "       NOT mc.cCtaContCod IS NULL and mc.cCtaContcod <> '' " _
               & "GROUP BY m.cMovNro " _
               & "HAVING Sum(mc.nMovImporte) <> 0 "
            Linea lsCabTipo, "ASIENTOS DESCUADRADOS ", 2
    
      Case gValidaConvesionME
           If gcConvTipo = "1" Then
            psSql = "SELECT m.cmovnro," _
               & "     SUM(CASE WHEN substring(mc.cCtaContCod,3,1) = '1' THEN mc.nMovImporte" _
               & "              WHEN substring(mc.cCtaContCod,1,1) IN ('4','5','6') and substring(mc.cCtaContCod,3,1) = '2' THEN mc.nMovImporte END) as nDife " _
               & "FROM   mov m join movcta mc on m.nMovnro = mc.nMovnro " _
               & "WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and " _
               & "       substring(cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & "' and " _
               & "       EXISTS (SELECT nMovNro FROM movcta mcr WHERE (substring(mcr.cCtaContCod,3,1) = '1' or substring(mcr.cCtaContCod,1,1) IN ('4','5','6') ) and mcr.nmovnro = mc.nmovnro) and " _
               & "       EXISTS (SELECT nMovNro FROM movcta mcr WHERE  substring(mcr.cCtaContCod,3,1) = '2' and mcr.nmovnro = mc.nmovnro) " _
               & "GROUP BY m.cMovNro " _
               & "HAVING SUM(CASE WHEN substring(mc.cCtaContCod,3,1) = '1' THEN mc.nMovImporte " _
               & "                WHEN substring(mc.cCtaContCod,1,1) IN ('4','5','6') and substring(mc.cCtaContCod,3,1) = '2' THEN mc.nMovImporte END) <> 0 "
            Else
               psSql = ""
               psSql = psSql & "SELECT m.cmovnro, me.nMovMEImporte nMovMEImporte, mc.nMovImporte nMovImporte, "
               psSql = psSql & "       " & pnTipoCambio & " nTipoCambio, convert(decimal(15,2), me.nMovMEImporte * " & pnTipoCambio & " ) Conversion "
               psSql = psSql & "FROM   Mov m join movcta mc on m.nMovnro = mc.nMovnro "
               psSql = psSql & "       join movme me on me.nmovnro = mc.nmovnro and me.nMovItem = mc.nMovItem "
               psSql = psSql & "WHERE  m.nMovEstado = '10' and not m.nMovFlag in (1,5) and "
               psSql = psSql & "       mc.cCtaContCod like '[1278]_2%' and "
               psSql = psSql & "       substring(cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & "' "
               psSql = psSql & "       and not convert(decimal(15,2), me.nMovMEImporte * " & pnTipoCambio & ")  - mc.nMovImporte between -0.02 and 0.02 "
               psSql = psSql & "       and not m.cMovNro LIKE '%XXX_' "
            End If
            Linea lsCabTipo, "ASIENTOS DE MONEDA EXTRANJERA ", 2
      
      Case gValidaCuentasNoExistentes
            psSql = "SELECT m.cMovNro, mc.* " _
                 & "FROM   movcta mc JOIN mov m ON m.nmovnro = mc.nmovnro LEFT JOIN " & sCentralCom & "CtaCont c ON c.cCtaContCod = mc.cCtaContCod " _
                 & "WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and c.cCtaContCod IS NULL and mc.cCtaContCod <> '' and " _
                 & "       substring(cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & "' " _
                 & "order by mc.cCtaContCod"
            Linea lsCabTipo, "ASIENTOS CON CUENTAS NO EXISTENTES ", 2
      
      Case gValidaCuentasNoExistentes2
            psSql = "SELECT m.cMovNro, mc.* FROM movcta mc JOIN MOV m ON m.nMovNro = mc.nMovNro JOIN " & sCentralCom & "CtaCont c ON c.cCtaContCod = mc.cCtaContCod " _
                 & "WHERE  c.cCtaContDesc = '**' and m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and " _
                 & "       substring(m.cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & "' " _
                 & "order by mc.cCtaContCod"
            Linea lsCabTipo, "ASIENTOS CON CUENTAS ** ", 2
            
      Case gValidaCuentasAnaliticas
            psSql = "SELECT m.cMovNro, mc.nMovNro, mc.nMovItem, mc.cCtaContcod, count(*) " _
                 & "FROM   movcta mc join mov m    on m.nmovnro = mc.nmovnro " _
                 & "       join " & sCentralCom & "CtaCont c  ON c.cCtaContCod like mc.cCtaContCod + '%' " _
                 & "WHERE  m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and " _
                 & "       substring(m.cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & "' and " _
                 & "       mc.cCtaContCod <> '' And nCtaEstado = 1 " _
                 & "GROUP BY m.cMovNro, mc.nMovNro, mc.nMovItem, mc.cCtaContcod " _
                 & "HAVING count(*) > 1 ORDER BY m.cmovnro "
            Linea lsCabTipo, "ASIENTOS CON CUENTAS ANALITICAS ", 2
            
      Case gValidaCuentasDeOrden
            psSql = "Select substring(m.cmovnro,1,8) as cDia, substring(m.cmovnro,9,6) as cHora, substring(m.cmovnro,15,2) as cAge, right(m.cmovnro,4) as cTipo, SUBSTRING(cCtaContCod,3,1) , " _
                 & "   ISNULL(sum(CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte END),0) as nDebe, " _
                 & "   ISNULL(sum(CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte*-1 END),0) as nHaber " _
                 & "From movcta mc join mov m on m.nmovnro = mc.nmovnro " _
                 & "where m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and " _
                 & "      substring(m.cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & "' " _
                 & "      and mc.cctacontcod like '8[12]%' " _
                 & "group by substring(m.cmovnro,1,8), substring(m.cmovnro,9,6), substring(m.cmovnro,15,2), right(m.cmovnro,4), SUBSTRING(cCtaContCod,3,1)  " _
                 & "having ISNULL(sum(CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte END),0) - ISNULL(sum(CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte*-1 END),0) <> 0 " _
                 & "UNION ALL "
         psSql = psSql & "Select substring(m.cmovnro,1,8) as cDia, substring(m.cmovnro,9,6) as cHora, substring(m.cmovnro,15,2) as cAge, right(m.cmovnro,4) as cTipo, SUBSTRING(cCtaContCod,3,1) , " _
                 & "   ISNULL(sum(CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte END),0) as nDebe, " _
                 & "   ISNULL(sum(CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte*-1 END),0) as nHaber " _
                 & "From movcta mc join mov m on m.nmovnro = mc.nmovnro " _
                 & "where m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and " _
                 & "      substring(m.cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & "' " _
                 & "      and mc.cctacontcod like '8[34]%' " _
                 & "group by substring(m.cmovnro,1,8), substring(m.cmovnro,9,6), substring(m.cmovnro,15,2), right(m.cmovnro,4), SUBSTRING(cCtaContCod,3,1)  " _
                 & "having ISNULL(sum(CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte END),0) - ISNULL(sum(CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte*-1 END),0) <> 0 " _
                 & "UNION ALL "
         psSql = psSql & "Select substring(m.cmovnro,1,8) as cDia, substring(m.cmovnro,9,6) as cHora, substring(m.cmovnro,15,2) as cAge, right(m.cmovnro,4) as cTipo, SUBSTRING(cCtaContCod,3,1) , " _
                 & "   ISNULL(sum(CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte END),0) as nDebe, " _
                 & "   ISNULL(sum(CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte*-1 END),0) as nHaber " _
                 & "From movcta mc join mov m on m.nmovnro = mc.nmovnro " _
                 & "where m.nMovEstado = '" & gMovEstContabMovContable & "' and not m.nMovFlag in (" & gMovFlagEliminado & "," & gMovFlagModificado & ") and " _
                 & "      substring(m.cMovNro,1,8) BETWEEN '" & psFechaIni & "' and '" & psFechaFin & "' " _
                 & "      and mc.cctacontcod like '8[56]%' " _
                 & "group by substring(m.cmovnro,1,8), substring(m.cmovnro,9,6), substring(m.cmovnro,15,2), right(m.cmovnro,4), SUBSTRING(cCtaContCod,3,1)  " _
                 & "having ISNULL(sum(CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte END),0) - ISNULL(sum(CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte*-1 END),0) <> 0 "
                 
            Linea lsCabTipo, "VALIDACION DE CUENTAS DE ORDEN  ", 2
            
      Case gValidaSaldosContables
         psSql = "Select cs.cCtaContCod, cs.dCtaSaldoFecha, cs.nCtaSaldoImporte from CtaSaldo cs, CtaCont c " _
               & "where cs.dCtaSaldoFecha = (Select MAX(dCtaSaldoFecha) FROM CtaSaldo cs1 WHERE cs1.cCtaContCod = cs.cCtaContCod and cs1.dCtaSaldoFecha <= '" & psFechaFin & "') " _
               & "  and nCtaSaldoImporte <> 0 and c.cCtaContCod LIKE cs.cCtaContCod + '%' And c.nCtaEstado = 1" _
               & "group by cs.cCtaContCod, cs.dCtaSaldoFecha, cs.nCtaSaldoImporte " _
               & "Having count(*) > 1 "
            Linea lsCabTipo, "VALIDACION DE SALDOS CONTABLES ", 2
      Case gValidaCuentasSinPadre
        
        psSql = "select cCtaContCod, cCtaContDesc "
        psSql = psSql & " From CtaCont"
        psSql = psSql & " Where nctaestado = 1 "
        psSql = psSql & " and len(cctacontcod) >2 "
        psSql = psSql & " and substring(cCtaContCod, 1, len(cctacontcod) -2) not in(  select cCtaContCod "
        psSql = psSql & " From CtaCont "
        psSql = psSql & " Where nctaestado = 1 "
        psSql = psSql & " ) order by substring(cCtaContCod, 3,1), cCtaContCod "
        
            Linea lsCabTipo, "VALIDACION DE CUENTAS SIN PADRE", 2
        
   End Select
   Set prs = dbConec.CargaRecordSet(psSql)
   If prs.EOF Then
      If Not pbProcCierre Then
         sTexto = lsCabTipo
         Linea sTexto, " *** NO SE ENCONTRARON OBSERVACIONES *** ", 2
         nLin = nLin + 2
      End If
   Else
      sTexto = lsCabTipo
      Do While Not prs.EOF
         DoEvents
         If nLin > pnLinPage - 6 Then
            Linea sTexto, Cabecera(sTitulo, P, "", pnLinPage, , , lsEmprLogo), 0
            nLin = 4
         End If
         Select Case pnTipoValida
            Case gValidaCuadreAsiento
               Linea sTexto, prs!cMovNro & "  " & PrnVal(prs!nImporte, 16, 2) & "    " & PrnVal(prs!nImporteME, 16, 2)
            Case gValidaConvesionME
               
               If gcConvTipo = "1" Then
                  Linea sTexto, prs!cMovNro & "  " & PrnVal(prs!nDife, 16, 2)
               Else
                  Linea sTexto, prs!cMovNro & "  " & PrnVal(prs!nMovMEImporte, 16, 2) & "  " & PrnVal(prs!nMovImporte, 16, 2) & "  " & PrnVal(prs!Conversion, 16, 2)
               End If
            Case gValidaCuentasNoExistentes
               Linea sTexto, prs!cMovNro & " " & prs!nMovItem & " " & prs!cCtaContCod & " " & PrnVal(prs!nMovImporte, 16, 2)
            Case gValidaCuentasNoExistentes2
               Linea sTexto, prs!cMovNro & " " & prs!nMovItem & " " & prs!cCtaContCod & " " & PrnVal(prs!nMovImporte, 16, 2)
            Case gValidaCuentasAnaliticas
               Linea sTexto, prs!cMovNro & "    " & prs!cCtaContCod
            Case gValidaCuentasDeOrden
               Linea sTexto, prs!cdia & "-" & prs!cHora & "-" & prs!cAge & "-" & prs!cTipo & "  " & PrnVal(prs!nDebe, 16, 2) & " " & PrnVal(prs!nHaber, 16, 2) & " " & PrnVal(prs!nDebe - prs!nHaber, 16, 2)
            Case gValidaSaldosContables
               Linea sTexto, Justifica(prs!cCtaContCod, 24) & "  " & prs!dCtaSaldoFecha & "   " & PrnVal(prs!nCtaSaldoImporte, 15, 2)
            Case gValidaCuentasSinPadre
               Linea sTexto, Justifica(prs!cCtaContCod, 24) & "  " & prs!cCtaContDesc
         End Select
         prs.MoveNext
      Loop
   End If
   RSClose prs
   If sTexto <> "" Then
      ImprimeValidaBalance = lsCabecera & sTexto
   End If
Exit Function
ImprimeValidaBalanceErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:ImprimeValidaBalance Method")
End Function

Public Function GetTipCambioBalance(psMovFecha As String) As Currency
Dim oCon As DConecta
Set oCon = New DConecta
 
On Error GoTo getTipCambioBalanceErr
   psSql = "SELECT nMovOtroImporte FROM MovOtrosItem MO JOIN Mov M ON M.nMovNro = MO.nMovNro WHERE M.nMovEstado = '" & gMovEstContabMovContable & "' and not M.nMovFlag IN ('" & gMovFlagEliminado & "','" & gMovFlagDeExtorno & "','" & gMovFlagExtornado & "','" & gMovFlagModificado & "') and " _
         & "       M.cMovNro LIKE '" & psMovFecha & "%' and cmovotrovariable = 'TC2'"
   Set prs = dbConec.CargaRecordSet(psSql)
   If Not prs.EOF Then
      GetTipCambioBalance = prs!nMovOtroImporte
   Else
      If gbBitCentral Then
          psSql = "SELECT nValFijo FROM " & sCentralCom & "TipoCambio WHERE LEFT(convert(varchar(8),dFecCamb,112),6) LIKE '" & Left(psMovFecha, 6) & "%' "
          Set prs = dbConec.CargaRecordSet(psSql)
      Else
          oCon.AbreConexionRemota "01"
          psSql = "SELECT nValFijo FROM DBcomunes..TipCambio WHERE LEFT(convert(varchar(8),dFecCamb,112),6) LIKE '" & Left(psMovFecha, 6) & "%' "
          Set prs = oCon.CargaRecordSet(psSql)
      End If
      
      If Not prs.EOF Then
         GetTipCambioBalance = prs!nValFijo
      Else
         GetTipCambioBalance = 0
      End If
   End If
   prs.Close: Set prs = Nothing
Exit Function
getTipCambioBalanceErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:getTipCambioBalance Method")
End Function

Public Function GetUtilidadAcumulada(psTipoBala As String, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, Optional pbGetUltimo As Boolean = False, Optional pbMsg As Boolean = True) As Currency
On Error GoTo GetUtilidadAcumuladaErr
GetUtilidadAcumulada = 0
If pbGetUltimo Then
   psSql = "SELECT nUtilidadImporte FROM UtilidadAcumula U " _
     & "WHERE  cBalanceCate = '" & psTipoBala & "' " & IIf(pnMoneda = -1, "", " and cBalanceTipo = '" & pnMoneda & "'") _
     & " and cUtilidadAnio+cUtilidadMes = (SELECT MAX(cUtilidadAnio+cUtilidadMes) FROM UtilidadAcumula WHERE cBalanceCate = U.cBalanceCate and cBalanceTipo = U.cBalanceTipo and cUtilidadAnio+cUtilidadMes <= '" & Format(pnAnio, "0000") & Format(pnMes, "00") & "' )"
Else
   psSql = "SELECT nUtilidadImporte FROM UtilidadAcumula " _
     & "WHERE  cBalanceCate = '" & psTipoBala & "' and cBalanceTipo = '" & pnMoneda _
     & "' and cUtilidadMes = '" & Format(pnMes, "00") & "' and " _
     & "cUtilidadAnio = '" & Format(pnAnio, "0000") & "'"
End If
Set prs = dbConec.CargaRecordSet(psSql)
If prs.EOF Then
   If pbMsg Then
      RSClose prs
'      Err.Raise 50001, "NBalanceCont: ValidaBalance", _
'          "No se encontraron datos de Utilidad Acumulada del Mes anterior." & oImpresora.gPrnSaltoLinea & "Por favor consultar con Sistemas...!"
      MsgBox "No se encontraron datos de Utilidad Acumulada del Mes anterior"
   End If
   Exit Function
End If
GetUtilidadAcumulada = prs!nUtilidadImporte
prs.Close
Set prs = Nothing
Exit Function
GetUtilidadAcumuladaErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:GetUtilidadAcumulada Method")
End Function

Public Function getImporteBalanceMes(psCtaCod As String, pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer) As Currency
On Error GoTo getImporteBalanceMesErr
getImporteBalanceMes = 0
psSql = "SELECT SubString(cCtaContCod,1,1) as cCtaContCod, " _
     & "       nSaldoFinImporte as nImporte " _
     & "From BalanceEstad WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
     & "     ' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
     & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' and cCtaContCod = '" & psCtaCod & "' "
Set prs = dbConec.CargaRecordSet(psSql)
If Not prs.EOF Then
   getImporteBalanceMes = prs!nImporte
End If
Exit Function
getImporteBalanceMesErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:getImporteBalanceMes Method")
End Function

Public Function ValidaBalance(lSoloUtilidad As Boolean, pdFechaIni As Date, pdFechaFin As Date, pnTipoBala As Integer, pnMoneda As Integer) As String
Dim nUtilidad As Currency
Dim nUtilidadMes As Currency
Dim nRei As Currency
Dim nDeduccion As Currency
Dim sValida    As String
Dim n5 As Currency, n4 As Currency
Dim n62 As Currency, n63 As Currency, n64 As Currency, n65 As Currency, n66 As Currency
nUtilidad = 0
nUtilidadMes = 0
If Month(pdFechaIni) > 1 Then
   nUtilidad = GetUtilidadAcumulada(Format(pnTipoBala, "#"), pnMoneda, Format(Month(pdFechaIni - 1), "00"), Format(Year(pdFechaIni - 1), "0000"))
End If

n5 = getImporteBalanceMes("5", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n62 = getImporteBalanceMes("62", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n63 = getImporteBalanceMes("63", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n64 = getImporteBalanceMes("64", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n65 = getImporteBalanceMes("65", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n66 = getImporteBalanceMes("66", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
n4 = getImporteBalanceMes("4", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nUtilidadMes = n5 + n62 + n64 - (n4 + n63 + n65)

'69
nRei = getImporteBalanceMes("69", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion = getImporteBalanceMes("67", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion = nDeduccion + getImporteBalanceMes("68", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
nDeduccion = nDeduccion * -1

If Not lSoloUtilidad Then
   Linea sValida, Cabecera("C A L C U L O   D E   L A   U T I L I D A D", 0, "", 80, , , lsEmprLogo)
   If pnMoneda = 0 Then
      Linea sValida, Centra("C O N S O L I D A D O", 80)
   End If
   Linea sValida, Centra("AL " & pdFechaFin, 80), 3
End If
Linea sValida, String(2, oImpresora.gPrnSaltoLinea) & Space(5) & Justifica("UTILIDAD ACUMULADA AL " & CDate(pdFechaIni) - 1, 40) & ": " & PrnVal(nUtilidad, 16, 2)
Linea sValida, Space(5) & Justifica("UTILIDAD DEL MES DE " & Mid(pdFechaFin, 4, 10), 40) & ": " & PrnVal(nUtilidadMes - nUtilidad, 16, 2)
Linea sValida, Space(45) & "--------------------"
Linea sValida, Space(5) & Justifica("UTILIDAD ACUMULADA AL " & pdFechaFin, 40) & ": " & PrnVal(nUtilidadMes, 16, 2)
Linea sValida, Space(5) & Justifica("R.E.I. " & Right(pdFechaFin, 4), 40) & ": " & PrnVal(nRei, 16, 2)
Linea sValida, Space(45) & "===================="
Linea sValida, Space(5) & Justifica("UTILIDAD ANTES DE IMPUESTO", 40) & ": " & PrnVal(nUtilidadMes + nRei, 16, 2)
Linea sValida, Space(5) & Justifica("IMP. A LA RENTA Y PARTICIP.LAB.", 40) & ": " & PrnVal(nDeduccion, 16, 2)
Linea sValida, Space(45) & "--------------------"
Linea sValida, Space(5) & Justifica("UTILIDAD(PERDIDA) NETA AL " & pdFechaFin, 40) & ": " & PrnVal(nRei + nUtilidadMes + nDeduccion, 16, 2)
Linea sValida, Space(45) & "====================", 2

If Not lSoloUtilidad Then

   Dim nActivo As Currency
   Dim nPasivo As Currency
   Dim nPatri  As Currency
   
   'Eliminamos si Existe la Utilidad Acumulada del Mes
   Dim dBalance As New DbalanceCont
   dBalance.EliminaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), True
   dBalance.InsertaUtilidadAcumulada pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni), nUtilidadMes, True
   dBalance.EjecutaBatch
   
   nActivo = getImporteBalanceMes("1", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPasivo = getImporteBalanceMes("2", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   nPatri = getImporteBalanceMes("3", pnTipoBala, pnMoneda, Month(pdFechaIni), Year(pdFechaIni))
   
   Linea sValida, BON & Centra(" CONSISTENCIA DE CUADRE DEL BALANCE ", 80) & BOFF
   Select Case pnMoneda
      Case 0: Linea sValida, Centra(" ( CONSOLIDADO ) ", 80), 2
      Case 1: Linea sValida, Centra(" ( MONEDA NACIONAL ) ", 80), 2
      Case 2: Linea sValida, Centra(" ( MONEDA EXTRANJERA ) ", 80), 2
   End Select
   Linea sValida, Justifica("ACTIVO", 40) & PrnVal(nActivo, 16, 2)
   Linea sValida, Justifica("PASIVO", 40) & Space(18) & PrnVal(nPasivo, 16, 2)
   Linea sValida, Justifica("PATRIMONIO", 40) & Space(18) & PrnVal(nPatri, 16, 2)
   Linea sValida, Justifica("UTILIDAD (PERDIDA) NETA", 40) & Space(18) & PrnVal(nRei + nUtilidadMes + nDeduccion, 16, 2)
   Linea sValida, Space(39) & String(36, "-")
   Linea sValida, Space(40) & PrnVal(nActivo, 16, 2) & "  " & PrnVal(nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion, 16, 2)
   Linea sValida, Space(39) & String(36, "="), 2
   
   Linea sValida, Space(39) & String(18, "-")
   Linea sValida, Justifica("DIFERENCIA", 40) & PrnVal(nActivo - (nPasivo + nPatri + nRei + nUtilidadMes + nDeduccion), 16, 2)
   Linea sValida, Space(39) & String(18, "-")
End If
ValidaBalance = sValida
End Function

Public Function BalanceGeneradoHisto(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer) As Boolean
On Error GoTo BalanceGeneradoErr
BalanceGeneradoHisto = True
psSql = "SELECT TOP 1 cCtaContCod as cCtaContCod " _
     & "FROM   BalanceEstad " _
     & "WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
     & "' and cBalanceMes = '" & Format(pnMes, "00") & "' " _
     & "  and cBalanceAnio = '" & Format(pnAnio, "0000") & "' "
Set prs = dbConec.CargaRecordSet(psSql)
If prs.EOF Then
   BalanceGeneradoHisto = False
End If
prs.Close: Set prs = Nothing
Exit Function
BalanceGeneradoErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:BalanceGenerado Method")
End Function

Public Function BalanceGeneradoFormaAB(psTipo As String, psFecha As String) As Boolean
On Error GoTo BalanceGeneradoFormaABErr
BalanceGeneradoFormaAB = True
psSql = "SELECT cFecha From BalanceGen Where cFecha='" & psFecha & "' and cBalanceCate = '" & psTipo & "'"
Set prs = dbConec.CargaRecordSet(psSql)
If prs.EOF Then
   BalanceGeneradoFormaAB = False
End If
prs.Close: Set prs = Nothing
Exit Function
BalanceGeneradoFormaABErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:BalanceGeneradoFormaAB Method")
End Function

Public Function BalanceGeneradoSectorial(pdfecha As Date) As Boolean
On Error GoTo BalanceGeneradoSectorialErr
BalanceGeneradoSectorial = True
   psSql = "Select TOP 1 * from BalanceSectorial Where nBalanceMes = " & Month(pdfecha) & " and nBalanceAnio = " & Year(pdfecha)
   Set prs = dbConec.CargaRecordSet(psSql)
   If prs.EOF Then
      BalanceGeneradoSectorial = False
   End If
prs.Close: Set prs = Nothing
Exit Function
BalanceGeneradoSectorialErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:BalanceGeneradoFormaAB Method")
End Function

Public Function LeeBalanceHisto(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, Optional psCtaIni As String = "", Optional psCtaFin As String = "", Optional pnDigitos As Integer = 0, Optional pbSoloAnaliticas As Boolean = False, Optional pbConCalculoUtilidad As Boolean = False, Optional pbSinAgencias As Boolean = False) As Recordset
Dim lsCtaFiltro As String
On Error GoTo LeeBalanceHistoErr

If psCtaIni <> "" Then
    lsCtaFiltro = " and b.cCtaContCod >= '" & psCtaIni & "'"
End If
If psCtaFin <> "" Then
    lsCtaFiltro = lsCtaFiltro & " and b.cCtaContCod <= '" & psCtaFin & "'"
End If

If pbConCalculoUtilidad = False Then
   If pbSinAgencias Then
   
      psSql = "SELECT b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte " & _
              "  FROM BalanceEstad b join " & sCentralCom & "CtaCont c ON c.cCtaContCod = b.cCtaContCod " & _
                 IIf(pbSoloAnaliticas, " join CtaCont c1 ON c1.cCtaContCod LIKE c.cCtaContCod + '%' ", "") & _
              " WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "' and " & _
              "       cBalanceMes  = '" & Format(pnMes, "00") & "' and " & _
              "       cBalanceAnio = '" & Format(pnAnio, "0000") & "' and " & _
              "       b.cCtaContCod not in (SELECT b.cCtaContCod from BalanceEstad b join " & sCentralCom & "CtaCont c ON c.cCtaContCod = b.cCtaContCod " & _
              "       inner join CtaContSinAge a on b.cCtaContCod like a.cCtaContCod+'%' " & _
              "       WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "' and " & _
              "       cBalanceMes  = '" & Format(pnMes, "00") & "' and " & _
              "       cBalanceAnio = '" & Format(pnAnio, "0000") & "' and len(b.cCtaContCod)>len(a.cCtaContCod) ) " & _
              " " & lsCtaFiltro _
            & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "") _
            & " ORDER BY b.cCtaContCod "
            
   Else
      psSql = "SELECT b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte " _
            & "FROM   BalanceEstad b join " & sCentralCom & "CtaCont c ON c.cCtaContCod = b.cCtaContCod " _
            & IIf(pbSoloAnaliticas, " join CtaCont c1 ON c1.cCtaContCod LIKE c.cCtaContCod + '%' ", "") _
            & "WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
            & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
            & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' " & lsCtaFiltro _
            & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "") _
            & " ORDER BY b.cCtaContCod "
   End If
Else
    psSql = "SELECT b.cCtaContCod cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte " _
         & "FROM   BalanceEstad b join " & sCentralCom & "CtaCont c ON c.cCtaContCod = b.cCtaContCod " _
         & IIf(pbSoloAnaliticas, " join CtaCont c1 ON c1.cCtaContCod LIKE c.cCtaContCod + '%' ", "") _
         & "WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' And " _
         & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' " & lsCtaFiltro _
         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "") _
    
'    psSql = psSql & " Union " _
'         & " SELECT '61' cCtaContCod, (Select cCtaContDesc from ctacont where cctacontcod = '61') cCtaContDesc , 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, Sum(Case b.cCtaContCod When '5' then nSaldoFinImporte else nSaldoFinImporte * -1 End) nSaldoFinImporte" _
'         & " FROM   BalanceEstad b WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
'         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
'         & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' And b.cCtaContCod in ('4','5') " & lsCtaFiltro _
'         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "")
'
'    psSql = psSql & " Union " _
'         & " SELECT '6101' cCtaContCod, (Select cCtaContDesc from ctacont where cctacontcod = '6101') cCtaContDesc , 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, Sum(Case b.cCtaContCod When '5' then nSaldoFinImporte else nSaldoFinImporte * -1 End) nSaldoFinImporte" _
'         & " FROM   BalanceEstad b WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
'         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
'         & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' And b.cCtaContCod in ('4','5') " & lsCtaFiltro _
'         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "")
'
'    psSql = psSql & " Union " _
'             & " SELECT '66' cCtaContCod, (Select cCtaContDesc from ctacont where cctacontcod = '66') cCtaContDesc , 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, Sum(Case When b.cCtaContCod In ('5','62','64') then nSaldoFinImporte else nSaldoFinImporte * -1 End) nSaldoFinImporte" _
'         & " FROM   BalanceEstad b WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
'         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
'         & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' And b.cCtaContCod in ('4','5','62','63','64','65') " & lsCtaFiltro _
'         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "") _
'
'    psSql = psSql & " Union " _
'             & " SELECT '6601' cCtaContCod, (Select cCtaContDesc from ctacont where cctacontcod = '6601') cCtaContDesc , 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, Sum(Case When b.cCtaContCod In ('5','6201','6401') then nSaldoFinImporte else nSaldoFinImporte * -1 End) nSaldoFinImporte" _
'         & " FROM   BalanceEstad b WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
'         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
'         & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' And b.cCtaContCod in ('4','5','6201','6301','6401','6501') " & lsCtaFiltro _
'         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "") _
'
'    psSql = psSql & " Union " _
'         & " SELECT '6901' cCtaContCod, (Select cCtaContDesc from ctacont where cctacontcod = '6901') cCtaContDesc , 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, Sum(Case When b.cCtaContCod In ('5','62','64') then nSaldoFinImporte else nSaldoFinImporte * -1 End) nSaldoFinImporte" _
'         & " FROM   BalanceEstad b WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
'         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
'         & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' And b.cCtaContCod in ('4','5','62','63','64','65','67','68') " & lsCtaFiltro _
'         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "") _
'
'    psSql = psSql & " Union " _
'         & " SELECT '69' cCtaContCod, (Select cCtaContDesc from ctacont where cctacontcod = '69') cCtaContDesc , 0 nSaldoIniImporte, 0 nDebe, 0 nHaber, Sum(Case When b.cCtaContCod In ('5','62','64','69') then nSaldoFinImporte else nSaldoFinImporte * -1 End) nSaldoFinImporte" _
'         & " FROM   BalanceEstad b WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
'         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
'         & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' And b.cCtaContCod in ('4','5','62','63','64','65','67','68','69') " & lsCtaFiltro _
'         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "")
         
    psSql = psSql & " Union " _
             & " SELECT '39' cCtaContCod, (Select cCtaContDesc from ctacont where cctacontcod = '39') cCtaContDesc , Sum(Case When b.cCtaContCod In ('5','62','64','69') then nSaldoIniImporte else nSaldoIniImporte * -1 End) nSaldoIniImporte, Sum(Case When b.cCtaContCod In ('5','62','64','69') then nDebe * -1 else nDebe  End) nDebe, Sum(Case When b.cCtaContCod In ('5','62','64','69') then nHaber else nHaber * -1 End) nHaber, Sum(Case When b.cCtaContCod In ('5','62','64','69') then nSaldoFinImporte else nSaldoFinImporte * -1 End) nSaldoFinImporte" _
         & " FROM   BalanceEstad b WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
         & "cBalanceAnio = '" & Format(pnAnio, "0000") & "' And b.cCtaContCod in ('4','5','62','63','64','65','67','68','69') " & lsCtaFiltro _
         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "") _
         & "  "

    psSql = psSql & " Union " _
             & "  SELECT '3911' cCtaContCod, (Select cCtaContDesc from ctacont where cctacontcod = '3911') cCtaContDesc , Sum(Case When b.cCtaContCod like '5_1%' Or b.cCtaContCod like '621%' Or b.cCtaContCod Like '641%' Or b.cCtaContCod Like '691%' then nSaldoIniImporte else nSaldoIniImporte * -1 End) nSaldoIniImporte, Sum(Case When b.cCtaContCod like '5_1%' Or b.cCtaContCod like '621%' Or b.cCtaContCod Like '641%' Or b.cCtaContCod Like '691%' then nDebe * -1 else nDebe  End) nDebe, Sum(Case When b.cCtaContCod like '5_1%' Or b.cCtaContCod like '621%' Or b.cCtaContCod Like '641%' Or b.cCtaContCod Like '691%' then nHaber else nHaber * -1 End) nHaber, Sum(Case When b.cCtaContCod like '5_1%' Or b.cCtaContCod like '621%' Or b.cCtaContCod Like '641%' Or b.cCtaContCod Like '691%' then nSaldoFinImporte else nSaldoFinImporte * -1 End) nSaldoFinImporte" _
         & " FROM   BalanceEstad b WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' And " _
         & " cBalanceAnio = '" & Format(pnAnio, "0000") & "' And " _
         & " (b.cCtaContCod Like '4_1%' Or b.cCtaContCod Like '5_1%' Or b.cCtaContCod Like '621%' Or " _
         & "  b.cCtaContCod Like '621%' Or b.cCtaContCod Like '631%' Or b.cCtaContCod Like '641%' Or" _
         & "  b.cCtaContCod Like '651%' Or b.cCtaContCod Like '671%' Or b.cCtaContCod Like '681%' Or" _
         & "  b.cCtaContCod Like '691%') " & lsCtaFiltro _
         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "")

    psSql = psSql & " Union " _
             & "  SELECT '3921' cCtaContCod, (Select cCtaContDesc from ctacont where cctacontcod = '3921') cCtaContDesc , Sum(Case When b.cCtaContCod like '5_2%' Or b.cCtaContCod like '622%' Or b.cCtaContCod Like '642%' Or b.cCtaContCod Like '692%' then nSaldoIniImporte else nSaldoIniImporte * -1 End) nSaldoIniImporte, Sum(Case When b.cCtaContCod like '5_2%' Or b.cCtaContCod like '622%' Or b.cCtaContCod Like '642%' Or b.cCtaContCod Like '692%' then nDebe * -1 else nDebe  End) nDebe, Sum(Case When b.cCtaContCod like '5_2%' Or b.cCtaContCod like '62%' Or b.cCtaContCod Like '642%' Or b.cCtaContCod Like '692%' then nHaber else nHaber * -1 End) nHaber, Sum(Case When b.cCtaContCod like '5_2%' Or b.cCtaContCod like '622%' Or b.cCtaContCod Like '642%' Or b.cCtaContCod Like '692%' then nSaldoFinImporte else nSaldoFinImporte * -1 End) nSaldoFinImporte" _
         & " FROM   BalanceEstad b WHERE len(b.cCtaContCod) <= " & pnDigitos & " and cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda _
         & "' and cBalanceMes = '" & Format(pnMes, "00") & "' And " _
         & " cBalanceAnio = '" & Format(pnAnio, "0000") & "' And " _
         & " (b.cCtaContCod Like '4_2%' Or b.cCtaContCod Like '5_2%' Or b.cCtaContCod Like '622%' Or " _
         & "  b.cCtaContCod Like '622%' Or b.cCtaContCod Like '632%' Or b.cCtaContCod Like '642%' Or" _
         & "  b.cCtaContCod Like '652%' Or b.cCtaContCod Like '672%' Or b.cCtaContCod Like '682%' Or" _
         & "  b.cCtaContCod Like '692%') " & lsCtaFiltro _
         & IIf(pbSoloAnaliticas, " GROUP BY b.cCtaContCod, c.cCtaContDesc, nSaldoIniImporte, nDebe, nHaber, nSaldoFinImporte HAVING count(*) = 1", "") _
         & " ORDER BY cCtaContCod "

End If

   Set LeeBalanceHisto = dbConec.CargaRecordSet(psSql)
Exit Function
LeeBalanceHistoErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:LeeBalanceHisto Method")
End Function


Public Function TotalizaBalanceHisto(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer) As Recordset
On Error GoTo TotalizaBalanceHistoErr
psSql = "SELECT SUM(nDebe) as nDebe, SUM(nHaber) as nHaber FROM BalanceEstad " _
     & "WHERE cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "' " _
     & "  and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
     & "      cBalanceAnio = '" & Format(pnAnio, "0000") & "' and LEN(cCtaContCod) = 1 "
Set TotalizaBalanceHisto = dbConec.CargaRecordSet(psSql)
Exit Function
TotalizaBalanceHistoErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:LeeBalanceHisto Method")
End Function

Public Function GetBalanceSituacion(pnTipoBala As Integer, pnMoneda As Integer, pnMes As Integer, pnAnio As Integer, psCaracter As String) As Recordset
On Error GoTo GetBalanceSituacionErr
psSql = "SELECT Substring(B.cCtaContCod,1,2) as cCtaContCod,nSaldoIniImporte as nInicial, nDebe as nDebe, nHaber as nHaber " _
     & "FROM   BalanceEstad B JOIN " & sCentralCom & "CtaContClase cls ON B.cCtaContcod LIKE RTRIM(cls.cCtaContCod) + '%' " _
     & "WHERE  cBalanceCate = '" & pnTipoBala & "' and cBalanceTipo = '" & pnMoneda & "' " _
     & "   and cBalanceMes = '" & Format(pnMes, "00") & "' and " _
     & "       cBalanceAnio = '" & Format(pnAnio, "0000") & "' and cls.cCtaCaracter = '" & psCaracter & "' " _
     & "and Len(B.cCtaContCod) = 2 and (nSaldoIniImporte <> 0 or nDebe <> 0 or nHaber<> 0) " _
     & "ORDER  BY B.cCtaContCod "
Set GetBalanceSituacion = dbConec.CargaRecordSet(psSql)
Exit Function
GetBalanceSituacionErr:
   Call RaiseError(MyUnhandledError, "NBalanceCont:GetBalanceSituacion Method")
End Function

Public Function ImprimeBalanceSituacion(pdFechaIni As Date, pdFechaFin As Date, pdfecha As Date, pnTipoBala As Integer, pnMoneda As Integer, nTotDebe As Currency, nTotHaber As Currency, Optional pnConCierreAnual As Integer = 0, Optional pbSinAge As Boolean = False, Optional vTituloHist As Boolean = False) As String
Dim P       As Integer
Dim sTexto  As String
Dim ntInicial As Currency, ntDebe As Currency, ntHaber As Currency
Dim lOrden  As Boolean
Dim nFinal  As Currency
Dim nFinalME As Currency
Dim nTipCambio As Currency
lOrden = False
P = 0
sTexto = ""
If pnMoneda = 2 Then
   nTipCambio = GetTipCambioBalance(Format(pdFechaFin, "yyyymmdd"))
End If

Linea sTexto, ImprimeBalanceSituacionCabecera(pdFechaIni, pdFechaFin, nTipCambio, P, pnTipoBala, pnMoneda, pdfecha, vTituloHist), 4
Linea sTexto, "( I ) ACTIVO Y SALDOS DEUDORES "
Linea sTexto, "-------------------------------", 2
Set prs = GetBalanceSituacion(pnTipoBala, pnMoneda, Month(pdFechaIni) + pnConCierreAnual, Year(pdFechaIni), "D")
Do While Not prs.EOF
      nFinal = prs!nInicial + prs!nDebe - prs!nHaber
      If pnMoneda = 2 Then
         nFinalME = Round(nFinal / nTipCambio, 2)
      End If
      Dim oGen As New DGeneral
      Linea sTexto, CON & " " & prs!cCtaContCod & Space(8) & Mid(oGen.CuentaNombre(prs!cCtaContCod) & Space(60), 1, 60) & " " _
               & PrnVal(prs!nInicial, 16, 2) & " " _
               & PrnVal(prs!nDebe, 16, 2) & " " _
               & PrnVal(prs!nHaber, 16, 2) & " " _
               & PrnVal(nFinal, 16, 2) & " " _
               & IIf(pnMoneda = 2, PrnVal(nFinalME, 16, 2), "") _
               & COFF
      Set oGen = Nothing
   ntInicial = ntInicial + prs!nInicial
   ntDebe = ntDebe + prs!nDebe
   ntHaber = ntHaber + prs!nHaber
   prs.MoveNext
   If prs.EOF Then Exit Do
   If Mid(prs!cCtaContCod, 1, 1) = "8" And Not lOrden Then
      Linea sTexto, oImpresora.gPrnSaltoLinea & CON & Space(15) & "TOTAL ACTIVO Y SALDOS DEUDORES       " & Space(20) & PrnVal(ntInicial, 16, 2) & " " & PrnVal(ntDebe, 16, 2) & " " & PrnVal(ntHaber, 16, 2) & " " & PrnVal(ntInicial + ntDebe - ntHaber, 16, 2) & COFF, 3
      Linea sTexto, "( II ) CUENTAS DE ORDEN DEUDORAS "
      Linea sTexto, "---------------------------------", 2
      ntInicial = 0: ntDebe = 0: ntHaber = 0
      lOrden = True
   End If
Loop
Linea sTexto, oImpresora.gPrnSaltoLinea & CON & Space(15) & "TOTAL CUENTAS DE ORDEN DEUDORAS      " & Space(20) & PrnVal(ntInicial, 16, 2) & " " & PrnVal(ntDebe, 16, 2) & " " & PrnVal(ntHaber, 16, 2) & " " & PrnVal(ntInicial + ntDebe - ntHaber, 16, 2) & COFF

'Ahora las Cuentas Acreedoras
lOrden = False
ntInicial = 0: ntDebe = 0: ntHaber = 0
Linea sTexto, ImprimeBalanceSituacionCabecera(pdFechaIni, pdFechaFin, nTipCambio, P, pnTipoBala, pnMoneda, pdfecha, vTituloHist), 4
Linea sTexto, "( III ) PASIVO Y SALDOS ACREEDORES"
Linea sTexto, "-----------------------------------", 2
Set prs = GetBalanceSituacion(pnTipoBala, pnMoneda, Month(pdFechaIni) + pnConCierreAnual, Year(pdFechaFin), "A")
Do While Not prs.EOF
   nFinal = prs!nInicial + prs!nHaber - prs!nDebe
   If pnMoneda = 2 Then
      nFinalME = Round(nFinal / nTipCambio, 2)
   End If
   
   Set oGen = New DGeneral
   Linea sTexto, CON & " " & prs!cCtaContCod & Space(8) & Mid(oGen.CuentaNombre(prs!cCtaContCod) & Space(60), 1, 60) & " " _
         & PrnVal(prs!nInicial, 16, 2) & " " _
         & PrnVal(prs!nDebe, 16, 2) & " " _
         & PrnVal(prs!nHaber, 16, 2) & " " _
         & PrnVal(nFinal, 16, 2) & " " _
         & IIf(pnMoneda = 2, PrnVal(nFinalME, 16, 2), "") _
         & COFF
   Set oGen = Nothing
   ntInicial = ntInicial + prs!nInicial
   ntDebe = ntDebe + prs!nDebe
   ntHaber = ntHaber + prs!nHaber
   prs.MoveNext
   If prs.EOF Then Exit Do
   If Mid(prs!cCtaContCod, 1, 1) = "8" And Not lOrden Then
      Linea sTexto, oImpresora.gPrnSaltoLinea & CON & Space(15) & "TOTAL PASIVO Y SALDOS ACREEDORES     " & Space(20) & PrnVal(ntInicial, 16, 2) & " " & PrnVal(ntDebe, 16, 2) & " " & PrnVal(ntHaber, 16, 2) & " " & PrnVal(ntInicial - ntDebe + ntHaber, 16, 2) & COFF, 3
      Linea sTexto, "( IV ) CUENTAS DE ORDEN ACREEDORAS"
      Linea sTexto, "-----------------------------------", 2
      ntInicial = 0: ntDebe = 0: ntHaber = 0
      lOrden = True
   End If
Loop
Linea sTexto, oImpresora.gPrnSaltoLinea & CON & Space(15) & "TOTAL CUENTAS DE ORDEN ACREEDORAS    " & Space(20) & PrnVal(ntInicial, 16, 2) & " " & PrnVal(ntDebe, 16, 2) & " " & PrnVal(ntHaber, 16, 2) & " " & PrnVal(ntInicial - ntDebe + ntHaber, 16, 2) & COFF, 2
Linea sTexto, BON & Space(15) & "TOTAL GENERAL" & Space(5) & "DEBE  : " & PrnVal(Val(Format(nTotDebe)), 16, 2)
Linea sTexto, Space(15) & "             " & Space(5) & "HABER : " & PrnVal(Val(Format(nTotHaber)), 16, 2) & BOFF
Linea sTexto, ValidaBalance(True, pdFechaIni, pdFechaFin, pnTipoBala, pnMoneda)
If prs.State = adStateOpen Then prs.Close: Set prs = Nothing
ImprimeBalanceSituacion = sTexto

End Function

Private Function ImprimeBalanceSituacionCabecera(pdFechaIni As Date, pdFechaFin As Date, pnTpoCambio As Currency, P As Integer, pnTipoBala As Integer, pnMoneda As Integer, pdfecha As Date, vTituloHist As Boolean) As String
Dim sTexto As String

'Linea sTexto, Cabecera(" B A L A N C E    D E    S I T U A C I O N   (" & IIf(pnTipoBala = 1, "HISTORICO", "AJUSTADO") & ") ", P, IIf(pnMoneda = 1, "S/.", IIf(pnMoneda = 2, "$", "")), 80, , Format(pdfecha, gsFormatoFecha), lsEmprLogo), 0
Linea sTexto, Cabecera(" B A L A N C E    D E    S I T U A C I O N   (" & IIf(vTituloHist, "HISTORICO", "AJUSTADO") & ") ", P, IIf(pnMoneda = 1, "S/.", IIf(pnMoneda = 2, "$", "")), 80, , Format(pdfecha, gsFormatoFecha), lsEmprLogo), 0

If pnMoneda = 0 Then
   Linea sTexto, Centra(" C O N S O L I D A D O "), 2
End If
Linea sTexto, Centra("( DEL " & pdFechaIni & " AL " & pdFechaFin & ")") & BOFF
If pnMoneda = 2 Then
   Linea sTexto, " TIPO DE CAMBIO : " & Format(pnTpoCambio, "###,##0.00##")
End If
Linea sTexto, CON
Linea sTexto, " =============================================================================================================================================" & IIf(pnMoneda = 2, "=================", "")
Linea sTexto, "  CUENTA                                                                         SALDO                    MOVIMIENTO                  SALDO   "
Linea sTexto, "  CONTABLE         D E S C R I P C I O N                                         INICIAL    ----------------------------------      ACUMULADO " & IIf(pnMoneda = 2, "      ACUMULADO ", "")
Linea sTexto, "                                                                                                    DEBE            HABER             M.N.    " & IIf(pnMoneda = 2, "         M.E.   ", "")
Linea sTexto, " ---------------------------------------------------------------------------------------------------------------------------------------------" & IIf(pnMoneda = 2, "-----------------", "") & COFF
ImprimeBalanceSituacionCabecera = sTexto
End Function

Public Function ImprimeBalanceComprobacion(pdFechaIni As Date, pdFechaFin As Date, pdfecha As Date, pnTipoBala As Integer, pnMoneda As Integer, pnLinPage As Integer, nTotDebe As Currency, nTotHaber As Currency, Optional psCtaIni As String = "", Optional psCtaFin As String = "", Optional pnDigitos As Integer = 0, Optional pbSoloAnaliticas As Boolean = False, Optional pnCierreAnio As Integer = 0, Optional pbSinAge As Boolean = False) As String
Dim N As Integer
Dim nLin   As Integer, P As Integer
Dim nTotD  As Currency, nTotH As Currency, nSaldo As Currency
Dim lOk    As Boolean
Dim sCond  As String
Dim sTexto As String
Dim lsImpre As String
Dim lnCont  As Long
Dim sObj   As String
Dim nImporte As Currency
Dim rsAge  As New ADODB.Recordset
Dim sImporteME   As String
Dim nTipCambio As Currency
Dim bHistorico As Boolean

lsImpre = ""
lnCont = 0
nLin = 66
If pnMoneda = 2 Then
   nTipCambio = GetTipCambioBalance(Format(pdFechaFin, "yyyymmdd"))
End If

If (pnTipoBala = 2 And pdFechaIni >= CDate("01/01/2005")) Or (pnTipoBala = 1) Then
   bHistorico = True
ElseIf pnTipoBala = 2 And pdFechaIni < CDate("01/01/2005") Then
   bHistorico = False
End If

Set prs = LeeBalanceHisto(pnTipoBala, pnMoneda, Month(pdFechaIni) + pnCierreAnio, Year(pdFechaIni), psCtaIni, psCtaFin, pnDigitos, pbSoloAnaliticas, , pbSinAge)

Do While Not prs.EOF
   DoEvents
   sTexto = sTexto & ImprimeBalanceComprobacionCabecera(pdFechaIni, pdFechaFin, nTipCambio, P, pnTipoBala, pnMoneda, pdfecha, nLin, pnLinPage, bHistorico)
   If Len(prs!cCtaContCod) <= 4 Then
      Linea sTexto, ""
      nLin = nLin + 1
      lnCont = lnCont + 1
        If lnCont Mod 600 = 0 Then
           lsImpre = lsImpre & sTexto
           sTexto = ""
        End If
   End If
   If pnMoneda = 2 Then
      sImporteME = PrnVal(Round(prs!nSaldoFinImporte / nTipCambio, 2), 16, 2)
   End If
   Linea sTexto, CON & " " & Mid(prs!cCtaContCod & Space(20), 1, 20) & " " & Mid(prs!cCtaContDesc & Space(50), 1, 50) & " " _
            & PrnVal(prs!nSaldoIniImporte, 16, 2) & " " _
            & PrnVal(prs!nDebe, 16, 2) & " " _
            & PrnVal(prs!nHaber, 16, 2) & " " _
            & PrnVal(prs!nSaldoFinImporte, 16, 2) & " " _
            & IIf(pnMoneda = 2, sImporteME, "") _
            & COFF
   nLin = nLin + 1
   lnCont = lnCont + 1
   If lnCont Mod 600 = 0 Then
      lsImpre = lsImpre & sTexto
      sTexto = ""
   End If
   prs.MoveNext
Loop
Linea sTexto, CON & " --------------------------------------------------------------------------------------------------------------------------------------------" & IIf(pnMoneda = 2, String(17, "-"), "")
Linea sTexto, BON & Space(88) & "TOTALES  " & PrnVal(nTotDebe, 14, 2) & " " & PrnVal(nTotHaber, 14, 2) & BOFF & COFF, 2
Linea sTexto, " ============================================================================================================================================" & IIf(pnMoneda = 2, String(17, "="), "") & COFF
ImprimeBalanceComprobacion = lsImpre & sTexto
End Function

Private Function ImprimeBalanceComprobacionCabecera(pdFechaIni As Date, pdFechaFin As Date, pnTpoCambio As Currency, P As Integer, pnTipoBala As Integer, pnMoneda As Integer, pdfecha As Date, ByRef nLin As Integer, ByRef pnLinPage As Integer, vHisto As Boolean) As String
Dim sTexto As String, cMoneda As String

If nLin > pnLinPage - 6 Then
'--------------------------------------
' CONDICIONES ANTERIORES AL 24/02/2005
'--------------------------------------
'   If pnMoneda = 1 Then
'      Linea sTexto, CabeceraItemPage(" BALANCE  DE  COMPROBACION  MONEDA NACIONAL (" & IIf(pnTipoBala = 1, "HISTORICO", "AJUSTADO") & ") ", P, "", 80, , Format(pdfecha, gsFormatoFechaView), lsEmprLogo), 0
'   ElseIf pnMoneda = 0 Then
'      Linea sTexto, CabeceraItemPage(" BALANCE  DE  COMPROBACION  CONSOLIDADO (" & IIf(pnTipoBala = 1, "HISTORICO", "AJUSTADO") & ") ", P, "", 80, , Format(pdfecha, gsFormatoFechaView), lsEmprLogo), 0
'   ElseIf pnMoneda = 2 Then
'      Linea sTexto, CabeceraItemPage(" BALANCE  DE  COMPROBACION  MONEDA EXTRANJERA  (" & IIf(pnTipoBala = 1, "HISTORICO", "AJUSTADO") & ") ", P, "", 80, , Format(pdfecha, gsFormatoFechaView), lsEmprLogo), 0
'   End If

'-----------------------------------------------
' CONDICIONES A PARTIR DEL 24/02/2005 - ESÑM
'-----------------------------------------------
Select Case pnMoneda
    Case 0
         cMoneda = "CONSOLIDADO"
    Case 1
         cMoneda = "MONEDA NACIONAL"
    Case 2
         cMoneda = "MONEDA EXTRANJERA"
End Select
Linea sTexto, CabeceraItemPage(" BALANCE  DE  COMPROBACION  " & cMoneda & " (" & IIf(vHisto, "HISTORICO", "AJUSTADO") & ") ", P, "", 80, , Format(pdfecha, gsFormatoFechaView), lsEmprLogo), 0
'--------------------------------------------------
   
   'If pnMoneda = 0 Then
   '   Linea sTexto, Centra(" CONSOLIDADO ")
   'End If
   Linea sTexto, Centra("( DEL " & pdFechaIni & " AL " & pdFechaFin & ")") & BOFF
   If pnMoneda = 2 Then
      Linea sTexto, " TIPO DE CAMBIO : " & Format(pnTpoCambio, "###,##0.00##")
   End If
   Linea sTexto, CON, 0
   Linea sTexto, " =============================================================================================================================================" & IIf(pnMoneda = 2, String(17, "="), "")
   Linea sTexto, "  CUENTA                                                                         SALDO              DEBE            HABER             SALDO   " & IIf(pnMoneda = 2, "      SALDO ME", "")
   Linea sTexto, "  CONTABLE         D E S C R I P C I O N                                         INICIAL    ----------------------------------      ACUMULADO " & IIf(pnMoneda = 2, "     ACUMULADO", "")
   'Linea sTexto, "                                                                                                    DEBE            HABER                     "
   Linea sTexto, " ---------------------------------------------------------------------------------------------------------------------------------------------" & IIf(pnMoneda = 2, String(17, "-"), "") & COFF
   nLin = 7
End If
ImprimeBalanceComprobacionCabecera = sTexto
End Function

Public Function ImprimeBalanceFormaAB(psOpeCod As String, pdfecha As Date, psBalanceCate As String, pbSoles As Boolean, psEmpresa As String) As String
Dim ContLin As Integer
Dim CadTit As String
Dim Salto As Boolean
Dim FinLinea As Boolean
Dim CadImp   As String
Dim lnDivide As Integer
If pbSoles Then
    lnDivide = 1
Else
    lnDivide = 1000
End If
    CadTit = ""
    Dim dBalance As New NBalanceCont
    Set prs = dBalance.CargaBalanceGeneral(psOpeCod, Format(pdfecha, "yyyymmdd"), psBalanceCate)
    Set dBalance = Nothing
        CadTit = Trim(prs!cDescrip)
        ContLin = 0
        Salto = True
        FinLinea = False
        Do While Not prs.EOF
            ContLin = ContLin + 1
            If ContLin >= 53 Then
                ContLin = 0
                CadImp = CadImp + oImpresora.gPrnSaltoPagina
                If Mid(Trim(prs!cDescrip), 1, 2) = "LN" Then
                    CadTit = Trim(prs!cDescrip)
                End If
                Linea CadImp, ImpreBalanceFormaABCabecera(CadTit, psBalanceCate, pbSoles, pdfecha, psEmpresa), 0
                Salto = True
                FinLinea = True
            End If
                If Trim(prs!cDescrip) = "-" Then
                    CadImp = CadImp + String(120, "-") + oImpresora.gPrnSaltoLinea
                    Salto = False
                Else
                    If Mid(Trim(prs!cDescrip), 1, 2) = "LN" Then
                        If Not FinLinea Then
                            If Salto Then
                                Salto = False
                            Else
                                CadImp = CadImp + oImpresora.gPrnSaltoPagina
                            End If
                            CadTit = Trim(prs!cDescrip)
                            Linea CadImp, ImpreBalanceFormaABCabecera(CadTit, psOpeCod, pbSoles, pdfecha, psEmpresa), 0
                            ContLin = 0
                        Else
                            FinLinea = False
                        End If
                    Else
                        CadImp = CadImp + Mid(RTrim(prs!cDescrip) + Space(45), 1, 45) + Space(18 - Len(Format(Round(prs!nMN / lnDivide, 2), gsFormatoNumeroView))) + Format(Round(prs!nMN / lnDivide, 2), gsFormatoNumeroView) + Space(18 - Len(Format(Round(prs!nME / lnDivide, 2), gsFormatoNumeroView))) + Format(Round(prs!nME / lnDivide, 2), gsFormatoNumeroView) + Space(18 - Len(Format(Round(prs!nTotal / lnDivide, 2), gsFormatoNumeroView))) + Format(Round(prs!nTotal / lnDivide, 2), gsFormatoNumeroView) + Space(18 - Len(Format(Round(prs!nTotAj / lnDivide, 2), gsFormatoNumeroView))) + Format(Round(prs!nTotAj / lnDivide, 2), gsFormatoNumeroView) + oImpresora.gPrnSaltoLinea
                        FinLinea = False
                        Salto = False
                    End If
                End If
            prs.MoveNext
        Loop
    RSClose prs
    ImprimeBalanceFormaAB = CadImp
End Function

Private Function ImpreBalanceFormaABCabecera(ByVal sTit As String, ByVal sTipBal As String, pbSoles As Boolean, pdfecha As Date, psEmpresa As String) As String
Dim CadImp As String
Dim sTit1 As String
Dim sTit2 As String
Dim sTit3 As String
Dim sTit4 As String
Dim I As Integer
Dim nPos As Integer
    sTit1 = psEmpresa
    sTit2 = ""
    For I = 3 To Len(sTit)
        If Mid(sTit, I, 1) <> "\" Then
            sTit2 = sTit2 & Mid(sTit, I, 1)
        Else
            nPos = I
            Exit For
        End If
    Next I
    'Item
    sTit3 = Mid(sTit, nPos + 1, (Len(sTit) - nPos))
    'Titulo de Moneda
    If pbSoles Then
       sTit4 = "( Expresado en Nuevos Soles )"
    Else
       sTit4 = "( Expresado en Miles de Nuevos Soles )"
    End If
    If sTit2 = "ACTIVO" Or sTit2 = "PASIVO" Then
        CadImp = CadImp & PrnSet("MI", 20) & oImpresora.gPrnBoldON & "FORMA ""A""" & Space(60 - 24) & "B A L A N C E    G E N E R A L" & oImpresora.gPrnSaltoLinea
    Else
        CadImp = CadImp & PrnSet("MI", 20) & oImpresora.gPrnBoldON & "FORMA ""B""" & Space(60 - 28) & "G A N A C I A S   Y   P E R D I D A S" & oImpresora.gPrnSaltoLinea
    End If
    CadImp = CadImp & Space(60 - (Len(sTit1) / 2)) & sTit1 & oImpresora.gPrnSaltoLinea
    CadImp = CadImp & Space(60 - (Len(sTit3 & pdfecha) / 2)) & sTit3 & pdfecha & oImpresora.gPrnSaltoLinea
    CadImp = CadImp + Space(60 - (Len(sTit4) / 2)) + sTit4 + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
    CadImp = CadImp + Space(10) + sTit2 + Space(43 - Len(sTit2)) + "Moneda" + Space(8) + "Equivalente" + Space(10) + "TOTAL" + Space(12) + "Total Ajustado" + oImpresora.gPrnSaltoLinea
    CadImp = CadImp + Space(52) + "Nacional" + Space(9) + "en M.E." + Space(28) + "Por Inflacion" + PrnSet("B-") + oImpresora.gPrnSaltoLinea
    ImpreBalanceFormaABCabecera = CadImp
End Function

Public Function CalculaSaldoCuentaAD(psCtaCod As String, psMoneda As Integer, pbSaldoAcreedor As Boolean, Optional psBalanceCate As String = "3") As Currency
Dim lsCtaCaracter As String
Dim CadSql   As String, CadSql2 As String
Dim RSalAD   As ADODB.Recordset
Dim lsCtaDos As String
Dim oCta     As New DCtaCont
   If Len(Trim(psCtaCod)) > 1 Then
       lsCtaDos = Mid(psCtaCod, 1, 2)
   Else
       lsCtaDos = Mid(psCtaCod, 1, 1)
   End If
   
   Set RSalAD = oCta.CargaCtaContClase(lsCtaDos)
   If Not RSalAD.EOF Then
       lsCtaCaracter = Trim(RSalAD!cCtaCaracter)
   End If
   RSalAD.Close
   CadSql2 = ""
   If pbSaldoAcreedor Then  'si es saldo acreedor
       If lsCtaCaracter = "A" Then ' si es cuenta Acreedora
           CadSql2 = " And z.nSaldoIniImporte > 0"
       Else                    'Si es deudora
           CadSql2 = " And z.nSaldoIniImporte < 0"
       End If
   Else                'Si es Saldo deudor
       If lsCtaCaracter = "A" Then ' si es cuenta Acreedora
           CadSql2 = " And z.nSaldoIniImporte < 0"
       Else                    'Si es deudora
           CadSql2 = " And z.nSaldoIniImporte > 0"
       End If
   End If
   
   CadSql = ""
       If Len(psCtaCod) >= 4 Then
           CadSql = " cCtaContCod like '" & lsCtaDos & psMoneda & Mid(psCtaCod, 4, Len(psCtaCod) - 3) + "%'"
       Else
           If Len(psCtaCod) = 1 Then
               CadSql = " cCtaContCod like '" & lsCtaDos + "_" & psMoneda & "%'"
           Else
               CadSql = " cCtaContCod like '" & Left(psCtaCod, 1) + "_" & psMoneda & "%'"
           End If
       End If
   
   CadSql = " SELECT SUM(z.nSaldoIniImporte) as Saldo " _
          & "FROM  BalanceTemp as z " _
          & "WHERE cBalanceCate = '" & psBalanceCate & "' and z.cCtaClasifica = '" & lsCtaCaracter & "' And " & CadSql & CadSql2
   Set RSalAD = dbConec.CargaRecordSet(CadSql)
       If IsNull(RSalAD!Saldo) Then
           CalculaSaldoCuentaAD = 0
       Else
           CalculaSaldoCuentaAD = Format(RSalAD!Saldo, "#0.0000")
       End If
   RSalAD.Close
   Set RSalAD = Nothing
   Set oCta = Nothing
   
End Function

Public Function CalculaSaldoCuenta(psCtaCod As String, psMoneda As String, Optional psBalanceCate As String = "3", Optional psTipo As String) As Currency
Dim lsCtaCaracter As String
Dim CadSql   As String
Dim RSalAD   As ADODB.Recordset
Dim lsCtaDos As String
   If Len(Trim(psCtaCod)) > 1 Then
       lsCtaDos = Mid(psCtaCod, 1, 2)
   Else
       lsCtaDos = Mid(psCtaCod, 1, 1)
   End If
   
   CadSql = ""
       If Len(psCtaCod) >= 4 Then
           CadSql = " cCtaContCod like '" & lsCtaDos & psMoneda & Mid(psCtaCod, 4, Len(psCtaCod) - 3) + "%'"
       Else
           If Len(psCtaCod) = 1 Then
               CadSql = " cCtaContCod like '" & lsCtaDos + "_" & psMoneda & "%'"
           Else
               CadSql = " cCtaContCod like '" & Left(psCtaCod, 2) & psMoneda & "%'"
           End If
       End If
   
   CadSql = "SELECT SUM(nSaldoIniImporte) as nTotal " _
        & "FROM BalanceTemp WHERE cBalanceCate = '" & psBalanceCate & "' and " & CadSql
        
        If Len(Trim(psTipo)) > 0 Then
            CadSql = CadSql & " AND cBalanceTipo='" & psTipo & "' "
        End If
   
   Set RSalAD = dbConec.CargaRecordSet(CadSql)
       If IsNull(RSalAD!nTotal) Then
           CalculaSaldoCuenta = 0
       Else
           CalculaSaldoCuenta = Format(RSalAD!nTotal, "#0.0000")
       End If
   RSalAD.Close
   Set RSalAD = Nothing
End Function

Public Function CargaBalanceSectorial(psFecha As String) As Recordset
On Error GoTo CargaBalanceSectorialErr
   psSql = "SELECT EB.cCodigo, EB.cDescrip, B.nMNAj, B.nMNHist, B.nME " _
         & "FROM   RepBaseFormula EB Inner JOIN BalanceSectorial B ON EB.cCodigo = B.cCodigo " _
         & "WHERE  B.nBalanceMes = " & Month(CDate(psFecha)) & " and B.nBalanceAnio = " & Year(CDate(psFecha)) & " " _
         & "ORDER BY EB.cCodigo"
   Set CargaBalanceSectorial = dbConec.CargaRecordSet(psSql, adLockReadOnly)
   Exit Function
CargaBalanceSectorialErr:
   Call RaiseError(MyUnhandledError, "DBalanceCont:CargaBalanceSectorial Method")
End Function

Public Function ImprimeBalanceSectorial(psFecha As String, psEmpresa As String, lbArchivo As Boolean) As String
Dim CadImp As String
Dim Cont As Integer
Dim P    As Integer
Dim lsFormatoNumero As String
    CadImp = ""
    
    Set prs = CargaBalanceSectorial(psFecha)
    Cont = 60
    P = 0
      Do While Not prs.EOF
         If Cont >= 56 And Not lbArchivo Then
            Cont = 7
            P = P + 1
            If P > 1 Then
               Linea CadImp, oImpresora.gPrnSaltoPagina, 0
            End If
                lsFormatoNumero = "#,#0"
                Linea CadImp, Space(35) + UCase(psEmpresa)
                Linea CadImp, Space(38) + "BALANCE SECTORIAL POR AGENTES ECONOMICOS"
                Linea CadImp, Space(50) + "AL " + Format(CDate(psFecha), gsFormatoFechaView)
                Linea CadImp, Space(37) + "(Saldos Expresados en Miles de Nuevos Soles)", 2
                Linea CadImp, "Código de Cuentas" + Space(10) + "Descripcion de la Cuenta" + Space(44) + "Moneda Nacional" + Space(10) + "MONEDA"
                Linea CadImp, Space(91) + "Ajustado" + Space(5) + "Historico" + Space(7) + "Extranjera", 2
         End If
         
         If lbArchivo Then
            CadImp = CadImp + Trim(prs!cCodigo) + Space(12 - Len(Format(prs!nMNAj, lsFormatoNumero))) + Format(prs!nMNAj, lsFormatoNumero) + Space(12 - Len(Format(prs!nMNHist, lsFormatoNumero))) + Format(prs!nMNHist, lsFormatoNumero) + Space(12 - Len(Format(prs!nME, lsFormatoNumero))) + Format(prs!nME, lsFormatoNumero) + oImpresora.gPrnSaltoLinea
         Else
            CadImp = CadImp + Trim(prs!cCodigo) + Space(5) + Mid(prs!cDescrip + Space(60), 1, 60) + Space(15 - Len(Format(prs!nMNAj, lsFormatoNumero))) + Format(prs!nMNAj, lsFormatoNumero) + Space(15 - Len(Format(prs!nMNHist, lsFormatoNumero))) + Format(prs!nMNHist, lsFormatoNumero) + Space(15 - Len(Format(prs!nME, lsFormatoNumero))) + Format(prs!nME, lsFormatoNumero) + oImpresora.gPrnSaltoLinea
         End If
         Cont = Cont + 1
         prs.MoveNext
      Loop
    prs.Close: Set prs = Nothing
    ImprimeBalanceSectorial = CadImp
    CadImp = ""
End Function




