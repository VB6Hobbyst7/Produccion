VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DPresupuesto"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Actualiza en base los Datos de el Presupuesto
Option Base 0
Option Explicit

Private lsServerComunes As String
Private lsServerPersona As String
Private lsServerAdministracion As String
Private lsServerNegocio As String
Private lsServerImagenes As String

'set this to 0 to disable debug code in this class
#Const DebugMode = 0
#If DebugMode Then
    'local variable to hold the serialized class ID that was created in Class_Initialize
    Private mlClassDebugID As Long
#End If

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'Elimina  un Item de el Presupuesto
Public Function EliminaRubroCta(psAnio As String, psPresupuestoCod As String, psRubro As String, psCtaContCod As String) As Boolean
   Dim sqlA As String
   Dim oCon As DConecta
   Set oCon = New DConecta
   On Error GoTo EliminaRubroCtaErr
       
   If oCon.AbreConexion() Then
      sqlA = " Delete PresuRubroCta" _
           & " WHERE nPresuCod = " & psPresupuestoCod & " and nPresuAnio = " & psAnio _
           & "   and cPresuRubCod = '" & psRubro & "' And cCtaContCod = '" & psCtaContCod & "' "
      oCon.Ejecutar sqlA
      EliminaRubroCta = True
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
EliminaRubroCtaErr:
    EliminaRubroCta = False
    Call RaiseError(MyUnhandledError, "DPresupuesto: EliminaRubroCta Method")
End Function

'Elimina  un Item de el Presupuesto
Public Function AgregaRubroCta(psAnio As String, psPresupuestoCod As String, psRubro As String, psCtaContCod As String) As Boolean
   Dim sqlA As String
   Dim oCon As DConecta
   On Error GoTo AgregaRubroCtaErr
   
   Set oCon = New DConecta
   If oCon.AbreConexion() Then
       sqlA = " Insert Into PresuRubroCta (nPresuCod, nPresuAnio, cPresuRubCod, cCtaContCod) " _
            & " Values(" & psPresupuestoCod & ", " & psAnio & ", '" & psRubro & "', '" & psCtaContCod & "')"
       oCon.Ejecutar sqlA
       AgregaRubroCta = True
       oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
AgregaRubroCtaErr:
    AgregaRubroCta = False
    Call RaiseError(MyUnhandledError, "DPresupuesto:AgregaRubroCta Method")
End Function

'Elimina  un Item de el Presupuesto
Public Function EliminaPresupuesto(psPresupuestoCod As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo EliminaPresupuestoErr
        
   If oCon.AbreConexion() Then
      sqlA = " Delete PresuClase Where nPresuCod = '" & psPresupuestoCod & "'"
      oCon.Ejecutar sqlA
      EliminaPresupuesto = True
      oCon.CierraConexion
   End If
    
   Set oCon = Nothing
   Exit Function
EliminaPresupuestoErr:
    EliminaPresupuesto = False
    Call RaiseError(MyUnhandledError, "DPresupuesto:EliminaPresupuesto Method")
End Function

'Devuelve todos los Items de PresuRubroAnio
Public Function GetRubrosPresupuesto(psPeriodo As String, psPresupuestoCod As String) As Recordset
   Dim sqlA As String
   Dim oCon As DConecta
   Set oCon = New DConecta
   On Error GoTo GetRubrosPresupuestoErr

   If oCon.AbreConexion() Then
       sqlA = " SELECT P.cPresuRubCod, P.cPresuRubDescripcion" _
            & " FROM PresuRubroAnio P JOIN PresuClase PRE ON P.nPresuCod = PRE.nPresuCod " _
            & " WHERE P.nPresuAnio = " & nVal(psPeriodo) & "  AND P.nPresuCod = '" & psPresupuestoCod & "' AND cFlag Is Null" _
            & " ORDER BY P.cPresuRubCod "
       Set GetRubrosPresupuesto = oCon.CargaRecordSet(sqlA)
       oCon.CierraConexion
   End If
   Set oCon = Nothing
   
   Exit Function
GetRubrosPresupuestoErr:
    Call RaiseError(MyUnhandledError, "DPresupuesto: GetRubrosPresupuesto Method")
End Function

'Devuelve todos los Items de la Asistencia Medica Familiar
Public Function GetRubrosPresupuestoDet(psPeriodo As String, psPresupuestoCod As String, Optional psRubro As String = "00") As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetRubrosPresupuestoDetErr

    If oCon.AbreConexion() Then
        sqlA = " SELECT R.cPresuRubCod, R.cPresuRubDescripcion" _
             & " FROM PresuRubroAnio R" _
             & " WHERE R.nPresuAnio = " & nVal(psPeriodo) & "  AND R.nPresuCod = '" & psPresupuestoCod & "' And R.cPresuRubCod like '" & psRubro & "__' AND cFlag Is Null" _
             & " ORDER BY R.cPresuRubCod"
        Set GetRubrosPresupuestoDet = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetRubrosPresupuestoDetErr:
    Call RaiseError(MyUnhandledError, "DPresupuesto: GetRubrosPresupuestoDet Method")
End Function

'Devuelve todos los Items de la Asistencia Medica Familiar
Public Function GetPresupuestoTpo(psPresupuestoCod As String) As Integer
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetPresupuestoErr

    If oCon.AbreConexion() Then
        sqlA = " SELECT R.nPresuTpo FROM PresuClase R" _
             & " WHERE R.nPresuCod = '" & psPresupuestoCod & "'"
        GetPresupuestoTpo = oCon.CargaRecordSet(sqlA)!nPresuTpo
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetPresupuestoErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function


'Actualiza un Item de el Presupuesto
Public Function ModificaPresupuesto(psPresupuestoCod As String, psPresupuestoNombre As String, psPresupuestoTipo As String, pbPredeterminado As Boolean) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo ModificaAsisMedPrivErr

    If oCon.AbreConexion() Then
        sqlA = " Update PresuClase " _
             & " Set cPresuDescripcion = '" & psPresupuestoNombre & "', nPresuTpo = " & psPresupuestoTipo & ", bDefecto = " & IIf(pbPredeterminado, 1, 0) & "  " _
             & " Where nPresuCod = '" & psPresupuestoCod & "'"
        oCon.Ejecutar sqlA
        
        If pbPredeterminado Then
            sqlA = " Update PresuClase " _
                 & " Set bDefecto = 0 " _
                 & " Where nPresuCod <> '" & psPresupuestoCod & "'"
            oCon.Ejecutar sqlA
        End If
        
        ModificaPresupuesto = True
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaAsisMedPrivErr:
    ModificaPresupuesto = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:ModificaAsisMedPriv Method")
End Function

'Actualiza un Item de el Presupuesto
Public Function AgregaPresupuesto(psPresupuestoCod As String, psPresupuestoNombre As String, psPresupuestoTipo As String, pbPredeterminado As Boolean) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo ModificaAsisMedPrivErr
    If oCon.AbreConexion() Then
        sqlA = " Insert PresuClase (nPresuCod, cPresuDescripcion, nPresuTpo, bDefecto)" _
             & " Values('" & psPresupuestoCod & "','" & psPresupuestoNombre & "'," & psPresupuestoTipo & "," & IIf(pbPredeterminado, 1, 0) & ")"
        oCon.Ejecutar sqlA
        
        If pbPredeterminado Then
            sqlA = " Update PresuClase " _
                 & " Set bDefecto = 0 " _
                 & " Where nPresuCod <> '" & psPresupuestoCod & "'"
            oCon.Ejecutar sqlA
        End If
        
        AgregaPresupuesto = True
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
ModificaAsisMedPrivErr:
    AgregaPresupuesto = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:ModificaAsisMedPriv Method")
End Function

'Actualiza un Rubro
Public Function AgregaRubro(psPresupuestoCod As String, psAnio As String, psRubroCod As String, psRubroDescripcion As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo AgregaRubroErr

   If oCon.AbreConexion() Then
      sqlA = " INSERT INTO PresuRubroAnio " _
           & " (nPresuCod, nPresuAnio, cPresuRubCod, cPresuRubDescripcion, cMoneda, " _
           & "  nPresuRubMonAnt,nPresuRubMonPre,nPresuRubMonCre1," _
           & "  nPresuRubMonCre2,cFlag) " _
           & " VALUES ( " & psPresupuestoCod & ", " & psAnio & ", '" & psRubroCod & "','" & psRubroDescripcion & "', '1', " _
           & "  0, 0, 0, 0, NULL) "
      oCon.Ejecutar sqlA
      AgregaRubro = True
      oCon.CierraConexion
   End If
    Set oCon = Nothing
    Exit Function
AgregaRubroErr:
    AgregaRubro = False
    Call RaiseError(MyUnhandledError, "DPresupuesto: AgregaRubro Method")
End Function

'Actualiza un Rubro
Public Function ModificaRubro(psPresupuestoCod As String, psAnio As String, psRubroCod As String, psRubroDescripcion As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo ModificaRubroErr

    If oCon.AbreConexion() Then
       sqlA = " UPDATE PresuRubroAnio SET cPresuRubDescripcion = '" & psRubroDescripcion & "'" _
            & " WHERE  nPresuCod = " & psPresupuestoCod & " and nPresuAnio = " & psAnio & " and cPresuRubCod = '" & psRubroCod & "' "
       oCon.Ejecutar sqlA
       ModificaRubro = True
       oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaRubroErr:
    ModificaRubro = False
    Call RaiseError(MyUnhandledError, "DPresupuesto: ModificaRubro Method")
End Function

'Actualiza un Rubro
Public Function ModificaMontoRubro(psPresupuestoCod As String, psAnio As String, psRubroCod As String, pnMonIni As Currency, pnMonto As Currency, pnMonCre1 As Currency, pnMonCre2 As Currency) As Boolean
   Dim sqlA As String
   Dim oCon As DConecta
   Set oCon = New DConecta
   On Error GoTo ModificaMontoRubroErr

   If oCon.AbreConexion() Then
      sqlA = " Update PresuRubroAnio " _
           & " SET    nPresuRubMonAnt = " & pnMonIni & ", nPresuRubMonPre = " & pnMonto & ", nPresuRubMonCre1 = " & pnMonCre1 & ", nPresuRubMonCre2 = " & pnMonCre2 & "" _
           & " WHERE  nPresuCod = " & psPresupuestoCod & " and nPresuAnio = " & psAnio & " and cPresuRubCod = '" & psRubroCod & "' "
      oCon.Ejecutar sqlA
      ModificaMontoRubro = True
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
ModificaMontoRubroErr:
    ModificaMontoRubro = False
    Call RaiseError(MyUnhandledError, "DPresupuesto: ModificaMontoRubro Method")
End Function

'Actualiza un Rubro
Public Function ModificaMontoRubroMes(psPresupuestoCod As String, psPeriodo As String, psRubroCod As String, pnMes As Currency, pnMontoMes As Currency) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo ModificaAsisMedPrivErr

    If oCon.AbreConexion() Then
          sqlA = " UPDATE PresuRubroMes" _
               & " Set nPresuRubMEsMonIni = " & pnMontoMes & "" _
               & " WHERE nPresuCod = " & psPresupuestoCod & " and  nPresuAnio = " & psPeriodo _
               & "   AND cPresuRubCod = '" & psRubroCod & "' And nPresuMes = " & pnMes
        oCon.Ejecutar sqlA
        ModificaMontoRubroMes = True
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaAsisMedPrivErr:
    ModificaMontoRubroMes = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:ModificaAsisMedPriv Method")
End Function

Public Function EliminaRubro(psPresupuestoCod As String, psPeriodo As String, psRubroCod As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo EliminaRubroErr

    If oCon.AbreConexion() Then
        sqlA = "DELETE PresuRubroCta WHERE nPresuAnio = " & nVal(psPeriodo) & " AND nPresuCod = " & psPresupuestoCod & " AND cPresuRubCod = '" & psRubroCod & "' "
        oCon.Ejecutar sqlA
        
        'sqlA = "UPDATE PresuRubroAnio SET cFlag = 'X' " & _
               " WHERE nPresuAnio = " & nVal(psPeriodo) & " AND nPresuCod = " & psPresupuestoCod & " AND " & _
               "       cPresuRubCod = '" & psRubroCod & "' "
        sqlA = "DELETE PresuRubroAnio " & _
               " WHERE nPresuAnio = " & nVal(psPeriodo) & " AND nPresuCod = " & psPresupuestoCod & " AND " & _
               "       cPresuRubCod = '" & psRubroCod & "' "
      
        oCon.Ejecutar sqlA
        EliminaRubro = True
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
EliminaRubroErr:
    EliminaRubro = False
    Call RaiseError(MyUnhandledError, "DPresupuesto:EliminaRubro Method")
End Function

'Actualiza un Rubro
Public Function AgregaMontoRubroMes(psPresupuestoCod As String, psPeriodo As String, psRubroCod As String, pnMes As Currency, pnMontoMes As Currency) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo ModificaAsisMedPrivErr

    If oCon.AbreConexion() Then
          'sqlA = " Declare @cad varchar(6);" _
               & " Select @Cad = nPresuAnio from PlaPresupuesto Where nPeriodo = '" & nVal(psPeriodo) & "' And nPresuCod = '" & psPresupuestoCod & "';" _
               & " INSERT INTO PresuRubroMes (nPresuCod, nPresuAnio, cPresuRubCod,nPresuMes,nPresuRubMesMonIni,nPresuRubroMesMonEje,nPresuRubroMesMonRes)" _
               & " VALUES ('" & psPresupuestoCod & "','" & psPeriodo & "','" & psRubroCod & "','" & pnMes & "'," & pnMontoMes & ", 0, 0)"
          sqlA = " INSERT INTO PresuRubroMes (nPresuCod, nPresuAnio, cPresuRubCod,nPresuMes,nPresuRubMesMonIni,nPresuRubMesMonEje,nPresuRubMesMonRes)" _
               & " VALUES ('" & psPresupuestoCod & "','" & psPeriodo & "','" & psRubroCod & "','" & pnMes & "'," & pnMontoMes & ", 0, 0) "
        
        oCon.Ejecutar sqlA
        AgregaMontoRubroMes = True
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaAsisMedPrivErr:
    AgregaMontoRubroMes = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:ModificaAsisMedPriv Method")
End Function

'Devuelve todos los Items de la Asistencia Medica Familiar
Public Function GetPresupuesto(Optional pbCampos As Boolean = False) As ADODB.Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetPresupuestoErr

    If oCon.AbreConexion() Then
        If Not pbCampos Then
            sqlA = " Select P.nPresuCod, P.cPresuDescripcion, cConsDescripcion +  Space(40) + Convert(Varchar(10),P.nPresuTpo) cTpo, bDefecto  " _
                 & " From PresuClase P" _
                 & " Left Join Constante CO On P.nPresuTpo  = CO.nConsValor And CO.nConsCod = '" & gPPPresupuestoTpo & "'" _
                 & " Order By P.nPresuCod"
        Else
            sqlA = " Select P.nPresuCod, P.cPresuDescripcion + Case bDefecto When 1 Then ' *Por Defecto*' Else '' End cPresuDescripcion " _
                 & " From PresuClase P" _
                 & " Order By P.nPresuCod"
        End If
        Set GetPresupuesto = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetPresupuestoErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function

Public Function GetPresupuestoDefecto() As Integer
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo GetPresupuestoErr

    If oCon.AbreConexion() Then
        sqlA = " Select P.nPresuCod From PresuClase P Where bDefecto = 1 "
        Set rs = oCon.CargaRecordSet(sqlA)
        
        If rs.EOF And rs.BOF Then
            GetPresupuestoDefecto = -1
        Else
            GetPresupuestoDefecto = rs.Fields(0)
        End If
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetPresupuestoErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function


'Devuelve todos los Items de la Asistencia Medica Familiar
Public Function GetPresupuestoEjec(psPresupuestoCod As String, psPeriodo As String, pbEjecutado As Boolean, Optional pnProy As Integer = 0, Optional pnBala As Integer = 0, Optional pnTipo As Integer = 1, Optional pnTipoSaldo As Integer = 0, Optional pnMoneda As Integer, Optional pbConsiderarCierreAño As Boolean = True, Optional pnMovContablesAdd As Integer = 0) As ADODB.Recordset
'Valores de PNTIPO
' 1: Mensual
' 2: Trimestral
' 3: Semestral

    Dim tmpSql As String
    Dim tmpPry As String
    Dim tmpSqlMovCnt As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetPresupuestoErr
    Dim sqlAdd As String
    Dim sqlMoneda As String
    
    If pnMoneda = 1 Then
        sqlAdd = "'1','3'"
        sqlMoneda = "[13]"
    ElseIf pnMoneda = 2 Then
        sqlAdd = "'2'"
        sqlMoneda = "2"
    ElseIf pnMoneda = 3 Then
        sqlAdd = "'1','3','2'"
        sqlMoneda = "_"
    ElseIf pnMoneda = 4 Then
        sqlAdd = "'6'"
        sqlMoneda = "6"
    ElseIf pnMoneda = 5 Then
        sqlAdd = "'1','3','6'"
        sqlMoneda = "[136]"
    ElseIf pnMoneda = 6 Then
        sqlAdd = "'2','6'"
        sqlMoneda = "[26]"
    ElseIf pnMoneda = 7 Then
        sqlAdd = "'1','2','3','6'"
        sqlMoneda = "_"
    End If
    
    
    If oCon.AbreConexion Then
      
        If Not pbEjecutado Then
            tmpSql = " SELECT r.cPresuRubCod cCodRub, r.cPresuRubDescripcion cDesRub, r.nPresuRubMonAnt nMonIni, r.nPresuRubMonPre nMonto, r.nPresuRubMonCre1 nMonCre1, r.nPresuRubMonCre2 nMonCre2, (r.nPresuRubMonPre + r.nPresuRubMonCre1 + r.nPresuRubMonCre2) Total " & _
                   "   FROM PresuRubroAnio R WITH (NOLOCK) WHERE r.nPresuAnio = " & nVal(psPeriodo) & " And nPresuCod  = " & nVal(psPresupuestoCod) & " " & _
                   "   And r.cFlag Is Null  ORDER BY r.cPresuRubCod "
        Else
            Dim sServidor As String
            sServidor = "" 'oCon.StringServidorRemoto("07", "02", False)
            
            If pnTipo = 1 Then
               tmpSql = "SELECT Presu.cPresuRubCod cCodRub, Presu.cPresuRubDescripcion cDesRub, Sum(Eje.nMonIni) nMonIni, Sum(Eje.nMonto) nMonto, " _
                      & "       Sum(Eje.nMonCre1) nMonCre1, Sum(Eje.nMonCre2) nMonCre2, " _
                      & "       Sum(Eje.Total) Total, " & IIf(pnTipoSaldo = 2, " SUM(nSaldoIni) nSaldoIni, ", "") _
                      & "       Sum(Ene)  Ene, Sum(MovEne)  MovEne, Sum(MovCEne) MovCEne, Sum(DMEne) DMEne, " _
                      & "       Sum(Feb)  Feb, Sum(MovFeb)  MovFeb, Sum(MovCFeb) MovCFeb, Sum(DMFeb) DMFeb, " _
                      & "       Sum(Mar)  Mar, Sum(MovMar)  MovMar, Sum(MovCMar) MovCMar, Sum(DMMar) DMMar, " _
                      & "       Sum(Abr)  Abr, Sum(MovAbr)  MovAbr, Sum(MovCAbr) MovCAbr, Sum(DMAbr) DMAbr, " _
                      & "       Sum(May)  May, Sum(MovMay)  MovMay, Sum(MovCMay) MovCMay, Sum(DMMay) DMMay, " _
                      & "       Sum(Jun)  Jun, Sum(MovJun)  MovJun, Sum(MovCJun) MovCJun, Sum(DMJun) DMJun, " _
                      & "       Sum(Jul)  Jul, Sum(MovJul)  MovJul, Sum(MovCJul) MovCJul, Sum(DMJul) DMJul, " _
                      & "       Sum(Ago)  Ago, Sum(MovAgo)  MovAgo, Sum(MovCAgo) MovCAgo, Sum(DMAgo) DMAgo, " _
                      & "       Sum(Seti) Seti, Sum(MovSet) MovSet, Sum(MovCSet) MovCSet, Sum(DMSet) DMSet, " _
                      & "       Sum(Oct)  Oct, Sum(MovOct)  MovOct, Sum(MovCOct) MovCOct, Sum(DMOct) DMOct, " _
                      & "       Sum(Nov)  Nov, Sum(MovNov)  MovNov, Sum(MovCNov) MovCNov, Sum(DMNov) DMNov, " _
                      & "       Sum(Dic)  Dic, Sum(MovDic)  MovDic, Sum(MovCDic) MovCDic, Sum(DMDic) DMDic " _
                      & "FROM PresuRubroAnio Presu WITH (NOLOCK) JOIN ( "
            End If
            
            If pnTipo = 2 Then
               tmpSql = "SELECT Presu.cPresuRubCod cCodRub, Presu.cPresuRubDescripcion cDesRub, Sum(Eje.nMonIni) nMonIni, Sum(Eje.nMonto) nMonto, " _
                      & "       Sum(Eje.nMonCre1) nMonCre1, Sum(Eje.nMonCre2) nMonCre2, " _
                      & "       Sum(Eje.Total) Total, " & IIf(pnTipoSaldo = 2, " SUM(nSaldoIni) nSaldoIni, ", "") _
                      & "       Sum(Ene+Feb+Mar) Primero, Sum(MovEne+MovFeb+MovMar)  MovPrimero, Sum(MovCEne+MovCFeb+MovcMar) MovCPrimero, Sum(DMEne+DMFeb+DMMar) DMPrimero, " _
                      & "       Sum(Abr+May+Jun) Segundo, Sum(MovAbr+MovMay+MovJun)  MovSegundo, Sum(MovCAbr+MovCMay+MovCJun) MovCSegundo, Sum(DMAbr+DMMay+DMJun) DMSegundo, " _
                      & "       Sum(Jul+Ago+Seti) Tercero, Sum(MovJul+MovAgo+MovSet) MovTercero, Sum(MovCJul+MovCAgo+MovCSet) MovCTercero, Sum(DMJul+DMAgo+DMSet) DMTercero, " _
                      & "       Sum(Oct+Nov+Dic) Cuarto , Sum(MovOct+MovNov+MovDic)  MovCuarto,  Sum(MovCOct+MovCNov+MovCDic) MovCCuarto , Sum(DMOct+DMNov+DMDic) DMCuarto  " _
                      & "FROM PresuRubroAnio Presu WITH (NOLOCK) JOIN ( "
            End If
            If pnTipo = 3 Then
               tmpSql = "SELECT Presu.cPresuRubCod cCodRub, Presu.cPresuRubDescripcion cDesRub, Sum(Eje.nMonIni) nMonIni, Sum(Eje.nMonto) nMonto, " _
                      & "       Sum(Eje.nMonCre1) nMonCre1, Sum(Eje.nMonCre2) nMonCre2, " _
                      & "       Sum(Eje.Total) Total, " & IIf(pnTipoSaldo = 2, " SUM(nSaldoIni) nSaldoIni ", "") _
                      & "       Sum(Ene+Feb+Mar+Abr+May+Jun) Primero, Sum(MovEne+MovFeb+MovMar+MovAbr+MovMay+MovJun) MovPrimero,  Sum(MovCEne+MovCFeb+MovCMar+MovCAbr+MovCMay+MovCJun) MovCPrimero, Sum(DMEne+DMFeb+DMMar+DMAbr+DMMay+DMJun) DMPrimero, " _
                      & "       Sum(Jul+Ago+Seti+Oct+Nov+Dic) Segundo, Sum(MovJul+MovAgo+MovSet+MovOct+MovNov+MovDic) MovSegundo, Sum(MovCJul+MovCAgo+MovCSet+MovCOct+MovCNov+MovCDic) MovCSegundo, Sum(DMJul+DMAgo+DMSet+DMOct+DMNov+DMDic) DMSegundo  " _
                      & "FROM PresuRubroAnio Presu WITH (NOLOCK) JOIN ( "
            End If
                'Movimientos CON proyeccion
                'Agregar el contable como condicional
                tmpSql = tmpSql _
                       & "SELECT R.cPresuRubCod, R.nPresuAnio, R.nPresuCod, R.nPresuRubMonAnt nMonIni, R.nPresuRubMonPre nMonto, R.nPresuRubMonCre1 nMonCre1, R.nPresuRubMonCre2 nMonCre2,  " _
                       & " (R.nPresuRubMonPre + R.nPresuRubMonCre1 + R.nPresuRubMonCre2) Total, " & IIf(pnTipoSaldo = 2, " Sum(nSaldoIni) nSaldoIni, ", "") _
                       & " sum(CASE WHEN P.nPresuMes = 1  THEN P.nPresuRubMesMonIni ELSE 0 END) Ene, (IsNull(Cta.MovEne,0) + IsNull(Pro.ProEne,0)) MovEne,  IsNull(MOVC.MovEne,0) MovCEne, (sum(CASE WHEN P.nPresuMes = 1  THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovEne,0) + IsNull(Pro.ProEne,0))) DMEne," _
                       & " sum(CASE WHEN P.nPresuMes = 2  THEN P.nPresuRubMesMonIni ELSE 0 END) Feb, (IsNull(Cta.MovFeb,0) + IsNull(Pro.ProFeb,0)) MovFeb,  IsNull(MOVC.MovFeb,0) MovCFeb, (sum(CASE WHEN P.nPresuMes = 2  THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovFeb,0) + IsNull(Pro.ProFeb,0))) DMFeb," _
                       & " sum(CASE WHEN P.nPresuMes = 3  THEN P.nPresuRubMesMonIni ELSE 0 END) Mar, (IsNull(Cta.MovMar,0) + IsNull(Pro.ProMar,0)) MovMar,  IsNull(MOVC.MovMar,0) MovCMar, (sum(CASE WHEN P.nPresuMes = 3  THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovMar,0) + IsNull(Pro.ProMar,0))) DMMar," _
                       & " sum(CASE WHEN P.nPresuMes = 4  THEN P.nPresuRubMesMonIni ELSE 0 END) Abr, (IsNull(Cta.MovAbr,0) + IsNull(Pro.ProAbr,0)) MovAbr,  IsNull(MOVC.MovAbr,0) MovCAbr, (sum(CASE WHEN P.nPresuMes = 4  THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovAbr,0) + IsNull(Pro.ProAbr,0))) DMAbr," _
                       & " sum(CASE WHEN P.nPresuMes = 5  THEN P.nPresuRubMesMonIni ELSE 0 END) May, (IsNull(Cta.MovMay,0) + IsNull(Pro.ProMay,0)) MovMay,  IsNull(MOVC.MovMay,0) MovCMay, (sum(CASE WHEN P.nPresuMes = 5  THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovMay,0) + IsNull(Pro.ProMay,0))) DMMay," _
                       & " sum(CASE WHEN P.nPresuMes = 6  THEN P.nPresuRubMesMonIni ELSE 0 END) Jun, (IsNull(Cta.MovJun,0) + IsNull(Pro.ProJun,0)) MovJun,  IsNull(MOVC.MovJun,0) MovCJun, (sum(CASE WHEN P.nPresuMes = 6  THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovJun,0) + IsNull(Pro.ProJun,0))) DMJun," _
                       & " sum(CASE WHEN P.nPresuMes = 7  THEN P.nPresuRubMesMonIni ELSE 0 END) Jul, (IsNull(Cta.MovJul,0) + IsNull(Pro.ProJul,0)) MovJul,  IsNull(MOVC.MovJul,0) MovCJul, (sum(CASE WHEN P.nPresuMes = 7  THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovJul,0) + IsNull(Pro.ProJul,0))) DMJul," _
                       & " sum(CASE WHEN P.nPresuMes = 8  THEN P.nPresuRubMesMonIni ELSE 0 END) Ago, (IsNull(Cta.MovAgo,0) + IsNull(Pro.ProAgo,0)) MovAgo,  IsNull(MOVC.MovAgo,0) MovCAgo, (sum(CASE WHEN P.nPresuMes = 8  THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovAgo,0) + IsNull(Pro.ProAgo,0))) DMAgo," _
                       & " sum(CASE WHEN P.nPresuMes = 9  THEN P.nPresuRubMesMonIni ELSE 0 END) Seti, (IsNull(Cta.MovSet,0) + IsNull(Pro.ProSet,0)) MovSet, IsNull(MOVC.MovSet,0) MovCSet, (sum(CASE WHEN P.nPresuMes = 9  THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovSet,0) + IsNull(Pro.ProSet,0))) DMSet," _
                       & " sum(CASE WHEN P.nPresuMes = 10 THEN P.nPresuRubMesMonIni ELSE 0 END) Oct, (IsNull(Cta.MovOct,0) + IsNull(Pro.ProOct,0)) MovOct,  IsNull(MOVC.MovOct,0) MovCOct, (sum(CASE WHEN P.nPresuMes = 10 THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovOct,0) + IsNull(Pro.ProOct,0))) DMOct," _
                       & " sum(CASE WHEN P.nPresuMes = 11 THEN P.nPresuRubMesMonIni ELSE 0 END) Nov, (IsNull(Cta.MovNov,0) + IsNull(Pro.ProNov,0)) MovNov,  IsNull(MOVC.MovNov,0) MovCNov, (sum(CASE WHEN P.nPresuMes = 11 THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovNov,0) + IsNull(Pro.ProNov,0))) DMNov," _
                       & " sum(CASE WHEN P.nPresuMes = 12 THEN P.nPresuRubMesMonIni ELSE 0 END) Dic, (IsNull(Cta.MovDic,0) + IsNull(Pro.ProDic,0)) MovDic,  IsNull(MOVC.MovDic,0) MovCDic, (sum(CASE WHEN P.nPresuMes = 12 THEN P.nPresuRubMesMonIni ELSE 0 END) - (IsNull(Cta.MovDic,0) + IsNull(Pro.ProDic,0))) DMDic" _
                       & " FROM PresuRubroAnio R WITH (NOLOCK) "
                       
                If pnTipoSaldo = 2 Then
                    tmpSql = tmpSql & " LEFT JOIN ( " _
                           & "  SELECT RCs.nPresuAnio , RCs.nPresuCod , rcs.cPresuRubCod , SUM(cs.nCtaSaldoImporte) nSaldoIni " _
                           & "  FROM PresuRubroCta RCs WITH (NOLOCK) JOIN " & sServidor & "CtaSaldo cs WITH (NOLOCK) ON cs.cCtaContCod like SubString(RCs.cCtaContCod,1,2) + '" & sqlMoneda & "' + SubString(RCs.cCtaContCod,4,20) + '%' " _
                           & "  WHERE  RCs.nPresuAnio = " & nVal(psPeriodo) & " And RCs.nPresuCod = " & nVal(psPresupuestoCod) & " And cs.dCtaSaldoFecha = " _
                           & "      (SELECT Max(dCtaSaldoFecha) FROM " & sServidor & "CtaSaldo sdo1 WITH (NOLOCK) " _
                           & "       WHERE sdo1.cCtaContCod = cs.cCtaContCod and sdo1.dCtaSaldoFecha <= '" & IIf(psPeriodo = "2004", Format(CDate("01/05/2007") - 1, gsFormatoFecha), Format(CDate("01/01/" & psPeriodo) - 1, gsFormatoFecha)) & "' ) " _
                           & "  GROUP BY RCs.nPresuAnio , RCs.nPresuCod , rcs.cPresuRubCod " _
                           & ") Ini ON R.nPresuAnio = Ini.nPresuAnio AND R.nPresuCod = Ini.nPresuCod AND R.cPresuRubCod = Ini.cPresuRubCod "
                End If
                
                tmpSql = tmpSql & " LEFT JOIN PresuRubroMes P WITH (NOLOCK) ON R.nPresuAnio = P.nPresuAnio AND P.cPresuRubCod = R.cPresuRubCod AND P.nPresuCod = R.nPresuCod " _
                       & " LEFT JOIN" _
                       & "   (SELECT Mov.nPresuAnio , Mov.nPresuCod , Mov.cPresuRubCod , Mov.cAbrev, " _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____01%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovEne," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____02%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovFeb," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____03%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovMar," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____04%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovAbr," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____05%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovMay," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____06%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovJun," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____07%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovJul," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____08%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovAgo," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____09%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovSet," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____10%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovOct," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____11%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovNov," _
                       & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____12%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovDic" _
                       & "      FROM ( SELECT RC.nPresuAnio , RC.nPresuCod , rc.cPresuRubCod , tip.cAbrev, M.cMovNro, mc.nMovImporte " _
                       & "             FROM PresuRubroCta RC WITH (NOLOCK) " _
                       & "              JOIN MovCta MC WITH (NOLOCK) ON MC.cCtaContCod like SubString(RC.cCtaContCod,1,2) + '" & sqlMoneda & "' + SubString(RC.cCtaContCod,4,20) + '%' " _
                       & "              JOIN Mov    M  WITH (NOLOCK) ON MC.nMovNro = M.nMovNro" _
                       & "              JOIN (Select cCtaContCod cValor, cCtaCaracter cAbrev from CtaContClase WITH (NOLOCK)) Tip ON substring(MC.cCtaContCod,1,len(Tip.cValor)) = Tip.cValor "
            
            If pnMovContablesAdd = 1 Then
            
                tmpSqlMovCnt = " LEFT JOIN" _
                             & "   (SELECT Mov.nPresuAnio , Mov.nPresuCod , Mov.cPresuRubCod , Mov.cAbrev, " _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____01%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovEne," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____02%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovFeb," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____03%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovMar," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____04%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovAbr," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____05%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovMay," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____06%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovJun," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____07%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovJul," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____08%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovAgo," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____09%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovSet," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____10%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovOct," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____11%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovNov," _
                             & "      (IsNull(Sum(CASE WHEN Mov.cMovNro like '____12%' THEN (CASE WHEN Mov.cAbrev = 'D' THEN (Mov.nMovImporte) ELSE (Mov.nMovImporte * -1) END) ELSE 0 END),0)) MovDic" _
                             & "      FROM ( SELECT RC.nPresuAnio , RC.nPresuCod , rc.cPresuRubCod , tip.cAbrev, M.cMovNro, mc.nMovImporte " _
                             & "             FROM PresuRubroCta RC WITH (NOLOCK) " _
                             & "              JOIN MovCta MC WITH (NOLOCK) ON MC.cCtaContCod like SubString(RC.cCtaContCod,1,2) + '" & sqlMoneda & "' + SubString(RC.cCtaContCod,4,20) + '%' " _
                             & "              JOIN Mov    M  WITH (NOLOCK) ON MC.nMovNro = M.nMovNro" _
                             & "              JOIN (Select cCtaContCod cValor, cCtaCaracter cAbrev from CtaContClase WITH (NOLOCK)) Tip ON substring(MC.cCtaContCod,1,len(Tip.cValor)) = Tip.cValor "
                
                '
                tmpSqlMovCnt = tmpSqlMovCnt & _
                    "           WHERE RC.nPresuAnio = " & nVal(psPeriodo) & " And RC.nPresuCod = " & nVal(psPresupuestoCod) & " And " & IIf(Left(psPeriodo, 4) = "2004", "m.cMovNro > '20070430'", "m.cMovNro LIKE '" & Left(psPeriodo, 4) & "%'") & " " & _
                    "               AND SubString(mc.cCtaContCod,3,1) IN (" & sqlAdd & ") " & _
                    "               AND m.nMovFlag Not In (" & MovFlag.gMovFlagEliminado & "," & MovFlag.gMovFlagModificado & ") AND m.nMovEstado = " & MovEstado.gMovEstContabMovContable
            
                
                tmpSqlMovCnt = tmpSqlMovCnt & _
                    "     ) Mov " & _
                    "     GROUP BY Mov.nPresuAnio, Mov.nPresuCod, Mov.cPresuRubCod, Mov.cAbrev " & _
                    " ) MOVC ON r.nPresuAnio = MOVC.nPresuAnio and r.nPresuCod = MOVC.nPresuCod and r.cPresuRubCod = MOVC.cPresuRubCod "
                
            Else
                tmpSqlMovCnt = " LEFT JOIN " _
                        & " ( SELECT RC.nPresuAnio, RC.nPresuCod, RC.cPresuRubCod," _
                        & " 0 MovEne, 0 MovFeb, 0 MovMar, 0 MovAbr, 0 MovMay, 0 MovJun," _
                        & " 0 MovJul, 0 MovAgo, 0 MovSet, 0 MovOct, 0 MovNov, 0 MovDic " _
                        & " FROM PresuRubroAnio RC WITH (NOLOCK) " _
                        & " " _
                        & " ) MOVC ON R.nPresuAnio = MOVC.nPresuAnio AND R.nPresuCod = MOVC.nPresuCod and R.cPresuRubCod = MOVC.cPresuRubCod "
                
                'tmpSqlMovCnt = tmpSqlMovCnt _
                    & " WHERE R.nPresuAnio = " & nVal(psPeriodo) & " And R.nPresuCod = " & nVal(psPresupuestoCod) & " AND R.cFlag Is Null " _
                    & " GROUP BY R.cPresuRubCod, R.nPresuAnio, R.nPresuCod, R.nPresuRubMonAnt, R.nPresuRubMonPre, R.nPresuRubMonCre1, R.nPresuRubMonCre2, " & IIf(pnTipoSaldo = 2, " nSaldoIni, ", "") _
                    & "        Cta.MovEne, Cta.MovFeb, Cta.MovMar, Cta.MovAbr, Cta.MovMay, Cta.MovJun, Cta.MovJul, Cta.MovAgo, Cta.MovSet, Cta.MovOct, Cta.MovNov, Cta.MovDic, " _
                    & "        Pro.ProEne, Pro.ProFeb, Pro.ProMar, Pro.ProAbr, Pro.ProMay, Pro.ProJun, Pro.ProJul, Pro.ProAgo, Pro.ProSet, Pro.ProOct, Pro.ProNov, Pro.ProDic " _

            End If
                                   
            
            If pnProy = 1 Then
                tmpPry = " Left Join" _
                       & " (SELECT RC.nPresuAnio, RC.cPresuRubCod," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____01%' THEN (MSD.nMonto) ELSE 0 END),0)) ProEne," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____02%' THEN (MSD.nMonto) ELSE 0 END),0)) ProFeb," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____03%' THEN (MSD.nMonto) ELSE 0 END),0)) ProMar," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____04%' THEN (MSD.nMonto) ELSE 0 END),0)) ProAbr," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____05%' THEN (MSD.nMonto) ELSE 0 END),0)) ProMay," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____06%' THEN (MSD.nMonto) ELSE 0 END),0)) ProJun," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____07%' THEN (MSD.nMonto) ELSE 0 END),0)) ProJul," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____08%' THEN (MSD.nMonto) ELSE 0 END),0)) ProAgo," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____09%' THEN (MSD.nMonto) ELSE 0 END),0)) ProSet," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____10%' THEN (MSD.nMonto) ELSE 0 END),0)) ProOct," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____11%' THEN (MSD.nMonto) ELSE 0 END),0)) ProNov," _
                       & " Abs(IsNull(Sum(CASE WHEN M.cMovNro like '____12%' THEN (MSD.nMonto) ELSE 0 END),0)) ProDic" _
                       & " FROM MovServicios MSD WITH (NOLOCK) " _
                       & " JOIN Mov M WITH (NOLOCK) ON MSD.nMovNro = M.nMovNro  " & IIf(pbConsiderarCierreAño, "", " And M.cOpeCod Not Like '70185%' ") & " " _
                       & " JOIN MovCta MC WITH (NOLOCK) ON M.nMovNro = MC.nMovNro" _
                       & " JOIN PresuRubroCta RC WITH (NOLOCK) ON MC.cCtaContCod Like Left(RC.cCtaContCod,2) + '" & sqlMoneda & "' + Substring(RC.cCtaContCod,4,20) " _
                       & " WHERE M.cMovNro like '" & nVal(psPeriodo) & "%' AND  M.nMovEstado= '" & gMovEstPresupProyectado & "' AND M.nMovFlag Not In (" & gMovFlagEliminado & "," & gMovFlagModificado & ") And RC.nPresuCod = " & nVal(psPresupuestoCod) & "" _
                       & " GROUP BY RC.nPresuAnio, RC.cPresuRubCod " _
                       & " ) Pro  ON R.nPresuAnio = Pro.nPresuAnio AND R.cPresuRubCod = Pro.cPresuRubCod"
                       
               tmpPry = tmpPry _
                   & " WHERE R.cFlag Is Null And R.nPresuAnio = " & nVal(psPeriodo) & " And r.nPresuCod = " & nVal(psPresupuestoCod) _
                   & " GROUP BY R.cPresuRubCod, R.nPresuAnio, R.nPresuCod, R.nPresuRubMonAnt, R.nPresuRubMonPre, R.nPresuRubMonCre1, R.nPresuRubMonCre2, " _
                   & "        Cta.MovEne, Cta.MovFeb, Cta.MovMar, Cta.MovAbr, Cta.MovMay, Cta.MovJun, Cta.MovJul, Cta.MovAgo, Cta.MovSet, Cta.MovOct, Cta.MovNov, Cta.MovDic, " _
                   & "        Pro.ProEne, Pro.ProFeb, Pro.ProMar, Pro.ProAbr, Pro.ProMay, Pro.ProJun, Pro.ProJul, Pro.ProAgo, Pro.ProSet, Pro.ProOct, Pro.ProNov, Pro.ProDic, "
                tmpPry = tmpPry & " MOVC.MovEne , MOVC.MovFeb, MOVC.MovMar, MOVC.MovAbr, MOVC.MovMay, MOVC.MovJun, MOVC.MovJul, MOVC.MovAgo, MOVC.MovSet, MOVC.MovOct, MOVC.MovNov, MOVC.MovDic"
                
            Else   'Movimientos SIN proyeccion
            
                tmpPry = " LEFT JOIN " _
                       & " ( SELECT RC.nPresuAnio, RC.nPresuCod, RC.cPresuRubCod," _
                       & " 0 ProEne, 0 ProFeb, 0 ProMar, 0 ProAbr, 0 ProMay, 0 ProJun," _
                       & " 0 ProJul, 0 ProAgo, 0 ProSet, 0 ProOct, 0 ProNov, 0 ProDic " _
                       & " FROM PresuRubroAnio RC WITH (NOLOCK) " _
                       & " " _
                       & " ) Pro ON R.nPresuAnio = Pro.nPresuAnio AND R.nPresuCod = Pro.nPresuCod and R.cPresuRubCod = Pro.cPresuRubCod "
               tmpPry = tmpPry _
                   & " WHERE R.nPresuAnio = " & nVal(psPeriodo) & " And R.nPresuCod = " & nVal(psPresupuestoCod) & " AND R.cFlag Is Null " _
                   & " GROUP BY R.cPresuRubCod, R.nPresuAnio, R.nPresuCod, R.nPresuRubMonAnt, R.nPresuRubMonPre, R.nPresuRubMonCre1, R.nPresuRubMonCre2, " & IIf(pnTipoSaldo = 2, " nSaldoIni, ", "") _
                   & "        Cta.MovEne, Cta.MovFeb, Cta.MovMar, Cta.MovAbr, Cta.MovMay, Cta.MovJun, Cta.MovJul, Cta.MovAgo, Cta.MovSet, Cta.MovOct, Cta.MovNov, Cta.MovDic, " _
                   & "        Pro.ProEne, Pro.ProFeb, Pro.ProMar, Pro.ProAbr, Pro.ProMay, Pro.ProJun, Pro.ProJul, Pro.ProAgo, Pro.ProSet, Pro.ProOct, Pro.ProNov, Pro.ProDic, " _
                   & "        MOVC.MovEne, MOVC.MovFeb, MOVC.MovMar, MOVC.MovAbr, MOVC.MovMay, MOVC.MovJun, MOVC.MovJul, MOVC.MovAgo, MOVC.MovSet, MOVC.MovOct, MOVC.MovNov, MOVC.MovDic"
            End If

            'DISTRIBUIDO
            'Agregar la parte contable de esta funcion
            tmpSql = tmpSql & _
                "           WHERE RC.nPresuAnio = " & nVal(psPeriodo) & " And RC.nPresuCod = " & nVal(psPresupuestoCod) & " And m.cMovNro LIKE '" & Left(psPeriodo, 4) & "%' " & _
                "               AND SubString(mc.cCtaContCod,3,1) IN (" & sqlAdd & ") " & _
                "               AND m.nMovFlag Not In (" & MovFlag.gMovFlagEliminado & "," & MovFlag.gMovFlagModificado & ") AND " & IIf(pnBala = 1, "m.nMovEstado = " & MovEstado.gMovEstContabMovContable & "  " & IIf(pbConsiderarCierreAño, "", " And M.cOpeCod Not Like '70185%' ") & " ", " (m.nMovEstado= " & MovEstado.gMovEstPresupAceptado & " OR " & _
                "                                   (m.nMovEstado = " & MovEstado.gMovEstPresupAceptado & " AND Not Exists(Select Distinct m2.nMovEstado " & _
                "                                                           From mov m2 WITH (NOLOCK) join movref mr WITH (NOLOCK) " & _
                "                                                               On m2.nmovnro = mr.nmovnroref " & _
                "                                                               And mr.nmovnro = mc.nmovnro " & _
                "                                                           Where m2.nMovFlag <> '" & gMovFlagEliminado & "' And m2.nMovEstado = '" & gMovEstPresupAceptado & "')  ) ) ")
            
            'SI pnBala = 0 no SELECCIONADO
            If pnBala = 0 Then 'Debemos unir las Ordenes de Compra y Servicios
               tmpSql = tmpSql _
                       & "    UNION   " _
                       & "    SELECT RC.nPresuAnio , RC.nPresuCod , rc.cPresuRubCod , tip.cAbrev, m.cMovNro, mc.nMovImporte " _
                       & "    FROM PresuRubroCta RC WITH (NOLOCK) " _
                       & "      JOIN MovCta MC WITH (NOLOCK) ON MC.cCtaContCod like SubString(RC.cCtaContCod,1,2) + '" & sqlMoneda & "' + SubString(RC.cCtaContCod,4,20) + '%' " _
                       & "      JOIN Mov    M  WITH (NOLOCK) ON MC.nMovNro = M.nMovNro" _
                       & "      JOIN (Select cCtaContCod cValor, cCtaCaracter cAbrev from CtaContClase WITH (NOLOCK) ) Tip ON substring(MC.cCtaContCod,1,len(Tip.cValor)) = Tip.cValor " _
                       & "   WHERE RC.nPresuAnio = " & nVal(psPeriodo) & " And RC.nPresuCod = " & nVal(psPresupuestoCod) & " AND m.cMovNro LIKE '" & Left(psPeriodo, 4) & "%' AND SubString(mc.cCtaContCod,3,1) IN (" & sqlAdd & ") " _
                       & "      AND M.nMovFlag Not In (" & gMovFlagEliminado & "," & gMovFlagModificado & ") And m.nMovEstado = '" & gMovEstPresupAceptado & "' "
            
            End If
            
            tmpSql = tmpSql & _
                "     ) Mov " & _
                "     GROUP BY Mov.nPresuAnio, Mov.nPresuCod, Mov.cPresuRubCod, Mov.cAbrev " & _
                " ) Cta ON r.nPresuAnio = cta.nPresuAnio and r.nPresuCod = cta.nPresuCod and r.cPresuRubCod = cta.cPresuRubCod "
                
            tmpSql = tmpSql & tmpSqlMovCnt & tmpPry
               
            tmpSql = tmpSql _
                & " ) Eje ON Eje.nPresuAnio = Presu.nPresuAnio AND Eje.nPresuCod = Presu.nPresuCod and Eje.cPresuRubCod LIKE Presu.cPresuRubCod + '%' " _
                & " WHERE Presu.nPresuAnio = " & nVal(psPeriodo) & " And Presu.nPresuCod = " & nVal(psPresupuestoCod) & " AND Presu.cFlag Is Null " _
                & " GROUP BY Presu.cPresuRubCod, Presu.cPresuRubDescripcion " _
                & " ORDER BY Presu.cPresuRubCod "
            
        End If
       
        Set GetPresupuestoEjec = oCon.CargaRecordSet(tmpSql)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetPresupuestoErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function

'Devuelve todos los Items de la Asistencia Medica Familiar
Public Function GetRubroCta(psPeriodo As String, psPresupuestoCod As String, psRubro As String) As ADODB.Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetPresupuestoErr

    If oCon.AbreConexion() Then
        sqlA = " SELECT cPresuRubCod cCodRub, cCtaContCod cCtaCnt FROM PresuRubroCta Where " _
             & " nPresuAnio = " & nVal(psPeriodo) & " And nPresuCod = " & nVal(psPresupuestoCod) & " " _
             & " And cPresuRubCod = '" & psRubro & "'" _
             & " Order By cCtaContCod "
        Set GetRubroCta = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetPresupuestoErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function

'Devuelve todos los Items de la Asistencia Medica Familiar
Public Function GetNroPresupuesto() As Integer
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetPresupuestoErr

    If oCon.AbreConexion() Then
        sqlA = " Select IsNull(Max(nPresuCod),0) + 1 Campo From PresuClase"
        GetNroPresupuesto = oCon.CargaRecordSet(sqlA).Fields(0)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetPresupuestoErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function
   
'Devuelve todos los Items de la Asistencia Medica Familiar
'##ModelId=3A96A676015F
Public Function GetPresupRubroMes(psPeriodo As String, psPresupuestoCod As String, psMoneda As String, psRubro As String) As ADODB.Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetPresupuestoErr
    
    If oCon.AbreConexion() Then
        sqlA = " SELECT t.nConsValor, t.cConsDescripcion cNomTab, p.nPresuRubMesMonIni " _
             & " FROM Constante T" _
             & " LEFT JOIN PresuRubroMes  P ON t.nConsValor = p.nPresuMes and P.nPresuAnio = " & nVal(psPeriodo) & " and P.nPresuCod = " & nVal(psPresupuestoCod) & " and p.cPresuRubCod = '" & psRubro & "' " _
             & " LEFT JOIN PresuRubroAnio R ON P.nPresuAnio = R.nPresuAnio And R.nPresuCod = P.nPresuCod " _
             & "       AND p.cPresuRubCod = R.cPresuRubCod " _
             & " WHERE t.nConsCod like '1010' " _
             & " ORDER BY t.nConsValor"

        Set GetPresupRubroMes = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetPresupuestoErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function
    
'Devuelve Rubros del Presupuesto
Public Function GetPresupRubro(psPeriodo As String, psPresupuestoCod As String, Optional psPresuRubCod As String = "", Optional psLockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetPresupuestoErr

    If oCon.AbreConexion() Then
        sqlA = " SELECT r.cPresuRubCod, r.cPresuRubDescripcion, r.nPresuRubMonAnt nMonIni, r.nPresuRubMonPre nMonto, r.nPresuRubMonCre1 nMonCre1, r.nPresuRubMonCre2 nMonCre2," _
             & " (r.nPresuRubMonPre + r.nPresuRubMonCre1 + r.nPresuRubMonCre2) nTotal" _
             & " FROM PresuRubroAnio r " _
             & " WHERE nPresuAnio = " & nVal(psPeriodo) & " AND r.nPresuCod = " & nVal(psPresupuestoCod) & " AND r.cFlag Is Null " _
             & IIf(psPresuRubCod = "", "", " and cPresuRubCod " & IIf(InStr(psPresuRubCod, "%") = 0 And InStr(psPresuRubCod, "_") = 0, "=", "LIKE") & " '" & psPresuRubCod & "' ") _
             & " ORDER BY r.cPresuRubCod "
        Set GetPresupRubro = oCon.CargaRecordSet(sqlA, psLockType)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetPresupuestoErr:
    Call RaiseError(Err.MyUnhandledError, "DPresupuesto:GetPresupRubro Method")
End Function

Public Function IsUltNivel(ByVal psPeriodo As String, ByVal psPresupuestoCod As String, ByVal pCodRub As String) As Boolean
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim tmpSql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    tmpSql = " SELECT R.cPresuRubCod, R.cPresuRubDescripcion FROM PresuRubroAnio R INNER JOIN PresuRubroMes P ON R.nPresuCod = P.nPresuCod and R.nPresuAnio = P.nPresuAnio " _
           & " WHERE R.nPresuAnio = " & nVal(psPeriodo) & " AND P.nPresuCod = " & nVal(psPresupuestoCod) & "  And R.cPresuRubCod like '" & pCodRub & "__' AND cFlag Is Null"
    
    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    If (tmpReg.BOF Or tmpReg.EOF) Then
        IsUltNivel = True
    Else
        IsUltNivel = False
    End If
End Function

Public Function IsMonto(psPeriodo As String, psPresupuestoCod As String, psRubro As String) As Boolean
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim tmpSql As String
    IsMonto = False
    
    tmpSql = " SELECT p.nPresuRubMesMonIni nMonIni " _
           & " FROM PresuRubroMes P" _
           & " WHERE p.nPresuAnio = " & nVal(psPeriodo) & " AND P.nPresuCod =  " & nVal(psPresupuestoCod) & "  AND p.cPresuRubCod =  '" & psRubro & "'"
    oCon.AbreConexion
    Set tmpReg = oCon.CargaRecordSet(tmpSql)
    If Not (tmpReg.BOF Or tmpReg.EOF) Then
        IsMonto = True
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Function

#If DebugMode Then
    Public Property Get ClassDebugID() As Long
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function GetPresupTipRep(pnTipo As Integer, psPeriodo As String, psPresupuestoCod As String) As ADODB.Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetPresupuestoErr

    If oCon.AbreConexion() Then
    If pnTipo = 1 Then
            sqlA = " SELECT rr.cCodRub, rr.cDesRub, sum(res.nMonIni) nMonIni, sum(res.nMonto) nMonto, sum(res.nMonCre1) nMonCre1, sum(res.nMonCre2) nMonCre2, SUM(res.nMonto + res.nMonCre1 + res.nMonCre2) Total, SUM(Enero) Enero, SUM(Febrero) Febrero, SUM(Marzo) Marzo, SUM(Abril) Abril, SUM(Mayo) Mayo, SUM(Junio) Junio, SUM(Julio) Julio, SUM(Agosto) Agosto, SUM(Setiembre) Setiembre, SUM(Octubre) Octubre, SUM(Noviembre) Noviembre, SUM(Diciembre) Diciembre" _
                  & " FROM (SELECT DISTINCT cPresuRubCod cCodRub, cPresuRubDescripcion cDesRub FROM PresuRubroAnio  WHERE nPresuAnio = " & nVal(psPeriodo) & " AND nPresuCod = " & nVal(psPresupuestoCod) & ") RR " _
                  & " LEFT JOIN (SELECT r.cPresuRubCod cCodRub, r.cPresuRubDescripcion cDesRub, r.nPresuRubMonAnt nMonIni, r.nPresuRubMonPre nMonto, r.nPresuRubMonCre1 nMonCre1, r.nPresuRubMonCre2 nMonCre2," _
                  & "          Sum(CASE WHEN p.nPresuMes = '01' THEN p.nPresuRubMesMonIni ELSE 0 END) Enero," _
                  & "          Sum(CASE WHEN p.nPresuMes = '02' THEN p.nPresuRubMesMonIni ELSE 0 END) Febrero," _
                  & "          Sum(CASE WHEN p.nPresuMes = '03' THEN p.nPresuRubMesMonIni ELSE 0 END) Marzo," _
                  & "          Sum(CASE WHEN p.nPresuMes = '04' THEN p.nPresuRubMesMonIni ELSE 0 END) Abril," _
                  & "          Sum(CASE WHEN p.nPresuMes = '05' THEN p.nPresuRubMesMonIni ELSE 0 END) Mayo," _
                  & "          Sum(CASE WHEN p.nPresuMes = '06' THEN p.nPresuRubMesMonIni ELSE 0 END) Junio," _
                  & "          Sum(CASE WHEN p.nPresuMes = '07' THEN p.nPresuRubMesMonIni ELSE 0 END) Julio," _
                  & "          Sum(CASE WHEN p.nPresuMes = '08' THEN p.nPresuRubMesMonIni ELSE 0 END) Agosto," _
                  & "          Sum(CASE WHEN p.nPresuMes = '09' THEN p.nPresuRubMesMonIni ELSE 0 END) Setiembre," _
                  & "          Sum(CASE WHEN p.nPresuMes = '10' THEN p.nPresuRubMesMonIni ELSE 0 END) Octubre," _
                  & "          Sum(CASE WHEN p.nPresuMes = '11' THEN p.nPresuRubMesMonIni ELSE 0 END) Noviembre," _
                  & "          Sum(CASE WHEN p.nPresuMes = '12' THEN p.nPresuRubMesMonIni ELSE 0 END) Diciembre" _
                  & " FROM PresuRubroAnio R" _
                  & "      LEFT JOIN PresuRubroMes P ON r.nPresuCod = p.nPresuCod and r.nPresuAnio = p.nPresuAnio And R.cPresuRubCod = P.cPresuRubCod" _
                  & "     " _
                  & " WHERE P.nPresuCod = " & nVal(psPresupuestoCod) & " and P.nPresuAnio = " & nVal(psPeriodo) & " And r.cFlag Is Null" _
                  & " GROUP BY r.cPresuRubCod, r.cPresuRubDescripcion, r.nPresuRubMonAnt, r.nPresuRubMonPre, r.nPresuRubMonCre1, r.nPresuRubMonCre2) res" _
                  & " ON SUBSTRING(res.cCodRub,1,LEN(rr.cCodRub)) = rr.cCodRub" _
                  & " GROUP BY rr.cCodRub, rr.cDesRub  ORDER BY rr.cCodRub"
        ElseIf pnTipo = 2 Then
             sqlA = " SELECT rr.cCodRub, rr.cDesRub, sum(res.nMonIni) nMonIni, sum(res.nMonto) nMonto, sum(res.nMonCre1) nMonCre1, sum(res.nMonCre2) nMonCre2, SUM(res.nMonto + res.nMonCre1 + res.nMonCre2) Total, SUM(Primero) Primero, SUM(Segundo) Segundo, SUM(Tercero) Tercero, SUM(Cuarto) Cuarto" _
                  & " FROM (SELECT DISTINCT cPresuRubCod cCodRub, cPresuRubDescripcion cDesRub FROM PresuRubroAnio  WHERE nPresuAnio = " & nVal(psPeriodo) & " AND nPresuCod = " & nVal(psPresupuestoCod) & " ) RR" _
                  & " LEFT JOIN (SELECT r.cPresuRubCod cCodRub, r.cPresuRubDescripcion cDesRub, r.nPresuRubMonAnt nMonIni, r.nPresuRubMonPre nMonto, r.nPresuRubMonCre1 nMonCre1, r.nPresuRubMonCre2 nMonCre2," _
                  & "          Sum(CASE WHEN p.nPresuMes = '01' Or p.nPresuMes = '02' Or p.nPresuMes = '03' THEN p.nPresuRubMesMonIni ELSE 0 END) Primero," _
                  & "          Sum(CASE WHEN p.nPresuMes = '04' Or p.nPresuMes = '05' Or p.nPresuMes = '06' THEN p.nPresuRubMesMonIni ELSE 0 END) Segundo," _
                  & "          Sum(CASE WHEN p.nPresuMes = '07' Or p.nPresuMes = '08' Or p.nPresuMes = '09' THEN p.nPresuRubMesMonIni ELSE 0 END) Tercero," _
                  & "          Sum(CASE WHEN p.nPresuMes = '10' Or p.nPresuMes = '11' Or p.nPresuMes = '12' THEN p.nPresuRubMesMonIni ELSE 0 END) Cuarto" _
                  & " FROM PresuRubroAnio R LEFT JOIN PresuRubroMes P ON r.nPresuCod = p.nPresuCod and r.nPresuAnio = p.nPresuAnio And R.cPresuRubCod = P.cPresuRubCod" _
                  & " WHERE R.nPresuCod = " & nVal(psPresupuestoCod) & " and R.nPresuAnio = " & nVal(psPeriodo) & " And r.cFlag Is Null" _
                  & " GROUP BY r.cPresuRubCod, r.cPresuRubDescripcion, r.nPresuRubMonAnt, r.nPresuRubMonPre, r.nPresuRubMonCre1, r.nPresuRubMonCre2) res" _
                  & " ON SUBSTRING(res.cCodRub,1,LEN(rr.cCodRub)) = rr.cCodRub" _
                  & " GROUP BY rr.cCodRub, rr.cDesRub  ORDER BY rr.cCodRub"
        Else
             sqlA = " SELECT rr.cCodRub, rr.cDesRub, sum(res.nMonIni) nMonIni, sum(res.nMonto) nMonto, sum(res.nMonCre1) nMonCre1, sum(res.nMonCre2) nMonCre2, SUM(res.nMonto + res.nMonCre1 + res.nMonCre2) Total, SUM(Primero) Primero, SUM(Segundo) Segundo" _
                  & " FROM (SELECT DISTINCT cPresuRubCod cCodRub, cPresuRubDescripcion cDesRub FROM PresuRubroAnio  WHERE nPresuAnio = " & nVal(psPeriodo) & " AND nPresuCod = " & nVal(psPresupuestoCod) & ") RR " _
                  & " LEFT JOIN (SELECT r.cPresuRubCod cCodRub, r.cPresuRubDescripcion cDesRub, r.nPresuRubMonAnt nMonIni, r.nPresuRubMonPre nMonto, r.nPresuRubMonCre1 nMonCre1, r.nPresuRubMonCre2 nMonCre2," _
                  & "          Sum(CASE WHEN p.nPresuMes = '01' Or p.nPresuMes = '02' Or p.nPresuMes = '03' Or p.nPresuMes = '04' Or p.nPresuMes = '05' Or p.nPresuMes = '06' THEN p.nPresuRubMesMonIni ELSE 0 END) Primero," _
                  & "          Sum(CASE WHEN p.nPresuMes = '07' Or p.nPresuMes = '08' Or p.nPresuMes = '09' Or p.nPresuMes = '10' Or p.nPresuMes = '11' Or p.nPresuMes = '12' THEN p.nPresuRubMesMonIni ELSE 0 END) Segundo" _
                  & " FROM PresuRubroAnio R" _
                  & " LEFT JOIN PresuRubroMes P ON r.nPresuAnio = p.nPresuAnio And R.cPresuRubCod = P.cPresuRubCod" _
                  & " WHERE R.nPresuCod = " & nVal(psPresupuestoCod) & " and R.nPresuAnio = " & nVal(psPeriodo) & " And r.cFlag Is Null" _
                  & " GROUP BY r.cPresuRubCod, r.cPresuRubDescripcion, r.nPresuRubMonAnt, r.nPresuRubMonPre, r.nPresuRubMonCre1, r.nPresuRubMonCre2) res" _
                  & " ON SUBSTRING(res.cCodRub,1,LEN(rr.cCodRub)) = rr.cCodRub" _
                  & " GROUP BY rr.cCodRub, rr.cDesRub  ORDER BY rr.cCodRub"
        End If
        Set GetPresupTipRep = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
GetPresupuestoErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function


Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
End Sub

'Actualiza un Rubro
Public Function DuplicaClase(psPresupuestoCod As String, psAnio As String, psPresupuestoCodCopia As String, psAnioCopia As String) As Boolean
   Dim sqlA As String
   Dim oCon As DConecta
   Set oCon = New DConecta
   On Error GoTo ModificaMontoRubroErr

   If oCon.AbreConexion() Then

      sqlA = " Delete Presurubrocta where nPresuCod = " & psPresupuestoCodCopia & " and nPresuAnio = " & psAnioCopia
      oCon.Ejecutar sqlA
      
      sqlA = " Delete Presurubromes where nPresuCod = " & psPresupuestoCodCopia & " and nPresuAnio = " & psAnioCopia
      oCon.Ejecutar sqlA

      sqlA = " Delete Presurubroanio where nPresuCod = " & psPresupuestoCodCopia & " and nPresuAnio = " & psAnioCopia
      oCon.Ejecutar sqlA

      sqlA = " Insert Presurubroanio(nPresuCod, nPresuAnio, cPresuRubCod, cPresuRubDescripcion, cMoneda, nPresuRubMonAnt, nPresuRubMonPre, nPresuRubMonCre1, nPresuRubMonCre2, cFlag)" _
            & " Select " & psPresupuestoCodCopia & " nPresuCod, " & psAnioCopia & " nPresuAnio, cPresuRubCod, cPresuRubDescripcion, cMoneda, nPresuRubMonAnt, nPresuRubMonPre, nPresuRubMonCre1, nPresuRubMonCre2, cFlag " _
            & "        From PresuRubroAnio Where nPresuCod = " & psPresupuestoCod & " And nPresuAnio = " & psAnio
      oCon.Ejecutar sqlA

       sqlA = " Insert PresuRubroMes(nPresuCod, nPresuAnio, cPresuRubCod, nPresuMes, nPresuRubMesMonIni, nPresuRubMesMonEje, nPresuRubMesMonRes)" _
            & " Select " & psPresupuestoCodCopia & " nPresuCod, " & psAnioCopia & " nPresuAnio, cPresuRubCod, nPresuMes, nPresuRubMesMonIni, nPresuRubMesMonEje, nPresuRubMesMonRes" _
            & "        From Presurubromes where nPresuCod = " & psPresupuestoCod & " and nPresuAnio = " & psAnio
      oCon.Ejecutar sqlA
      
      sqlA = " Insert Presurubrocta(nPresuCod, nPresuAnio, cPresuRubCod, cCtaContCod)" _
           & " Select " & psPresupuestoCodCopia & " nPresuCod, " & psAnioCopia & "  nPresuAnio, cPresuRubCod, cCtaContCod" _
           & "        From Presurubrocta Where nPresuCod = " & psPresupuestoCod & " and nPresuAnio = " & psAnio
      oCon.Ejecutar sqlA
      
      
      DuplicaClase = True
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
ModificaMontoRubroErr:
    DuplicaClase = False
    Call RaiseError(MyUnhandledError, "DPresupuesto: ModificaMontoRubro Method")
End Function



