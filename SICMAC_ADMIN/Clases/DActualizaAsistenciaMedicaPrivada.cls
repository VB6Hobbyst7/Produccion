VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DActualizaAsistMedicaPrivada"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A96A1E2038A"
'Actualiza en base los Datos de la Asistencia Medica Privada
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
    '##ModelId=3AB902F30274
    Private mlClassDebugID As Long
#End If

Public Sub Inicio(psServerComunes As String, psServerPersona As String, psServerAdministracion As String, psServerNegocio As String, psServerImagenes As String)
    lsServerComunes = psServerComunes
    lsServerPersona = psServerPersona
    lsServerAdministracion = psServerAdministracion
    lsServerNegocio = psServerNegocio
    lsServerImagenes = psServerImagenes
End Sub

'##ModelId=3AB902F303C9
Private Sub Class_Terminate()
    #If DebugMode Then
    'the class is being destroyed
    Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " is terminating"
    #End If
End Sub

'##ModelId=3AB902F3038D
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

    #If DebugMode Then
        'get the next available class ID, and print out
        'that the class was created successfully
        mlClassDebugID = GetNextClassDebugID()
        Debug.Print "'" & TypeName(Me) & "' instance " & CStr(mlClassDebugID) & " created"
    #End If
End Sub

'Agrega un nuevo Item de la Asistencia Medica Privada
'##ModelId=3A96A66403E4
'Public Function AgregaAsisMedPriv(psRHAsistMedPrivCod As String, psRHAsistMedPrivDescripcion As String, pnRHAsistMedPrivMonto As Currency, pnRHAsistMedPrivMontoDesc As Currency, psUltimaActualizacion As String) As Boolean
'    Dim sqlA As String
'    Dim oCon As DConecta
'    Set oCon = New DConecta
'    On Error GoTo AgregaAsisMedPrivErr
'
'    If oCon.AbreConexion() Then
'        sqlA = " Insert " & lsServerComunes & "RHAsistMedPrivTabla (cRHAsistMedPrivCod,cRHAsistMedPrivDescripcion,nRHAsistMedPrivMonto,nRHAsistMedPrivMonto,cUltimaActualizacion)" _
'             & " Values('" & psRHAsistMedPrivCod & "','" & psRHAsistMedPrivDescripcion & "'," & pnRHAsistMedPrivMonto & "," & pnRHAsistMedPrivMontoDesc & ",'" & psUltimaActualizacion & "')"
'        oCon.Ejecutar sqlA
'        AgregaAsisMedPriv = True
'        oCon.CierraConexion
'    Else
'        AgregaAsisMedPriv = False
'    End If
'    Set oCon = Nothing
'    Exit Function
'AgregaAsisMedPrivErr:
'    AgregaAsisMedPriv = False
'    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:AgregaAsisMedPriv Method")
'End Function

'Elimina  un Item de la Asistencia Medica Privada
'##ModelId=3A96A6650038
Public Function EliminaAsisMedPriv(psRHAsistMedPrivCod As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo EliminaAsisMedPrivErr
        
    If oCon.AbreConexion() Then
        sqlA = " Delete " & lsServerComunes & "RHAsistMedPrivTabla Where cRHAsistMedPrivCod = '" & psRHAsistMedPrivCod & "'"
        oCon.Ejecutar sqlA
        EliminaAsisMedPriv = True
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaAsisMedPrivErr:
    EliminaAsisMedPriv = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:EliminaAsisMedPriv Method")
End Function

'Elimina  un Item de la Asistencia Medica Privada
'##ModelId=3A96A6650038
Public Function TipoUsado(psRHAsistMedPrivCod As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo EliminaAsisMedPrivErr
        
    If oCon.AbreConexion() Then
        sqlA = " Select cRHAsistMedPrivCod From RHAsistenciaMedica Where cRHAsistMedPrivCod = '" & psRHAsistMedPrivCod & "'"
        Set rs = oCon.CargaRecordSet(sqlA)
        
        If rs.EOF And rs.BOF Then
            TipoUsado = False
        Else
            TipoUsado = True
        End If
        
        oCon.CierraConexion
    End If
    
    Set oCon = Nothing
    Exit Function
EliminaAsisMedPrivErr:
    TipoUsado = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:EliminaAsisMedPriv Method")
End Function


'Actualiza un Item de la Asistencia Medica Privada
'##ModelId=3A96A6650074
Public Function ModificaAsisMedPriv(psRHAsistMedPrivCod As String, psRHAsistMedPrivDescripcion As String, pnRHAsistMedPrivMonto As Currency, pnRHAsistMedPrivMontoDesc As Currency, psUltimaActualizacion As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo ModificaAsisMedPrivErr

    If oCon.AbreConexion() Then
        sqlA = " Update " & lsServerComunes & "RHAsistMedPrivTabla " _
             & " Set cRHAsistMedPrivDescripcion = '" & psRHAsistMedPrivDescripcion & "', nRHAsistMedPrivMonto = " & pnRHAsistMedPrivMonto & ", nRHAsistMedPrivMontoDesc = " & pnRHAsistMedPrivMontoDesc & ", cUltimaActualizacion = '" & psUltimaActualizacion & "'" _
             & " Where cRHAsistMedPrivCod = '" & psRHAsistMedPrivCod & "'"
        oCon.Ejecutar sqlA
        ModificaAsisMedPriv = True
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Exit Function
ModificaAsisMedPrivErr:
    ModificaAsisMedPriv = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:ModificaAsisMedPriv Method")
End Function

'''Public Function ModificaAsisMedPriv(psnCodCatalogo As String, pscCodCatEPS As String, psRHAsistMedPrivDescripcion As String, pnnCantPer As Integer, nCostoTotal As Currency, pnCostoIGV As Currency, pnCostoNeto As Currency, psUltimaActualizacion As String, pnAdeQuin As Currency) As Boolean
'''    Dim sqlA As String
'''    Dim oCon As DConecta
'''    Set oCon = New DConecta
'''    On Error GoTo ModificaAsisMedPrivErr
'''    If oCon.AbreConexion() Then
'''        sqlA = " Update RHCatalogoPlanEPS  " _
'''               & " set  cCodCatEPS = '" & pscCodCatEPS & "',  " _
'''               & " cDesCatalogo = '" & psRHAsistMedPrivDescripcion & "', " _
'''               & " nCantPer =" & pnnCantPer & " ,  " _
'''               & " nCostoTotal = " & nCostoTotal & " , " _
'''               & " nCostoIGV = " & pnCostoIGV & ", " _
'''               & " nCostoNeto = " & pnCostoNeto & "," _
'''               & " cUltimaActualizacion = '" & psUltimaActualizacion & "', " _
'''               & " cCodValQuin = " & pnAdeQuin & " " _
'''               & " where  nCodCatalogo = '" & psnCodCatalogo & "'  "
'''        oCon.Ejecutar sqlA
'''        ModificaAsisMedPriv = True
'''        oCon.CierraConexion
'''    End If
'''    Set oCon = Nothing
'''    Exit Function
'''ModificaAsisMedPrivErr:
'''    ModificaAsisMedPriv = False
'''    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:ModificaAsisMedPriv Method")
'''End Function

'Devuelve todos los Items de la Asistencia Medica Familiar
'##ModelId=3A96A676015F
Public Function GetAsisMedPriv(Optional pbTipoArbol As Boolean = False) As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetAsisMedPrivErr

    If oCon.AbreConexion() Then
        If Not pbTipoArbol Then
            sqlA = " Select '', cRHAsistMedPrivCod as Codigo, LTrim(RTrim(cRHAsistMedPrivDescripcion)) As Descripción, " _
                 & " Case nRHAsistMedPrivMonto When 0 Then '0.00' Else Convert(Varchar(15),Convert(Decimal(8,2),nRHAsistMedPrivMonto)) End As Monto," _
                 & " Case nRHAsistMedPrivMontoDesc When 0 Then '0.00' Else Convert(Varchar(15),Convert(Decimal(8,2),nRHAsistMedPrivMontoDesc)) End As Descuent," _
                 & " cUltimaActualizacion As Actualizacion From " & lsServerComunes & "RHAsistMedPrivTabla Order by cRHAsistMedPrivCod"
        Else
            sqlA = " Select cRHAsistMedPrivCod as Codigo, cRHAsistMedPrivDescripcion as Descripción,  Len(cRHAsistMedPrivCod) Nivel From " & lsServerComunes & "RHAsistMedPrivTabla Order by cRHAsistMedPrivCod"
        End If
        Set GetAsisMedPriv = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetAsisMedPrivErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function

'Devuelve todos los Items de la Asistencia Medica Familiar
'##ModelId=3A96A676015F
Public Function GetAsisMedPrivTabla() As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetAsisMedPrivTablaErr

    If oCon.AbreConexion() Then
        sqlA = " Select cRHAsistMedPrivDescripcion as Descripción, cRHAsistMedPrivCod as Codigo" _
             & " From " & lsServerComunes & "RHAsistMedPrivTabla Order by cRHAsistMedPrivCod"
        Set GetAsisMedPrivTabla = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetAsisMedPrivTablaErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetAsisMedPriv Method")
End Function

'Devuelve el monto asignado a un Item de asistencia Medico Familiar
'##ModelId=3A96A676019B
Public Function GetValorAsisMedPriv(psRHAsistMedPrivCod As String, pbValor As Boolean) As Double
    Dim sqlA As String
    Dim rsA As ADODB.Recordset
    Set rsA = New ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetValorAsisMedPrivErr
    
    If oCon.AbreConexion() Then
        If pbValor Then
            sqlA = "Select nRHAsistMedPrivMonto as Monto from " & lsServerComunes & "RHAsistMedPrivTabla Where cRHAsistMedPrivCod = '" & psRHAsistMedPrivCod & "'"
        Else
            sqlA = "Select nRHAsistMedPrivMontoDesc as Monto from " & lsServerComunes & "RHAsistMedPrivTabla Where cRHAsistMedPrivCod = '" & psRHAsistMedPrivCod & "'"
        End If
        Set rsA = oCon.CargaRecordSet(sqlA)
    
        If Not (rsA.BOF And rsA.EOF) Then
           GetValorAsisMedPriv = rsA!Monto
        Else
            GetValorAsisMedPriv = -1
        End If
    
        rsA.Close
        Set rsA = Nothing
    End If
    Exit Function
GetValorAsisMedPrivErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetValorAsisMedPriv Method")
End Function

#If DebugMode Then
    '##ModelId=3AB902F302C5
    Public Property Get ClassDebugID() As Long
        'if we are in debug mode, surface this property that consumers can query
        ClassDebugID = mlClassDebugID
    End Property
#End If

Public Function GetRepAsistenciaMedica(pgdFecSis As Date, pgsEmpresa As String, pgsNomAge As String, pgsAMPEmpresa As String, pbFamiliares As Boolean, Optional psCodPersona As String = "", Optional psNombre As String = "", Optional psPoliza As String = "", Optional psFecha As String = "", Optional pnMontoComImpre As Currency = 0, Optional pnMontoAjuste As Currency = 0) As String
    Dim lsCadena As String
    Dim lsAgencia As String * 25
    Dim lsSinRecargo As String * 15
    Dim lsConRecargo As String * 15
    Dim lsInclusion As String * 15
    Dim lsTotPagar As String * 15
    Dim lsPorAgencia As String * 15
    Dim lsComPorEmision As String * 15
    Dim lsTotPagarConING As String * 15
    
    Dim lnSinRecargo As Currency
    Dim lnConRecargo As Currency
    Dim lnInclusion As Currency
    Dim lnTotPagar As Currency
    Dim lnPorAgencia As Currency
    Dim lnComPorEmision As Currency
    Dim lnTotPagarConING As Currency
    
    Dim sqlR As String
    Dim rsR As ADODB.Recordset
    Dim lsTitulo As String
    Dim lnItem As Long
    Dim lnPagina As Long
    Set rsR = New ADODB.Recordset
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    
    If Not pbFamiliares Then
        lsTitulo = "POLIZA DE ASISTENCIA MEDICA DE " & Format(pgdFecSis, "mmmm") & " - " & Format(pgdFecSis, "yyyy") & "- " & pgsAMPEmpresa
        sqlR = "Select * from AsistenciaEmp Order By cAgenciaAsig"
    Else
        lsTitulo = "POLIZA DE ASISTENCIA MEDICA DE " & Format(pgdFecSis, "mmmm") & " - " & Format(pgdFecSis, "yyyy") & "- " & pgsAMPEmpresa
        sqlR = "Select * from AsistenciaFam Order By cAgenciaAsig"
    End If
    
    Set rsR = oCon.CargaRecordSet(sqlR)
    
    If Not (rsR.EOF And rsR.BOF) Then
        If psCodPersona <> "" Then
            If Not pbFamiliares Then
                lsTitulo = " POLIZA DE : " & psCodPersona & " - " & psNombre & " - PERSONAL"
            Else
                lsTitulo = " POLIZA DE : " & psCodPersona & " - " & psNombre & " - FAMILIARES"
            End If
        End If
        
        lsCadena = lsCadena & CabeceraPagina(lsTitulo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
        lsCadena = lsCadena & Encabezado("Agencia;15; ;5;S/Recargo;20; ;8;C/Recargo;10; ;4;Inclusion;10; ;4;Total Pag;10; ;5;% Agencia;10; ;2;Com. Emision;15;Tot.Pag.IGV;12;", 0)
        While Not rsR.EOF
            lsAgencia = rsR!Age
            RSet lsSinRecargo = Format(rsR!MontoI, "#,#00.0000")
            RSet lsConRecargo = Format(0, "#,##0.0000")
            RSet lsInclusion = Format(0, "#,##0.0000")
            RSet lsTotPagar = Format(rsR!MontoI, "#,##0.0000")
            RSet lsPorAgencia = Format(rsR!PorAge, "#,##0.00")
            RSet lsComPorEmision = Format(rsR!MontoAddAge, "#,##0.0000")
            
            
            If rsR!cAgenciaAsig = "01" Then
                RSet lsTotPagarConING = Format(rsR!Total + pnMontoAjuste, "#,##0.00")
                lnTotPagarConING = lnTotPagarConING + rsR!Total + pnMontoAjuste
            Else
                RSet lsTotPagarConING = Format(rsR!Total, "#,##0.00")
                lnTotPagarConING = lnTotPagarConING + IIf(IsNull(rsR!Total), 0, rsR!Total)
            End If
            
            lnSinRecargo = lnSinRecargo + IIf(IsNull(rsR!MontoI), 0, rsR!MontoI)
            lnConRecargo = lnConRecargo + 0
            lnInclusion = lnInclusion + 0
            lnTotPagar = lnTotPagar + IIf(IsNull(rsR!MontoI), 0, rsR!MontoI)
            lnPorAgencia = lnPorAgencia + IIf(IsNull(rsR!PorAge), 0, rsR!PorAge)
            lnComPorEmision = lnComPorEmision + IIf(IsNull(rsR!MontoAddAge), 0, rsR!MontoAddAge)
            
            lsCadena = lsCadena & lsAgencia & lsSinRecargo & lsConRecargo & lsInclusion & lsTotPagar & lsPorAgencia & lsComPorEmision & lsTotPagarConING & oImpresora.gPrnSaltoLinea
            
            lnItem = lnItem + 1
            If lnItem = 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina(lsTitulo, lnPagina, lnItem, pgsNomAge, pgsEmpresa, pgdFecSis, "")
                lsCadena = lsCadena & Encabezado("Agencia;15; ;5;S/Recargo;20; ;8;C/Recargo;10; ;4;Inclusion;10; ;4;Total Pag;10; ;5;% Agencia;10; ;2;Com. Emision;15;Tot.Pag.IGV;12;", 0)
            End If
            rsR.MoveNext
        Wend
        
        If pnMontoComImpre <> 0 Then
            lsAgencia = "COMISION X IMPRESION"
            RSet lsSinRecargo = Format(0, "#,#00.0000")
            RSet lsConRecargo = Format(0, "#,##0.0000")
            RSet lsInclusion = Format(0, "#,##0.0000")
            RSet lsTotPagar = Format(0, "#,##0.0000")
            RSet lsPorAgencia = Format(0, "#,##0.00")
            RSet lsComPorEmision = Format(0, "#,##0.0000")
            RSet lsTotPagarConING = Format(pnMontoComImpre, "#,##0.00")
            
            lnTotPagarConING = lnTotPagarConING + pnMontoComImpre
            lsCadena = lsCadena & lsAgencia & lsSinRecargo & lsConRecargo & lsInclusion & lsTotPagar & lsPorAgencia & lsComPorEmision & lsTotPagarConING & oImpresora.gPrnSaltoLinea
        End If
        
        lsCadena = lsCadena & String(131, "=") & oImpresora.gPrnSaltoLinea
        
        lsAgencia = ""
        RSet lsSinRecargo = Format(lnSinRecargo, "#,#00.0000")
        RSet lsConRecargo = Format(0, "#,##0.0000")
        RSet lsInclusion = Format(0, "#,##0.0000")
        RSet lsTotPagar = Format(lnTotPagar, "#,##0.0000")
        RSet lsPorAgencia = Format(lnPorAgencia, "#,##0.00")
        RSet lsComPorEmision = Format(lnComPorEmision, "#,##0.0000")
        RSet lsTotPagarConING = Format(lnTotPagarConING, "#,##0.00")
        lsCadena = lsCadena & lsAgencia & lsSinRecargo & lsConRecargo & lsInclusion & lsTotPagar & lsPorAgencia & lsComPorEmision & lsTotPagarConING & oImpresora.gPrnSaltoLinea
        
        If psCodPersona <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & " POLIZA NRO = " & psPoliza & "     EMITIDA EL : " & psFecha & oImpresora.gPrnSaltoLinea
        If psCodPersona <> "" Then lsCadena = lsCadena & " TOTAL A PAGAR  = " & lsTotPagarConING
    End If
    
    GetRepAsistenciaMedica = lsCadena
End Function

Public Function GeneraAsientoAsistenciaMedica(ByVal sOpe As String, pbFamiliares As Boolean, psPersona As String, psPoliza As String, psFecha As String, pgsfecsis As Date, pgsUser As String, pgsCodAge As String, pnAjuste As Currency) As String
    Dim rsPla As ADODB.Recordset
    Dim sMov As String, sGlosa As String, sCta As String
    Set rsPla = New ADODB.Recordset
    Dim nDebe As Double, nHaber As Double, nOtro As Double
    Dim nDH As Double, nMonto As Double
    Dim nItem As Integer
    Dim sSQLC As String
    Dim lnDiff As Currency
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    Dim oMov As DMov
    Set oMov = New DMov
    Dim nMovNro As Long
    sSQLC = ""
    
    oCon.AbreConexion
    
    If rsPla.State = 1 Then rsPla.Close
    sql = "Select cMovNro from mov where cmovnro like '" & Format(pgsfecsis, gsFormatoMovFecha) & "%' And cOpeCod = '" & sOpe & "' And nMovFlag = 0"
    Set rsPla = oCon.CargaRecordSet(sql)
    
    If Not (rsPla.EOF And rsPla.BOF) Then
        If MsgBox("El asiento contable ya fue generado. Desea Re-Imprimirlo ? ", vbQuestion + vbYesNo) = vbNo Then
            GeneraAsientoAsistenciaMedica = ""
        Else
            GeneraAsientoAsistenciaMedica = rsPla.Fields(0)
        End If
        Exit Function
    End If
    
    If rsPla.State = 1 Then rsPla.Close
    sql = "Select cOpeDesc From " & gcCentralCom & "OpeTpo Where cOpeCod = '" & sOpe & "'"
    Set rsPla = oCon.CargaRecordSet(sql)
    Set rsPla.ActiveConnection = Nothing
    sGlosa = ""
    If Not (rsPla.EOF And rsPla.BOF) Then sGlosa = rsPla("cOpeDesc")
    rsPla.Close
    
    If Not pbFamiliares Then
        sql = "  Select  (sum(case debehab when 'D' then monto else 0 end) " _
            & " - sum(case debehab when 'H' then monto else 0 end))" _
            & " Diff" _
            & " From" _
            & "     (Select T.Cuenta, T.DebeHab, convert(money, rOUND(SUM(Monto),2)) Monto" _
            & "         From" _
            & "         (Select CC.cRHConceptoCod Concepto, CC.cOpeCtaDH DebeHab, Cuenta = Replace(CC.cCtaContCod,'AG',Right(PD.ID,2)),rOUND(PD.nMonto,6) Monto" _
            & "             FROM RHConceptoCta CC" _
            & "             Inner Join ( Select total nMOnto, cAgenciaAsig ID , '" & gRHConcDctoAMP & "' cConcepCod" _
            & "             From asistenciaEmp) PD ON CC.cRHConceptoCod = PD.cConcepCod Where cOpeCod = '" & sOpe & "')" _
            & " T Group by T.Cuenta, T.DebeHab) as aaa"
        
        Set rsPla = oCon.CargaRecordSet(sql)
        lnDiff = rsPla!Diff
        rsPla.Close
        
        sql = " Select T.Cuenta, T.DebeHab, ROUND(SUM(Monto),2) Monto" _
            & " FROM (Select CC.cRHConceptoCod Concepto, CC.cOpeCtaDH DebeHab, Cuenta = Replace(CC.cCtaContCod,'AG',Right(PD.ID,2)),rOUND(PD.nMonto,6) Monto" _
            & " FROM RHConceptoCta CC" _
            & " Inner Join ( Select total nMonto, cAgenciaAsig ID , '" & gRHConcDctoAMP & "' cConcepCod from asistenciaemp) PD" _
            & " ON CC.cRHConceptoCod = PD.cConcepCod" _
            & " Where cOpeCod = '" & sOpe & "' ) T Group by T.Cuenta, T.DebeHab" _
            & " Order by T.DebeHab,T.Cuenta"
    Else
        sql = "  Select  (sum(case debehab when 'D' then monto else 0 end) " _
            & " - sum(case debehab when 'H' then monto else 0 end))" _
            & " Diff" _
            & " From" _
            & "     (Select T.Cuenta, T.DebeHab, convert(money, rOUND(SUM(Monto),2)) Monto" _
            & "         From" _
            & "         (Select CC.cRHConceptoCod Concepto, CC.cOpeCtaDH DebeHab, Cuenta = Replace(CC.cCtaContCod,'AG',Right(PD.ID,2)),rOUND(PD.nMonto,6) Monto" _
            & "             FROM RHConceptoCta CC" _
            & "             Inner Join ( Select total nMOnto, cAgenciaAsig ID , '" & gRHConcDctoAMP & "' cConcepCod" _
            & "             From asistenciaFam) PD ON CC.cRHConceptoCod = PD.cConcepCod Where cOpeCod = '" & sOpe & "')" _
            & " T Group by T.Cuenta, T.DebeHab) as aaa"
        
        Set rsPla = oCon.CargaRecordSet(sql)
        lnDiff = rsPla!Diff
        rsPla.Close
        
        sql = " Select T.Cuenta, T.DebeHab, ROUND(SUM(Monto),2) Monto" _
            & " FROM (Select CC.cRHConceptoCod Concepto, CC.cOpeCtaDH DebeHab, Cuenta = Replace(CC.cCtaContCod,'AG',Right(PD.ID,2)),rOUND(PD.nMonto,6) Monto" _
            & " FROM RHConceptoCta CC" _
            & " Inner Join ( Select total nMonto, cAgenciaAsig ID , '" & gRHConcDctoAMP & "' cConcepCod from asistenciafam) PD" _
            & " ON CC.cRHConceptoCod = PD.cConcepCod" _
            & " Where cOpeCod = '" & sOpe & "' ) T Group by T.Cuenta, T.DebeHab" _
            & " Order by T.DebeHab,T.Cuenta"
    End If

    Set rsPla = oCon.CargaRecordSet(sql)
    
    Set rsPla.ActiveConnection = Nothing
    If Not (rsPla.EOF And rsPla.BOF) Then
        
        sMov = oMov.GeneraMovNro(pgsfecsis, pgsCodAge, pgsUser)
        sGlosa = sGlosa & " DE " & UCase(Format$(pgsfecsis, gsFormatoFechaView))
        
        oMov.BeginTrans
            oMov.InsertaMov sMov, sOpe, sGlosa
            nMovNro = oMov.GetnMovNro(sMov)
            
            oMov.InsertaMovDoc nMovNro, TpoDocFactura, psPoliza, psFecha
            
            nItem = 0
            Do While Not rsPla.EOF
                nItem = nItem + 1
                nDH = IIf(rsPla("DebeHab") = "D", 1, -1)
                sCta = rsPla("Cuenta")
                
                If Right(sCta, 2) = "01" Then
                    nMonto = rsPla("Monto") + pnAjuste
                ElseIf Right(sCta, 2) = "07" Then
                    nMonto = IIf(IsNull(rsPla("Monto")), 0, rsPla("Monto")) - lnDiff
                Else
                    nMonto = rsPla("Monto")
                End If
                
                If Left(sCta, 2) = "25" Then
                    nMonto = rsPla("Monto") + pnAjuste
                End If
                
                oMov.InsertaMovGasto nMovNro, psPersona, ""
                If nMonto <> 0 Then oMov.InsertaMovCta nMovNro, nItem, sCta, nMonto * nDH
                
                rsPla.MoveNext
            Loop
        oMov.CommitTrans
    End If
    rsPla.Close
    Set rsPla = Nothing
    GeneraAsientoAsistenciaMedica = sMov
End Function

Public Function GetDevAMP(ByVal psPersonas As String) As ADODB.Recordset
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim sql As String
    Dim nMovNro As Long
    
    oCon.AbreConexion
    
    sql = " Select Isnull(MontoPer,0) MontoPer , IsNull(MontoFam,0) MontoFam From" _
        & " (Select Round(count(*) * (dbo.GetAsistTit()) * 0.03 * dbo.GetIGV(1),2) As MontoPer From RRHH Where nRHEstado LIke '[78]%' And cPersCod In (" & psPersonas & ")) As AAA," _
        & " (Select (Sum(nRHAsistMedPrivMonto) - Round((count(*) * (dbo.GetAsistTit())),2)) * 0.03 * dbo.GetIGV(1) MontoFam" _
        & " From RRHH RH" _
        & " Inner JOin RHEmpleado RHE On RH.cPersCod = RHE.cPersCod" _
        & " Inner Join RHAsistMedPrivTabla TC On TC.cRHAsistMedPrivCod = RHE.cRHEmplAMPCod" _
        & " where RH.nRHEstado LIke '[78]%' And nRHAsistMedPrivMonto  <> 0 And RH.cPersCod In (" & psPersonas & ")) AS BBB"
    Set GetDevAMP = oCon.CargaRecordSet(sql)
End Function





Public Function GetRHCatalogoPlanEPS(pscCodCatEPS As String) As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetRHCatalogoPlanEPSErr

    If oCon.AbreConexion() Then
        sqlA = "select nCodCatalogo as [Cod Plan],cCodCatEPS as [Cod Cat.],cDesCatalogo as [Descripcion Plan] , " _
             & " nCantPer as [Cant Per] , nCostoTotal as Total, nCostoIGV as IGV , nCostoNeto AS Neto,cCodValQuin as Desc_Quin,cUltimaActualizacion " _
             & " from RHCatalogoPlanEPS where cCodCatEPS  ='" & pscCodCatEPS & "'"
        Set GetRHCatalogoPlanEPS = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetRHCatalogoPlanEPSErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHCatalogoPlanEPS Method")
End Function


Public Function AgregaAsisMedPriv(psRHAsistMedPrivCod As String, psRHAsistMedPrivDescripcion As String, pnRHAsistMedPrivMonto As Currency, pnRHAsistMedPrivMontoDesc As Currency, psUltimaActualizacion As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo AgregaAsisMedPrivErr
        
    If oCon.AbreConexion() Then
        sqlA = " Insert " & lsServerComunes & "RHAsistMedPrivTabla (cRHAsistMedPrivCod,cRHAsistMedPrivDescripcion,nRHAsistMedPrivMonto,nRHAsistMedPrivMonto,cUltimaActualizacion)" _
             & " Values('" & psRHAsistMedPrivCod & "','" & psRHAsistMedPrivDescripcion & "'," & pnRHAsistMedPrivMonto & "," & pnRHAsistMedPrivMontoDesc & ",'" & psUltimaActualizacion & "')"
        oCon.Ejecutar sqlA
        AgregaAsisMedPriv = True
        oCon.CierraConexion
    Else
        AgregaAsisMedPriv = False
    End If
    Set oCon = Nothing
    Exit Function
AgregaAsisMedPrivErr:
    AgregaAsisMedPriv = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:AgregaAsisMedPriv Method")
End Function


Public Function GetRHGeneraCodPlanEPS(pscCodCatEPS As String) As String
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo GetRHGeneraCodPlanEPSErr
    If oCon.AbreConexion() Then
        sqlA = "Select right('0000' + convert (varchar(3),max(substring (nCodCatalogo,3,3))+ 1),3) as Num " _
               & " from RHCatalogoPlanEPS where cCodCatEPS  ='" & pscCodCatEPS & "'"
        Set rs = oCon.CargaRecordSet(sqlA)
        
        If rs.EOF Then
            GetRHGeneraCodPlanEPS = pscCodCatEPS + "001"
            Else
            GetRHGeneraCodPlanEPS = rs!Num
        End If
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    Set rs = Nothing
    
    Exit Function
GetRHGeneraCodPlanEPSErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHGeneraCodPlanEPS Method")
End Function

Public Function GetRHCategorias() As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetRHCategoriasErr

    If oCon.AbreConexion() Then
        sqlA = "Select cCodCatEPS,cDescripcion,2 from RHCatPlanEPS where nEstado = 0 " _
               & "  Union " _
               & "  Select '00','Categorias Plan EPS ',1 "

        
        Set GetRHCategorias = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetRHCategoriasErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHCategorias Method")
End Function

Public Function EliminaRHPersonaPlan(psCodPersona As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo AgregaAsisMedPrivErr
    If oCon.AbreConexion() Then
        'sqlA = " Insert " & lsServerComunes & "RHAsistMedPrivTabla (cRHAsistMedPrivCod,cRHAsistMedPrivDescripcion,nRHAsistMedPrivMonto,nRHAsistMedPrivMonto,cUltimaActualizacion)" _
        '     & " Values('" & psRHAsistMedPrivCod & "','" & psRHAsistMedPrivDescripcion & "'," & pnRHAsistMedPrivMonto & "," & pnRHAsistMedPrivMontoDesc & ",'" & psUltimaActualizacion & "')"
        '(psnCodCatalogo As String, pscCodCatEPS As String, psRHAsistMedPrivDescripcion As String,
        'pnnCantPer As Integer, nCostoTotal As Currency, pnCostoIGV As Currency, pnCostoNeto As Currency, psUltimaActualizacion As String)
        
        sqlA = " Delete from  RHPersonaPlan  where   cPersCod ='" & psCodPersona & "' "
        oCon.Ejecutar sqlA
        
        oCon.CierraConexion
    Else
        EliminaRHPersonaPlan = False
    End If
    Set oCon = Nothing
    Exit Function
AgregaAsisMedPrivErr:
    EliminaRHPersonaPlan = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:EliminaRHPersonaPlan Method")
End Function



Public Function AgregaRHPersonaPlan(pRSPlanEPS As ADODB.Recordset, psCodPersona As String, psUltimaActualizacion As String) As Boolean
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo AgregaAsisMedPrivErr
    If oCon.AbreConexion() Then
        'sqlA = " Insert " & lsServerComunes & "RHAsistMedPrivTabla (cRHAsistMedPrivCod,cRHAsistMedPrivDescripcion,nRHAsistMedPrivMonto,nRHAsistMedPrivMonto,cUltimaActualizacion)" _
        '     & " Values('" & psRHAsistMedPrivCod & "','" & psRHAsistMedPrivDescripcion & "'," & pnRHAsistMedPrivMonto & "," & pnRHAsistMedPrivMontoDesc & ",'" & psUltimaActualizacion & "')"
        '(psnCodCatalogo As String, pscCodCatEPS As String, psRHAsistMedPrivDescripcion As String,
        'pnnCantPer As Integer, nCostoTotal As Currency, pnCostoIGV As Currency, pnCostoNeto As Currency, psUltimaActualizacion As String)
        
        sqlA = " Delete from  RHPersonaPlan  where   cPersCod ='" & psCodPersona & "' "
        oCon.Ejecutar sqlA
        
        sqlA = ""
        While Not pRSPlanEPS.EOF
                
                If pRSPlanEPS.Fields(0) <> "" Then
                    sqlA = " insert into RHPersonaPlan  " _
                   & " (cPersCod,nCodCatalogo,cUltimaActualizacion) " _
                   & "  values('" & psCodPersona & "','" & pRSPlanEPS.Fields(0) & "','" & psUltimaActualizacion & "') "
                   oCon.Ejecutar sqlA
                   AgregaRHPersonaPlan = True
                   Else
                End If
               
                
                
                pRSPlanEPS.MoveNext
        Wend
        
        oCon.CierraConexion
    Else
        AgregaRHPersonaPlan = False
    End If
    Set oCon = Nothing
    Exit Function
AgregaAsisMedPrivErr:
    AgregaRHPersonaPlan = False
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:AgregaAsisMedPriv Method")
End Function

Public Function GetRHCatPlanEPS() As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetRHCatPlanEPSErr

    If oCon.AbreConexion() Then
        'sqlA = "Select cCodCatEPS,cDescripcion,2 from RHCatPlanEPS where nEstado = 0 " _
        '       & "  Union " _
        '       & "  Select '00','Categorias Plan EPS ',1 "
        
    sqlA = " Select cCodCatEPS,cDescripcion,2 from RHCatPlanEPS where nEstado = 0 " _
         & " Union" _
         & " Select '00','Categorias Plan EPS ',1 " _
         & "Union " _
         & " Select nCodCatalogo,cDesCatalogo,3 from RHCatalogoPlanEPS "
        
        

        
        Set GetRHCatPlanEPS = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetRHCatPlanEPSErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHCatPlanEPS Method")
End Function

Public Function GetRHPersonaPlan(psCodPersona As String) As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetRHPersonaPlanErr

    If oCon.AbreConexion() Then
        
            sqlA = " select  RP.nCodCatalogo,cDesCatalogo from RHPersonaPlan RP,RHCatalogoPlanEPS  RE " _
            & " where    cPersCod = '" & psCodPersona & "' and " _
            & " RP.nCodCatalogo = RE.nCodCatalogo "
        
        
        
        Set GetRHPersonaPlan = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetRHPersonaPlanErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHPersonaPlan Method")
End Function

Public Function GetRHListadoEPS(pnConsulta As Integer, psPeriodo As String) As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetRHListadoEPSErr

    If oCon.AbreConexion() Then
        sqlA = " SP_RHLISTADOEPS " & pnConsulta & " ,'" & psPeriodo & "'"
        Set GetRHListadoEPS = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetRHListadoEPSErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHListadoEPS Method")
End Function

Public Function GetRHExistCodEPS(psFecha As String) As Integer
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    On Error GoTo GetRHExistCodEPSErr

    If oCon.AbreConexion() Then
        sqlA = " SELECT  count(*) as Cant FROM  RHCALCULOEPS where substring(cCodCalculoEps,1,8)='20050429'"
        Set rs = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
        GetRHExistCodEPS = rs!Cant
    End If
    
    Set oCon = Nothing
    Set rs = Nothing
    
    Exit Function
GetRHExistCodEPSErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHExistCodEPS Method")
End Function



Public Function GetRHCodcalculoEPS() As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetRHCodcalculoEPSErr

    If oCon.AbreConexion() Then
        sqlA = " SELECT  cCodCalculoEps,cDescripcion,cUltimaActualizacion FROM  RHCALCULOEPS  WHERE NESTADOEPS=0 "
        Set GetRHCodcalculoEPS = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetRHCodcalculoEPSErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHCodcalculoEPS Method")
End Function


Public Function GetRHDetallecalculoEPS(psCodCalculoEps As String) As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetRHDetallecalculoEPSErr
    
    If oCon.AbreConexion() Then
        sqlA = " SELECT  P.cPersCod , cPersNombre as Nombres, " _
        & " nSueldo as Sueldo,nSueldo_x_225 as [2.25 EPS],nCantPersonas AS CantPers,nPlanSinIGV P_S_IGV,nPromedio AS Promedio,nNeto Neto, " _
        & " nPagaEmpleado [Pag.Emple],nPagaEmpresa [Pag.Empre],nAdicionalHijos as [A.Hijos],nAdicionalPadres as [A.Padres],nTotalEmpleado as [T.Emple]," _
        & " nDesQuincena as [DQuin] , nDesMensual as [DMens], nDesQuincenaUno as [Des 1-15], nDescQuincenaDos as [Des 15-30], nSaldo  as Saldo" _
        & " FROM  RHCALCULOEPSDET A, PERSONA P WHERE cCodCalculoEps ='" & psCodCalculoEps & "' " _
        & " AND  A.CPERSCOD= P.CPERSCOD " _
        & " ORDER BY cPersNombre ASC "
        Set GetRHDetallecalculoEPS = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetRHDetallecalculoEPSErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHDetallecalculoEPS Method")
End Function


Public Function GetRHCalculoEPS(pnPromedio As Currency, psPeriodo As String, pnQuincena As Currency, pnEmpresa As Currency) As Recordset
    Dim sqlA As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    On Error GoTo GetRHCalculoEPSErr

    If oCon.AbreConexion() Then
        sqlA = " SP_RHCALCULAEPS " & pnPromedio & " ,'" & psPeriodo & "'," & pnQuincena & ", " & pnEmpresa & ""
        Set GetRHCalculoEPS = oCon.CargaRecordSet(sqlA)
        oCon.CierraConexion
    End If
    Set oCon = Nothing
    
    Exit Function
GetRHCalculoEPSErr:
    Call RaiseError(MyUnhandledError, "DActualizaAsistenciaMedicaPrivada:GetRHCalculoEPS Method")
End Function


