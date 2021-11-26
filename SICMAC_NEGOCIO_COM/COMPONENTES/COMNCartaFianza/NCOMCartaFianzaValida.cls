VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCartaFianzaValida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************
'* Clase que contiene las validaciones de las operaciones de Credito Pignoraticio
'* LAYG - 01/07/2001
'****************************
Option Explicit

Function nCFGarantiasGravada(ByVal psCtaCod As String, ByVal pnTipoCambioFijo As Double, ByRef psmensaje As String) As Double
' Obtiene el Monto de las Garantias Gravadas
Dim loConecta As COMConecta.DCOMConecta
Dim lsSql As String
Dim lrValida As New ADODB.Recordset
Dim oDGar As New COMDCredito.DCOMGarantia 'EJVG20150713
    'EJVG20150713 ***
    'lsSql = "SELECT isNull(SUM (Case When nMoneda = 1 then CG.nGravado " _
         & "                         When nMoneda = 2 then CG.nGravado * " & pnTipoCambioFijo & " end ),0) ValGrav " _
         & " FROM ColocGarantia CG " _
         & " Where CG.cCtaCod = '" & psCtaCod & "' "
    psmensaje = oDGar.CadenaGarantiasPendienteMigracion(psCtaCod)
    Set oDGar = Nothing
    If Len(psmensaje) > 0 Then Exit Function
    
    lsSql = "EXEC stp_sel_ERS0632014_RecuperaCoberturaCF '" & psCtaCod & "'," & pnTipoCambioFijo
    'END EJVG *******
    Set loConecta = New COMConecta.DCOMConecta
        loConecta.AbreConexion
        Set lrValida = loConecta.CargaRecordSet(lsSql)
        loConecta.CierraConexion
    Set loConecta = Nothing

    If (lrValida.BOF Or lrValida.EOF) Or (lrValida!ValGrav = 0) Then
         psmensaje = "El Carta Fianza No Tiene Garantias Gravadas"
         nCFGarantiasGravada = 0
    Else
        If Mid(psCtaCod, 9, 1) = "1" Then
            nCFGarantiasGravada = lrValida!ValGrav
        ElseIf Mid(psCtaCod, 9, 1) = "2" Then
            nCFGarantiasGravada = lrValida!ValGrav / pnTipoCambioFijo
        End If
    End If
    Set lrValida = Nothing
End Function


Public Function nCFPermisoAprobacion(ByVal psPersCod As String, ByVal psProducto As String) As Boolean
Dim lsSql As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As New ADODB.Recordset
Dim lbOk As Boolean

lsSql = "SELECT cCodNiv  FROM ColocCredPersNivelesApr  " _
      & "WHERE cCodNiv in ( SELECT cCodNiv FROM ColocCredNivelesApr " _
      & "                       WHERE  cProduct = '" & psProducto & "' )" _
      & " AND cPersCod ='" & psPersCod & "' "

    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSql)
    If Not (lrValida.BOF And lrValida.EOF) Then ' Encuentra el permiso
        lbOk = True
    Else
        lbOk = False
    End If
    Set lrValida = Nothing
    nCFPermisoAprobacion = lbOk
End Function
'MADM 20120117 - pInd
Public Function RecuperaPersonasEstadoCF(ByVal pnEstadoCF As Variant, Optional ByVal pMatprod As Variant = Nothing, Optional gdfecSis As Date, Optional pInd As Integer = 0, Optional ByVal psCtaCod As String = "", Optional ByVal pbComision As Boolean = False, Optional ByVal pdFechaAut As Date = "01/01/1900") As ADODB.Recordset 'WIOR 20130313 AGREGO pbComision y pdFechaAut
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta
Dim vsEstadoCred As String
Dim vsProductos As String
Dim i As Integer

On Error GoTo ErrorRecuperaPersonasEstadoCF
    vsEstadoCred = ""
    vsProductos = ""
    For i = 0 To UBound(pnEstadoCF)
        vsEstadoCred = vsEstadoCred & pnEstadoCF(i) & ","
    Next i
    vsEstadoCred = Mid(vsEstadoCred, 1, Len(vsEstadoCred) - 1)
    
 If pInd = 1 Then
    If IsArray(pMatprod) Then
        For i = 0 To UBound(pMatprod)
            vsProductos = vsProductos & pMatprod(i) & ","
        Next i
        vsProductos = Mid(vsProductos, 1, Len(vsProductos) - 1)
    End If
    
    lsSql = "exec stp_sel_pago_CartaFianza '" & vsEstadoCred & "','" & vsProductos & "','" & Format(gdfecSis, "yyyymmdd") & "'"

'WIOR 20120806**********
ElseIf pInd = 2 Then
    lsSql = "exec stp_sel_Mod_CartaFianza '" & vsEstadoCred & "','" & Format(gdfecSis, "yyyymmdd") & "','" & psCtaCod & "'"
ElseIf pInd = 3 Then
    lsSql = "exec stp_sel_ModPagada_CartaFianza '" & vsEstadoCred & "','" & Format(gdfecSis, "yyyymmdd") & "'"
'WIOR FIN **************
Else
    vsProductos = ""
    If IsArray(pMatprod) Then
    vsProductos = "('"
        For i = 0 To UBound(pMatprod)
            vsProductos = vsProductos & pMatprod(i) & "','"
        Next i
    vsProductos = Mid(vsProductos, 1, Len(vsProductos) - 2)
    End If
    vsProductos = vsProductos & ")"
    
    lsSql = "Select P.cPersCod, P.cPersNombre,C.cCtaCod  from Persona P inner join ProductoPersona PP ON P.cPersCod = PP.cPersCod"
    lsSql = lsSql & " Inner join ColocCartaFianza C ON PP.cCtaCod = C.cCtaCod "
    lsSql = lsSql & " Inner join Producto Prd ON PP.cCtaCod = Prd.cCtaCod "
    lsSql = lsSql & " WHERE PP.nPrdPersRelac = " & gColRelPersTitular & " AND Prd.nPrdEstado in (" & Trim(vsEstadoCred) & ") "
    If Len(vsProductos) > 1 Then
        lsSql = lsSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN " & vsProductos
    End If
    lsSql = lsSql & " Union"
    lsSql = lsSql & " Select P.cPersCod, P.cPersNombre,C.cCtaCod"
    lsSql = lsSql & " from Persona P"
    lsSql = lsSql & " inner join ProductoPersona PP ON P.cPersCod = PP.cPersCod"
    lsSql = lsSql & " Inner join ColocCartaFianza C ON PP.cCtaCod = C.cCtaCod"
    lsSql = lsSql & " Inner join Producto Prd ON PP.cCtaCod = Prd.cCtaCod"
    lsSql = lsSql & " WHERE PP.nPrdPersRelac = 20 AND Prd.nPrdEstado in (2020)"
    lsSql = lsSql & " AND nRenovacion <>0 and datediff (day,dVencimiento,'" & Format(gdfecSis, " mm/dd/yyyy") & "') = 0"
    If Len(vsProductos) > 1 Then
        lsSql = lsSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN " & vsProductos
    End If
    'WIOR 20130313 *************************************************************
    If pbComision Then
        lsSql = lsSql & " Union"
        lsSql = lsSql & " Select P.cPersCod, P.cPersNombre,C.cCtaCod  from Persona P"
        lsSql = lsSql & " Inner Join ProductoPersona PP ON P.cPersCod = PP.cPersCod"
        lsSql = lsSql & " Inner Join ColocCartaFianza C ON PP.cCtaCod = C.cCtaCod"
        lsSql = lsSql & " Inner Join Producto Prd ON PP.cCtaCod = Prd.cCtaCod"
        lsSql = lsSql & " Inner Join ColocCFAutRenovacion CFA ON CFA.cCtaCod=Prd.cCtaCod and CFA.nEstado=0 AND CFA.dAutorizado=1"
        lsSql = lsSql & " AND CFA.nMovNro=0 and CFA.nRenovacion=ISNULL(C.nRenovacion,0)"
        lsSql = lsSql & " Where PP.nPrdPersRelac = 20"
        lsSql = lsSql & " AND Prd.nPrdEstado in (2020)"
        lsSql = lsSql & " AND LEFT(CFA.cMovAut,8)='" & Format(pdFechaAut, "yyyymmdd") & "'"
        lsSql = lsSql & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN " & vsProductos
    End If
    'WIOR 20130313 *************************************************************
    lsSql = lsSql & " order by P.cPersNombre"

End If
    Set loConecta = New COMConecta.DCOMConecta
        loConecta.AbreConexion
        Set RecuperaPersonasEstadoCF = loConecta.CargaRecordSet(lsSql)
        loConecta.CierraConexion
    Set loConecta = Nothing
    
    Exit Function
    
ErrorRecuperaPersonasEstadoCF:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function

'By capi 24022009 se adiciono parametro fecha de proceso
'Public Function nCFPagoComision(ByVal psCta As String) As Double
Public Function nCFPagoComision(ByVal psCta As String, Optional ByVal pdFecSis As Date, Optional ByVal pnTipo As Integer = 0) As Double
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta
Dim lrValida As New ADODB.Recordset
Dim lnComisPag As Double

'By capi 24022009 para que valide si el pago de comision corresponde a la actual transaccion
'
'lsSQL = "SELECT IsNull(Sum(nMonto),0) ComisPag FROM MovCol MC join Mov M on MC.nMovNro = M.nMovNro " _
'      & "WHERE cCtaCod ='" & psCta & "' And MC.cOpeCod ='" & gColCFOpeComisEfe & "' " _
'      & " AND M.nMovFlag in (" & gMovFlagVigente & " ) "
'
If pnTipo = 0 Then
lsSql = "SELECT IsNull(Sum(nMonto),0) ComisPag FROM MovCol MC join Mov M on MC.nMovNro = M.nMovNro " _
      & "WHERE cCtaCod ='" & psCta & "' And MC.cOpeCod ='" & gComiCredEmisionRenovCF & "' " _
      & " AND M.nMovFlag in (" & gMovFlagVigente & " ) " _
      & " AND datediff(day," & Format(pdFecSis, "DD/MM/YYYY") & ",cast(left(M.cmovnro,8) as datetime))>=0 "
ElseIf pnTipo = 1 Then

lsSql = "SELECT IsNull(Sum(nMonto),0) ComisPag FROM MovCol MC join Mov M on MC.nMovNro = M.nMovNro " _
      & " WHERE cCtaCod ='" & psCta & "' And MC.cOpeCod ='" & gComiCredModifCF & "' " _
      & " AND M.nMovFlag in (" & gMovFlagVigente & " ) " _
      & " AND datediff(day,'" & Format(pdFecSis, "yyyymmdd") & "',cast(left(M.cmovnro,8) as datetime))>=0 "

End If
    Set loConecta = New COMConecta.DCOMConecta
        loConecta.AbreConexion
        Set lrValida = loConecta.CargaRecordSet(lsSql)
        loConecta.CierraConexion
    Set loConecta = Nothing
   
    lnComisPag = lrValida!ComisPag
    
    Set lrValida = Nothing
    nCFPagoComision = lnComisPag
    
End Function

Public Function nBuscaOperacionesCFParaExtorno(ByVal psFecTrans As String, ByVal psOpeExtor As String, _
        Optional ByVal psCtaCod As String = "@", Optional ByVal pbOrdenAsc As Boolean = False, Optional ByRef psmensaje As String) As Recordset

Dim lsSql As String
Dim loRegValida As COMDColocPig.DCOMColPFunciones
Dim lrValida As New ADODB.Recordset
    
    'JUEZ 20160304 ************************************************
    'lsSql = " SELECT M.cMovNro, M.nMovNro, M.cOpeCod, M.nMovFlag, " _
    '      & " MC.cCtaCod, MC.nMonto, Ope.cOpeDesc  " _
    '      & " FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
    '      & " LEFT JOIN OpeTpo Ope ON MC.cOpeCod = Ope.cOpeCod " _
    '      & " WHERE M.cMovNro LIKE '" & psFecTrans & "%' " _
    '      & " AND M.nMovFlag = " & gMovFlagVigente & " " _
    '      & " AND M.cOpeCod in " & psOpeExtor & "  "
    'If psCtaCod <> "@" Then
    '    lsSql = lsSql & " AND MC.cCtaCod ='" & psCtaCod & "' "
    'End If
    'lsSql = lsSql & " ORDER BY M.nMovNro DESC  "
    
    lsSql = "exec stp_sel_BuscaOperacionesCFParaExtorno '" & psFecTrans & "','','" & psCtaCod & "'"
    'END JUEZ *****************************************************
    
    Set loRegValida = New COMDColocPig.DCOMColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSql)
    If lrValida Is Nothing Then
        psmensaje = "ERROR: al Buscar datos para Extorno "
        Exit Function
    End If
    
    Set nBuscaOperacionesCFParaExtorno = lrValida
    Set lrValida = Nothing
End Function


'******************
'CAAU - 29/10/2002
'******************
Function RecuperaRelacinesInternas(ByRef Codpersona As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta

lsSql = "Select  P1.cPersNombre as Cliente,"
lsSql = lsSql & " (Select C.cConsDescripcion as Describe from Constante C where PP.nPrdPersRelac=C.nConsValor and nConsCod=3002) as Relacion,"
lsSql = lsSql & " CF.cCtaCod as Carta,"
lsSql = lsSql & " (Select C.cConsDescripcion from Constante C where  Pro.nPrdEstado=nConsValor and nConsCod=3001) as Estado,"
lsSql = lsSql & " (Select nMonto from ColocCFEstado where cCtaCod=CF.cCtaCod and nPrdEstado= Pro.nPrdEstado ) Saldo, "
'lsSQL = lsSQL & " Pro.nSaldo as saldo,"
lsSql = lsSql & " (Select C.cConsDescripcion from Constante C where  substring(CF.cCtacod,9,1)=nConsValor and nConsCod=1011) as Moneda"
lsSql = lsSql & " From ColocCartaFianza CF"
lsSql = lsSql & " Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod"
lsSql = lsSql & " Inner Join ProductoPersona PP on PP.cCtaCod = CF.cCtaCod"
lsSql = lsSql & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod"
lsSql = lsSql & " Where PP.nPrdPersRelac = 20 and Pro.nPrdEstado = 2001 And P1.cPersCod =" & Codpersona

On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RecuperaRelacinesInternas = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    'End If
    Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Function RecuperaFuenteIngreso(ByRef Codpersona As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta
lsSql = " Select  PFI.cPersCod as CodigoPers, PFI.cUltimaActualizacion as UltimaAct, PFI.cRazSocDirecc as Direccion, PFI.cRazSocDescrip as Razon, "
lsSql = lsSql & " PFI.dPersFIinicio as FechaInicio,"
lsSql = lsSql & " (Select C.cConsDescripcion from Constante C Where PFI.nPersTipFte = nConsValor and nConsCod=1028) as Tipo_Ingreso,"
lsSql = lsSql & " (Select cCIIUdescripcion from CIIU  Where P.cPersCIIU = cCIIUcod) as Actividad"
lsSql = lsSql & " from PersFteIngreso PFI inner join  Persona P on P.cPersCod  = PFI.cPersCod"
lsSql = lsSql & " where PFI.cUltimaActualizacion=(Select max(cUltimaActualizacion) from PersFteIngreso where cPersCod=PFI.cPErsCod)"
lsSql = lsSql & " and PFI.cPersCod=" & Codpersona
On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RecuperaFuenteIngreso = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Function RecuperaDatosAcreedor(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta

 lsSql = " Select CF.cCtaCod as Carta, P1.cPersCod as Codigo, P1.cPersNombre as Nombre, P1.cPersDireccDomicilio as Direccion, "
 lsSql = lsSql & "P1.nPersPersoneria as TipoPers, "
 lsSql = lsSql & "(Select C.cConsDescripcion from Constante C where  substring(CF.cCtacod,6,3)=nConsValor and nConsCod=1001) as Tipo, "
 lsSql = lsSql & "(Select C.cConsDescripcion from Constante C where  Pro.nPrdEstado=nConsValor and nConsCod=3001) as Estado, "
 lsSql = lsSql & "DocDNI = case when P1.nPersPersoneria = 1 then (Select cPersIDnro from PersID where P1.cPersCod=cPersCod and cPersIdTpo=1) end, "
 lsSql = lsSql & "DocRUC = case when P1.nPersPersoneria <> 1 then (Select cPersIDnro from PersID where P1.cPersCod=cPersCod and cPersIdTpo=2) end "
 lsSql = lsSql & "From ColocCartaFianza CF "
 lsSql = lsSql & "Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod "
 lsSql = lsSql & "Inner Join ProductoPersona  PP on PP.cCtaCod = CF.cCtaCod "
 lsSql = lsSql & "Inner Join Persona P1 on P1.cPersCod = PP.cPersCod "
 lsSql = lsSql & "Where PP.nPrdPersRelac = 35 and CF.cCtaCod=" & psCtaCod
On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RecuperaDatosAcreedor = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Function RecuperaDatosJuridica(ByRef Codpersona As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta
lsSql = "SELECT PJ.cPersCod, "
lsSql = lsSql & " (SELECT cPersJurTpoDesc From PersJurTpo Where PJ.cPersJurTpo=cPersJurTpoCod ) as DescribeJur,"
lsSql = lsSql & " (SELECT cConsDescripcion From Constante Where PJ.cPersJurMagnitud=nConsValor and nConsCod=1004) as Magnitud"
lsSql = lsSql & " FROM PERSONAJUR PJ WHERE PJ.cPersCod='" & Codpersona & "'"
On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RecuperaDatosJuridica = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Function RecuperaDatosT(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta
lsSql = "Select CF.cCtaCod as Carta, P1.cPersCod as Codigo, P1.cPersNombre as Nombre, P1.cPersDireccDomicilio as Direccion, "
lsSql = lsSql & " P1.nPersPersoneria as TipoPers,P1.dPersNacCreac as Nacimiento, "
lsSql = lsSql & " (Select C.cConsDescripcion from Constante C where  substring(CF.cCtacod,6,3)=nConsValor and nConsCod=1001) as Tipo, "
lsSql = lsSql & " (Select C.cConsDescripcion from Constante C where  Pro.nPrdEstado=nConsValor and nConsCod=3001) as Estado, "
lsSql = lsSql & " DocDNI = case when P1.nPersPersoneria = 1 then (Select cPersIDnro from PersID where P1.cPersCod=cPersCod and cPersIdTpo=1) end, "
lsSql = lsSql & " DocRUC = case when P1.nPersPersoneria <> 1 then (Select cPersIDnro from PersID where P1.cPersCod=cPersCod and cPersIdTpo=2) end, "
lsSql = lsSql & " nSaldo "
lsSql = lsSql & " From ColocCartaFianza CF "
lsSql = lsSql & " Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod "
lsSql = lsSql & " Inner Join ProductoPersona  PP on PP.cCtaCod = CF.cCtaCod "
lsSql = lsSql & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod "
lsSql = lsSql & " Where PP.nPrdPersRelac = 20 and CF.cCtaCod=" & psCtaCod
On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RecuperaDatosT = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function

ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function

Function RepcuperaOtrosDt(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta

lsSql = lsSql & "Select CFE.dPrdEstado as Fecha,"
lsSql = lsSql & "( Select C.cConsDescripcion as Describe from Constante C where CFE.nPrdEstado=C.nConsValor and nConsCod=3001) Condicion, "
lsSql = lsSql & "( Select C.cConsDescripcion as Describe from Constante C where CFE.nMotivoRechazo=C.nConsValor and nConsCod=3024) Motivo "
lsSql = lsSql & "from ColocCFEstado CFE where CFE.nPrdEstado=2003 "
lsSql = lsSql & "and CFE.cCtaCod=" & psCtaCod

On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RepcuperaOtrosDt = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing

    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function

ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Function RecuperaPersonaRelacion(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta
'WIOR 20120330-INICIO
lsSql = " declare @cPersCodTitular as varchar(13) "
lsSql = lsSql & " set @cPersCodTitular  =(select cPersCod from ProductoPersona where cCtaCod ='" & psCtaCod & "' and nPrdPersRelac=20) "
'WIOR - FIN

lsSql = lsSql & "Select P.cPersCod as CodPersona,"
'lsSQL = lsSQL & " P.cPersNombre as Nombre "
'WIOR 20120330-INICIO
lsSql = lsSql & " case when P.cPersCod=@cPersCodTitular and nPrdPersRelac=38 and CCF.cConsorcio IS NOT null then 'CONSORCIO: '+isnull(CCF.cConsorcio,'') Else isnull(P.cPersNombre,'') end Nombre "
'WIOR - FIN
'lsSQL = lsSQL & " , cCtaCod, "
lsSql = lsSql & " , PP.cCtaCod, " 'WIOR 20120330
lsSql = lsSql & " ( Select C.cConsDescripcion as Describe from Constante C where PP.nPrdPersRelac=C.nConsValor and nConsCod=3002) as Relacion"
lsSql = lsSql & " from ProductoPersona PP"
lsSql = lsSql & " Inner Join Persona P on P.cPersCod=PP.cPersCod "
lsSql = lsSql & " LEFT JOIN ColocCartaFianza CCF ON PP.cCtaCod =CCF.cCtaCod  " 'WIOR 20120330
lsSql = lsSql & " where PP.nPrdPersRelac!=28 and "
'lsSQL = lsSQL & " cCtaCod='" & psCtaCod & "'"
lsSql = lsSql & " PP.cCtaCod='" & psCtaCod & "'" 'WIOR 20120330

On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RecuperaPersonaRelacion = loConecta.CargaRecordSet(lsSql)
    loConecta.CierraConexion
    Set loConecta = Nothing

    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '   MsgBox "No existen Personas Relacionadas", vbInformation, "AVISO"
    'End If
    
    Exit Function

ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Function RecuperaHistorial(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta


lsSql = "Select  CFE.nMonto as Monto, CFE.cCtaCod as CF,"
lsSql = lsSql & "(Select C.cConsDescripcion from Constante C where  CFE.nPrdEstado=nConsValor and nConsCod=3001) as Estado"
lsSql = lsSql & ", CFE.dPrdEstado as Fecha from ColocCFEstado CFE where cCtaCod=" & psCtaCod
On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RecuperaHistorial = loConecta.CargaRecordSet(lsSql)

    loConecta.CierraConexion
    Set loConecta = Nothing

    'If RecuperaHistorial.BOF And RecuperaHistorial.EOF Then
    '    MsgBox "No existen Historial", vbInformation, "CMACT"
    'End If
    
    Exit Function

ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Function RecuperaGarantias(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta

'lsSql = "Select CG.cNumGarant as NroGarantia, CF.cCtaCod as CF, "
'lsSql = lsSql & "(Select C.cConsDescripcion from Constante C where  G.nTpoGarantia = nConsValor and nConsCod=1027 )as Tipo_Garn, "
'lsSql = lsSql & "(Select cDocDesc from Documento where G.cTpoDoc=nDocTpo) as Tipo_Doc, "
'lsSql = lsSql & "CG.nMoneda as Moneda, CG.nGravado as Gravado, G.cDescripcion as Descripcion "
'lsSql = lsSql & "from ColocGarantia CG "
'lsSql = lsSql & "Inner Join ColocCartaFianza CF on CF.cCtacod = CG.cCtaCod "
'lsSql = lsSql & "Inner Join Garantias G on CG.cNumGarant = G.cNumgarant "
'lsSql = lsSql & "where CF.cCtaCod=" & psCtaCod

lsSql = "EXEC stp_sel_ERS0632014_RecuperaGarantiasxCartaFianza '" & psCtaCod & "'" 'EJVG20151013

On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RecuperaGarantias = loConecta.CargaRecordSet(lsSql)

    loConecta.CierraConexion
    Set loConecta = Nothing

    'If RecuperaGarantias.BOF And RecuperaGarantias.EOF Then
    '    MsgBox "No existen", vbInformation, "CMACT"
    'End If
    
    Exit Function

ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Function RecuperaDatosGeneralesCF(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSql As String
Dim loConecta As COMConecta.DCOMConecta
'By Capi Nov2007
Dim lnEstado As Integer
Dim lrs As ADODB.Recordset

lsSql = "Select nPrdEstado  From Producto Where cCtaCod='" & psCtaCod & "'"
    
    Set loConecta = New COMConecta.DCOMConecta
        loConecta.AbreConexion
    Set lrs = loConecta.CargaRecordSet(lsSql)
        lnEstado = lrs!nPrdEstado
        loConecta.CierraConexion
    Set loConecta = Nothing
    Set lrs = Nothing
'End By

'MAVM 201006 BAS II se cambio TIPO
lsSql = "Select CF.cCtaCod as Carta, CF.dVencimiento as Vence, CF.dAsignacion as F_Asignacion,ce.nMonto,CE.dPrdEstado,CE.dVenc, "
lsSql = lsSql & " (Select C.cConsDescripcion from Constante C where  Pro.nPrdEstado=nConsValor and nConsCod=3001) as Estado, "
lsSql = lsSql & " ISNULL((Select CO.cConsDescripcion from Colocaciones C inner join Constante CO on C.cTpoCredCod = CO.nConsValor where C.cCtaCod = CF.cCtacod and nConsCod = 3034), '') as Tipo, "
lsSql = lsSql & " (Select C.cConsDescripcion from Constante C where  substring(CF.cCtacod,9,1)=nConsValor and nConsCod=1011) as Moneda, "
lsSql = lsSql & " P1.cPersNombre as Cliente, P2.cPersNombre as Analista, CF.cFinalidad as Finalidad, "
lsSql = lsSql & " (Select C.cConsDescripcion from Constante C Where CF.nModalidad=nConsValor and nConsCod=" & gColCFModalidad & ") as Modalidad, "
lsSql = lsSql & " (Select C.cConsDescripcion from Constante C Where CF.nCondicion=nConsValor and nConsCod=3015) as Condicion, "
lsSql = lsSql & " (select  cRazSocDescrip as Desripcion from PersFteIngreso"
lsSql = lsSql & " inner join ColocFteIngreso CFF on PersFteIngreso.cNumFuente=CFF.cNumFuente"
lsSql = lsSql & " where cUltimaActualizacion=(Select top 1 max(cUltimaActualizacion)  from PersFteIngreso  where cPerscod=P1.cPersCod)"
lsSql = lsSql & " and  CFF.cCtaCod=CF.cCtaCod) as F_Ingreso,"
lsSql = lsSql & " (Select A.cPersNombre from Persona A Inner Join ProductoPersona P  "
lsSql = lsSql & " on A.cPersCod=P.cPersCod Where P.cCtaCod=CF.cCtaCod and nPrdPersRelac=29 ) as Apoderado    "
'JOEP20181221 CP
lsSql = lsSql & ",CF.cOtrsModalidades as OtrsModalidades "
lsSql = lsSql & ",CF.nModalidad as nModalidad "
'JOEP20181221 CP
lsSql = lsSql & "From ColocCartaFianza CF  Inner Join ColocCFEstado CE on ce.cctacod=CF.cctacod   "
'By Capi Nov2007
If lnEstado >= 2002 Then
    lsSql = lsSql & "And nPrdEstado=2002"
End If
'End By
lsSql = lsSql & "Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod "
lsSql = lsSql & "Inner Join ProductoPersona PP on PP.cCtaCod = CF.cCtaCod "
lsSql = lsSql & "Inner Join ProductoPersona PP2 on PP2.cCtaCod=CF.cCtaCod "
lsSql = lsSql & "Inner Join Persona P1 on P1.cPersCod = PP.cPersCod "
lsSql = lsSql & "Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod "
lsSql = lsSql & "Where PP.nPrdPersRelac = 20 and  PP2.nPrdPersRelac=28 and "
lsSql = lsSql & "CF.cCtaCod='" & psCtaCod & "'"
On Error GoTo ErrorConexion
    Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set RecuperaDatosGeneralesCF = loConecta.CargaRecordSet(lsSql)

    loConecta.CierraConexion
    Set loConecta = Nothing
    
    Exit Function

ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function

Public Function dObtieneCFianzaDePersona(ByVal psPersCod As String, _
        ByVal psEstados As String, Optional ByVal psAgencia As String = "") As Recordset

'Obtiene Listado Creditos Pig. de una Persona
Dim lrs As ADODB.Recordset
Dim lsSql As String
Dim lsAgencia As String
Dim loConecta As COMConecta.DCOMConecta

If Trim(psAgencia) = "" Then
    lsAgencia = "__"
Else
    lsAgencia = Trim(psAgencia)
End If

On Error GoTo dError

'MAVM 20100617 BAS II
'lsSQL = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
'        & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, " _
'        & "UPPER(T3.cConsDescripcion) cMoneda " _
'        & "FROM ProductoPersona PP INNER JOIN Producto P ON P.cCtaCod = PP.cCtaCod " _
'        & "INNER JOIN Constante T ON P.nPrdEstado = T.nConsValor " _
'        & "INNER JOIN Constante T1 ON PP.nPrdPersRelac = T1.nConsValor  " _
'        & "INNER JOIN Constante T2 ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor)" _
'        & "INNER JOIN Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
'        & "" _
'        & "WHERE PP.cPersCod = '" & psPersCod & "' " _
'        & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
'        & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda _
'        & "AND P.cCtaCod like '___" & lsAgencia & "[12]21" & "%' " _
'        & "AND P.nPrdEstado in ( " & psEstados & " ) "
'
'lsSQL = lsSQL & " ORDER BY PP.cCtaCod"

lsSql = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
        & " T1.cConsDescripcion cRelacion, UPPER(ISNULL(T2B.cConsDescripcion, '')) cProducto, " _
        & " UPPER(T3.cConsDescripcion) cMoneda " _
        & " FROM ProductoPersona PP INNER JOIN Producto P ON P.cCtaCod = PP.cCtaCod " _
        & " INNER JOIN Constante T ON P.nPrdEstado = T.nConsValor " _
        & " INNER JOIN Constante T1 ON PP.nPrdPersRelac = T1.nConsValor  " _
        & " INNER join Colocaciones C on P.cCtaCod = C.cCtaCod " _
        & " LEFT JOIN Constante T2B on C.cTpoCredCod = T2B.nConsValor and T2B.nConsCod = 3034 and T2B.nConsValor <>3034 " _
        & " INNER JOIN Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) "
lsSql = lsSql & " LEFT JOIN ColocCartaFianza CCF ON PP.cCtaCod =CCF.cCtaCod " 'WIOR 20120330
lsSql = lsSql & " WHERE PP.cPersCod = '" & psPersCod & "' " _
        & " AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
        & " T3.nConsCod = " & gMoneda _
        & " AND (P.cCtaCod like '___" & lsAgencia & "[12]21" & "%' Or P.cCtaCod like '___" & lsAgencia & "514" & "%') " _
        & " AND P.nPrdEstado in ( " & psEstados & " ) "
lsSql = lsSql & " and((PP.nPrdPersRelac in(38) AND CCF.cConsorcio IS null ) or (PP.nPrdPersRelac in(20) AND CCF.cConsorcio IS not null ) or (PP.nPrdPersRelac =20 AND CCF.cConsorcio IS null)) " 'WIOR 20120330
lsSql = lsSql & " ORDER BY PP.cCtaCod"

Set loConecta = New COMConecta.DCOMConecta
    loConecta.AbreConexion
    Set lrs = loConecta.CargaRecordSet(lsSql)
    Set dObtieneCFianzaDePersona = lrs
    Set lrs = Nothing
Set loConecta = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function

'CAAU
Public Function GetCF_Poliza(ByVal psCodCta As String) As String
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim coConex As COMConecta.DCOMConecta
Set coConex = New COMConecta.DCOMConecta
On Error GoTo ErrorConexion
coConex.AbreConexion
Sql = "select nPoliza from coloccfpoliza where cCtaCod = '" & psCodCta & "'"
Set rs = coConex.CargaRecordSet(Sql)
'Co.cierraConexion
If rs.EOF And rs.BOF Then
   GetCF_Poliza = "0000000"
Else
   GetCF_Poliza = Format(rs!nPoliza, "0000000")
End If
coConex.CierraConexion
Set rs = Nothing
Set coConex = Nothing
Exit Function
ErrorConexion:
    GetCF_Poliza = "0000000"
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function

Public Function Get_Agencia_CF(ByVal psCodCta As String) As String
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim coConex As COMConecta.DCOMConecta
Set coConex = New COMConecta.DCOMConecta
On Error GoTo ErrorConexion
coConex.AbreConexion
Sql = "SELECT cageDireccion Direccion,  (SELECT cUbiGeoDescripcion FROM UBICACIONGEOGRAFICA U where U.cUbiGeoCod = A.cUbiGeoCod ) Ciudad"
Sql = Sql & " FROM AGENCIAS A WHERE cAgeCod = '" & Mid(psCodCta, 4, 2) & "'"
Set rs = coConex.CargaRecordSet(Sql)
'Co.cierraConexion
If rs.EOF And rs.BOF Then
   Get_Agencia_CF = ""
Else
   Get_Agencia_CF = rs!Direccion & " - " & rs!Ciudad
End If
coConex.CierraConexion
Set rs = Nothing
Set coConex = Nothing
Exit Function

ErrorConexion:
    Get_Agencia_CF = ""
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function


Public Function GrabaCF_Poliza(ByVal psCodCta As String, ByVal pdFecSis As Date) As String
Dim Sql As String
Dim nPoliza As Long
Dim coConex As COMConecta.DCOMConecta
Set coConex = New COMConecta.DCOMConecta
On Error GoTo ErrorConexion
Dim rs As New ADODB.Recordset
coConex.AbreConexion
coConex.BeginTrans
Sql = "select max(nPoliza) Nro from coloccfpoliza where substring(cCtaCod,4,2) = '" & Mid(psCodCta, 4, 2) & "'"
Set rs = coConex.CargaRecordSet(Sql)
'Co.cierraConexion
If rs.EOF And rs.BOF Then
   nPoliza = 1
Else
   nPoliza = IIf(IsNull(rs!Nro), 0, rs!Nro) + 1
End If
Sql = "Insert coloccfpoliza"
Sql = Sql & " (cCtaCod,nPoliza,dFecha) Values"
Sql = Sql & " ('" & psCodCta & "','" & nPoliza & "','" & Format(pdFecSis, "MM/DD/YYYY") & "')"
coConex.Ejecutar (Sql)
Sql = "select max(nPoliza) Nro from coloccfpoliza where left(cCtaCod,2) = '" & Left(psCodCta, 2) & "'"
Set rs = coConex.CargaRecordSet(Sql)
GrabaCF_Poliza = Format(rs!Nro, "0000000")
coConex.CommitTrans
coConex.CierraConexion
Set rs = Nothing
Set coConex = Nothing
Exit Function
ErrorConexion:
    GrabaCF_Poliza = "0000000"
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function

'ARCV 04-07-2007
'Public Function GrabaCF_Poliza_NEW(ByVal psCodCta As String, ByVal pdFecSis As Date, ByVal pnNumPoliza As Long)
Public Function GrabaCF_Poliza_NEW(ByVal psCodCta As String, ByVal pdFecSis As Date, ByVal pnNumPoliza As Long) As Long  'WIOR 20120407

Dim Sql As String
Dim rs As ADODB.Recordset 'WIOR 20120407
Dim coConex As COMConecta.DCOMConecta
Set coConex = New COMConecta.DCOMConecta

On Error GoTo ErrorConexion 'WIOR 20120407

coConex.AbreConexion

Sql = "Insert coloccfpoliza"
Sql = Sql & " (cCtaCod,nPoliza,dFecha) Values"
Sql = Sql & " ('" & psCodCta & "'," & pnNumPoliza & ",'" & Format(pdFecSis, "MM/DD/YYYY") & "')"

'WIOR 20120407
coConex.Ejecutar (Sql)
Sql = "exec stp_sel_CFPolizaXCTA '" & psCodCta & "'"
Set rs = coConex.CargaRecordSet(Sql)
GrabaCF_Poliza_NEW = rs!nPoliza
Set rs = Nothing
'FIN


coConex.CierraConexion
Set coConex = Nothing

Exit Function
ErrorConexion: 'WIOR 20120407
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
'WIOR 20120618 ********************************************************
Public Sub GrabarEnvioFolios(ByVal psCodEnvio As String, ByVal psAgeCod As String, ByVal dFechaEnvio As Date, ByVal pnDesde As Long, ByVal pnHasta As Long, ByVal pnSFolio As Long)
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Sql = "insert into ColocCFFolioEnvio(nCodEnvio,cAgeCod,dFechaEnvio,nDesde,nHasta,nSFolio,nEstado)" & _
    "values ('" & psCodEnvio & "','" & psAgeCod & "','" & Format(dFechaEnvio, "MM/DD/YYYY") & "'," & _
    pnDesde & "," & pnHasta & "," & pnSFolio & ",0)"
oConecta.Ejecutar (Sql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Sub

Public Function ObtenerEnvioFolios(ByVal psEstado As String, Optional ByVal psAgencia As String = "") As ADODB.Recordset
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Sql = "SELECT CF.nCodEnvio,AG.cAgeDescripcion ,CF.dFechaEnvio,CF.nDesde ,CF.nHasta ,CF.nSFolio,CF.nEstado" & _
    " FROM ColocCFFolioEnvio CF " & _
    " INNER JOIN Agencias AG ON AG.cAgeCod=CF.cAgeCod  " & _
    " WHERE CF.nEstado IN (SELECT Valor  FROM dbo.fnc_getTblValoresNumerico ('" & psEstado & "'))" & _
    " AND CF.cAgeCod LIKE '%" & psAgencia & "%' order by CF.nCodEnvio desc"

Set ObtenerEnvioFolios = oConecta.CargaRecordSet(Sql)
oConecta.CierraConexion
Set oConecta = Nothing

Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Public Sub QuitarEnvioFolios(ByVal psCodEnvio As String)
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Sql = "delete ColocCFFolioEnvio where nCodEnvio = '" & psCodEnvio & "'"
oConecta.Ejecutar (Sql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Sub
Public Function ObtenerFolios(ByVal pnDesde As Long, ByVal pnHasta As Long) As ADODB.Recordset
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Sql = "SELECT AG.cAgeDescripcion,CF.nDesde,CF.nHasta" & _
    " FROM ColocCFFolioEnvio CF " & _
    " INNER JOIN Agencias AG ON AG.cAgeCod=CF.cAgeCod " & _
    " WHERE ((nDesde>=" & pnDesde & " and nDesde <=" & pnHasta & ") OR (nHasta>=" & pnDesde & " and nHasta <=" & pnHasta & ")) " & _
    " Union " & _
    " select  AG.cAgeDescripcion, MIN(CF.nPoliza) nDesde, MAX(CF.nPoliza) nHasta " & _
    " from ColocCFPoliza CF " & _
    " INNER JOIN Agencias AG ON AG.cAgeCod=SUBSTRING(CF.cCtaCod ,4,2) " & _
    " where nPoliza >=" & pnDesde & " and nPoliza<=" & pnHasta & _
    " group by AG.cAgeDescripcion"

Set ObtenerFolios = oConecta.CargaRecordSet(Sql)
oConecta.CierraConexion
Set oConecta = Nothing

Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Public Sub ActualizarEnvioFolios(ByVal psCodEnvio As String, ByVal pnEstado As Integer)
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Sql = "update ColocCFFolioEnvio set nEstado = " & pnEstado & " where nCodEnvio = '" & psCodEnvio & "'"
oConecta.Ejecutar (Sql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Sub
Public Function ObtenerCtaCod(ByVal pnPoliza As Long) As String
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim coConex As COMConecta.DCOMConecta
Set coConex = New COMConecta.DCOMConecta
On Error GoTo ErrorConexion
coConex.AbreConexion
Sql = "select cCtaCod from coloccfpoliza where nPoliza = " & pnPoliza & ""
Set rs = coConex.CargaRecordSet(Sql)

If rs.EOF And rs.BOF Then
   ObtenerCtaCod = ""
Else
   ObtenerCtaCod = rs!cCtaCod
End If
coConex.CierraConexion
Set rs = Nothing
Set coConex = Nothing
Exit Function
ErrorConexion:
    ObtenerCtaCod = ""
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Public Sub ConfirmarEnvioFolios(ByVal psCodEnvio As String, ByVal pdFecha As Date)
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Sql = "exec stp_ins_RegistrarCFConfirmadas '" & psCodEnvio & "','" & Format(pdFecha, "yyyymmdd") & "'"
oConecta.Ejecutar (Sql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Sub
Public Sub ActualizarFolio(ByVal psNumFolio As Long, Optional ByVal pnEstado As Integer = 0, Optional ByVal psCtaCod As String = "", Optional ByVal pdFechas As Date = "01/01/1900", Optional ByVal psAgencia As String = "")
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Sql = "update ColocCFFolioEmision set "
If pnEstado <> 0 Then
    Sql = Sql & " nEstado = " & pnEstado & ","
End If

If psCtaCod <> "" Then
    Sql = Sql & " cCtaCod = " & psCtaCod & ","
End If

If pdFechas <> "01/01/1900" Then
    Sql = Sql & " dEstado = '" & Format(pdFechas, "yyyymmdd") & "',"
End If

Sql = Mid(Sql, 1, Len(Sql) - 1)
Sql = Sql & " where nNumFolio=" & psNumFolio

If psAgencia <> "" Then
    Sql = Sql & " and SUBSTRING(nCodEnvio,4,2) = '" & psAgencia & "'"
End If

oConecta.Ejecutar (Sql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Sub
Public Function ObtenerFoliosEmision(ByVal psNumFolio As Long) As ADODB.Recordset
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Sql = "SELECT nCodEnvio,nNumFolio,cCtaCod ,dEstado ,nEstado  FROM ColocCFFolioEmision " & _
    " Where nNumFolio = " & psNumFolio & ""

Set ObtenerFoliosEmision = oConecta.CargaRecordSet(Sql)
oConecta.CierraConexion
Set oConecta = Nothing

Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function

Public Function ObtenerNumFolioAEmitir(ByVal psAgencia As String, ByVal pnEstado As Integer, Optional ByVal pbSinFolio As Boolean = False) As ADODB.Recordset
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

If Not pbSinFolio Then
    Sql = "select ISNULL(MIN(CFEM.nNumFolio ),0) nNumFolio from ColocCFFolioEmision CFEM " & _
        " INNER JOIN ColocCFFolioEnvio CFENV ON CFEM.nCodEnvio =CFENV.nCodEnvio " & _
        " WHERE CFEM.nEstado =" & pnEstado & " AND SUBSTRING(CFEM.nCodEnvio,4,2)='" & psAgencia & "'"
Else
    Sql = "select nCodEnvio,cAgeCod,dFechaEnvio ,nDesde ,nHasta,nSFolio,nSFolioUtl ,nEstado " & _
        " from ColocCFFolioEnvio where dFechaEnvio =(select MIN(dFechaEnvio) from ColocCFFolioEnvio " & _
        " where cAgeCod ='" & psAgencia & "' and nEstado =" & pnEstado & ") AND nSFolio>0 and nSFolio>nSFolioUtl"
End If

Set ObtenerNumFolioAEmitir = oConecta.CargaRecordSet(Sql)
oConecta.CierraConexion
Set oConecta = Nothing

Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function

Public Sub QuitarEmisionFolio(ByVal psCtaCod As String)
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Sql = "delete from ColocCFPoliza where cCtaCod ='" & psCtaCod & "'"
oConecta.Ejecutar (Sql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Sub

Public Sub ActualizarFolioSinNumero(ByVal psCodEnvio As String, Optional ByVal pnSFolio As Integer = 0)
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion
Sql = "UPDATE ColocCFFolioEnvio SET nSFolioUtl=nSFolioUtl+" & pnSFolio & " WHERE nCodEnvio='" & psCodEnvio & "'"
oConecta.Ejecutar (Sql)
oConecta.CierraConexion
Set oConecta = Nothing
Exit Sub
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Sub
Public Function UltimoRegistroEnvio(Optional ByVal psAgencias As String = "", Optional ByVal pnEstado As Integer = 0, Optional ByVal psNumFolio As Long = 0, Optional ByVal pnSFolio As Integer = 0) As ADODB.Recordset
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

If psNumFolio <> 0 Then
    Sql = "SELECT nCodEnvio  FROM ColocCFFolioEnvio WHERE nHasta =" & psNumFolio & "  AND nSFolio =nSFolioUtl"
Else
    Sql = "SELECT nCodEnvio  FROM ColocCFFolioEnvio CFEN WHERE nSFolio=nSFolioUtl+" & pnSFolio & _
        " AND dFechaEnvio =(select MIN(dFechaEnvio) from ColocCFFolioEnvio " & _
        " where cAgeCod ='" & psAgencias & "' and nEstado =" & pnEstado & ") " & _
        " AND nHasta =(SELECT ISNULL(MAX(nNumFolio),0)  FROM ColocCFFolioEmision " & _
        " WHERE nCodEnvio =CFEN.nCodEnvio AND nEstado <>0 )"
End If

Set UltimoRegistroEnvio = oConecta.CargaRecordSet(Sql)
oConecta.CierraConexion
Set oConecta = Nothing

Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function

Public Function ConsultarEnviosFolios(ByVal pnTipo As Integer, Optional ByVal psAgencias As String = "", Optional ByVal pdFechaA As Date = "01/01/1900", Optional ByVal pdFechaB As Date = "01/01/1900", Optional ByVal psEstado As String = "", Optional ByVal psCodEnvio As String = "") As ADODB.Recordset
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

If pnTipo = 1 Then
    Sql = " SELECT CF.nCodEnvio,AG.cAgeDescripcion ,CF.dFechaEnvio,CF.nDesde ,CF.nHasta ,CF.nSFolio " & _
        " ,CF.nSFolioUtl,CONS.cConsDescripcion Estado FROM ColocCFFolioEnvio CF " & _
        " INNER JOIN Agencias AG ON AG.cAgeCod=CF.cAgeCod " & _
        " inner join Constante CONS ON CONS.nConsValor =CF.nEstado AND CONS.nConsCod =9996 " & _
        " WHERE CF.nEstado IN (SELECT Valor  FROM dbo.fnc_getTblValoresNumerico ('" & psEstado & "')) " & _
        " AND CF.cAgeCod ='" & psAgencias & "'" & _
        " and CF.dFechaEnvio between '" & Format(pdFechaA, "yyyymmdd") & "' and '" & Format(pdFechaB, "yyyymmdd") & "'" & _
        " order by CF.dFechaEnvio"
ElseIf pnTipo = 2 Then
    Sql = " SELECT CFe.nCodEnvio,AG.cAgeDescripcion,CFE.nNumFolio,CFE.cCtaCod ,CFE.dEstado " & _
        " ,CONS.cConsDescripcion Estado " & _
        " FROM  ColocCFFolioEmision CFE " & _
        " INNER JOIN Agencias AG ON AG.cAgeCod=substring(CFE.nCodEnvio,4,2) " & _
        " inner join Constante CONS ON CONS.nConsValor =CFE.nEstado AND CONS.nConsCod =9997 " & _
        " WHERE CFE.nCodEnvio ='" & psCodEnvio & "'" & _
        " order by CFe.nNumFolio "
End If

Set ConsultarEnviosFolios = oConecta.CargaRecordSet(Sql)
oConecta.CierraConexion
Set oConecta = Nothing

Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
Public Function ExisteRemeseas(ByVal psAgencia As String) As ADODB.Recordset
On Error GoTo ErrorConexion
Dim Sql As String
Dim oConecta As COMConecta.DCOMConecta
Set oConecta = New COMConecta.DCOMConecta
oConecta.AbreConexion

Sql = "select count(nCodEnvio) nCodEnvio from ColocCFFolioEnvio where cAgeCod ='" & psAgencia & "'"

Set ExisteRemeseas = oConecta.CargaRecordSet(Sql)
oConecta.CierraConexion
Set oConecta = Nothing

Exit Function
ErrorConexion:
    Err.Raise Err.Number, "Error Carta Fianza ", Err.Description
End Function
'WIOR FIN *************************************************************
'WIOR 20120806*********************************************************
Public Sub t()
Dim loBase As COMDCredito.DCOMCredActBD
Dim lnMovNro As Long
Dim lsOpeCod As String

lsOpeCod = gColCFOpeEmision

'On Error GoTo ErrorModCF
Set loBase = New COMDCredito.DCOMCredActBD



End Sub
'WIOR FIN *************************************************************
