VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCartaFianzaValida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'****************************
'* Clase que contiene las validaciones de las operaciones de Credito Pignoraticio
'* LAYG - 01/07/2001
'****************************
Option Explicit

Function nCFGarantiasGravada(ByVal psCtaCod As String, ByVal pnTipoCambioFijo As Double) As Double
' Obtiene el Monto de las Garantias Gravadas
Dim loConecta As DConecta
Dim lsSQL As String
Dim lrValida As ADODB.Recordset

    lsSQL = "SELECT isNull(SUM (Case When nMoneda = 1 then CG.nGravado " _
         & "                         When nMoneda = 2 then CG.nGravado * " & pnTipoCambioFijo & " end ),0) ValGrav " _
         & " FROM ColocGarantia CG " _
         & " Where CG.cCtaCod = '" & psCtaCod & "' "

    Set loConecta = New DConecta
        loConecta.AbreConexion
        Set lrValida = loConecta.CargaRecordSet(lsSQL)
        loConecta.CierraConexion
    Set loConecta = Nothing

    If (lrValida.BOF Or lrValida.EOF) Or (lrValida!ValGrav = 0) Then
        MsgBox "El Carta Fianza No Tiene Garantias Gravadas", vbInformation, "Aviso"
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
Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset
Dim lbOk As Boolean

lsSQL = "SELECT cCodNiv  FROM ColocCredPersNivelesApr  " _
      & "WHERE cCodNiv in ( SELECT cCodNiv FROM ColocCredNivelesApr " _
      & "                       WHERE  cProduct = '" & psProducto & "' )" _
      & " AND cPersCod ='" & psPersCod & "' "

    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If Not (lrValida.BOF And lrValida.EOF) Then ' Encuentra el permiso
        lbOk = True
    Else
        lbOk = False
    End If
    Set lrValida = Nothing
    nCFPermisoAprobacion = lbOk
End Function

Public Function RecuperaPersonasEstadoCF(ByVal pnEstadoCF As Variant, Optional ByVal pMatProd As Variant = Nothing) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta
Dim vsEstadoCred As String
Dim vsProductos As String
Dim i As Integer

    On Error GoTo ErrorRecuperaPersonasEstadoCF
    vsEstadoCred = ""
    For i = 0 To UBound(pnEstadoCF)
        vsEstadoCred = vsEstadoCred & pnEstadoCF(i) & ","
    Next i
    vsEstadoCred = Mid(vsEstadoCred, 1, Len(vsEstadoCred) - 1)
    vsProductos = ""
    If IsArray(pMatProd) Then
        vsProductos = "('"
        For i = 0 To UBound(pMatProd)
            vsProductos = vsProductos & pMatProd(i) & "','"
        Next i
        vsProductos = Mid(vsProductos, 1, Len(vsProductos) - 2)
    End If
    vsProductos = vsProductos & ")"
    
    lsSQL = "Select P.cPersCod, P.cPersNombre,C.cCtaCod  from Persona P inner join ProductoPersona PP ON P.cPersCod = PP.cPersCod"
    lsSQL = lsSQL & " Inner join ColocCartaFianza C ON PP.cCtaCod = C.cCtaCod "
    lsSQL = lsSQL & " Inner join Producto Prd ON PP.cCtaCod = Prd.cCtaCod "
    lsSQL = lsSQL & " WHERE PP.nPrdPersRelac = " & gColRelPersTitular & " AND Prd.nPrdEstado in (" & Trim(vsEstadoCred) & ") "
    If Len(vsProductos) > 1 Then
        lsSQL = lsSQL & " AND  SUBSTRING(Prd.cCtaCod,6,3) IN " & vsProductos
    End If
    lsSQL = lsSQL & " order by P.cPersNombre"
    
    Set loConecta = New DConecta
        loConecta.AbreConexion
        Set RecuperaPersonasEstadoCF = loConecta.CargaRecordSet(lsSQL)
        loConecta.CierraConexion
    Set loConecta = Nothing
    
    Exit Function
    
ErrorRecuperaPersonasEstadoCF:
    
End Function

Public Function nCFPagoComision(ByVal psCta As String) As Double
Dim lsSQL As String
Dim loConecta As DConecta
Dim lrValida As ADODB.Recordset
Dim lnComisPag As Double

lsSQL = "SELECT IsNull(Sum(nMonto),0) ComisPag FROM MovCol MC join Mov M on MC.nMovNro = M.nMovNro " _
      & "WHERE cCtaCod ='" & psCta & "' And MC.cOpeCod ='" & gColCFOpeComisEfe & "' " _
      & " AND M.nMovFlag in (" & gMovFlagVigente & " ) "

    
    Set loConecta = New DConecta
        loConecta.AbreConexion
        Set lrValida = loConecta.CargaRecordSet(lsSQL)
        loConecta.CierraConexion
    Set loConecta = Nothing
   
    lnComisPag = lrValida!ComisPag
    
    Set lrValida = Nothing
    nCFPagoComision = lnComisPag
End Function

Public Function nBuscaOperacionesCFParaExtorno(ByVal psFecTrans As String, ByVal psOpeExtor As String, _
        Optional ByVal psCtaCod As String = "@", Optional ByVal pbOrdenAsc As Boolean = False) As Recordset

Dim lsSQL As String
Dim loRegValida As dColPFunciones
Dim lrValida As ADODB.Recordset

    lsSQL = " SELECT M.cMovNro, M.nMovNro, M.cOpeCod, M.nMovFlag, " _
          & " MC.cCtaCod, MC.nMonto, Ope.cOpeDesc  " _
          & " FROM Mov M INNER JOIN MovCol MC ON M.nMovNro = MC.nMovNro " _
          & " LEFT JOIN OpeTpo Ope ON MC.cOpeCod = Ope.cOpeCod " _
          & " WHERE M.cMovNro LIKE '" & psFecTrans & "%' " _
          & " AND M.nMovFlag = " & gMovFlagVigente & " " _
          & " AND M.cOpeCod in " & psOpeExtor & "  "
    If psCtaCod <> "@" Then
        lsSQL = lsSQL & " AND MC.cCtaCod ='" & psCtaCod & "' "
    End If
    lsSQL = lsSQL & " ORDER BY M.nMovNro DESC  "
    
    Set loRegValida = New dColPFunciones
    Set lrValida = loRegValida.dObtieneRecordSet(lsSQL)
    If lrValida Is Nothing Then
        MsgBox "ERROR: al Buscar datos para Extorno ", vbInformation, "Aviso"
        Exit Function
    End If
    
    Set nBuscaOperacionesCFParaExtorno = lrValida
    Set lrValida = Nothing
End Function


'******************
'CAAU - 29/10/2002
'******************
Function RecuperaRelacinesInternas(ByRef Codpersona As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta

lsSQL = "Select  P1.cPersNombre as Cliente,"
lsSQL = lsSQL & " (Select C.cConsDescripcion as Describe from Constante C where PP.nPrdPersRelac=C.nConsValor and nConsCod=3002) as Relacion,"
lsSQL = lsSQL & " CF.cCtaCod as Carta,"
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C where  Pro.nPrdEstado=nConsValor and nConsCod=3001) as Estado,"
lsSQL = lsSQL & " (Select nMonto from ColocCFEstado where cCtaCod=CF.cCtaCod and nPrdEstado= Pro.nPrdEstado ) Saldo, "
'lsSQL = lsSQL & " Pro.nSaldo as saldo,"
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C where  substring(CF.cCtacod,9,1)=nConsValor and nConsCod=1011) as Moneda"
lsSQL = lsSQL & " From ColocCartaFianza CF"
lsSQL = lsSQL & " Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod"
lsSQL = lsSQL & " Inner Join ProductoPersona PP on PP.cCtaCod = CF.cCtaCod"
lsSQL = lsSQL & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod"
lsSQL = lsSQL & " Where PP.nPrdPersRelac = 20 and Pro.nPrdEstado = 2001 And P1.cPersCod =" & Codpersona

On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RecuperaRelacinesInternas = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    'End If
    Exit Function
ErrorConexion:
    MsgBox Err.Description
End Function
Function RecuperaFuenteIngreso(ByRef Codpersona As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta
lsSQL = " Select  PFI.cPersCod as CodigoPers, PFI.cUltimaActualizacion as UltimaAct, PFI.cRazSocDirecc as Direccion, PFI.cRazSocDescrip as Razon, "
lsSQL = lsSQL & " PFI.dPersFIinicio as FechaInicio,"
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C Where PFI.nPersTipFte = nConsValor and nConsCod=1028) as Tipo_Ingreso,"
lsSQL = lsSQL & " (Select cCIIUdescripcion from CIIU  Where P.cPersCIIU = cCIIUcod) as Actividad"
lsSQL = lsSQL & " from PersFteIngreso PFI inner join  Persona P on P.cPersCod  = PFI.cPersCod"
lsSQL = lsSQL & " where PFI.cUltimaActualizacion=(Select max(cUltimaActualizacion) from PersFteIngreso where cPersCod=PFI.cPErsCod)"
lsSQL = lsSQL & " and PFI.cPersCod=" & Codpersona
On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RecuperaFuenteIngreso = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function
ErrorConexion:
    MsgBox Err.Description
End Function
Function RecuperaDatosAcreedor(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta

 lsSQL = " Select CF.cCtaCod as Carta, P1.cPersCod as Codigo, P1.cPersNombre as Nombre, P1.cPersDireccDomicilio as Direccion, "
 lsSQL = lsSQL & "P1.nPersPersoneria as TipoPers, "
 lsSQL = lsSQL & "(Select C.cConsDescripcion from Constante C where  substring(CF.cCtacod,6,3)=nConsValor and nConsCod=1001) as Tipo, "
 lsSQL = lsSQL & "(Select C.cConsDescripcion from Constante C where  Pro.nPrdEstado=nConsValor and nConsCod=3001) as Estado, "
 lsSQL = lsSQL & "DocDNI = case when P1.nPersPersoneria = 1 then (Select cPersIDnro from PersID where P1.cPersCod=cPersCod and cPersIdTpo=1) end, "
 lsSQL = lsSQL & "DocRUC = case when P1.nPersPersoneria <> 1 then (Select cPersIDnro from PersID where P1.cPersCod=cPersCod and cPersIdTpo=2) end "
 lsSQL = lsSQL & "From ColocCartaFianza CF "
 lsSQL = lsSQL & "Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod "
 lsSQL = lsSQL & "Inner Join ProductoPersona  PP on PP.cCtaCod = CF.cCtaCod "
 lsSQL = lsSQL & "Inner Join Persona P1 on P1.cPersCod = PP.cPersCod "
 lsSQL = lsSQL & "Where PP.nPrdPersRelac = 35 and CF.cCtaCod=" & psCtaCod
On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RecuperaDatosAcreedor = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function
ErrorConexion:
    MsgBox Err.Description
End Function
Function RecuperaDatosJuridica(ByRef Codpersona As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta
lsSQL = "SELECT PJ.cPersCod, "
lsSQL = lsSQL & " (SELECT cPersJurTpoDesc From PersJurTpo Where PJ.cPersJurTpo=cPersJurTpoCod ) as DescribeJur,"
lsSQL = lsSQL & " (SELECT cConsDescripcion From Constante Where PJ.cPersJurMagnitud=nConsValor and nConsCod=1004) as Magnitud"
lsSQL = lsSQL & " FROM PERSONAJUR PJ WHERE PJ.cPersCod=" & Codpersona
On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RecuperaDatosJuridica = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function
ErrorConexion:
    MsgBox Err.Description
End Function
Function RecuperaDatosT(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta
lsSQL = "Select CF.cCtaCod as Carta, P1.cPersCod as Codigo, P1.cPersNombre as Nombre, P1.cPersDireccDomicilio as Direccion, "
lsSQL = lsSQL & " P1.nPersPersoneria as TipoPers,P1.dPersNacCreac as Nacimiento, "
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C where  substring(CF.cCtacod,6,3)=nConsValor and nConsCod=1001) as Tipo, "
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C where  Pro.nPrdEstado=nConsValor and nConsCod=3001) as Estado, "
lsSQL = lsSQL & " DocDNI = case when P1.nPersPersoneria = 1 then (Select cPersIDnro from PersID where P1.cPersCod=cPersCod and cPersIdTpo=1) end, "
lsSQL = lsSQL & " DocRUC = case when P1.nPersPersoneria <> 1 then (Select cPersIDnro from PersID where P1.cPersCod=cPersCod and cPersIdTpo=2) end "
lsSQL = lsSQL & " From ColocCartaFianza CF "
lsSQL = lsSQL & " Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod "
lsSQL = lsSQL & " Inner Join ProductoPersona  PP on PP.cCtaCod = CF.cCtaCod "
lsSQL = lsSQL & " Inner Join Persona P1 on P1.cPersCod = PP.cPersCod "
lsSQL = lsSQL & " Where PP.nPrdPersRelac = 20 and CF.cCtaCod=" & psCtaCod
On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RecuperaDatosT = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing
    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function

ErrorConexion:
    MsgBox Err.Description
End Function

Function RepcuperaOtrosDt(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta

lsSQL = lsSQL & "Select CFE.dPrdEstado as Fecha,"
lsSQL = lsSQL & "( Select C.cConsDescripcion as Describe from Constante C where CFE.nPrdEstado=C.nConsValor and nConsCod=3001) Condicion, "
lsSQL = lsSQL & "( Select C.cConsDescripcion as Describe from Constante C where CFE.nMotivoRechazo=C.nConsValor and nConsCod=3024) Motivo "
lsSQL = lsSQL & "from ColocCFEstado CFE where CFE.nPrdEstado=2003 "
lsSQL = lsSQL & "and CFE.cCtaCod=" & psCtaCod

On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RepcuperaOtrosDt = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing

    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '    MsgBox "No existen Personas Relacionadas", vbInformation, "CMACT"
    
    'End If
    Exit Function

ErrorConexion:
    MsgBox Err.Description
End Function
Function RecuperaPersonaRelacion(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta

lsSQL = "Select P.cPersCod as CodPersona, p.cPersNombre as Nombre, cCtaCod, "
lsSQL = lsSQL & "( Select C.cConsDescripcion as Describe from Constante C where PP.nPrdPersRelac=C.nConsValor and nConsCod=3002) as Relacion"
lsSQL = lsSQL & " from ProductoPersona PP"
lsSQL = lsSQL & " Inner Join Persona P on P.cPersCod=PP.cPersCod where PP.nPrdPersRelac!=28 and cCtaCod=" & psCtaCod

On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RecuperaPersonaRelacion = loConecta.CargaRecordSet(lsSQL)
    loConecta.CierraConexion
    Set loConecta = Nothing

    'If RecuperaPersonaRelacion.BOF And RecuperaPersonaRelacion.EOF Then
    '   MsgBox "No existen Personas Relacionadas", vbInformation, "AVISO"
    'End If
    
    Exit Function

ErrorConexion:
    MsgBox Err.Description

End Function
Function RecuperaHistorial(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta


lsSQL = "Select  CFE.nMonto as Monto, CFE.cCtaCod as CF,"
lsSQL = lsSQL & "(Select C.cConsDescripcion from Constante C where  CFE.nPrdEstado=nConsValor and nConsCod=3001) as Estado"
lsSQL = lsSQL & ", CFE.dPrdEstado as Fecha from ColocCFEstado CFE where cCtaCod=" & psCtaCod
On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RecuperaHistorial = loConecta.CargaRecordSet(lsSQL)

    loConecta.CierraConexion
    Set loConecta = Nothing

    'If RecuperaHistorial.BOF And RecuperaHistorial.EOF Then
    '    MsgBox "No existen Historial", vbInformation, "CMACT"
    'End If
    
    Exit Function

ErrorConexion:
    MsgBox Err.Description

End Function
Function RecuperaGarantias(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta

lsSQL = "Select CG.cNumGarant as NroGarantia, CF.cCtaCod as CF, "
lsSQL = lsSQL & "(Select C.cConsDescripcion from Constante C where  G.nTpoGarantia = nConsValor and nConsCod=1027 )as Tipo_Garn, "
lsSQL = lsSQL & "(Select cDocDesc from Documento where G.cTpoDoc=nDocTpo) as Tipo_Doc, "
lsSQL = lsSQL & "CG.nMoneda as Moneda, CG.nGravado as Gravado, G.cDescripcion as Descripcion "
lsSQL = lsSQL & "from ColocGarantia CG "
lsSQL = lsSQL & "Inner Join ColocCartaFianza CF on CF.cCtacod = CG.cCtaCod "
lsSQL = lsSQL & "Inner Join Garantias G on CG.cNumGarant = G.cNumgarant "
lsSQL = lsSQL & "where CF.cCtaCod=" & psCtaCod

On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RecuperaGarantias = loConecta.CargaRecordSet(lsSQL)

    loConecta.CierraConexion
    Set loConecta = Nothing

    'If RecuperaGarantias.BOF And RecuperaGarantias.EOF Then
    '    MsgBox "No existen", vbInformation, "CMACT"
    'End If
    
    Exit Function

ErrorConexion:
    MsgBox Err.Description

End Function


Function RecuperaDatosGeneralesCF(ByRef psCtaCod As String) As ADODB.Recordset
Dim lsSQL As String
Dim loConecta As DConecta

lsSQL = "Select CF.cCtaCod as Carta, CF.dVencimiento as Vence, CF.dAsignacion as F_Asignacion, "
lsSQL = lsSQL & " (Select nMonto from ColocCFEstado where cCtaCod=CF.cCtaCod and nPrdEstado=Pro.nPrdEstado ) Monto, "
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C where  Pro.nPrdEstado=nConsValor and nConsCod=3001) as Estado, "
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C where  substring(CF.cCtacod,6,3)=nConsValor and nConsCod=1001) as Tipo, "
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C where  substring(CF.cCtacod,9,1)=nConsValor and nConsCod=1011) as Moneda, "
lsSQL = lsSQL & " P1.cPersNombre as Cliente, P2.cPersNombre as Analista, CF.cFinalidad as Finalidad, "
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C Where CF.nModalidad=nConsValor and nConsCod=" & gColCFModalidad & ") as Modalidad, "
lsSQL = lsSQL & " (Select C.cConsDescripcion from Constante C Where CF.nCondicion=nConsValor and nConsCod=3015) as Condicion, "
lsSQL = lsSQL & " (select  cRazSocDescrip as Desripcion from PersFteIngreso"
lsSQL = lsSQL & " inner join ColocFteIngreso CFF on PersFteIngreso.cNumFuente=CFF.cNumFuente"
lsSQL = lsSQL & " where cUltimaActualizacion=(Select top 1 max(cUltimaActualizacion)  from PersFteIngreso  where cPerscod=P1.cPersCod)"
lsSQL = lsSQL & " and  CFF.cCtaCod=CF.cCtaCod) as F_Ingreso,"
lsSQL = lsSQL & " (Select A.cPersNombre from Persona A Inner Join ProductoPersona P  "
lsSQL = lsSQL & " on A.cPersCod=P.cPersCod Where P.cCtaCod=CF.cCtaCod and nPrdPersRelac=29 ) as Apoderado    "
lsSQL = lsSQL & "From ColocCartaFianza CF "
lsSQL = lsSQL & "Inner Join Producto Pro on CF.cCtaCod = Pro.cCtaCod "
lsSQL = lsSQL & "Inner Join ProductoPersona PP on PP.cCtaCod = CF.cCtaCod "
lsSQL = lsSQL & "Inner Join ProductoPersona PP2 on PP2.cCtaCod=CF.cCtaCod "
'lsSQL = lsSQL & "Inner Join ProductoPersona PP3 on PP3.cctaCod=CF.cCtaCod "
lsSQL = lsSQL & "Inner Join Persona P1 on P1.cPersCod = PP.cPersCod "
lsSQL = lsSQL & "Inner Join Persona P2 on P2.cPersCod = PP2.cPersCod "
'lsSQL = lsSQL & "Inner Join Persona P3 on P3.cPersCod = PP3.cPersCod "
lsSQL = lsSQL & "Where PP.nPrdPersRelac = 20 and  PP2.nPrdPersRelac=28 and "
'lsSQL = lsSQL & "PP3.nPrdPersRelac=29 and "
lsSQL = lsSQL & "CF.cCtaCod='" & psCtaCod & "'"
On Error GoTo ErrorConexion
    Set loConecta = New DConecta
    loConecta.AbreConexion
    Set RecuperaDatosGeneralesCF = loConecta.CargaRecordSet(lsSQL)

    loConecta.CierraConexion
    Set loConecta = Nothing
    
    Exit Function

ErrorConexion:
    MsgBox Err.Description
End Function

Public Function dObtieneCFianzaDePersona(ByVal psPersCod As String, _
        ByVal psEstados As String, Optional ByVal psAgencia As String = "") As Recordset

'Obtiene Listado Creditos Pig. de una Persona
Dim lrs As ADODB.Recordset
Dim lsSQL As String
Dim lsAgencia As String
Dim loConecta As DConecta

If Trim(psAgencia) = "" Then
    lsAgencia = "__"
Else
    lsAgencia = Trim(psAgencia)
End If

On Error GoTo dError

lsSQL = "SELECT PP.cCtaCod, P.nPrdEstado, PP.nPrdPersRelac, T.cConsDescripcion cEstado, " _
        & "T1.cConsDescripcion cRelacion, UPPER(T2.cConsDescripcion) cProducto, " _
        & "UPPER(T3.cConsDescripcion) cMoneda " _
        & "FROM ProductoPersona PP INNER JOIN Producto P ON P.cCtaCod = PP.cCtaCod " _
        & "INNER JOIN Constante T ON P.nPrdEstado = T.nConsValor " _
        & "INNER JOIN Constante T1 ON PP.nPrdPersRelac = T1.nConsValor  " _
        & "INNER JOIN Constante T2 ON SUBSTRING(PP.cCtaCod,6,3) = CONVERT(Varchar(3),T2.nConsValor)" _
        & "INNER JOIN Constante T3 ON SUBSTRING(PP.cCtaCod,9,1) = CONVERT(Varchar(1),T3.nConsValor) " _
        & "" _
        & "WHERE PP.cPersCod = '" & psPersCod & "' " _
        & "AND T1.nConsCod = " & gColocRelacPers & " AND T.nConsCod = " & gColocEstado & " AND " _
        & "T2.nConsCod = " & gProducto & " AND T3.nConsCod = " & gMoneda _
        & "AND P.cCtaCod like '___" & lsAgencia & "[12]21" & "%' " _
        & "AND P.nPrdEstado in ( " & psEstados & " ) "

lsSQL = lsSQL & " ORDER BY PP.cCtaCod"

Set loConecta = New DConecta
    loConecta.AbreConexion
    Set lrs = loConecta.CargaRecordSet(lsSQL)

    
    Set dObtieneCFianzaDePersona = lrs
    Set lrs = Nothing
Set loConecta = Nothing

Exit Function

dError:
    Err.Raise Err.Number, "Obtiene Datos Carta Fianza <<dObtieneCFianzaDePersona>>", Err.Description
    
End Function


