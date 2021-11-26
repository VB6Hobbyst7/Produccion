VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCaptaGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
Public oImp As COMFunciones.FCOMVarImpresion
Public gcPEN As New COMDConstantes.DCOMConstantes 'MARG ERS044-2016

Public Sub IniciaImpresora(Optional ByVal nImpresora As COMFunciones.Impresoras = gEPSON)
Set oImp = New COMFunciones.FCOMVarImpresion
oImp.Inicia nImpresora
End Sub

Public Function GetListadoCuentasCTS(ByVal sPersCod As String) As ADODB.Recordset
Dim sSql As String
Dim rs As New ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales

sSql = "select pp.cctacod,ct.ccodInst,cInstitucion=i.cpersnombre,cEstado=k.cconsdescripcion from producto p "
sSql = sSql & " join captaccts ct on ct.cctacod=p.cctacod"
sSql = sSql & " join productopersona pp on pp.cctacod=p.cctacod"
sSql = sSql & " join persona i on i.cperscod=ct.ccodinst"
sSql = sSql & " join (select * from constante where nconscod=2001) K on k.nconsvalor=p.nprdestado"
sSql = sSql & " where p.cctacod like '108__234%'"
sSql = sSql & " and p.nprdestado not in (1300,1400) and pp.nprdpersrelac=10"
sSql = sSql & " and pp.cperscod='" & sPersCod & "'"



'CUSCO
rs.CursorLocation = adUseClient
rs.Open sSql, clsMant.dbCmact, adOpenStatic, adLockReadOnly, adCmdText
If rs.EOF Then
    'Set rsCta = Nothing
    Set GetListadoCuentasCTS = Nothing
Else
    Set rs.ActiveConnection = Nothing
    Set GetListadoCuentasCTS = rs
    'Set rsCta = Nothing
End If

End Function

Public Function GetCTSPeriodo() As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetCTSPeriodo = clsMant.GetCTSPeriodo()
Set clsMant = Nothing
End Function

Public Sub EliminaBeneficiarios(ByVal sPersona As String)
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
clsMant.EliminaBeneficiarios sPersona
Set clsMant = Nothing
End Sub
'ferimoro CTI2: ERS039-2019 21082020 //JATO 20210407
Public Function recuperaCCI(ByVal nCctaCod As String) As String
    Dim sSql As String
    Dim oDCOMConecta As New COMConecta.DCOMConecta
    Dim rsCom As New ADODB.Recordset

    oDCOMConecta.AbreConexion
    sSql = "exec stp_sel_ERS_030_2019_RecuperaCCI '" & nCctaCod & "'"
    Set rsCom = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion

    If Not (rsCom.BOF And rsCom.EOF) Then
        recuperaCCI = rsCom("CCI")
    End If

Set oDCOMConecta = Nothing
Set rsCom = Nothing
End Function
'ferimoro CTI2: ERS039-2019 21082020 //JATO 20210407
Public Function recuperaSubProducto(ByVal nProducto As String, ByVal nTpoProducto As String) As String
    Dim sSql As String
    Dim oDCOMConecta As New COMConecta.DCOMConecta
    Dim rsCom As New ADODB.Recordset

    oDCOMConecta.AbreConexion
    sSql = "exec stp_sel_ERS_030_2019_RecuperaSubProducto '" & nProducto & "','" & nTpoProducto & "'"
    Set rsCom = oDCOMConecta.CargaRecordSet(sSql)
    oDCOMConecta.CierraConexion

    If Not (rsCom.BOF And rsCom.EOF) Then
        recuperaSubProducto = rsCom("cConsDescripcion")
    End If

Set oDCOMConecta = Nothing
Set rsCom = Nothing
End Function

'************TODOCOMPLETA Error en rsBenef("%")
'**********************************************

Public Function ActualizaBeneficiarios(ByVal sPersona As String, ByVal rsBenef As ADODB.Recordset, ByVal sMovNro As String) As Boolean
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean
Dim nParentesco As PersRelacion
Dim nValor As Double
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.EliminaBeneficiarios sPersona
Do While Not rsBenef.EOF
    nValor = CDbl(rsBenef("%"))
    nParentesco = CLng(Trim(Right(rsBenef("Parentesco"), 2)))
    clsMant.AgregaBeneficiario sPersona, rsBenef("Codigo"), nParentesco, nValor, sMovNro
    rsBenef.MoveNext
Loop
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
ActualizaBeneficiarios = True
Exit Function
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    ActualizaBeneficiarios = False
End Function

Public Function ActualizaTarjetaEstado(ByVal sTarjeta As String, ByVal sMovNro As String, _
        ByVal nEstado As CaptacTarjetaEstado, ByVal sGlosa As String, Optional sClave As String = "") As Boolean

Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean

Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.ActualizaTarjetaEstado sTarjeta, nEstado, sClave
clsMant.AgregaTarjetaEstado sTarjeta, sMovNro, nEstado, sGlosa
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
ActualizaTarjetaEstado = True
Exit Function
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    'MsgBox Err.Description, vbExclamation, "Error"
    Err.Raise Err.Number, "Error", Err.Description
    ActualizaTarjetaEstado = False
End Function

Public Function GetCuentasNoAsociadas(ByVal sPersCod As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetCuentasNoAsociadas = clsMant.GetCuentasNoAsociadas(sPersCod)
Set clsMant = Nothing
End Function
Public Function GetOrdenPagoEmitidas(ByVal sCuenta As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetOrdenPagoEmitidas = clsMant.GetOrdenPagoEmitidas(sCuenta)
Set clsMant = Nothing
End Function

Public Function GetDatosPersona(ByVal sPersona As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetDatosPersona = clsMant.GetDatosPersona(sPersona)
Set clsMant = Nothing
End Function

Public Function GetPersonaTarj(ByVal sPersona As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetPersonaTarj = clsMant.GetPersonaTarj(sPersona)
Set clsMant = Nothing
End Function

Public Function GetProductoPersona(ByVal sCuenta As String, Optional sFiltro As String = "") As ADODB.Recordset
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set GetProductoPersona = clsMant.GetProductoPersona(sCuenta, sFiltro)
    Set clsMant = Nothing
End Function

Public Function GetDatosCuenta(ByVal sCuenta As String) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
Dim nProd As Producto
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
nProd = CInt(Mid(sCuenta, 6, 3))
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Select Case nProd
    Case gCapAhorros
        Set rsCta = clsMant.GetDatosCuentaAho(sCuenta)
        
    Case gCapPlazoFijo
        Set rsCta = clsMant.GetDatosCuentaPF(sCuenta)
        
    Case gCapCTS
        Set rsCta = clsMant.GetDatosCuentaCTS(sCuenta)
        
End Select
Set GetDatosCuenta = rsCta
Set rsCta = Nothing
Set clsMant = Nothing
End Function
'*** BRGO 20111020 **********************************************
Public Function TieneCuentasCaptacxSubProducto(ByVal sPersCod As String, ByVal nProducto As Integer, ByVal nTpoPrograma As Integer, Optional ByVal psCodInst As String = "") As Boolean 'JUEZ 20140319 Se agregó psCodInst
    Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim rs As ADODB.Recordset
    
    Dim bValida As Boolean
    
    Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    bValida = False
    Set rs = oCap.GetCuentasCaptacxSubProducto(sPersCod, nProducto, nTpoPrograma, psCodInst)
    If rs.EOF <> True And rs.BOF <> True Then
        bValida = True
    End If
    TieneCuentasCaptacxSubProducto = bValida
    Set oCap = Nothing
    Set rs = Nothing
End Function
'*** END BRGO **************************************************

Public Function GetCuentasPersonaAI(Optional ByVal sPersCod As String) As ADODB.Recordset
  Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales, rsTemp As ADODB.Recordset
  Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
  
              Set rsTemp = clsMant.GetCuentasPersonaAI(sPersCod)
              Set GetCuentasPersonaAI = rsTemp
  
  Set clsMant = Nothing
End Function

Public Function GetTpoActI() As ADODB.Recordset
  Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales, rsTemp As ADODB.Recordset
  Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
  
              Set rsTemp = clsMant.GetTpoActI()
              Set GetTpoActI = rsTemp
  
  Set clsMant = Nothing
End Function

Public Function GetTpoDocActI() As ADODB.Recordset
  Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales, rsTemp As ADODB.Recordset
  Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
  
              Set rsTemp = clsMant.GetTpoDocActI()
              Set GetTpoDocActI = rsTemp
  
  Set clsMant = Nothing
End Function

Public Function GetDPTO() As ADODB.Recordset
  Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales, rsTemp As ADODB.Recordset
  Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
  
              Set rsTemp = clsMant.GetDPTO()
              Set GetDPTO = rsTemp
  
  Set clsMant = Nothing
End Function

Public Function GetPROV(ByVal cDpto As String) As ADODB.Recordset
  Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales, rsTemp As ADODB.Recordset
  Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
  
              Set rsTemp = clsMant.GetPROV(cDpto)
              Set GetPROV = rsTemp
  
  Set clsMant = Nothing
End Function

Public Function EsCtaConvenio(ByVal sCuenta As String) As Boolean
EsCtaConvenio = False
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    
     
    EsCtaConvenio = clsMant.EsCtaConvenio(sCuenta)
    
    Set clsMant = Nothing

End Function

Public Function ActualizaDatosCuenta(ByVal sCuenta As String, ByVal rsRel As ADODB.Recordset, _
        ByVal nFirmas As Integer, ByVal nTipoCuenta As ProductoCuentaTipo, _
        Optional sInstitucion As String = "", Optional bOrdPag As Boolean = False, _
        Optional sCuentaAboInt As String = "", Optional nFirmasMin As Integer, Optional sMovNro As String = "", _
        Optional rsEliminados As ADODB.Recordset, Optional sAlias As String = "", _
        Optional ByVal nTipoPrograma As Integer = 0, Optional ByVal pnPorRetAdiCTS As Integer = 0, Optional ByVal nSubasta As Integer, _
        Optional ByVal sGrupos As String, Optional ByVal sReglas As String) As Boolean
        'MIOL 20121109, SEGUN RFC098-2012-B SE AGREGO: nSubasta
        'By Capi Agosto2007 ultimo parametro
        'RIRO 20131102 SEGUN "CAMBIOS EN PODERES", SE AGREGO "sGrupos, sReglas"

Dim rsCta As ADODB.Recordset, rsTarj As ADODB.Recordset
Dim nProd As Producto
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim clsMovs As COMDCaptaGenerales.DCOMCaptaMovimiento 'By Capi Acta 014-2007
Dim bTrans As Boolean, COBLIGATORIO As String
Dim nRelacion As CaptacRelacPersona
Dim sPersona As String
Dim nOperacion As String 'By Capi Acta 014-2007
Dim sGlosa As String 'By Capi Acta 014-2007
Dim nMovNro As Long 'By Capi Acta 014-2007
Dim bCtaAboIntInact As Boolean, nSaldDispCtaAboIntInact As Double, nSaldCntCtaAboIntInact As Double 'JUEZ 20130909

nProd = CInt(Mid(sCuenta, 6, 3))
On Error GoTo ErrGraba
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMovs = New COMDCaptaGenerales.DCOMCaptaMovimiento 'By Capi Acta 014-2007
'JUEZ 20130909 *******************************************
If nProd = gCapPlazoFijo And sCuentaAboInt <> "" Then
    Dim rsCtaAbo As ADODB.Recordset
    Set rsCtaAbo = clsMant.GetDatosCuentaAho(sCuentaAboInt)
    If rsCtaAbo!bInactiva Then
        bCtaAboIntInact = True
        nSaldDispCtaAboIntInact = rsCtaAbo("nSaldoDisp")
        nSaldCntCtaAboIntInact = rsCtaAbo("nSaldo")
    End If
    Set rsCtaAbo = Nothing
End If
'END JUEZ ************************************************
clsMant.dbCmact.BeginTrans
bTrans = True
'By Capi Acta 014-2007
    nOperacion = "220303"
    sGlosa = sAlias
    clsMovs.AgregaMov sMovNro, nOperacion, sGlosa, 13
    nMovNro = clsMant.GetnMovNro(sMovNro)
    clsMovs.AgregaMovCap nMovNro, nOperacion, sCuenta, 0, 0, 0
'End By
'MIOL 20121109, SEGUN RFC098-2012-B ******************************************
'clsMant.ActualizaDatosCuenta sCuenta, bOrdPag, nFirmas, nTipoCuenta, sInstitucion, sCuentaAboInt, nFirmasMin, sAlias, nTipoPrograma, pnPorRetAdiCTS, sMovNro, nSubasta COMENTADO POR RIRO20131102
'clsMant.ActualizaDatosCuenta sCuenta, bOrdPag, nFirmas, nTipoCuenta, sInstitucion, sCuentaAboInt, nFirmasMin, sAlias, nTipoPrograma, pnPorRetAdiCTS, sMovNro
'END MIOL ********************************************************************

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
clsMant.ActualizaDatosCuenta sCuenta, bOrdPag, nFirmas, nTipoCuenta, sInstitucion, sCuentaAboInt, nFirmasMin, sAlias, nTipoPrograma, pnPorRetAdiCTS, sMovNro, nSubasta, sReglas

clsMant.EliminaProductoPersona sCuenta

' *** AGREGADO POR RIRO 20131102 SEGUN "CAMBIOS EN PODERES"
Dim letras() As String
Dim nContar As Integer
nContar = 0
letras = Split(sGrupos, "-")
' *** END RIRO

Do While Not rsRel.EOF
    sPersona = rsRel("Codigo")
    nRelacion = CLng(Trim(Right(rsRel("Relacion"), 4)))
    COBLIGATORIO = Left(rsRel("Obligatorio"), 1)
    ' RIRO 20131102 SEGUN "CAMBIOS EN PODERES", SE AGREGO PARAMETRO "letras(nContar)"
    clsMant.AgregaProductoPersona sCuenta, sPersona, nRelacion, COBLIGATORIO, letras(nContar)
    clsMant.AgregaBitacoraPPer sMovNro, sCuenta, sPersona, nRelacion, COBLIGATORIO, letras(nContar)
    nContar = nContar + 1 ' RIRO 20131102 SEGUN "CAMBIOS EN PODERES", SE AGREGO PARAMETRO "letras(nContar)"

    rsRel.MoveNext
Loop

'JUEZ 20130909 ***************************************************************************
If nProd = gCapPlazoFijo And bCtaAboIntInact Then
    clsMovs.AgregaMovCap nMovNro, gAhoEstInacAct, sCuentaAboInt, nSaldCntCtaAboIntInact, nSaldDispCtaAboIntInact, nSaldCntCtaAboIntInact
    clsMovs.AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuentaAboInt, gConcCapital, nSaldCntCtaAboIntInact
    clsMovs.ActualizaEstadoCuenta sCuentaAboInt, gCapEstActiva, CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4)), sMovNro
    clsMovs.ActualizaInactivaAct sCuentaAboInt
End If
'END JUEZ ********************************************************************************

clsMant.dbCmact.CommitTrans
Set clsMant = Nothing
ActualizaDatosCuenta = True
Exit Function
ErrGraba:
    If bTrans Then
        clsMant.dbCmact.RollbackTrans
    End If
    Set clsMant = Nothing
    ActualizaDatosCuenta = False
'    MsgBox Err.Description, vbExclamation, "Error"
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function GetCuentasPersona(ByVal sPersona As String, Optional nProd As Producto, _
        Optional bActivas As Boolean = False, Optional bBloqueadas As Boolean = False, _
        Optional nMoneda As Moneda, Optional bOrdPag As Boolean = False, _
        Optional sAgencia As String = "", Optional sPrograma As String = "", _
        Optional bCtasDesemb = False, _
        Optional psCadPrd As String = "", Optional ByVal pbBloqueadaTotal As Boolean = False) As ADODB.Recordset '** Juez 20120723 Se agregó bCtasDesemb
'Modificado por BRGO 20111111 - Se agregó el parámetro sPrograma
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetCuentasPersona = clsMant.GetCuentasPersona(sPersona, nProd, bActivas, bBloqueadas, nMoneda, bOrdPag, sAgencia, sPrograma, bCtasDesemb, , psCadPrd, pbBloqueadaTotal)
Set clsMant = Nothing
End Function

'** BRGO 14/06/2011 ********************************************
Public Function GetCuentaAhorroTitularesPF(ByVal nTpoCta As Integer, ByVal cPersCod As String, _
        ByVal nMoneda As Moneda) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetCuentaAhorroTitularesPF = clsMant.GetCuentaAhorroTitularesPF(nTpoCta, cPersCod, nMoneda)
Set clsMant = Nothing
End Function
'**************************************************************
Public Function GetCuentasAhorro(ByVal bTodas As Boolean, Optional cAgeCod As String = "", Optional psCctacod As String = "") As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetCuentasAhorro = clsMant.GetCuentasAhorro(bTodas, cAgeCod, psCctacod)
Set clsMant = Nothing
End Function

Public Sub ActualizaCtaDescInact(ByVal psCuenta As String, ByVal bdescinactiva As Boolean, ByVal psMovNro As String)
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
       Call clsMant.ActualizaCtaDescInact(psCuenta, bdescinactiva, psMovNro)
       Set clsMant = Nothing
End Sub

Public Function GetCapBloqueos(ByVal sCuenta As String, ByVal nTipoBloqueo As CaptacTipoBloqueo, ByVal nConstante As ConstanteCabecera) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetCapBloqueos = clsMant.GetCapBloqueos(sCuenta, nTipoBloqueo, nConstante)
Set clsMant = Nothing
End Function

Public Sub EliminaProductoPersona(ByVal sCuenta As String, ByVal sPersona As String)
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
clsMant.EliminaProductoPersona sCuenta
Set clsMant = Nothing
End Sub

Public Function AgregaTarjeta(ByVal sTarjeta As String, ByVal sClave As String, ByVal sMovNro As String, _
        ByVal sPersona As String, ByVal sCuenta As String, ByVal dCaduca As Date, ByRef psmensaje As String) As Boolean
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales

If Not clsMant.ExisteTarjeta(sTarjeta) Then
    clsMant.dbCmact.BeginTrans
    clsMant.AgregaTarjeta sTarjeta, sClave, dCaduca, sMovNro
    clsMant.AgregaCuentaTarj sTarjeta, sCuenta, sPersona
    clsMant.dbCmact.CommitTrans
    AgregaTarjeta = True
Else
'    MsgBox "Tarjeta ya registrada al cliente. Ingrese una tarjeta nueva", vbInformation, "Aviso"
    psmensaje = "Tarjeta ya registrada al cliente. Ingrese una tarjeta nueva"
    AgregaTarjeta = False
End If
Set clsMant = Nothing
End Function

Public Function AgregaTarjetaLOTE(ByVal sTarjeta As String, ByVal sClave As String, ByVal sMovNro As String, _
         ByVal rCuentas As ADODB.Recordset, ByVal dCaduca As Date, ByVal sPersCod As String, Optional ByRef psMensajeValidacion As String = "") As Boolean
         
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales, rsPers As ADODB.Recordset
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales

If Not clsMant.ExisteTarjeta(sTarjeta) Then
    
    If Not clsMant.EsTarjetaEmitida(sTarjeta) Then
        'MsgBox "Esta Tarjeta no ha sido emitida por la CMAC ICA", vbOKOnly + vbExclamation, "AVISO IMPORTANTE"
        psMensajeValidacion = "AVISO IMPORTANTE: Esta Tarjeta no ha sido emitida por la CMAC ICA"
        AgregaTarjetaLOTE = False
        Exit Function
    End If

    clsMant.dbCmact.BeginTrans
    
     clsMant.AgregaTarjeta sTarjeta, sClave, dCaduca, sMovNro
        
     While Not rCuentas.EOF
        'Set rsPers = New ADODB.Recordset
        'Set rsPers = clsmant.GetPersonasCuentaNA(rCuentas.Fields("Cuenta"), "", CInt(rCuentas.Fields("Personeria")))
         
        ' While Not rsPers.EOF
           ' clsmant.AgregaCuentaTarj sTarjeta, rCuentas.Fields("Cuenta"), rsPers.Fields("Codigo")
           ' rsPers.MoveNext
        ' Wend
         clsMant.AgregaCuentaTarj sTarjeta, rCuentas.Fields("Cuenta"), sPersCod
         
       ' Set rsPers = Nothing
        rCuentas.MoveNext
     Wend
    
    clsMant.dbCmact.CommitTrans
    AgregaTarjetaLOTE = True
Else
    ' MsgBox "Tarjeta ya registrada al cliente. Ingrese una tarjeta nueva", vbInformation, "Aviso"
    psMensajeValidacion = "Tarjeta ya registrada al cliente. Ingrese una tarjeta nueva"
    AgregaTarjetaLOTE = False
End If

Set clsMant = Nothing
End Function


Public Sub AgregaProductoPersona(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
clsMant.AgregaProductoPersona sCuenta, sPersona, nRelacion
Set clsMant = Nothing
End Sub

'ALPA 20091118*************************************************
Public Function ObtenerTREA(ByVal psCodTPro As String, ByVal nSaldo As Currency, ByVal nTEA As Currency, Optional nPlazo As Integer = 0, Optional nCostoMan As Currency = 0) As Currency
    Dim nSaldoIni As Currency
    Dim nSaldoFin As Currency
    Dim nMontInteres As Currency
    Dim nTEM As Double
    Dim nTREA As Double
    Dim Ix As Integer
    If nSaldo > 0 Then
        If nTEA > 0 Then
            If psCodTPro = gCapPlazoFijo Then
                nSaldoIni = nSaldo
                nSaldoFin = nSaldo
                nMontInteres = ((1 + nTEA) ^ (nPlazo / 360) - 1) * nSaldo
                nSaldoFin = nSaldoFin + nMontInteres
                nTREA = Format$(((nSaldoFin / nSaldoIni) ^ (360 / nPlazo) - 1), "#0.00")
            Else
                nSaldoIni = nSaldo
                nSaldoFin = nSaldo
'                nTEM = Format$(((1 + nTEA) ^ (1 / 12) - 1), "#0.000000")
'                For Ix = 1 To 12
'                   nSaldoFin = nSaldoFin + nSaldoFin * nTEM + nCostoMan
                   nSaldoFin = nSaldoFin + nSaldoFin * (nTEA / 100) - (nCostoMan * 12) 'APRI20190109 ERS077-2018
'                Next Ix
'                nTREA = Format$(((nSaldoFin / nSaldoIni) ^ (12 / 12) - 1), "#0.00")
                nTREA = Format$(((nSaldoFin / nSaldoIni) ^ (12 / 12) - 1) * 100, "#0.00") 'APRI20190109 ERS077-2018
             End If
        End If
    End If
    ObtenerTREA = nTREA
End Function

'**************************************************************
Public Function GeneraRegistroFirmas(ByVal sCuenta As String, sTipCuenta As String, _
    ByVal dFecApe As Date, ByVal bOrdPag As Boolean, ByVal rsProdPers As ADODB.Recordset, _
    ByVal sAgencia As String, ByVal dFecSis As Date, ByVal sUser As String, Optional ByVal nTpoPrograma As Integer = 0, Optional nCostoMan As Currency = 0) As String
         
    Dim sTipPer As String
    Dim sCodCli As String
    Dim sNomCli As String
    Dim sDirCli As String
    Dim sDocCli As String
    Dim sTelCli As String
    Dim sNomSec As String
    Dim sRazSoc As String
    Dim sDirEmp As String
    Dim sTelEmp As String
    Dim sDocEmp As String
    Dim sNumFte As String
    Dim sNomAge As String
    Dim sDesPro As String * 25, sDesTip As String * 25
    Dim sRelCta As String
    Dim sClte01 As String
    Dim sClte02 As String
    Dim sClte03 As String
    Dim sClte04 As String
    Dim nIngres As Currency
    Dim dFchNac As Date
    Dim nLinCab As Integer
    Dim nCntCli As Integer
    Dim i As Integer, nCarLin As Long
    Dim sMoneda As String
    Dim sCad As String
    Dim sNumPag As String
    Dim sOcupCli As String
    Dim sPromotor As String
    Dim sInstitucion As String
    
    
    sCad = ""
    
    Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
    Dim nItem As Integer
    Dim nPagina As Integer
    
    Dim sL1 As String * 74
    Dim sL2 As String * 74
    Dim sL3 As String * 74
    Dim sL4 As String * 74
    Dim sL5 As String * 74
    Dim sTitRp1 As String, sTitRp2 As String
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oImpreCap As New NCOMCaptaImpresion
    Dim oFunLetras As New COMFunciones.FCOMCadenas
    Dim oImpre As New COMFunciones.FCOMImpresion
    Dim oFunCad As New COMFunciones.FCOMCadenas
    
    'ARCV 28-10-2006
    Dim dFecnacimiento As Date
    Dim sEstadoCivil As String
    Dim sL6 As String * 74
    Dim sL7 As String * 74
    Dim sL8 As String * 74
'----------
    'APRI2018 ERS036-2017
    Dim sL9 As String * 74
    Dim sEnvioEstadoCta As String
    'END APRI
    
    Dim rsPers As ADODB.Recordset
    nTamMayor = 70
    nTamMenor = 38
    nTamAdd = 22
    nItem = 12
    nPagina = 1
    sNomAge = sAgencia
    nCarLin = nTamMayor + nTamMenor + 2
    
    '------------ AVMM 18-01-2007 ---------------------
    'Procediemitno para agregar la TNA/TEA al Reporte
     Dim oConect As COMConecta.DCOMConecta
     Dim rs As New ADODB.Recordset
     Dim rs1 As New ADODB.Recordset
     Dim nConvierteTNAaTEA As Double
     Dim nTasa As Double
     Dim nPlazo As Integer
     Dim sql As String
     Dim nSaldo As Double
     Dim nTREA As Double
     Dim nSaldoIn As Currency
     
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion Then
       sql = "select nTasaInteres,nSaldo From Producto Where cCtaCod='" & sCuenta & "'"
       Set rs = oConect.CargaRecordSet(sql)
       If Not (rs.EOF And rs.BOF) Then
           nTasa = rs!ntasainteres
           nSaldoIn = rs!nSaldo
           If CLng(Mid$(sCuenta, 6, 3)) = gCapPlazoFijo Then
               sql = "select nPlazo from CaptacPlazoFijo where cCtaCod='" & sCuenta & "'"
               Set rs1 = oConect.CargaRecordSet(sql)
               nPlazo = rs1!nPlazo
               'nConvierteTNAaTEA = Format$(((((((nTasa / 12) / 3000) + 1) ^ nPlazo) - 1) * 100), "#0.00")
               nConvierteTNAaTEA = Format$(((1 + nTasa / 36000) ^ 360 - 1) * 100, "#0.00")
               'ALPA 20091118***************************************
               nTREA = ObtenerTREA(Mid$(sCuenta, 6, 3), IIf(nSaldoIn = 0, 10, nSaldoIn), nConvierteTNAaTEA, nPlazo, nCostoMan)
               '****************************************************
           Else
               nConvierteTNAaTEA = Format$(((1 + nTasa / 36000) ^ 360 - 1) * 100, "#0.00")
               'ALPA 20091118***************************************
               nTREA = ObtenerTREA(Mid$(sCuenta, 6, 3), IIf(nSaldoIn = 0, 10, nSaldoIn), nConvierteTNAaTEA, , nCostoMan)
               '****************************************************
           End If
           nSaldo = rs!nSaldo
       Else
           nTasa = 0
           nConvierteTNAaTEA = 0
           nSaldo = 0
       End If
    End If
    
    'By Capi Agosto 2007(Acta 014-2007)
    If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
        Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        rsProdPers.MoveFirst
        Do While Not rsProdPers.EOF
            If CInt(Trim(Right(rsProdPers("Relacion"), 2))) = gCapRelPersPromotor Then
                sCodCli = rsProdPers("Codigo")
                Set rsPers = clsMant.GetDatosPersona(sCodCli)
                sPromotor = rsPers("Nombre")
            Else
                Set rsPers = clsMant.GetDatosCuentaCTS(sCuenta)
                sInstitucion = rsPers("cInstitucion")
            End If
            rsProdPers.MoveNext
        Loop
        Set clsMant = Nothing
        Set rsPers = Nothing
    End If
    
    
    
    'End
    '--------------------------------------------------
    Select Case CLng(Mid$(sCuenta, 6, 3))
        Case gCapAhorros
            sDesPro = "AHORRO" & IIf(bOrdPag, " CON ORDEN DE PAGO", "")
            If nTpoPrograma = 0 Then
                sDesPro = Trim(sDesPro) & " CORRIENTE"
            ElseIf nTpoPrograma = 1 Then
                sDesPro = Trim(sDesPro) & " ÑAÑITO"
'Comentado por BRGO 20111221 ************************
'            ElseIf nTpoPrograma = 2 Then
'                'MAVM 20110307 ***
'                'sDesPro = Trim(sDesPro) & " PANDERITO"
'                sDesPro = Trim(sDesPro) & " DIARIO"
'                '***
'            ElseIf nTpoPrograma = 3 Then
'                '*** PEAC 20090721
'                'sDesPro = Trim(sDesPro) & " PANDERO"
'                sDesPro = Trim(sDesPro) & " POCO A POCO AHORRO"
            ElseIf nTpoPrograma = 4 Then
                sDesPro = Trim(sDesPro) & " DESTINO"
            ElseIf nTpoPrograma = 5 Then
                sDesPro = " CUENTA SOÑADA"
            ElseIf nTpoPrograma = 6 Then
                sDesPro = " CAJA SUELDO"
            ElseIf nTpoPrograma = 7 Then
                sDesPro = Trim(sDesPro) & " ECOTAXI"
            '***Agregado por ELRO el 20130131, segun TI-ERS020-2013
            ElseIf nTpoPrograma = 8 Then
                sDesPro = Trim(sDesPro) & " CONVENIO"
            '***Fin Agregado por ELRO el 20130131******************
            End If
        Case gCapPlazoFijo
            If nTpoPrograma = 0 Then
                sDesPro = "PLAZO FIJO CLÁSICO"
            ElseIf nTpoPrograma = 1 Then
                sDesPro = "DEPÓSITO PLAZO PREMIUM"
            ElseIf nTpoPrograma = 2 Then
                sDesPro = "AHORRO DIARIO"
            ElseIf nTpoPrograma = 3 Then
                sDesPro = "POCO A POCO AHORRO"
            End If
        Case gCapCTS
            sDesPro = "C T S"
            If nTpoPrograma = 0 Then
                sDesPro = Trim(sDesPro) & " + CAJA SUELDO" '**JUEZ 20120216
            End If
    End Select
'**** END BRGO ***********************************************
    sDesTip = sTipCuenta
    sCad = ""
    
    '*** PEAC 20090727 - SE AGREGO EL EUROS
    '''sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", "NUEVOS SOLES", IIf(Mid$(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) 'marg ers044-2016
    sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), IIf(Mid$(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) 'marg ers044-2016
    
    sTitRp1 = "R E G I S T R O   D E   F I R M A S"
    sTitRp2 = "CUENTA N° " & Mid(sCuenta, 1, 3) & "-" & Mid(sCuenta, 4, 2) & "-" & Mid(sCuenta, 6, 3) & "-" & Mid(sCuenta, 9, 10)
    sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "PRODUCTO           : " & Trim(sDesPro) & Space(25 - Len(Trim(sDesPro))) & "AGENCIA : " & Trim(sAgencia) & Space(25 - Len(Trim(sAgencia))) & "Usuario     : " & Trim(sUser) & oImp.gPrnSaltoLinea
    'By Capi Agosto 2007(Acta 014-2007)
    If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
        sCad = sCad & "TIPO DE CUENTA     : " & Trim(sDesTip) & Space(25 - Len(Trim(sDesTip))) & "MONEDA  : " & Trim(sMoneda) & Space(25 - Len(Trim(sMoneda))) & "ENTIDAD     : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
        'ALPA 20091118*******************************
        'sCad = sCad & "MONTO DE APERTURA  : " & Trim(nSaldo) & Space(25 - Len(Trim(nSaldo))) & "T.N.A   : " & Trim(nTasa) & "%" & Space(25 - Len(Trim(nTasa)) - 1) & "PROMOTOR(A) : " & Trim(sPromotor) & oImp.gPrnSaltoLinea
        sCad = sCad & "MONTO DE APERTURA  : " & Trim(nSaldo) & Space(25 - Len(Trim(nSaldo))) & "T.R.E.A.: " & Trim(nTREA) & "%" & Space(25 - Len(Trim(nTasa)) - 1) & "PROMOTOR(A) : " & Trim(sPromotor) & oImp.gPrnSaltoLinea
        '********************************************
        
    Else
        sCad = sCad & "TIPO DE CUENTA     : " & Trim(sDesTip) & Space(25 - Len(Trim(sDesTip))) & "MONEDA  : " & sMoneda & oImp.gPrnSaltoLinea
        'ALPA 20091118******************************************
        'sCad = sCad & "MONTO DE APERTURA  : " & Trim(nSaldo) & Space(25 - Len(Trim(nSaldo))) & "T.N.A   : " & Trim(nTasa) & "%" & oImp.gPrnSaltoLinea
        sCad = sCad & "MONTO DE APERTURA  : " & Trim(nSaldo) & Space(25 - Len(Trim(nSaldo))) & "T.R.E.A.: " & Trim(nTREA) & "%" & oImp.gPrnSaltoLinea
        '*******************************************************
    End If
    'End
    sCad = sCad & "FECHA DE APERTURA  : " & Format$(dFecApe, "dd mmm yyyy") & Space(25 - Len(Trim(Format$(dFecApe, "dd mmm yyyy")))) & "T.E.A   : " & Trim(nConvierteTNAaTEA) & "%" & Space(28) & "FIRMA" & Space(28) & "HUELLA" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    nLinCab = 5
    sCodCli = ""

    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsPers = New ADODB.Recordset
    rsProdPers.MoveFirst
    Do While Not rsProdPers.EOF
        
        '-------------GRVA 17-07-2007 -----------------
        '----------Solo muesra Titulares---------------
        If CInt(Trim(Right(rsProdPers("Relacion"), 2))) <> gCapRelPersPromotor Then
            sCodCli = rsProdPers("Codigo")
            Set rsPers = clsMant.GetDatosPersona(rsProdPers("Codigo"))
            nCntCli = nCntCli + 1
            sNomCli = oImpre.ImpreCarEsp(oFunLetras.PstaNombre(rsPers("Nombre"), False))
            sRelCta = UCase(Trim(Left(rsProdPers("Relacion"), 30)))
            sDirCli = rsPers("Direccion") & ""
            sDocCli = rsPers("ID N°")
            sTelCli = rsPers("Fono")
            sNomSec = rsPers("Zona")
            sOcupCli = rsPers("cCIIUdescripcion")
            dFecnacimiento = Format(rsPers("dPersNacCreac"), "dd-mm-yyyy")
            sEstadoCivil = rsPers("cEstadoCivil")
            
            'APRI20180420 ERS036-2017
            sql = "select * From ProductoEnvioEstadoCta Where nEstado = 1 and cCtaCod='" & sCuenta & "' and cPersCod ='" & sCodCli & "'"
            Set rs = oConect.CargaRecordSet(sql)
            If Not (rs.EOF And rs.BOF) Then
                sEnvioEstadoCta = "SI"
            Else
                sEnvioEstadoCta = "NO"
            End If
            Set rs = Nothing
            'END APRI
            
            sL1 = "CODIGO   " & oImpre.FillText(Trim(sCodCli), 13, " ") & "  DNI/LE/CE   " & Trim(sDocCli)
            sL2 = Left(oImpre.ImpreCarEsp(oFunCad.PstaNombre(sNomCli, False)), nTamMenor)
            sL3 = Trim(sDirCli) & Space(3) & Trim(sNomSec)
            sL4 = "TELEFONO   " & sTelCli
            sL5 = "RC   " & sRelCta
            
            sL6 = "FECHA NACIMIENTO   " & dFecnacimiento
            sL7 = "ESTADO CIVIL       " & sEstadoCivil
            sL8 = "OCUPACION          " & sOcupCli
            sL9 = "ENVIO ESTADO DE CUENTA  " & sEnvioEstadoCta
            If nItem > 48 Then
                nPagina = nPagina + 1
                sCad = sCad & Chr(12)
                sCad = sCad & oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
                sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
                sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & Space(Len(sAgencia) - 9) & "T.N.A :" & nTasa & "%" & "   Usuario :" & sUser & oImp.gPrnSaltoLinea
                sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & Space(Len(sAgencia) - 6) & "T.E.A :" & nConvierteTNAaTEA & "%" & oImp.gPrnSaltoLinea
                sCad = sCad & "MONTO DE APERTURA  : " & nSaldo & oImp.gPrnSaltoLinea
                sCad = sCad & "FECHA DE APERTURA  : " & Format$(dFecApe, "dd mmm yyyy") & Space(55) & "FIRMA" & Space(25) & "HUELLA" & oImp.gPrnSaltoLinea
                sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
                nItem = 12
            End If
             'nTamMenor
             sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL2 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL3 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL4 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL5 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL6 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL7 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL8 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL9 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea 'APRI20180420 ERS036-2017
             sL1 = ""
             sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & oImpre.CentrarCadena(".......................................", nTamAdd + 15) & " ¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
             sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamAdd + 15) & "¦" & Space(nTamAdd + 1) & "¦") & oImp.gPrnSaltoLinea
             'sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-" & "+") & "+" & String(nTamAdd, "-") & "+" & oImp.gPrnSaltoLinea
             sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
             rsPers.Close
             Set rsPers = Nothing
        End If
        nItem = nItem + 12
        rsProdPers.MoveNext
    Loop
    Dim arreglocad As String
    If nItem > 48 Then
        nPagina = nPagina + 1
        sCad = sCad & oImp.gPrnSaltoPagina
        sCad = sCad & "" 'oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(Str(sNumPag)), "", dFecSis) & oImp.gPrnSaltoLinea
        arreglocad = oImpreCap.CabeRepoCaptac(" ", " ", 1, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis, oImp.gPrnSaltoLinea)
        sCad = sCad & arreglocad & oImp.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & Space(Len(sAgencia) - 9) & "T.N.A :" & nTasa & "%" & Space(8) & "   Usuario :" & sUser & oImp.gPrnSaltoLinea
        sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & Space(Len(sAgencia) - 6) & "T.E.A :" & nConvierteTNAaTEA & "%" & oImp.gPrnSaltoLinea
        sCad = sCad & "MONTO DE APERTURA  : " & nSaldo & oImp.gPrnSaltoLinea
        sCad = sCad & "FECHA DE APERTURA  : " & Format$(dFecApe, "dd mmm yyyy") & Space(55) & "FIRMA" & Space(25) & "HUELLA" & oImp.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
        nItem = 12
    End If

    sL1 = oImpre.CentrarCadena("INFORME DE ASESORIA LEGAL", nTamMayor)
    sL2 = oImpre.CentrarCadena(" FORMAS DE INTERVENCION", nTamMayor)
    sL3 = oImpre.CentrarCadena("[ ]ABRIR   [ ]CERRAR   [ ]RETIRAR", nTamMayor)
    sL4 = ""
    
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.CentrarCadena(sL1, nTamMayor + nTamMenor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.CentrarCadena(sL2, nTamMayor + nTamMenor + nTamAdd + 4) & "  ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(oImpre.CentrarCadena(sL3, nTamMayor + nTamMenor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.CentrarCadena(String(90, "_"), nTamMayor + nTamMenor + nTamAdd + 4) & "  ¦" & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.CentrarCadena(String(90, "_"), nTamMayor + nTamMenor + nTamAdd + 4) & "  ¦" & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.CentrarCadena(String(90, "_"), nTamMayor + nTamMenor + nTamAdd + 4) & "  ¦" & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.CentrarCadena(String(90, "_"), nTamMayor + nTamMenor + nTamAdd + 4) & "  ¦" & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.CentrarCadena(String(90, "_"), nTamMayor + nTamMenor + nTamAdd + 4) & "  ¦" & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.CentrarCadena(String(90, "_"), nTamMayor + nTamMenor + nTamAdd + 4) & "  ¦" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+" & oImp.gPrnSaltoLinea
    
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("_____________________________               _____________________________             _____________________________", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("      ASESOR DEL CLIENTE                      SUPERVISOR DE OPERACIONES                       ASESOR LEGAL         ", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-")
    sCad = sCad & oImp.gPrnSaltoPagina
    '----------COMENTADO POR CMAC-CUSCO -- AVMM -- 18/04/2006---------------
    'sCad = sCad & RegistroFirmasCANC(sCuenta, bOrdPag, rsProdPers, dFecSis)
    '-----------------------------------------------------------------------
    GeneraRegistroFirmas = sCad
End Function
'By Capi Acta 014-2007
Public Function GeneraRegistroLibretas(ByVal sCuenta As String, sTipCuenta As String, _
    ByVal dFecApe As Date, ByVal bOrdPag As Boolean, ByVal rsProdPers As ADODB.Recordset, _
    ByVal sAgencia As String, ByVal dFecSis As Date, ByVal sUser As String, Optional ByVal nTpoPrograma As Integer = 0) As String
          
    Dim sTipPer As String
    Dim sCodCli As String
    Dim sNomCli As String
    Dim sDirCli As String
    Dim sDocCli As String
    Dim sTelCli As String
    Dim sNomSec As String
    Dim sRazSoc As String
    Dim sDirEmp As String
    Dim sTelEmp As String
    Dim sDocEmp As String
    Dim sNumFte As String
    Dim sNomAge As String
    Dim sDesPro As String * 25, sDesTip As String * 25
    Dim sRelCta As String
    Dim sClte01 As String
    Dim sClte02 As String
    Dim sClte03 As String
    Dim sClte04 As String
    Dim nIngres As Currency
    Dim dFchNac As Date
    Dim nLinCab As Integer
    Dim nCntCli As Integer
    Dim i As Integer, nCarLin As Long
    Dim sMoneda As String
    Dim sCad As String
    Dim sNumPag As String
    Dim sOcupCli As String
    Dim sPromotor As String
    Dim sInstitucion As String
    
    
    sCad = ""
    
    Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
    Dim nItem As Integer
    Dim nPagina As Integer
    
    Dim sL1 As String * 74
    Dim sL2 As String * 74
    Dim sL3 As String * 74
    Dim sL4 As String * 74
    Dim sL5 As String * 74
    Dim sTitRp1 As String, sTitRp2 As String
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oImpreCap As New NCOMCaptaImpresion
    Dim oFunLetras As New COMFunciones.FCOMCadenas
    Dim oImpre As New COMFunciones.FCOMImpresion
    Dim oFunCad As New COMFunciones.FCOMCadenas
    
    'ARCV 28-10-2006
    Dim dFecnacimiento As Date
    Dim sEstadoCivil As String
    Dim sL6 As String * 74
    Dim sL7 As String * 74
    Dim sL8 As String * 74
'----------

    Dim rsPers As ADODB.Recordset
    nTamMayor = 70
    nTamMenor = 38
    nTamAdd = 22
    nItem = 12
    nPagina = 1
    sNomAge = sAgencia
    nCarLin = nTamMayor + nTamMenor + 2
    
   '*** modificado por BRGO 20111221 *************************
    Select Case CLng(Mid$(sCuenta, 6, 3))
        Case gCapAhorros
            sDesPro = "AHORRO" & IIf(bOrdPag, " CON ORDEN DE PAGO", "")
            If nTpoPrograma = 0 Then
                sDesPro = Trim(sDesPro) & " CORRIENTE"
            ElseIf nTpoPrograma = 1 Then
                sDesPro = Trim(sDesPro) & " ÑAÑITO"
            ElseIf nTpoPrograma = 4 Then
                sDesPro = Trim(sDesPro) & " DESTINO"
            ElseIf nTpoPrograma = 5 Then
                sDesPro = " CUENTA SOÑADA"
            ElseIf nTpoPrograma = 6 Then
                sDesPro = " CAJA SUELDO"
            ElseIf nTpoPrograma = 7 Then
                sDesPro = Trim(sDesPro) & " ECOTAXI"
            End If
        Case gCapPlazoFijo
            If nTpoPrograma = 0 Then
                sDesPro = "PLAZO FIJO CLÁSICO"
            ElseIf nTpoPrograma = 1 Then
                sDesPro = "DEPÓSITO PLAZO PREMIUM"
            ElseIf nTpoPrograma = 2 Then
                sDesPro = "AHORRO DIARIO"
            ElseIf nTpoPrograma = 3 Then
                sDesPro = "POCO A POCO AHORRO"
            End If
        Case gCapCTS
            sDesPro = "C T S"
            If nTpoPrograma = 0 Then
                sDesPro = Trim(sDesPro) & " + CAJA SUELDO" '**JUEZ 20120216
            End If
    End Select
    '*** END BRGO **************************************************
    
    sDesTip = sTipCuenta
    sCad = ""
    '*** PEAC 20090727 - SE AGREGO EL EUROS
    '''sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", "NUEVOS SOLES", IIf(Mid$(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) 'marg ers044-2016
    sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), IIf(Mid$(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) 'marg ers044-2016
    
    sTitRp1 = "...DATOS DEL CLIENTE..."
    sTitRp2 = "Cuenta N°... " & Mid(sCuenta, 1, 3) & "-" & Mid(sCuenta, 4, 2) & "-" & Mid(sCuenta, 6, 3) & "-" & Mid(sCuenta, 9, 10)
    sCad = sCad & sTitRp1 & oImp.gPrnSaltoLinea
    sCad = sCad & String(45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnNegritaON & sTitRp2 & oImp.gPrnNegritaOFF & oImp.gPrnSaltoLinea
    sCad = sCad & "Tp/Md " & Trim(sDesPro) & Space(5) & Trim(sDesTip) & oImp.gPrnSaltoLinea
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsPers = New ADODB.Recordset
    rsProdPers.MoveFirst
    Do While Not rsProdPers.EOF
        If CInt(Trim(Right(rsProdPers("Relacion"), 2))) <> gCapRelPersPromotor Then
            Set rsPers = clsMant.GetDatosPersona(rsProdPers("Codigo"))
            sNomCli = rsPers("Nombre")
            sDocCli = rsPers("ID N°")
            sCad = sCad & Trim(sNomCli) + Space(2) + Trim(sDocCli) & oImp.gPrnSaltoLinea
            rsPers.Close
            Set rsPers = Nothing
        End If
        nItem = nItem + 12
        rsProdPers.MoveNext
    Loop
    sCad = sCad & Trim(sAgencia) & oImp.gPrnSaltoLinea
    sCad = sCad & Format$(dFecApe, "dd mmm yyyy") & Space(20) & Trim(sUser) & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    GeneraRegistroLibretas = sCad
End Function





Public Function RegistroFirmasCANC(ByVal sCuenta As String, ByVal bOrdPag As Boolean, ByVal rsProdPers As ADODB.Recordset, Optional pdFecSis As Date = CDate("01/01/1900")) As String
    
Dim sTipPer As String
Dim sCodCli As String
Dim sNomCli As String
Dim sDirCli As String
Dim sDocCli As String
Dim sTelCli As String
Dim sNomSec As String
Dim sRazSoc As String
Dim sDirEmp As String
Dim sTelEmp As String
Dim sDocEmp As String
Dim sNumFte As String
Dim sNomAge As String
Dim sDesPro As String * 25, sDesTip As String * 25
Dim sRelCta As String
Dim sClte01 As String
Dim sClte02 As String
Dim sClte03 As String
Dim sClte04 As String
Dim nIngres As Currency
Dim dFchNac As Date
Dim nLinCab As Integer
Dim nCntCli As Integer
Dim i As Integer, nCarLin As Long
Dim sMoneda As String
Dim sCad As String
Dim sNumPag As String
sCad = ""
    
Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
Dim nItem As Integer
Dim nPagina As Integer

Dim sL1 As String * 74
Dim sL2 As String * 74
Dim sL3 As String * 74
Dim sL4 As String * 74
Dim sL5 As String * 74
Dim sTitRp1 As String, sTitRp2 As String
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oImpre As New COMFunciones.FCOMImpresion
Dim oImpreCapta As New NCOMCaptaImpresion
Dim oFunCad As New COMFunciones.FCOMCadenas

Dim rsPers As ADODB.Recordset
nTamMayor = 70
nTamMenor = 38
nTamAdd = 22
nItem = 12
nPagina = 1
'sNomAge = sAgencia
nCarLin = nTamMayor + nTamMenor - 10

Select Case CLng(Mid$(sCuenta, 6, 3))
    Case gCapAhorros
        sDesPro = "AHORRO" & IIf(bOrdPag, " CON ORDEN DE PAGO", "")
    Case gCapPlazoFijo
        sDesPro = "PLAZO FIJO"
    Case gCapCTS
        sDesPro = "C T S"
End Select

'sDesTip = sTipCuenta

If pdFecSis = "01/01/1900" Then
   pdFecSis = Date
End If

'''sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", "NUEVOS SOLES", "DOLARES") 'marg ers044-2016
sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
sTitRp1 = "R E G I S T R O   D E   F I R M A S  -  C A N C E L A C I O N   D E   C U E N T A  "
sTitRp2 = "CUENTA N° " & Mid(sCuenta, 1, 3) & "-" & Mid(sCuenta, 4, 2) & "-" & Mid(sCuenta, 6, 3) & "-" & Mid(sCuenta, 9, 10)
sCad = sCad & oImpreCapta.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), "", pdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & "_______________________" & oImp.gPrnSaltoLinea
sCad = sCad & "FECHA DE CANCELACION  : " & "__________" & Space(55) & "FIRMA" & Space(25) & "HUELLA" & oImp.gPrnSaltoLinea
sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
nLinCab = 5
sCodCli = ""

Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rsPers = New ADODB.Recordset
rsProdPers.MoveFirst
Do While Not rsProdPers.EOF
    sCodCli = rsProdPers("Codigo")
    Set rsPers = clsMant.GetDatosPersona(sCodCli)
    nCntCli = nCntCli + 1
    sNomCli = oImpre.ImpreCarEsp(oFunCad.PstaNombre(rsPers("Nombre"), False))
    sRelCta = UCase(Trim(Left(rsProdPers("Relacion"), 30)))
    sDirCli = rsPers("Direccion") & ""
    sDocCli = rsPers("ID N°")
    sTelCli = rsPers("Fono")
    sNomSec = rsPers("Zona")
        
    sL1 = "CODIGO   " & oImpre.FillText(Trim(sCodCli), 13, " ") & "DNI/LE/CE   " & Trim(sDocCli)
    sL2 = Left(oImpre.ImpreCarEsp(oFunCad.PstaNombre(sNomCli, False)), nTamMenor)
    sL3 = Trim(sDirCli) & Space(3) & Trim(sNomSec)
    sL4 = "TELEFONO   " & sTelCli
    sL5 = "RC   " & sRelCta
        
    If nItem > 48 Then
        nPagina = nPagina + 1
        sCad = sCad & Chr(12)
        sCad = sCad & oImpreCapta.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), "", pdFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor + nTamMenor + 2 + nTamAdd, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & "_______________________" & oImp.gPrnSaltoLinea
        sCad = sCad & "FECHA DE CANCELACION  : " & "__________" & Space(55) & "FIRMA" & Space(25) & "HUELLA" & oImp.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
        nItem = 12
    End If
    
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL2 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL3 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL4 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL5 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    
    sL1 = ""
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & oImpre.CentrarCadena(".........................................", nTamMenor - 1) & " ¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
         
    nItem = nItem + 12
    rsProdPers.MoveNext
    rsPers.Close
    Set rsPers = Nothing
Loop
Dim arreglocad As String
'If nItem > 48 Then
'    nPagina = nPagina + 1
'    sCad = sCad & Chr(12)
'    sCad = sCad & "" 'oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(Str(sNumPag)), "", dFecSis) & oImp.gPrnSaltoLinea
'    arreglocad = oImpreCap.CabeRepoCaptac(" ", " ", 1, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis)
'    sCad = sCad & arreglocad & oImp.gPrnSaltoLinea
'    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
'    sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & oImp.gPrnSaltoLinea
'    sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & oImp.gPrnSaltoLinea
'    sCad = sCad & "FECHA DE CANCELACION  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & oImp.gPrnSaltoLinea
'    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & oImp.gPrnSaltoLinea
'    nItem = 12
'End If
'
'sL1 = CentrarCadena("INFORME DE ASESORIA LEGAL", nTamMayor)
'sL2 = CentrarCadena("FORMAS DE INTERVENCION", nTamMayor)
'sL3 = CentrarCadena("[ ]ABRIR   [ ]CERRAR   [ ]RETIRAR", nTamMayor)
'sL4 = ""
'
'sCad = sCad & ImpreCarEsp(CentrarCadena(sL1, nTamMayor + nTamMenor + nTamAdd) & " ¦") & oImp.gPrnSaltoLinea
'sCad = sCad & ImpreCarEsp(CentrarCadena(sL2, nTamMayor + nTamMenor + nTamAdd) & " ¦") & oImp.gPrnSaltoLinea
'sCad = sCad & ImpreCarEsp(CentrarCadena(sL3, nTamMayor + nTamMenor + nTamAdd) & " ¦") & oImp.gPrnSaltoLinea
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
'sCad = sCad & CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
'sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+" & oImp.gPrnSaltoLinea
'
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena("_____________________________               _____________________________             _____________________________", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
'sCad = sCad & ImpreCarEsp("¦" & CentrarCadena("    AUXILIAR DE AHORROS                       ASISTENTE DE OPERACIONES                       ASESORIA LEGAL        ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMayor + nTamAdd + 43) & "¦") & oImp.gPrnSaltoLinea
sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    
RegistroFirmasCANC = sCad
End Function

Public Function GeneraRegistroFirmasCANCELACION(ByVal sCuenta As String, sTipCuenta As String, _
    ByVal dFecApe As Date, ByVal bOrdPag As Boolean, ByVal rsProdPers As ADODB.Recordset, _
    ByVal sAgencia As String, ByVal dFecSis As Date) As String
    
Dim sTipPer As String
Dim sCodCli As String
Dim sNomCli As String
Dim sDirCli As String
Dim sDocCli As String
Dim sTelCli As String
Dim sNomSec As String
Dim sRazSoc As String
Dim sDirEmp As String
Dim sTelEmp As String
Dim sDocEmp As String
Dim sNumFte As String
Dim sNomAge As String
Dim sDesPro As String * 25, sDesTip As String * 25
Dim sRelCta As String
Dim sClte01 As String
Dim sClte02 As String
Dim sClte03 As String
Dim sClte04 As String
Dim nIngres As Currency
Dim dFchNac As Date
Dim nLinCab As Integer
Dim nCntCli As Integer
Dim i As Integer, nCarLin As Long
Dim sMoneda As String
Dim sCad As String
Dim sNumPag As String
sCad = ""
    
Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
Dim nItem As Integer
Dim nPagina As Integer

Dim sL1 As String * 74
Dim sL2 As String * 74
Dim sL3 As String * 74
Dim sL4 As String * 74
Dim sL5 As String * 74
Dim sTitRp1 As String, sTitRp2 As String
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oImpre As New COMFunciones.FCOMImpresion
Dim oImpreCapta As New NCOMCaptaImpresion
Dim oFunCad As New COMFunciones.FCOMCadenas

Dim rsPers As ADODB.Recordset
nTamMayor = 74
nTamMenor = 40
nTamAdd = 22
nItem = 12
nPagina = 1
sNomAge = sAgencia
nCarLin = nTamMayor + nTamMenor + 2

Select Case CLng(Mid$(sCuenta, 6, 3))
    Case gCapAhorros
        sDesPro = "AHORRO" & IIf(bOrdPag, " CON ORDEN DE PAGO", "")
    Case gCapPlazoFijo
        sDesPro = "PLAZO FIJO"
    Case gCapCTS
        sDesPro = "C T S"
End Select

sDesTip = sTipCuenta

'''sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", "NUEVOS SOLES", "DOLARES") 'marg ers044-2016
sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
sTitRp1 = "R E G I S T R O   D E   F I R M A S"
sTitRp2 = "CUENTA N° " & Mid(sCuenta, 1, 3) & "-" & Mid(sCuenta, 4, 2) & "-" & Mid(sCuenta, 6, 3) & "-" & Mid(sCuenta, 9, 10)
sCad = sCad & oImpreCapta.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
sCad = sCad & String(nTamMayor + nTamMenor + 2 + nTamAdd, "-") & oImp.gPrnSaltoLinea
sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & oImp.gPrnSaltoLinea
sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & oImp.gPrnSaltoLinea
sCad = sCad & "FECHA DE CANCELACION  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & Space(35) & "HUELLA" & oImp.gPrnSaltoLinea
sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & String(nTamAdd, "-") & "+" & oImp.gPrnSaltoLinea
nLinCab = 5
sCodCli = ""

Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rsPers = New ADODB.Recordset
rsProdPers.MoveFirst
Do While Not rsProdPers.EOF
    sCodCli = rsProdPers("Codigo")
    Set rsPers = clsMant.GetDatosPersona(sCodCli)
    nCntCli = nCntCli + 1
    sNomCli = oImpre.ImpreCarEsp(oFunCad.PstaNombre(rsPers("Nombre"), False))
    sRelCta = UCase(Trim(Left(rsProdPers("Relacion"), 30)))
    sDirCli = rsPers("Direccion") & ""
    sDocCli = rsPers("ID N°")
    sTelCli = rsPers("Fono")
    sNomSec = rsPers("Zona")
        
    sL1 = "CODIGO   " & oImpre.FillText(Trim(sCodCli), 13, " ") & "DNI/LE/CE   " & Trim(sDocCli)
    sL2 = Left(oImpre.ImpreCarEsp(oFunCad.PstaNombre(sNomCli, False)), nTamMenor)
    sL3 = Trim(sDirCli) & Space(3) & Trim(sNomSec)
    sL4 = "TELEFONO   " & sTelCli
    sL5 = "RC   " & sRelCta
        
    If nItem > 48 Then
        nPagina = nPagina + 1
        sCad = sCad & oImp.gPrnSaltoPagina
        sCad = sCad & oImpreCapta.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor + nTamMenor + 2 + nTamAdd, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & oImp.gPrnSaltoLinea
        sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & oImp.gPrnSaltoLinea
        sCad = sCad & "FECHA DE CANCELACION  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & Space(35) & "HUELLA" & oImp.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & String(nTamAdd, "-") & "+" & oImp.gPrnSaltoLinea
        nItem = 12
    End If
    
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL2 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL3 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL4 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL5 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    
    sL1 = ""
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & oImpre.CentrarCadena("........................................", nTamMenor - 1) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp(sL1 & "¦" & Space(nTamMenor) & "¦" & Space(nTamAdd) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-" & "+") & "+" & String(nTamAdd, "-") & "+" & oImp.gPrnSaltoLinea
         
    nItem = nItem + 12
    rsProdPers.MoveNext
    rsPers.Close
    Set rsPers = Nothing
Loop
Dim arreglocad As String
If nItem > 48 Then
    nPagina = nPagina + 1
    sCad = sCad & oImp.gPrnSaltoPagina
    sCad = sCad & "" 'oImpreCap.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(Str(sNumPag)), "", dFecSis) & oImp.gPrnSaltoLinea
    arreglocad = oImpreCapta.CabeRepoCaptac(" ", " ", 1, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(sNumPag)), "", dFecSis, oImp.gPrnSaltoLinea)
    sCad = sCad & arreglocad & oImp.gPrnSaltoLinea
    sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "PRODUCTO           : " & sDesPro & "AGENCIA  " & sAgencia & oImp.gPrnSaltoLinea
    sCad = sCad & "TIPO DE CUENTA     : " & sDesTip & "MONEDA   " & sMoneda & oImp.gPrnSaltoLinea
    sCad = sCad & "FECHA DE CANCELACION  : " & Format$(dFecApe, "dd mmm yyyy") & Space(60) & "FIRMA" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor, "-") & "+" & String(nTamMenor, "-") & "+" & oImp.gPrnSaltoLinea
    nItem = 12
End If

sL1 = oImpre.CentrarCadena("INFORME DE ASESORIA LEGAL", nTamMayor)
sL2 = oImpre.CentrarCadena("FORMAS DE INTERVENCION", nTamMayor)
sL3 = oImpre.CentrarCadena("[ ]ABRIR   [ ]CERRAR   [ ]RETIRAR", nTamMayor)
sL4 = ""

sCad = sCad & oImpre.ImpreCarEsp(oImpre.CentrarCadena(sL1, nTamMayor + nTamMenor + nTamAdd) & " ¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.ImpreCarEsp(oImpre.CentrarCadena(sL2, nTamMayor + nTamMenor + nTamAdd) & " ¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.ImpreCarEsp(oImpre.CentrarCadena(sL3, nTamMayor + nTamMenor + nTamAdd) & " ¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.CentrarCadena(String(120, "_"), nTamMayor + nTamMenor + nTamAdd) & " ¦" & oImp.gPrnSaltoLinea
sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+" & oImp.gPrnSaltoLinea

sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("_____________________________               _____________________________             _____________________________", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("    AUXILIAR DE AHORROS                       ASISTENTE DE OPERACIONES                       ASESORIA LEGAL        ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd) & "¦") & oImp.gPrnSaltoLinea
sCad = sCad & String(nTamMayor, "-") & "-" & String(nTamMenor + nTamAdd, "-") & "-+"
    
GeneraRegistroFirmasCANCELACION = sCad
End Function

'ANDE 20171108 SE AGREGÓ EL PARAMETRO pdFecSis
Public Sub ActualizaBloqueos(ByVal sCuenta As String, ByVal rsRet As ADODB.Recordset, _
        ByVal rsTot As ADODB.Recordset, ByVal sMovNro As String, ByVal nEstado As CaptacEstado, Optional ByVal pdFecSis As Date)

Dim nEstadoCta As CaptacEstado
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales, clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim sComentario As String, sFlag As String, sMov As String
Dim nMotRet As CaptacMotBloqueoRet
Dim nMotTot As CaptacMotBloqueoTot
Dim bIniTran As Boolean, bEstado As Boolean, nMovNro As Long
Dim RSTMP As ADODB.Recordset
Dim sOperacion As String

'ANDE 20170726 AGREGADO DE FECHA
VerificarCobroComisionRetencionFondos rsRet, sCuenta, sMovNro, pdFecSis 'JUEZ 20160115
'END ANDE
Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento

Set RSTMP = GetDatosCuenta(sCuenta)

On Error GoTo ErrActBloqueo
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
clsMant.dbCmact.BeginTrans
bIniTran = True
nEstadoCta = gCapEstActiva

If Not rsRet Is Nothing Then
    rsRet.MoveFirst
    Do While Not rsRet.EOF
        nMotRet = CLng(Trim(Right(rsRet("Motivo"), 3)))
        bEstado = IIf(rsRet("Est") = "0", False, True)
        sComentario = rsRet("Comentario")
        sFlag = rsRet("Flag")
        sMov = rsRet("cMovNro")
        
        If sFlag = "N" And bEstado Then
            clsMant.NuevoBloqueoRetiro sCuenta, nMotRet, sComentario, sMovNro
            nEstadoCta = gCapEstBloqRetiro
            
            If nMotRet = gCapMotBlqRetGarantia Or nMotRet = 15 Then
                sOperacion = IIf(Mid(sCuenta, 6, 3) = "232", "201101", IIf(Mid(sCuenta, 6, 3) = "233", "211101", "221101"))
                clsCap.AgregaMov sMovNro, sOperacion, sComentario, gMovEstContabMovContable, gMovFlagVigente
                nMovNro = clsCap.GetnMovNro(sMovNro)
                clsCap.AgregaMovCap nMovNro, sOperacion, sCuenta, RSTMP.Fields("nsaldo"), RSTMP.Fields("nsaldodisp"), RSTMP.Fields("nsaldo")
                clsCap.AgregaMovCapDet nMovNro, sOperacion, sCuenta, gConcCapital, RSTMP.Fields("nsaldo")
            End If
            
        ElseIf sFlag = "A" And Not bEstado Then 'Desbloquea los que ya fueron desmarcados
            clsMant.ActualizaBloqueoRet sCuenta, sComentario, sMovNro, nMotRet, sMov
            clsMant.ActualizaRelacionAhoCred sCuenta, sMovNro, False 'FRHU20140308 RQ14008
            
            If nMotRet = gCapMotBlqRetGarantia Or nMotRet = 15 Then
                sOperacion = IIf(Mid(sCuenta, 6, 3) = "232", "201102", IIf(Mid(sCuenta, 6, 3) = "233", "211102", "221102"))
                clsCap.AgregaMov sMovNro, sOperacion, sComentario, gMovEstContabMovContable, gMovFlagVigente
                nMovNro = clsCap.GetnMovNro(sMovNro)
                clsCap.AgregaMovCap nMovNro, sOperacion, sCuenta, RSTMP.Fields("nsaldo"), RSTMP.Fields("nsaldodisp"), RSTMP.Fields("nsaldo")
                clsCap.AgregaMovCapDet nMovNro, sOperacion, sCuenta, gConcCapital, RSTMP.Fields("nsaldo")
            End If
            
        ElseIf sFlag = "A" And bEstado Then
            clsMant.ActualizaComentBlqRet sCuenta, sComentario, nMotRet, sMov
            nEstadoCta = gCapEstBloqRetiro
        End If
        
        
        rsRet.MoveNext
    Loop
End If
If Not rsTot Is Nothing Then
    rsTot.MoveFirst
    Do While Not rsTot.EOF
        nMotTot = CLng(Trim(Right(rsTot("Motivo"), 3)))
        bEstado = IIf(rsTot("Est") = "0", False, True)
        sComentario = rsTot("Comentario")
        sFlag = rsTot("Flag")
        sMov = rsTot("cMovNro")
        If sFlag = "N" And bEstado Then 'Agrega los nuevos bloqueos
            clsMant.NuevoBloqueototal sCuenta, nMotTot, sComentario, sMovNro
            nEstadoCta = gCapEstBloqTotal
            
            If nMotTot = gCapMotBlqRetGarantia Or nMotTot = 15 Then
                sOperacion = IIf(Mid(sCuenta, 6, 3) = "232", "201101", IIf(Mid(sCuenta, 6, 3) = "233", "211101", "221101"))
                clsCap.AgregaMov sMovNro, sOperacion, sComentario, gMovEstContabMovContable, gMovFlagVigente
                nMovNro = clsCap.GetnMovNro(sMovNro)
                clsCap.AgregaMovCap nMovNro, sOperacion, sCuenta, RSTMP.Fields("nsaldo"), RSTMP.Fields("nsaldodisp"), RSTMP.Fields("nsaldo")
                clsCap.AgregaMovCapDet nMovNro, sOperacion, sCuenta, gConcCapital, RSTMP.Fields("nsaldo")
            End If
            
        ElseIf sFlag = "A" And Not bEstado Then 'Desbloquea los que ya fueron desmarcados
            clsMant.ActualizaBloqueoTot sCuenta, sComentario, sMovNro, nMotTot, sMov
            
            If nMotTot = gCapMotBlqRetGarantia Or nMotTot = 15 Then
                sOperacion = IIf(Mid(sCuenta, 6, 3) = "232", "201102", IIf(Mid(sCuenta, 6, 3) = "233", "211102", "221102"))
                clsCap.AgregaMov sMovNro, sOperacion, sComentario, gMovEstContabMovContable, gMovFlagVigente
                nMovNro = clsCap.GetnMovNro(sMovNro)
                clsCap.AgregaMovCap nMovNro, sOperacion, sCuenta, RSTMP.Fields("nsaldo"), RSTMP.Fields("nsaldodisp"), RSTMP.Fields("nsaldo")
                clsCap.AgregaMovCapDet nMovNro, sOperacion, sCuenta, gConcCapital, RSTMP.Fields("nsaldo")
            End If
            
        ElseIf sFlag = "A" And bEstado Then
            clsMant.ActualizaComentBlqTot sCuenta, sComentario, nMotTot, sMov
            nEstadoCta = gCapEstBloqTotal
            
        End If
        rsTot.MoveNext
    Loop
End If
clsMant.ActualizaEstadoCuenta sCuenta, nEstadoCta
clsMant.dbCmact.CommitTrans
Set clsMant = Nothing

Exit Sub
ErrActBloqueo:
    If bIniTran Then
        clsMant.dbCmact.RollbackTrans
    End If
    Set clsMant = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Sub
'ANDE 20171108 SE AGREGÓ EL PARAMETRO pdFecSis
Public Sub ActualizaBloqueosParciales(ByVal sCuenta As String, ByVal rsRet As ADODB.Recordset, _
        ByVal sMovNro As String, ByVal nMontoAcumBlq As Double, ByVal pdFecSis As Date)

Dim nEstadoCta As CaptacEstado
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales, clsCap As COMDCaptaGenerales.DCOMCaptaMovimiento
Dim sComentario As String, sFlag As String, sMov As String
Dim nMotRet As CaptacMotBloqueoRet
Dim nMotTot As CaptacMotBloqueoTot
Dim bIniTran As Boolean, bEstado As Boolean, nMovNro As Long
Dim RSTMP As ADODB.Recordset
Dim sOperacion As String
Dim nMonto As Double

'ANDE 20170726 AGREGADO DE pdFecSis
VerificarCobroComisionRetencionFondos rsRet, sCuenta, sMovNro, pdFecSis, True, nMontoAcumBlq 'JUEZ 20160115
'END ANDE
Set clsCap = New COMDCaptaGenerales.DCOMCaptaMovimiento

Set RSTMP = GetDatosCuenta(sCuenta)

On Error GoTo ErrActBloqueo
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
clsMant.dbCmact.BeginTrans
bIniTran = True
If Not rsRet Is Nothing Then
    rsRet.MoveFirst
    Do While Not rsRet.EOF
        nMotRet = CLng(Trim(Right(rsRet("Motivo"), 3)))
        bEstado = IIf(rsRet("Est") = "0", False, True)
        sComentario = rsRet("Comentario")
        sFlag = rsRet("Flag")
        sMov = rsRet("cMovNro")
        nMonto = rsRet("Monto")
        
        If sFlag = "N" And bEstado Then 'Nuevo Bloqueo Parcial
            clsMant.NuevoBloqueoParcial sCuenta, nMotRet, sComentario, sMovNro, nMonto
        ElseIf sFlag = "A" And Not bEstado Then 'Desbloquea los que ya fueron desmarcados
            clsMant.ActualizaBloqueoParcial sCuenta, sComentario, sMovNro, nMotRet, sMov
        ElseIf sFlag = "A" And bEstado Then
            clsMant.ActualizaComentBlqParcial sCuenta, sComentario, nMotRet, sMov
        End If
        rsRet.MoveNext
    Loop
    Set rsRet = Nothing
End If

clsMant.ActualizaMontoBlqParcial sCuenta, nMontoAcumBlq
clsMant.dbCmact.CommitTrans
Set clsMant = Nothing

Exit Sub
ErrActBloqueo:
    If bIniTran Then
        clsMant.dbCmact.RollbackTrans
    End If
    Set clsMant = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Sub

Public Function GetCapTasaInteres(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
    ByVal nTipoTasa As CaptacTipoTasa, Optional nPlazo As Double = 0, Optional nValor As Double = 0, _
    Optional sCodAge As String = "", Optional bOrdPag As Boolean = False) As Double

Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
GetCapTasaInteres = clsMant.GetCapTasaInteres(nProducto, nMoneda, nTipoTasa, nPlazo, nValor, sCodAge, bOrdPag)
Set clsMant = Nothing
End Function

Public Function GetTasaNominal(ByVal nTasaEfecAn As Double, ByVal nPeriodo As Integer) As Double
GetTasaNominal = Round((((1 + nTasaEfecAn / 100) ^ (1 / nPeriodo)) - 1) * 100 * nPeriodo, 4) 'By Capi 04122007 para que trabaje con 4 decimales
End Function

Public Function GetTasaEfectiva(ByVal nTasaNomAn As Double, ByVal nPeriodo As Integer) As Double
'GetTasaEfectiva = (((1 + nTasaNomAn / 36000) ^ nPeriodo) - 1) * 100
GetTasaEfectiva = ((1 + nTasaNomAn / 36000) ^ nPeriodo - 1) * 100

End Function

Public Function GetInteresPF(ByVal nTasa As Double, ByVal nCapital As Double, _
                nPlazo As Long) As Double
    GetInteresPF = ((1 + nTasa / 36000) ^ nPlazo - 1) * nCapital
End Function

Public Function GetPFPlanRetInt(ByVal dFecApe As Date, ByVal nInteres As Double, ByVal nPlazo As Long, ByVal nMoneda As Moneda, _
            ByVal nIntFinPlazo As Double, ByVal nCapital As Double, ByVal nTasa As Double, Optional nDiasVal As Integer = 0, _
            Optional dFecVal As Date, Optional pnTotIntMes As Double, Optional sCuenta As String = "", Optional nCostoMan As Currency = 0, Optional pnInteresAdel As Double) As String
'By Capi 20042008 se agrego parametro pnTotIntMes
Dim sCad As String
Dim nCarLin As Integer, nDias As Long
Dim sTitRp1 As String, sTitRp2 As String
Dim sMoneda As String, sNumPag As String
Dim nLinPag As Integer, nCntPag As Integer
Dim dFecVenc As Date, dAuxiliar As Date
Dim sFecVenc As String, sFchtra As String
Dim sCodCta As String, sNumDoc As String
Dim sNroCuo As String, nAcum As Double, nUltInteres As Double
Dim bPrimero As Boolean, dFecCan As Date, sFecCan As String

Dim oImpre As New COMFunciones.FCOMImpresion
Dim ofun As New COMFunciones.FCOMCadenas
Dim clsCaptaImpre As New NCOMCaptaImpresion

Dim nTasaE As Double
Dim nTasaN As Double 'By Capi 30112007

sCad = ""
sNumPag = ""
nLinPag = 0
nCarLin = 85
'''sMoneda = IIf(nMoneda = gMonedaNacional, "SOLES", "DOLARES") 'marg ers044-2016
sMoneda = IIf(nMoneda = gMonedaNacional, StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
sTitRp1 = "CRONOGRAMA  DE  RETIRO  DE INTERESES"
sTitRp2 = ""
sCodCta = ""
nCntPag = 1
nAcum = 0

dAuxiliar = DateAdd("d", nDiasVal, dFecApe)
dFecVenc = DateAdd("d", nPlazo, dFecApe)
dFecCan = DateAdd("d", nPlazo + 1, dFecApe)
sFecVenc = Format$(dFecVenc, "dd mmmm yyyy")
sFecCan = Format$(dFecCan, "dd mmmm yyyy")

'By Capi 30112007 porque debe jalar la TNA Porque  el parametro nTasa es TEA, ademas es independiente al plazo
'nTasaN = Format$(GetTasaNominal(nTasa, 360), "#,##0.00")
nTasaE = Format$(GetTasaEfectiva(nTasa, 360), "#,##0.00")
'ALPA 20091118************************************************
Dim objCaptac As COMNCaptaGenerales.NCOMCaptaGenerales
Set objCaptac = New COMNCaptaGenerales.NCOMCaptaGenerales
Dim pnPlazoT As Integer
Dim nTREA As Currency
pnPlazoT = CInt(nPlazo)
'*** PEAC 20091214
'nTREA = objCaptac.ObtenerTREA(Mid$(sCuenta, 6, 3), nCapital, nTasaE, pnPlazoT, nCostoMan)
nTREA = objCaptac.ObtenerTREA(IIf(Len(Trim(sCuenta)) = 0, "233", Mid$(sCuenta, 6, 3)), nCapital, nTasaE, pnPlazoT, nCostoMan)
'*** FIN PEAC

'*************************************************************
sNumPag = ofun.FillNum(Trim(Str(nCntPag)), 4, " ")
sCad = sCad & clsCaptaImpre.CabeRepoCaptac("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, sNumPag, "", dFecVenc, oImp.gPrnSaltoLinea) & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & String(3, " ") & "PLAZO             : " & nPlazo & Space(3) & " DIAS" & oImp.gPrnSaltoLinea
sCad = sCad & String(3, " ") & "CAPITAL           : " & Format$(nCapital, "#,##0.00") & Space(3) & sMoneda & oImp.gPrnSaltoLinea
'ALPA 2009118*****
'sCad = sCad & String(3, " ") & "T.N.A             : " & nTasa & " %" & oImp.gPrnSaltoLinea 'By Capi 30112007
sCad = sCad & String(3, " ") & "T.R.E.A.          : " & nTREA & " %" & oImp.gPrnSaltoLinea 'By Capi 30112007
'*******************
sCad = sCad & String(3, " ") & "T.E.A             : " & nTasaE & " %" & oImp.gPrnSaltoLinea  'By Capi 30112007

If nDiasVal > 0 Then
    sCad = sCad & String(3, " ") & "FECHA APE. CHQ.   : " & Format$(dFecApe, "dd mmmm yyyy") & oImp.gPrnSaltoLinea
    sCad = sCad & String(3, " ") & "FECHA VAL. CHQ.   : " & Format$(dFecVal, "dd mmmm yyyy") & oImp.gPrnSaltoLinea
Else
    sCad = sCad & String(3, " ") & "FECHA APERTURA    : " & Format$(dFecApe, "dd mmmm yyyy") & oImp.gPrnSaltoLinea
End If
sCad = sCad & String(3, " ") & "FECHA VENCIMIENTO : " & sFecVenc & oImp.gPrnSaltoLinea
'ALPA 20091121************************************
'sCad = sCad & String(3, " ") & "FECHA CANCELACION : " & sFecCan & oImp.gPrnSaltoLinea
'*************************************************
sCad = sCad & String(3, " ") & "MONTO ADELANTADO  : " & Format$(pnInteresAdel, "#,##0.00") & " (Si retira el intereses adelantado)" & oImp.gPrnSaltoLinea
sCad = sCad & String(3, " ") & "MONTO FINAL PLAZO : " & Format$(nIntFinPlazo, "#,##0.00") & " (Sin retirar intereses mensualmente)" & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & "  FECHA OPERACION            IMPORTE   " & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
nLinPag = 11
dAuxiliar = DateAdd("d", 30, dAuxiliar)
bPrimero = True
While dFecVenc > dAuxiliar
    sFchtra = Format(dAuxiliar, "dd/mm/yyyy")
    If bPrimero Then
        nAcum = nAcum + nInteres
        sCad = sCad & "   " & sFchtra & "           " & ofun.JDNum(Trim(Str(nInteres)), 12, True, 9, 2) & oImp.gPrnSaltoLinea
        bPrimero = False
    Else
        sCad = sCad & "   " & sFchtra & "           " & ofun.JDNum(Trim(Str(nInteres)), 12, True, 9, 2) & oImp.gPrnSaltoLinea
        nAcum = nAcum + nInteres
    End If
    
    nLinPag = nLinPag + 1
    
    dAuxiliar = DateAdd("d", 30, dAuxiliar)
Wend
dAuxiliar = DateAdd("d", -30, dAuxiliar)
nDias = DateDiff("d", dAuxiliar, dFecVenc)
If nDias < 30 Then
    nUltInteres = GetInteresPF(nTasa, nCapital, nDias)
    nAcum = nAcum + nUltInteres
    sCad = sCad & "   " & Format(dFecVenc, "dd/mm/yyyy") & "           " & ofun.JDNum(Trim(Str(nUltInteres)), 12, True, 9, 2) & "    RENOVACION" & oImp.gPrnSaltoLinea
Else
    nAcum = nAcum + nInteres
    sCad = sCad & "   " & Format(dFecVenc, "dd/mm/yyyy") & "           " & ofun.JDNum(Trim(Str(nInteres)), 12, True, 9, 2) & "    RENOVACION" & oImp.gPrnSaltoLinea
End If
'By Capi 20042008
pnTotIntMes = Round(nAcum, 2)
'
sCad = sCad & String(nCarLin, "-") & oImp.gPrnSaltoLinea
sCad = sCad & String(3, " ") & "ACUMULADO " & oImpre.FillText(sMoneda, 8, " ") & "   " & ofun.JDNum(Trim(Str(nAcum)), 12, True, 9, 2) & oImp.gPrnSaltoLinea
sCad = sCad & String(nCarLin, "=") & oImp.gPrnSaltoLinea
GetPFPlanRetInt = sCad
End Function

Public Function GetBeneficiarios(ByVal sPersona As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetBeneficiarios = clsMant.GetBeneficiarios(sPersona)
Set clsMant = Nothing
End Function

Public Function GetTarjetaCuentas(ByVal sTarjeta As String, Optional nProducto As Producto = gCapTodoAhorro) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetTarjetaCuentas = clsMant.GetTarjetaCuentas(sTarjeta, nProducto)
Set clsMant = Nothing
End Function

Public Function GetTarjetaEstadoHist(ByVal sTarjeta As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetTarjetaEstadoHist = clsMant.GetTarjetaEstadoHist(sTarjeta)
Set clsMant = Nothing
End Function

Public Function ActualizaCuentaTarj(ByVal sTarjeta As String, ByVal sPersona As String, ByVal rsTarj As ADODB.Recordset) As Boolean
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean
Dim sCuenta As String
Dim nCuenta As Long
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
On Error GoTo ErrGrabar

clsMant.dbCmact.BeginTrans
bTrans = True

clsMant.EliminaCuentaTarjPersona sPersona, sTarjeta

    If Not (rsTarj Is Nothing) Then
        Do While Not rsTarj.EOF
            sCuenta = rsTarj("Cuenta")
            clsMant.AgregaCuentaTarjPersona sCuenta, sPersona, sTarjeta
            rsTarj.MoveNext
        Loop
    End If

clsMant.dbCmact.CommitTrans
Set clsMant = Nothing

ActualizaCuentaTarj = True

Exit Function
ErrGrabar:

    If bTrans And Err.Number <> 91 Then
        clsMant.dbCmact.RollbackTrans
    End If
'    If bTrans And Err.Number = 91 Then
'        clsmant.dbCmact.CommitTrans
'    End If
    Set clsMant = Nothing
'    MsgBox Err.Description, vbExclamation, "Error"
    Err.Raise Err.Number, "Error", Err.Description
    ActualizaCuentaTarj = False
    
End Function

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal nNumOP As Long) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetDatosOrdenPago = clsMant.GetDatosOrdenPago(sCuenta, nNumOP)
Set clsMant = Nothing
End Function

Public Function GetPersonaCuenta(ByVal sCuenta As String, Optional nTipRel As CaptacRelacPersona) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetPersonaCuenta = clsMant.GetPersonaCuenta(sCuenta, nTipRel)
Set clsMant = Nothing
End Function

Public Function GetPersonasCuentaNA(ByVal sCuenta As String, ByVal sPersCod As String, Optional ByVal nPersoneria As Integer) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetPersonasCuentaNA = clsMant.GetPersonasCuentaNA(sCuenta, sPersCod, nPersoneria)
Set clsMant = Nothing
End Function


Public Function GeneraConvenioTarjeta(ByVal psTextoCarta As String, sCuenta As String, ByVal dFecha As Date) As String
Dim rsPers As ADODB.Recordset
Dim lsTit As String, lsRep As String
Dim lsDireccion As String, lsDNI As String
Dim sCad As String, sPersona As String, saux As String
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim ofun As New COMFunciones.FCOMCadenas

sCad = Chr$(27) & Chr$(77)
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rsPers = clsMant.GetPersonaCuenta(sCuenta)
Set clsMant = Nothing

'Open App.Path & "\FormatoCarta\ConTar.txt" For Input As #1

'Do While Not EOF(1)
'    Line Input #1, saux
'    sCad = sCad & Chr$(10) & saux
'Loop
'Close #1

If Not (rsPers.EOF And rsPers.BOF) Then
    lsTit = ""
    lsRep = ""
    sPersona = ""
    Do While Not rsPers.EOF
        If sPersona <> rsPers("cPersCod") Then
            If rsPers("nPrdPersRelac") <> gCapRelPersTitular Then
                If lsRep = "" Then
                    lsRep = ofun.PstaNombre(Trim(rsPers("Nombre")))
                Else
                    lsRep = lsRep & oImp.gPrnSaltoLinea & Space(40) & ofun.PstaNombre(Trim(rsPers("Nombre")))
                End If
            Else
                If lsTit = "" Then
                    lsDNI = rsPers("ID") & " " & rsPers("ID N°")
                    lsTit = ofun.PstaNombre(rsPers("Nombre"))
                    lsDireccion = rsPers("Direccion") & ""
                Else
                    lsTit = lsTit & oImp.gPrnSaltoLinea & Space(52) & ofun.PstaNombre(rsPers("Nombre"))
                End If
            End If
            sPersona = rsPers("cPersCod")
        End If
        rsPers.MoveNext
    Loop
    
    sCad = psTextoCarta
    sCad = Replace(sCad, "<<NOMBRE>>", lsTit)
    sCad = Replace(sCad, "<<DNI>>", lsDNI)
    sCad = Replace(sCad, "<<DIRECCION>>", lsDireccion)
    sCad = Replace(sCad, "<<REPRESENTANTE>>", lsRep)
    sCad = Replace(sCad, "<<DIA>>", Format$(dFecha, "dddd dd"))
    sCad = Replace(sCad, "<<MES>>", Format$(dFecha, "mmmm"))
    sCad = Replace(sCad, "<<ANIO>>", Format$(dFecha, "yyyy"))
    
    
    
End If
rsPers.Close
Set rsPers = Nothing
GeneraConvenioTarjeta = sCad
End Function



Public Function GetCartasOP(ByVal sAge As String) As ADODB.Recordset
'Dim rsOP As ADODB.Recordset
Dim lsTit As String, lsRep As String
Dim lsDireccion As String, lsDNI As String
Dim sCad As String, sPersona As String, saux As String
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim LALISTA() As String
Dim lnFilas As Long, lnItem As Long
Dim lsCartaModelo As String
Dim lsCadImp As String

   Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales

   Set GetCartasOP = clsMant.GetDatosOP(sAge)
   Set clsMant = Nothing
     
  '  Set rsOP = clsMant.GetDatosOP(sAge)
    
    
  
'    lnFilas = 0
'    ReDim LALISTA(rsOP.RecordCount, 5)
'
'    While Not rsOP.EOF
'        lnFilas = lnFilas + 1
'                    LALISTA(lnFilas, 0) = rsOP!Nombre
'                    LALISTA(lnFilas, 1) = rsOP!Direccion
'                    LALISTA(lnFilas, 2) = rsOP!CCTACODANT
'                    LALISTA(lnFilas, 3) = rsOP!cCtaCod
'
'        rsOP.MoveNext
'     Wend
     
'      For lnItem = 1 To lnFilas
'        lsCartaModelo = psTextoCarta
'        lsCartaModelo = Replace(lsCartaModelo, "<<DIA>>", Format(dFecha, "dd"))
'        lsCartaModelo = Replace(lsCartaModelo, "<<MES>>", Format(dFecha, "MM"))
'        lsCartaModelo = Replace(lsCartaModelo, "<<ANIO>>", Format(dFecha, "YYYY"))
'        lsCartaModelo = Replace(lsCartaModelo, "<<NOMBRE>>", PstaNombre(LALISTA(lnItem, 0), False), , 1, vbTextCompare)
'        lsCartaModelo = Replace(lsCartaModelo, "<<DIRECCION>>", LALISTA(lnItem, 1), , 1, vbTextCompare)
'        lsCartaModelo = Replace(lsCartaModelo, "<<CODIGO ANTIGUO>>", LALISTA(lnItem, 2), , 1, vbTextCompare)
'        lsCartaModelo = Replace(lsCartaModelo, "<<CODIGO NUEVO>>", LALISTA(lnItem, 3), , 1, vbTextCompare)
'
'
'        lsCadImp = lsCadImp & lsCartaModelo & Chr(12)
'        'prgList.Value = vCont
'    Next lnItem




    
'    Erase LALISTA
    'nRemImprimeAvisoRemate = lsCadImp
          
'Set clsMant = Nothing
'        rsOP.Close
'        Set rsOP = Nothing
'        GeneraCartasOP = lsCadImp
End Function


Public Function GetNombreTitulares(ByVal sCuenta As String, Optional pbNombreCompleto As Boolean = False, Optional pnTitulares As Integer = 0, Optional pnMargenIzquierdo As Integer = 0) As String
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
GetNombreTitulares = clsMant.GetNombreTitulares(sCuenta, pbNombreCompleto, pnTitulares, pnMargenIzquierdo)
Set clsMant = Nothing
End Function

Public Sub EmiteRangoOrdenPago(ByVal sCuenta As String, ByVal nInicio As Long, _
        ByVal nFin As Long, ByVal sMovNro As String, Optional bDescuento As Boolean = True, _
        Optional nMontoDcto As Double = 0, Optional ByRef psMensajeValidacion As String = "")

Dim sRango As String
Dim bTrans As Boolean
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim nErrNum As Integer

nErrNum = 0
On Error GoTo ErrGraba
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales

sRango = clsMant.ExisteOrdenPagoEmitidas(sCuenta, nInicio, nFin)
If sRango = "" Then
    If bDescuento Then
        'Agregar codigo para realizar el retiro de la cuenta de ahorros correspondiente
    End If
    clsMant.dbCmact.BeginTrans
    bTrans = True
    clsMant.AgregaOrdenPagoEmitidas sCuenta, nInicio, nFin, sMovNro
    clsMant.dbCmact.CommitTrans
    Set clsMant = Nothing
    Exit Sub
Else
    nErrNum = 1
End If
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    If nErrNum = 1 Then
        'MsgBox "Rango ya emitido. Rango Referencia: " & Chr$(13) & sRango, vbInformation, "Aviso"
        psMensajeValidacion = "Rango ya emitido. Rango Referencia: " & sRango
    Else
        'MsgBox Err.Description, vbInformation, "Aviso"
        Err.Raise Err.Number, "Error", Err.Description
    End If
End Sub

Public Sub AnulaRangoOrdenPago(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nNumIni As Long, ByVal nNumFin As Long, ByVal sGlosa As String, _
        Optional ByRef psMensajeValidacion As String = "")

Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim nErrNum As Integer, nAnuladas As Integer
Dim i As Long
Dim rsEstado As ADODB.Recordset
Dim bTrans As Boolean
Dim nEstado As COMDConstantes.CaptacOrdPagoEstado
Dim sDesc As String, sMsgErr As String, sEstado As String
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales

On Error GoTo ErrGraba
nErrNum = 0
nAnuladas = 0
sDesc = "Rango Anulado : " & nNumIni & " - " & nNumFin
sGlosa = sGlosa & Chr$(13) & sDesc
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.AgregaMovNoContable sMovNro, sGlosa, gAhoOPAnulacion
For i = nNumIni To nNumFin
    If clsMant.ExisteOrdenPagoEmitidas(sCuenta, i, i) = "" Then
        sMsgErr = sMsgErr & "Orden Pago NO Emitida : " & i & Chr$(13)
    Else
        Set rsEstado = clsMant.GetEstadoOrdenPagoEmitida(sCuenta, i)
        If rsEstado.EOF And rsEstado.BOF Then
            nEstado = gCapOPEstEmitida
            sEstado = ""
        Else
            nEstado = rsEstado("nEstado")
            sEstado = rsEstado("Estado")
        End If
        If nEstado <> gCapOPEstEmitida And nEstado <> gCapOPEstExtornada Then
            sMsgErr = sMsgErr & "Orden Pago NO Válida : " & i & Space(2) & sEstado & Chr$(13)
        Else
            clsMant.AnulaOrdenPagoEmitida sCuenta, i, sMovNro, 0
            nAnuladas = nAnuladas + 1
        End If
    End If
Next i
If nAnuladas = 0 Then
    clsMant.dbCmact.RollbackTrans
Else
    clsMant.dbCmact.CommitTrans
End If
bTrans = False
Set clsMant = Nothing
If sMsgErr = "" Then
    Exit Sub
Else
    nErrNum = 1
End If
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    If nErrNum = 0 Then
        Err.Raise Err.Number, "Error", Err.Description
    ElseIf nErrNum = 1 Then
        psMensajeValidacion = "Ordenes Anuladas : " & nAnuladas & sMsgErr
    End If
End Sub

Public Sub AnulaOrdenPago(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nNumOP As Long, ByVal sGlosa As String, Optional nMonto As Double = 0, Optional ByRef psMensajeValidacion As String = "")

Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim nErrNum As Integer
Dim rsEstado As ADODB.Recordset
Dim bTrans As Boolean
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Dim sEstado As String
Dim nEstado As CaptacOrdPagoEstado
On Error GoTo ErrGraba
nErrNum = 0
If clsMant.ExisteOrdenPagoEmitidas(sCuenta, nNumOP, nNumOP) = "" Then
    nErrNum = 1
    GoTo ErrGraba
End If
Set rsEstado = clsMant.GetEstadoOrdenPagoEmitida(sCuenta, nNumOP)
If rsEstado.EOF And rsEstado.BOF Then
    nEstado = gCapOPEstEmitida
    sEstado = ""
Else
    nEstado = rsEstado("nEstado")
    sEstado = rsEstado("Estado")
End If
If nEstado <> gCapOPEstEmitida And nEstado <> gCapOPEstExtornada Then
    nErrNum = 2
    GoTo ErrGraba
End If
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.AgregaMovNoContable sMovNro, sGlosa, gAhoOPAnulacion
clsMant.AnulaOrdenPagoEmitida sCuenta, nNumOP, sMovNro, nMonto
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
Exit Sub
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    If nErrNum = 0 Then
        Err.Raise Err.Number, "Aviso", Err.Description
    ElseIf nErrNum = 1 Then
        psMensajeValidacion = "Orden Pago NO Emitida"
    ElseIf nErrNum = 2 Then
        psMensajeValidacion = "Orden Pago NO Válida. Estado " & sEstado
    End If
End Sub

Public Function GetSaldoFecha(ByVal sCuenta As String, ByVal dFecha As Date) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetSaldoFecha = clsMant.GetSaldoFecha(sCuenta, dFecha)
Set clsMant = Nothing
End Function

Public Function GetSaldoFecha2(ByVal cUser As String, ByVal dFecFin As Date) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetSaldoFecha2 = clsMant.GetSaldoFecha2(cUser, dFecFin)
Set clsMant = Nothing
End Function

Public Function GetMovimientosCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, Optional bMuestraExtornos As Boolean = True) As ADODB.Recordset
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set GetMovimientosCuenta = clsMant.GetMovimientosCuenta(sCuenta, dFecIni, dFecFin, nNumMov, bMuestraExtornos)
    Set clsMant = Nothing
End Function
'By Capi 11112008 caj aut
Public Function GetMovimientosPendientesCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, Optional bMuestraExtornos As Boolean = True) As ADODB.Recordset
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set GetMovimientosPendientesCuenta = clsMant.GetMovimientosPendientesCuenta(sCuenta, dFecIni, dFecFin, nNumMov, bMuestraExtornos)
    Set clsMant = Nothing
End Function
'JUEZ 20150205 ***********************************************
Public Function GetMovimientosDebitosAutoCuenta(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, Optional bMuestraExtornos As Boolean = True) As ADODB.Recordset
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set GetMovimientosDebitosAutoCuenta = clsMant.GetMovimientosDebitosAutoCuenta(sCuenta, dFecIni, dFecFin, nNumMov, bMuestraExtornos)
    Set clsMant = Nothing
End Function
'END JUEZ ****************************************************
Public Function GetMovimientosCuenta2(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, Optional bMuestraExtornos As Boolean = True, Optional cUser As String, Optional ByRef rsTemp As ADODB.Recordset) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetMovimientosCuenta2 = clsMant.GetMovimientosCuenta2(sCuenta, dFecIni, dFecFin, nNumMov, bMuestraExtornos, cUser, rsTemp)
Set clsMant = Nothing
End Function

Public Function GetMovimientosCuenta4(ByVal sCuenta As String, Optional dFecIni As Date, _
        Optional dFecFin As Date, Optional nNumMov As Long = 0, Optional bMuestraExtornos As Boolean = True, Optional cUser As String, Optional ByRef rsTemp As ADODB.Recordset) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetMovimientosCuenta4 = clsMant.GetMovimientosCuenta4(sCuenta, dFecIni, dFecFin, nNumMov, bMuestraExtornos, cUser, rsTemp)
Set clsMant = Nothing
End Function

Public Sub ActualizaSobregiro(ByVal sCuenta As String, ByVal nSobregiro As Long, _
        ByVal sMovNro As String, ByVal nEstadoCta As CaptacEstado, Optional bBloqueaCuenta As Boolean = False)
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean

Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales

On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True
If bBloqueaCuenta Then
    clsMant.NuevoBloqueoRetiro sCuenta, gCapMotBlqRetOrdenPagoRechazada, "Orden Sobregirada", sMovNro
    If nEstadoCta = gCapEstActiva Then
        clsMant.ActualizaEstadoCuenta sCuenta, gCapEstBloqRetiro
        clsMant.AgregaCaptacEstado sCuenta, gCapEstBloqRetiro, sMovNro
    End If
End If
clsMant.ActualizaSobregiro sCuenta, nSobregiro
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
Exit Sub
ErrGraba:
    If bTrans Then
        clsMant.dbCmact.RollbackTrans
    End If
    Set clsMant = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Sub

Public Function GetTarifaOrdenPago(ByVal nMoneda As Moneda) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetTarifaOrdenPago = clsMant.GetTarifaOrdenPago(nMoneda)
Set clsMant = Nothing
End Function

Public Function GetMaxOrdPagEmitida(ByVal sCuenta As String) As Long
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
GetMaxOrdPagEmitida = clsMant.GetMaxOrdPagEmitida(sCuenta)
Set clsMant = Nothing
End Function

Public Function GetHistOrdPagEmision(ByVal sCuenta As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetHistOrdPagEmision = clsMant.GetHistOrdPagEmision(sCuenta)
Set clsMant = Nothing
End Function

Public Sub AgregaSolicitudOrdPagTal(ByVal sCuenta As String, ByVal sMovNro As String, _
        ByVal nInicio As Long, ByVal nFin As Long, ByVal nNumTal As Long)
Dim oCapMant As COMDCaptaGenerales.DCOMCaptaGenerales

Dim bTrans As Boolean
Set oCapMant = New COMDCaptaGenerales.DCOMCaptaGenerales
bTrans = True
oCapMant.AgregaCapOrdPagEmision sCuenta, sMovNro, gCapTalOrdPagEstSolicitado, nInicio, nFin, nNumTal
oCapMant.AgregaCapOrdPagEstado sCuenta, sMovNro, gCapTalOrdPagEstSolicitado, nInicio
bTrans = False
Set oCapMant = Nothing
Exit Sub
ErrGraba:
    If bTrans Then oCapMant.dbCmact.RollbackTrans
    Set oCapMant = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Sub

Public Function GetCapOrdPagEmision(ByVal nEstado As CapOrdPagTalEstado) As ADODB.Recordset
Dim oCapMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set oCapMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetCapOrdPagEmision = oCapMant.GetCapOrdPagEmision(nEstado)
Set oCapMant = Nothing
End Function

Public Function GetCapOrdPagEmision_Cuenta(ByVal nEstado As CapOrdPagTalEstado, ByVal psCodCta As String) As ADODB.Recordset
Dim oCapMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set oCapMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetCapOrdPagEmision_Cuenta = oCapMant.GetCapOrdPagEmision_Cuenta(nEstado, psCodCta)
Set oCapMant = Nothing
End Function
Public Sub ActualizaCapOrdPagEmision(ByVal sCuenta As String, ByVal sMovNroNuevo As String, _
        ByVal nEstadoNuevo As CapOrdPagTalEstado, ByVal sMovNroAnt As String, _
        ByVal nEstadoAnt As CapOrdPagTalEstado, ByVal nInicio As Long)

Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
oCap.dbCmact.BeginTrans
oCap.ActualizaCapOrdPagEmision sCuenta, sMovNroNuevo, nEstadoNuevo, sMovNroAnt, nEstadoAnt
oCap.AgregaCapOrdPagEstado sCuenta, sMovNroNuevo, nEstadoNuevo, nInicio
oCap.dbCmact.CommitTrans
End Sub

Private Function GetTitulares(ByVal sCuenta As String) As String
Dim sNombre1 As String * 60, sNombre2 As String * 60, sNombre3 As String * 60
Dim sDocId1 As String * 11, sDocId2 As String * 11, sDocId3 As String * 11
Dim sNom As String * 60, sDoc As String * 11
Dim rsPers As ADODB.Recordset
'Dim oCapMant As NCOMCaptaMovimiento
Dim i As Integer
Dim ofun As New COMFunciones.FCOMCadenas

sNombre1 = Space(60)
sNombre2 = Space(60)
sNombre3 = Space(60)
sDocId1 = Space(11)
sDocId2 = Space(11)
sDocId3 = Space(11)
i = 0

'Set oCapMant = New NCOMCaptaMovimiento
Set rsPers = GetPersonaCuenta(sCuenta)
If Not (rsPers.EOF And rsPers.BOF) Then
    Do While Not rsPers.EOF
        sNom = ofun.PstaNombre(rsPers("Nombre"), True)
        sDoc = rsPers("ID N°")
        i = i + 1
        Select Case i
            Case 1
                sNombre1 = sNom
                sDocId1 = sDoc
            Case 2
                sNombre2 = sNom
                sDocId2 = sDoc
            Case 3
                sNombre3 = sNom
                sDocId3 = sDoc
        End Select
        rsPers.MoveNext
    Loop
    GetTitulares = sDocId1 & sDocId2 & sDocId3 & sNombre1 & sNombre2 & sNombre3
Else
    GetTitulares = ""
End If
rsPers.Close
Set rsPers = Nothing
End Function

Private Function GetMiembrosCta(ByVal sCuenta As String) As String
Dim sNombre1 As String * 60, sNombre2 As String * 60, sNombre3 As String * 40
Dim sDocId1 As String * 11, sDocId2 As String * 11, sDocId3 As String * 11
Dim sNom As String * 60, sDoc As String * 11
Dim sNom1 As String * 40
Dim rsPers As ADODB.Recordset
'Dim oCapMant As NCapMantenimiento
Dim i As Integer

sNombre1 = Space(60)
sNombre2 = Space(60)
'****POR MODIFICACION
'sNombre3 = Space(60)
sNombre3 = Space(40)
sDocId1 = Space(11)
sDocId2 = Space(11)
sDocId3 = Space(11)
i = 0

'Set oCapMant = New NCapMantenimiento
Dim ofun As New COMFunciones.FCOMCadenas

LblSeguir:

If i = 0 Then
    Set rsPers = GetPersonaCuenta(sCuenta, gCapRelPersTitular)
Else
    Set rsPers = GetPersonaCuenta(sCuenta, gCapRelPersRepTitular)
    
End If

If Not (rsPers.EOF And rsPers.BOF) Then
    Do While Not rsPers.EOF
        sNom = ofun.PstaNombre(rsPers("Nombre"), True)
        sNom1 = ofun.PstaNombre(rsPers("Nombre"), True)
        sDoc = rsPers("ID N°")
        
        If i = 0 Then
            If rsPers("nperspersoneria") <> 1 Then
                sNombre1 = IIf(Len(Trim(sNom)) < 60, Trim(sNom) + Space(60 - Len(Trim(sNom))), Left(Trim(sNom), 60))
                sDocId1 = IIf(Len(Trim(sDoc)) < 11, Trim(sDoc) + Space(11 - Len(Trim(sDoc))), Left(Trim(sDoc), 11))
                i = i + 1
               GoTo LblSeguir
            End If
        End If
      
        Select Case i
            Case 0
                sNombre1 = IIf(Len(Trim(sNom)) < 60, Trim(sNom) + Space(60 - Len(Trim(sNom))), Left(Trim(sNom), 60))
                sDocId1 = IIf(Len(Trim(sDoc)) < 11, Trim(sDoc) + Space(11 - Len(Trim(sDoc))), Left(Trim(sDoc), 11))
            Case 1
                sNombre2 = IIf(Len(Trim(sNom)) < 60, Trim(sNom) + Space(60 - Len(Trim(sNom))), Left(Trim(sNom), 60))
                sDocId2 = IIf(Len(Trim(sDoc)) < 11, Trim(sDoc) + Space(11 - Len(Trim(sDoc))), Left(Trim(sDoc), 11))
            Case 2
                'sNombre3 = IIf(Len(Trim(sNom)) < 60, Trim(sNom) + Space(60 - Len(Trim(sNom))), Left(Trim(sNom), 60))
                sNombre3 = IIf(Len(Trim(sNom1)) < 40, Trim(sNom1) + Space(40 - Len(Trim(sNom1))), Left(Trim(sNom1), 40))
                sDocId3 = IIf(Len(Trim(sDoc)) < 11, Trim(sDoc) + Space(11 - Len(Trim(sDoc))), Left(Trim(sDoc), 11))
            Case Else
                Exit Do
        End Select
          i = i + 1
        rsPers.MoveNext
    Loop
    
    'GetTitulares = sDocId1 & sDocId2 & sDocId3 & sNombre1 & sNombre2 & sNombre3
    'GetMiembrosCta = sNombre1 & sNombre2 & sNombre3 & sDocId1 & sDocId2 & sDocId3
    GetMiembrosCta = sNombre1 & sDocId1
Else
    GetMiembrosCta = ""
End If
rsPers.Close
Set rsPers = Nothing
End Function


Public Function GeneraArchivoEnvioOrdPagTal(ByVal psTextoCarta As String, ByVal dFecSis As Date, ByVal rsOrd As ADODB.Recordset, _
            ByVal sMovNro As String, Optional ByRef psMensajeValidacion As String = "") As String
Dim sRuta As String
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean
Dim sCuenta As String
Dim sInicio As String * 8, sFin As String * 8
Dim sNumTal As String * 2
Dim sOficCentral As String * 3
Dim sAgencia As String * 3
Dim sTitular As String
Dim sCadena As String, sFecha As String
Dim sMoneda As String * 1
Dim sCadImp As String
Dim sOPTipoTalon  As String
Dim sMovNroAnt As String
Dim sCatChe As String
Dim sDesMoneda As String

Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
oCap.dbCmact.BeginTrans
bTrans = True
On Error GoTo ErrGraba
Do Until rsOrd.EOF
    If rsOrd("Act") = "1" Then
    
        sCuenta = rsOrd("Cuenta")
        sInicio = rsOrd("Inicio")
        sFin = rsOrd("Fin")
        sNumTal = Format$(rsOrd("# Tal"), "00")
        sOPTipoTalon = rsOrd("nTipo")
        sMovNroAnt = rsOrd("cMovNro")
        sAgencia = Format$(CLng(Mid$(sCuenta, 4, 2)), "000")
        sOficCentral = "001"
        sMoneda = Mid$(sCuenta, 9, 1)
        sTitular = GetMiembrosCta(sCuenta)
        sCatChe = IIf(sOPTipoTalon = 1, "CHQx25", IIf(sOPTipoTalon = 2, "CHQx50", "CHQx100"))
        '''sDesMoneda = IIf(sOPTipoTalon = 1, "SOLES", "DOLARES") 'marg ers044-2016
        sDesMoneda = IIf(sOPTipoTalon = 1, StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), "DOLARES") 'marg ers044-2016
        If sTitular <> "" Then
            
            sCadena = sInicio & Right(sCuenta, 10) & sTitular & _
                      Space(22) & sAgencia & sAgencia & "0" & sNumTal & "0" & sOPTipoTalon & sMoneda
            
            sCadImp = sCadImp & sCadena & Chr$(10)
            
            
            oCap.ActualizaCapOrdPagEmision sCuenta, sMovNro, gCapTalOrdPagEstEnviado, sMovNroAnt, gCapTalOrdPagEstSolicitado
            oCap.AgregaCapOrdPagEstado sCuenta, sMovNro, gCapTalOrdPagEstEnviado, CLng(sInicio)
        Else
            
            psMensajeValidacion = "No se encontraron titulares para la cuenta " & sCuenta
        End If
        
        
    End If
    rsOrd.MoveNext
Loop

'Close #1

oCap.dbCmact.CommitTrans
bTrans = False
GeneraArchivoEnvioOrdPagTal = sCadImp

Exit Function
ErrGraba:
    If bTrans Then oCap.dbCmact.RollbackTrans
    Set oCap = Nothing
    'MsgBox Err.Description, vbExclamation, "Error"
    Err.Raise Err.Number, "Error", Err.Description
    GeneraArchivoEnvioOrdPagTal = ""
End Function

Public Sub ActualizaCapOrdPagEstado(ByVal dFecSis As Date, ByVal rsOrd As ADODB.Recordset, _
            ByVal sMovNro As String, ByVal nEstadoNuevo As CapOrdPagTalEstado, _
            ByVal nEstadoAnt As CapOrdPagTalEstado)

Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
'Dim oMovCap As COMDCaptaGenerales.DCOMCaptaMovimiento

Dim sCuenta As String
Dim sMovNroAnt As String
Dim nInicio As Long
Dim bTrans As Boolean
'Dim nMovNroAnt As Long 'BRGO

Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
'Set oMovCap = New COMDCaptaGenerales.DCOMCaptaMovimiento
oCap.dbCmact.BeginTrans
bTrans = True
On Error GoTo ErrGraba


Do While Not rsOrd.EOF
    If rsOrd("Act") = "1" Then
        sCuenta = rsOrd("Cuenta")
        sMovNroAnt = rsOrd("cMovNro")
        nInicio = rsOrd("Inicio")
        'nMovNroAnt = oMovCap.GetnMovNro(sMovNroAnt)
        oCap.ActualizaCapOrdPagEmision sCuenta, sMovNro, nEstadoNuevo, sMovNroAnt, nEstadoAnt
        oCap.AgregaCapOrdPagEstado sCuenta, sMovNro, nEstadoNuevo, nInicio
        'oMovCap.ActualizaEstadoMov nMovNroAnt, gMovFlagVigente
    End If
    rsOrd.MoveNext
Loop


oCap.dbCmact.CommitTrans
bTrans = False
Exit Sub
ErrGraba:
    If bTrans Then oCap.dbCmact.RollbackTrans
    Set oCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Sub
Public Sub ImprimeConevenioOP(ByVal rsOrd As ADODB.Recordset)
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Dim sCuenta As String
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
EliminaConvenioOP
rsOrd.MoveFirst
Do While Not rsOrd.EOF
    If rsOrd("Act") = "1" Then
        sCuenta = rsOrd("Cuenta")
        GeneraConvenioOP sCuenta  'ppoa  carga temporal
    End If
    rsOrd.MoveNext
Loop

'**************ejrs **********************
Dim rs As ADODB.Recordset
Dim oCon As COMConecta.DCOMConecta
Dim sql  As String
Set oCon = New COMConecta.DCOMConecta
Set rs = New ADODB.Recordset
Dim oImpreCap As New NCOMCaptaImpresion

'OJO: Verificar como manejar esta OPCION
'MsgBox "Coloque papel para el convenio de Orden de Pago", vbInformation, "Aviso"

oCon.AbreConexion
sql = "SELECT * FROM tmpConvenioOP  "
Set rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing

'OJO 02-11-2005

'Call oImpreCap.CargaDataReport(DRContratoOP)

'With DRContratoOP
'   Set .DataSource = rs
'      .DataMember = ""
'      .WindowState = vbMaximized
'      .Title = "CONVENIO ORDEN DE PAGO"
'      .inicio
'      .Refresh
'      .Show 1
' End With
'DRContratoOP.Show 1

Set oImpreCap = Nothing
End Sub
Public Function GetCuentaAbonoIF(ByVal sPersCod As String, ByVal nMoneda As Moneda) As String
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
GetCuentaAbonoIF = oCap.GetCuentaAbonoIF(sPersCod, nMoneda)
Set oCap = Nothing
End Function

Public Function BuscaCreditosPendientesPago(ByVal sCuenta As String) As Boolean
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
BuscaCreditosPendientesPago = oCap.BuscaCreditosPendientesPago(sCuenta)
Set oCap = Nothing
End Function

Public Function GetChequesOperaciones(Optional sCuenta As String = "", Optional sNumCheque As String = "", Optional ByVal nEstado As Integer = -1) As ADODB.Recordset
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetChequesOperaciones = oCap.GetChequesOperaciones(sCuenta, sNumCheque, nEstado)
Set oCap = Nothing
End Function

Public Function GetChequesRegistrados(ByVal sNumCheque As String) As ADODB.Recordset
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetChequesRegistrados = oCap.GetChequesRegistrados(sNumCheque)
Set oCap = Nothing
End Function

Public Function GetChequesValorizadosInmediato(ByVal dFecha As Date, Optional sCuenta As String = "", _
            Optional sNumCheque As String = "") As ADODB.Recordset
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetChequesValorizadosInmediato = oCap.GetChequesValorizadosInmediato(dFecha, sCuenta, sNumCheque)
Set oCap = Nothing
End Function

'GIPO ERS070-2016 Sobrantes de Adjudicados
Public Function obtenerTiposDOI() As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerTiposDOI = oDCOMCaptaGenerales.devolverTipoDOI
Set oDCOMCaptaGenerales = Nothing
End Function
'GIPO ERS070-2016
Public Function obtenerBeneficiarioSobrantesDeAdjudicados(ByVal psPersIDnro As String, ByVal psTipoDOI As String) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerBeneficiarioSobrantesDeAdjudicados = oDCOMCaptaGenerales.devolverBeneficiarioSobrantesDeAdjudicados(psPersIDnro, psTipoDOI)
Set oDCOMCaptaGenerales = Nothing
End Function

'GIPO ERS070-2016 20-04-2017
Public Function obtenerFechaGeneracionCarta() As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerFechaGeneracionCarta = oDCOMCaptaGenerales.obtenerFechaGeneracionCarta()
Set oDCOMCaptaGenerales = Nothing
End Function

'GIPO ERS070-2016 26-04-2016
Public Function obtenerAutorizacionDevolucion(ByVal psPersIDnro As String, ByVal pscAgeCod As String, ByVal psUser As String, ByVal fechaHoy As Date) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerAutorizacionDevolucion = oDCOMCaptaGenerales.obtenerAutorizacionDevolucion(psPersIDnro, pscAgeCod, psUser, fechaHoy)
Set oDCOMCaptaGenerales = Nothing
End Function

'GIPO ERS070-2016
Public Function registrarDevolucionDeSobranteAdjudicado(ByVal fecha As String, ByVal codAge As String, ByVal user As String, _
                                                        ByVal cUsuarioVisto As String, ByVal psOpeCod As String, ByVal prsDebitos As ADODB.Recordset, _
                                                        ByVal psTotal As Currency, ByVal ITF As Currency, ByVal cDOI As String) As String
    Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim cMovNro As String
    cMovNro = oDCOMCaptaGenerales.RegistrarDevolucionSobrantesAdj(fecha, codAge, user, cUsuarioVisto, psOpeCod, prsDebitos, psTotal, ITF, cDOI)
    registrarDevolucionDeSobranteAdjudicado = cMovNro
End Function

'GIPO ERS070-2016 26-04-2017
Public Sub registrarAutorizacionDevolucionSobrante(ByVal fecha As Date, ByVal codAge As String, ByVal cUser As String, _
            ByVal nMontoDevolucion As Double, ByVal cPersCodCliente As String)
    Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim cMovNro As String
    cMovNro = oDCOMCaptaGenerales.GeneraMovNro(fecha, codAge, cUser)
    Call oDCOMCaptaGenerales.AgregarAutorizacionDevolucionSobrante(cMovNro, nMontoDevolucion, cPersCodCliente)
End Sub

Public Function GetChequesEnValorizados(sCuenta) As ADODB.Recordset
Dim oCap As COMDCaptaGenerales.DCOMCaptaGenerales
Set oCap = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetChequesEnValorizados = oCap.GetChequesEnValorizacion(sCuenta)
Set oCap = Nothing
End Function

Public Function GetCreditosPendientes(ByVal sPersona As String) As ADODB.Recordset
'Dim rsCred As ADODB.Recordset
'Dim rsCta As ADODB.Recordset
'Dim nProducto As Producto
'Set rsCred = New ADODB.Recordset
'
'Dim oMant As NCapMantenimiento
'Set oMant = New NCapMantenimiento
'Set rsCta = oMant.GetProductoPersonaConsol(sPersona)
'Set oMant = Nothing
'
'Do While Not rsCta.EOF
'
'Loop
'GeneraRecordsetCred rsCred
'rsCred.Open
'CreditosPendientes sPersona, rsCred
'Set GetCreditosPendientes = rsCred
End Function

Public Function CreditosPendientes(ByVal sPersona As String, ByRef rsCred As ADODB.Recordset) As Boolean
'Dim rsCta As ADODB.Recordset
'Dim nAgencias As Long, nLen As Long, nCuentasCre As Long, nCuentasPrd As Long, nCuentasJud As Long
'Dim sBuffer As String, sAgencia As String, sTipo As String
'Dim sSimbMoneda As String, sCuentas As String, sVencidos As String
'Dim I As Integer, J As Integer, nDiasAtraso As Integer
'Dim nDeudaPrend As Double, nMinimoPagar As Double
'Dim bVencidos As Boolean
'Dim nValor1 As Double, nValor2 As Double
'sAgencia = ""
'sBuffer = ""
'nCuentasCre = 0
'
'GetParametrosValores gsBusqCredPendVenc, nValor1, nValor2
'If nValor1 = 1 Then
'    bVencidos = True
'    nDiasAtraso = Int(nValor2)
'Else
'    bVencidos = False
'End If
'Set rsCta = New ADODB.Recordset
'rsCta.CursorLocation = adUseClient
'For I = 1 To UBound(CuentasPersonaCre, 1)
'    sCuentas = "("
'    sAgencia = Left(CuentasPersonaCre(I), 2)
'    For J = I To UBound(CuentasPersonaCre, 1)
'        sBuffer = Left(CuentasPersonaCre(J), 2)
'        If sBuffer = sAgencia Then
'            sCuentas = sCuentas & "'" & CuentasPersonaCre(J) & "',"
'        Else
'            Exit For
'        End If
'    Next J
'    sCuentas = Left(sCuentas, Len(sCuentas) - 1)
'    sCuentas = sCuentas & ")"
'    I = J - 1
'    If AbreConeccion(sAgencia & "0000000000") Then
'        'Verifica primero los creditos pymes pendientes de pago
'        If bVencidos Then
'            sVencidos = "nDiasAtraso > " & nDiasAtraso & " And "
'        Else
'            sVencidos = ""
'        End If
'        VSQL = "Select cCodCta, nDiasAtraso, cRefinan, cNota1, Convert(Varchar(10),dFecVenc,103) dFecVenc, Count(*) nNroCuo, " _
'            & "SUM(nCuota) nCuota, SUM(nMora) nMora, SUM(nGastos) nGastos From ( " _
'            & "Select cCodCta, nDiasAtraso, cRefinan, cNota1, cNroCuo, dFecVenc, nCuota, nMora, SUM(nGastos) nGastos From ( " _
'            & "Select C.cCodCta, C.nDiasAtraso, C.cRefinan, C.cNota1, P.cNroCuo, nCuota = P.nCapital-P.nCapPag+P.nInteres-P.nIntComPag+P.nIntGra+P.nIntGraPag, " _
'            & "nMora = ISNULL(P.nMora-P.nMoraPag,0), nGastos = ISNULL(PG.nMonNeg-PG.nMonPag,0), dFecVenc = (Select Min(dFecVenc) " _
'            & "From PlandesPag P1 Where P1.cCodCta = P.cCodCta And P1.cEstado = 'P' And P1.cTipo = 'C' And DateDiff(dd,P1.dFecVenc,'" & Format$(Date, "mm/dd/yyyy") & "') > 0) " _
'            & "From Credito C INNER JOIN PlandesPag P LEFT JOIN PlanGastos PG ON " _
'            & "P.cCodCta = PG.cCodCta And Convert(Int,P.cNroCuo) = PG.nNroCuo ON C.cCodCta = P.cCodCta " _
'            & "Where " & sVencidos & " C.cEstado = 'F' And P.cEstado = 'P' And P.cTipo = 'C' And " _
'            & "((PG.cAplicado = 'C' And PG.cEstado = 'P') Or (PG.nMonNeg-PG.nMonPag) IS NULL) " _
'            & "And C.cCodCta IN " & sCuentas & " And DateDiff(dd,P.dFecVenc,'" & Format$(Date, "mm/dd/yyyy") & "') > 0 " _
'            & ") T Group by cCodCta, nDiasAtraso, cRefinan, cNota1, cNroCuo, dFecVenc, nCuota, nMora " _
'            & ") T1 Group by cCodCta, nDiasAtraso, cRefinan, cNota1, Convert(Varchar(10),dFecVenc,103)"
'
'        rsCta.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
'        Set rsCta.ActiveConnection = Nothing
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasCre = nCuentasCre + 1
'                rsCred.AddNew
'                rsCred("Cuenta") = Trim(rsCta("cCodCta"))
'                rsCred("Estado") = IIf(rsCta("cRefinan") = "R", "RF", "")
'                rsCred("Vencimiento") = CDate(rsCta("dFecVenc"))
'                rsCred("CuoVenc") = Trim(rsCta("nNroCuo"))
'                rsCred("Cuota") = rsCta("nCuota")
'                rsCred("Mora") = rsCta("nMora")
'                rsCred("Gastos") = rsCta("nGastos")
'                rsCred("MinPagar") = rsCta("nCuota") + rsCta("nGastos") + rsCta("nMora")
'                rsCred.Update
'                rsCta.MoveNext
'            Loop
'        End If
'        rsCta.Close
'    End If
'Next I
'
''Verifica los Créditos Prendarios
'sAgencia = ""
'sBuffer = ""
'nCuentasPrd = 0
'For I = 1 To UBound(CuentasPersonaPrd, 1)
'    sCuentas = "("
'    sAgencia = Left(CuentasPersonaPrd(I), 2)
'    For J = I To UBound(CuentasPersonaPrd, 1)
'        sBuffer = Left(CuentasPersonaPrd(J), 2)
'        If sBuffer = sAgencia Then
'            sCuentas = sCuentas & "'" & CuentasPersonaPrd(J) & "',"
'        Else
'            Exit For
'        End If
'    Next J
'    sCuentas = Left(sCuentas, Len(sCuentas) - 1)
'    sCuentas = sCuentas & ")"
'    I = J - 1
'    If AbreConeccion(sAgencia & "0000000000") Then
'        If bVencidos Then
'            sVencidos = "And DateDiff(dd,C.dFecVenc,'" & Format$(Date, "mm/dd/yyyy") & "') > " & nDiasAtraso
'        Else
'            sVencidos = ""
'        End If
'        VSQL = "Select C.cCodCta, C.dFecVenc, C.nValTasac, C.nTasaIntVenc, C.nCostCusto, C.cEstado, C.nSaldoCap, " _
'            & "C.nTasaInt, C.nPlazo, " _
'            & "nTasaCustodia = (Select nValor1 From Parametro Where cCodPar = '10102'), " _
'            & "nTasaImpuesto = (Select nValor1 From Parametro Where cCodPar = '10104'), " _
'            & "nCostoPrepRemate = (Select nValor1 From Parametro Where cCodPar = '10107') " _
'            & "From CredPrenda C Where C.cEstado IN ('1','4','6','7') " & sVencidos & " " _
'            & " And C.cCodCta IN " & sCuentas & " Order by C.cCodCta"
'
'        rsCta.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
'        Set rsCta.ActiveConnection = Nothing
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasPrd = nCuentasPrd + 1
'                nDeudaPrend = CalculaDeudaPrendario(rsCta("nSaldoCap"), rsCta("dFecVenc"), rsCta("nValTasac"), rsCta("nTasaIntVenc"), _
'                                    rsCta("nCostCusto"), rsCta("nTasaImpuesto"), rsCta("cEstado"), rsCta("nCostoPrepRemate"))
'                nMinimoPagar = CalculaMinimoPagar(rsCta("nTasaInt"), rsCta("nPlazo"), rsCta("nValTasac"), rsCta("nTasaCustodia"), _
'                                    rsCta("nSaldoCap"), rsCta("cEstado"), rsCta("dFecVenc"), rsCta("nTasaIntVenc"), rsCta("nTasaCustodia"), _
'                                    rsCta("nTasaImpuesto"), rsCta("nCostoPrepRemate"))
'                rsCred.AddNew
'                rsCred("Cuenta") = Trim(rsCta("cCodCta"))
'                rsCred("Estado") = IIf(rsCta("cEstado") = "6", "RE", "")
'                rsCred("Vencimiento") = rsCta("dFecVenc")
'                rsCred("CuoVenc") = 1
'                rsCred("Cuota") = rsCta("nSaldoCap")
'                rsCred("Mora") = nDeudaPrend - rsCta("nSaldoCap")
'                rsCred("Gastos") = 0
'                rsCred("MinPagar") = nMinimoPagar
'                rsCred.Update
'                rsCta.MoveNext
'            Loop
'        End If
'        rsCta.Close
'    End If
'
'Next I
'
'nCuentasJud = 0
'
'If UBound(CuentasPersonaJud, 1) > 0 Then
'    sCuentas = "("
'    For I = 1 To UBound(CuentasPersonaJud, 1)
'        sCuentas = sCuentas & "'" & CuentasPersonaJud(I) & "',"
'    Next I
'    sCuentas = Left(sCuentas, Len(sCuentas) - 1)
'    sCuentas = sCuentas & ")"
'    If AbreConeccion(sAgencia & "0000000000") Then
'        'Verificamos los Creditos Judiciales concentrados en la agencia principal
'        VSQL = "Select C.cCodCta, C.cCondicion From CredCJudi C " _
'            & "Where C.cEstado = 'V' And C.cCodCta IN " & sCuentas
'        rsCta.Open VSQL, dbCmactN, adOpenStatic, adLockReadOnly, adCmdText
'        Set rsCta.ActiveConnection = Nothing
'        If Not (rsCta.EOF And rsCta.BOF) Then
'            Do While Not rsCta.EOF
'                nCuentasJud = nCuentasJud + 1
'                rsCred.AddNew
'                rsCred("Cuenta") = Trim(rsCta("cCodCta"))
'                rsCred("Estado") = "JU"
'                rsCred("Vencimiento") = Null
'                rsCred("CuoVenc") = 0
'                rsCred("Cuota") = 0
'                rsCred("Mora") = 0
'                rsCred("Gastos") = 0
'                rsCred("MinPagar") = 0
'                rsCred.Update
'                rsCta.MoveNext
'            Loop
'        End If
'        rsCta.Close
'    End If
'End If
'
'If (nCuentasCre + nCuentasPrd + nCuentasJud) = 0 Then
'    CreditosPendientes = False
'Else
'    CreditosPendientes = True
'End If

End Function


Public Function GeneraConvenioOrdenPago(ByVal psTexto As String, ByVal sCuenta As String, ByVal dFecha As Date) As String
Dim rsPers As ADODB.Recordset
Dim lsTit As String, lsRep As String
Dim lsDireccion As String, lsDNI As String
Dim sCad As String, sPersona As String, saux As String
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales

Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set rsPers = clsMant.GetPersonaCuenta(sCuenta)
Dim ofun As New COMFunciones.FCOMCadenas

Set clsMant = Nothing

'COMENTADO AVMM
'Open App.Path & "\FormatoCarta\ConOP.txt" For Input As #1
'Do While Not EOF(1)
'    Line Input #1, saux
'    sCad = sCad & Chr$(10) & saux
'Loop
'Close #1

If Not (rsPers.EOF And rsPers.BOF) Then
    lsTit = ""
    lsRep = ""
    sPersona = ""
    Do While Not rsPers.EOF
        If sPersona <> rsPers("cPersCod") Then
            If rsPers("nPrdPersRelac") <> gCapRelPersTitular Then
                If lsRep = "" Then
                    lsRep = ofun.PstaNombre(Trim(rsPers("Nombre")))
                Else
                    lsRep = lsRep & oImp.gPrnSaltoLinea & Space(40) & ofun.PstaNombre(Trim(rsPers("Nombre")))
                End If
            Else
                If lsTit = "" Then
                    lsDNI = rsPers("ID") & " " & rsPers("ID N°")
                    lsTit = ofun.PstaNombre(rsPers("Nombre"))
                    lsDireccion = rsPers("Direccion") & ""
                Else
                    lsTit = lsTit & oImp.gPrnSaltoLinea & Space(52) & ofun.PstaNombre(rsPers("Nombre"))
                End If
            End If
            sPersona = rsPers("cPersCod")
        End If
        rsPers.MoveNext
    Loop
    
    'Datos de la Orden------------------------------------------------------------------------------------------------
      
    Dim lsInicio As String
    Dim lsFin As String
    Dim lsFecha As String
    Dim lsNumTal As String
    Dim lsPersNombre As String
    Dim lsTipo As String
    Dim lsMovNro As String
    
    'Dim oMant As NCapMantenimiento
    'Set oMant = New NCapMantenimiento
    
    Dim lrsOrdenPago As ADODB.Recordset
    Set lrsOrdenPago = New ADODB.Recordset
    
    
    Set lrsOrdenPago = GetOrdenPagoEmitidas(sCuenta)
                            
                            
                            
                            
    lsInicio = lrsOrdenPago("nInicio")
    lsFin = lrsOrdenPago("nFin")
    lsFecha = Format$(lrsOrdenPago("Fecha"), "dd/mm/yyyy")
    
'    lsNumTal = lrsOrdenPago("nNumTal")
'    lsPersNombre = PstaNombre(lrsOrdenPago("cPersNombre"), False)
'    lsTipo = lrsOrdenPago("nTipo")
'    lsMovNro = lrsOrdenPago("cMovNro")
    
    '-----------------------------------------------------------------------------------------------------------------
    sCad = psTexto
    sCad = Replace(sCad, "<<Inicio>>", Format(lsInicio, "#########"))
    sCad = Replace(sCad, "<<Fin>>", Format(lsFin, "#########"))
    sCad = Replace(sCad, "<<Fecha>>", lsFecha)
    sCad = Replace(sCad, "<<NumTal>>", Format(lsNumTal, "#########"))
    sCad = Replace(sCad, "<<PersNombre>>", lsPersNombre)
    sCad = Replace(sCad, "<<Tipo>>", Format(lsTipo, "#########"))
    sCad = Replace(sCad, "<<MovNro>>", lsMovNro)
    
    sCad = Replace(sCad, "<<NOMBRE>>", lsTit)
    sCad = Replace(sCad, "<<DNI>>", lsDNI)
    sCad = Replace(sCad, "<<DIRECCION>>", lsDireccion)
    sCad = Replace(sCad, "<<REPRESENTANTE>>", lsRep)
    sCad = Replace(sCad, "<<DIA>>", Format$(dFecha, "dddd dd"))
    sCad = Replace(sCad, "<<MES>>", Format$(dFecha, "mmmm"))
    sCad = Replace(sCad, "<<ANIO>>", Format$(dFecha, "yyyy"))
    
End If
rsPers.Close
Set rsPers = Nothing
GeneraConvenioOrdenPago = sCad
End Function

Private Sub GeneraConvenioOP(sCuenta As String)

    
    On Error GoTo GeneraMovNroErr
    
    Dim rs As ADODB.Recordset
    Dim sql As String
    Dim oConn As COMConecta.DCOMConecta
    
    Set oConn = New COMConecta.DCOMConecta
    
    Dim oError As New COMConecta.COMErrorHandling
    
    If oConn.AbreConexion = False Then Exit Sub
    sql = "SP_CapConvenioOrdenPago '" & sCuenta & "'"
    Set rs = oConn.Ejecutar(sql)
    Set rs = Nothing
    
    Exit Sub
GeneraMovNroErr:
    Call oError.RaiseError(oError.MyUnhandledError, "GeneraConvenioOP")
                
End Sub

Private Sub EliminaConvenioOP()

    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As COMConecta.DCOMConecta
    
    Set rs = New ADODB.Recordset
    Set oConect = New COMConecta.DCOMConecta
            
                    
    If oConect.AbreConexion = False Then Exit Sub
    
    sql = "Delete From tmpConvenioOP"
    
    Set rs = oConect.CargaRecordSet(sql)
    
    oConect.CierraConexion
    Set oConect = Nothing

End Sub

Public Sub BloqueCuentaTotal(ByVal sCuenta As String, ByVal nMotivo As CaptacMotBloqueoTot, _
        ByVal sComentario As String, ByVal sMovNro As String)
Dim clsBlq As COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean

bTrans = False
On Error GoTo ErrBlq

Set clsBlq = New COMDCaptaGenerales.DCOMCaptaGenerales

clsBlq.dbCmact.BeginTrans
bTrans = True
clsBlq.NuevoBloqueototal sCuenta, nMotivo, sComentario, sMovNro
clsBlq.ActualizaEstadoCuenta sCuenta, gCapEstBloqTotal
clsBlq.dbCmact.CommitTrans
Exit Sub
ErrBlq:
    If bTrans Then clsBlq.dbCmact.RollbackTrans
    Set clsBlq = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Sub
'By Capi 07082008 se adiciono parametros opcionales bTasaPactada y bCambioTasa
Public Sub CambioTasa(ByVal sCuenta As String, ByVal nTasaAnt As Double, ByVal nTasa As Double, ByVal cMovNro As String, Optional ByVal bTasaPactada As Boolean = True, Optional ByVal bCambioTasa As Boolean)
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim bTrans As Boolean
    
    bTrans = False
    On Error GoTo ErrTasa
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    clsCap.dbCmact.BeginTrans
    bTrans = True
    'By Capi 13082008
    If bCambioTasa = True Then
        clsCap.CambiaTasaInteres sCuenta, nTasaAnt, nTasa, cMovNro, bTasaPactada
    Else
        clsCap.CambiaTasaInteres sCuenta, nTasa, nTasa, cMovNro, bTasaPactada
    End If
    '
    clsCap.dbCmact.CommitTrans
    Set clsCap = Nothing
    Exit Sub
ErrTasa:
    If bTrans = True Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Sub

Private Sub Class_Initialize()
IniciaImpresora
End Sub

Private Sub Class_Terminate()
Set oImp = Nothing
End Sub
'JACA 20111018************************************************************************
Public Function obtenerTarifasOpeEnOtrasAgencias(psAgencia As String, psTipoAhorro As String, pnMoneda As Integer) As Recordset
    Dim clsCap As New COMDCaptaGenerales.DCOMCaptaGenerales
   
    On Error GoTo ErrorMsg
    
    Set obtenerTarifasOpeEnOtrasAgencias = clsCap.getTarifasOpeEnOtrasAgencias(psAgencia, psTipoAhorro, pnMoneda)
    
    Set clsCap = Nothing
    Exit Function
ErrorMsg:
    
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description

End Function
Public Function guardarTarifaOpeEnOtrasAge(pRs As Recordset, psAgencias As String, psTpoPrograma As String, pnMoneda As Integer, psMovNro As String) As Boolean
    Dim objCapD As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim i As Integer
    Dim bTrans As Boolean
    
    Set objCapD = New COMDCaptaGenerales.DCOMCaptaGenerales
    
    On Error GoTo ErrorMsg
    
    objCapD.dbCmact.BeginTrans
    bTrans = True
    
    'Se actualizan los estados a inactivos de las tarifas a modificar
    objCapD.updEstadoTarifaOpeEnOtrasAge psAgencias, psTpoPrograma, pnMoneda
    
    'Guardamos el nuevo tarifario
    For i = 0 To pRs.RecordCount - 1
        If pRs!ok = 1 Then
            objCapD.insertarTarifaOpeEnOtrasAge pRs(6), pRs(7), pnMoneda, pRs(3), pRs(4), pRs(5), psMovNro
        End If
        pRs.MoveNext
    Next i
    
    objCapD.dbCmact.CommitTrans
    bTrans = False
    
    guardarTarifaOpeEnOtrasAge = True
    Set objCapD = Nothing
    
    Exit Function
ErrorMsg:
    
    If bTrans Then objCapD.dbCmact.RollbackTrans
    Set objCapD = Nothing
    guardarTarifaOpeEnOtrasAge = False
    Err.Raise Err.Number, "Error", Err.Description

End Function
'Public Function obtenerValTarifaOpeEnOtrasAge(psAgencia As String, pnTipoAhorro As Integer, pnMoneda As Integer) As Recordset
Public Function obtenerValTarifaOpeEnOtrasAge(psCtaCod As String, Optional ByVal pnOperacion As Integer = 0) As Recordset 'APRI20190109 ERS077-2018
    Dim clsCap As New COMDCaptaGenerales.DCOMCaptaGenerales
   
    On Error GoTo ErrorMsg
    
    'Set obtenerValTarifaOpeEnOtrasAge = clsCap.getValTarifaOpeEnOtrasAge(psAgencia, pnTipoAhorro, pnMoneda)
    'Set obtenerValTarifaOpeEnOtrasAge = clsCap.getValTarifaOpeEnOtrasAge(psCtaCod) 'APRI20190109 ERS077-2018
    Set obtenerValTarifaOpeEnOtrasAge = clsCap.getValTarifaOpeEnOtrasAge(psCtaCod, pnOperacion) 'APRI20210403
    
    Set clsCap = Nothing
    Exit Function
ErrorMsg:
    
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description

End Function
Public Function obtenerOperaciones(Optional ByVal psFiltro As String = "") As Recordset
    Dim clsCap As New COMDCaptaGenerales.DCOMCaptaGenerales
   
    On Error GoTo ErrorMsg
    
    Set obtenerOperaciones = clsCap.getOperaciones(psFiltro)
    
    Set clsCap = Nothing
    Exit Function
ErrorMsg:
    
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description

End Function
Public Function obtenerTarifasOperaciones(psOpeCod As String, psAgencia As String, psTipoAhorro As String) As Recordset
    Dim clsCap As New COMDCaptaGenerales.DCOMCaptaGenerales
   
    On Error GoTo ErrorMsg
    
    Set obtenerTarifasOperaciones = clsCap.getTarifasOperaciones(psOpeCod, psAgencia, psTipoAhorro)
    
    Set clsCap = Nothing
    Exit Function
ErrorMsg:
    
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description

End Function
Public Function guardarTarifasOperaciones(pRs As Recordset, psOpeCod As String, psAgencias As String, psTpoPrograma As String, psMovNro As String) As Boolean
    Dim objCapD As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim i As Integer
    Dim bTrans As Boolean
    
    Set objCapD = New COMDCaptaGenerales.DCOMCaptaGenerales
    
    On Error GoTo ErrorMsg
    
    objCapD.dbCmact.BeginTrans
    bTrans = True
    
    
    'Se actualizan los estados a inactivos de las tarifas a modificar
    objCapD.updEstadoTarifaOperaciones psOpeCod, psAgencias, psTpoPrograma
    
    'Guardamos el nuevo tarifario
    For i = 0 To pRs.RecordCount - 1
        If pRs!ok = 1 Then
            'objCapD.insertarTarifaOperaciones pRs(8), pRs(6), pRs(7), pRs(4), pRs(5), psMovNro
            objCapD.insertarTarifaOperaciones pRs!copecod, pRs!cAgeCod, pRs!nTpoPrograma, pRs(4), pRs(5), pRs(6), psMovNro
        End If
        pRs.MoveNext
    Next i
    
    objCapD.dbCmact.CommitTrans
    bTrans = False
    
    guardarTarifasOperaciones = True
    Set objCapD = Nothing
    
    Exit Function
ErrorMsg:
    
    If bTrans Then objCapD.dbCmact.RollbackTrans
    Set objCapD = Nothing
    guardarTarifasOperaciones = False
    Err.Raise Err.Number, "Error", Err.Description

End Function

Public Function obtenerNroMovimientos(ByVal psCtaCod As String, ByVal psFecha As Date) As Integer
    Dim sql As String
    Dim clsCap As New COMDCaptaGenerales.DCOMCaptaGenerales
     On Error GoTo ErrorMsg
    
     obtenerNroMovimientos = clsCap.getNroMovimientos(psCtaCod, psFecha)
    
    Set clsCap = Nothing
    Exit Function
ErrorMsg:
    
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description
    
End Function
Public Function obtenerComisionxNroOperaciones(psOpeCod As String, psAgencia As String, pnTipoAhorro As Integer) As Recordset
    Dim clsCap As New COMDCaptaGenerales.DCOMCaptaGenerales
   
    On Error GoTo ErrorMsg
    
    Set obtenerComisionxNroOperaciones = clsCap.getComisionxNroOperaciones(psOpeCod, psAgencia, pnTipoAhorro)
    
    Set clsCap = Nothing
    Exit Function
ErrorMsg:
    
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description

End Function
'JACA END*****************************************************************************
'***Agregado por ELRO el 20120627, según OYP-RFC024-2012
Public Function obtenerCajaBancosOperacionesPendientes(ByVal pdFechaVoucher As Date, _
                                                       ByVal psIFTpo As String, _
                                                       ByVal psPersCod As String, _
                                                       ByVal psCtaIFCod As String, _
                                                       ByVal pnMovImporte As Currency, _
                                                       ByVal psMoneda As String) As ADODB.Recordset
    Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales

    On Error GoTo ErrorMsg
    Set obtenerCajaBancosOperacionesPendientes = oDCOMCaptaGenerales.devolverCajaBancosOperacionesPendientes(Format(pdFechaVoucher, "yyyyMMdd"), _
                                                                                                             psIFTpo, _
                                                                                                             psPersCod, _
                                                                                                             psCtaIFCod, _
                                                                                                             pnMovImporte, _
                                                                                                             psMoneda)
    Set oDCOMCaptaGenerales = Nothing
    Exit Function

ErrorMsg:
    Set oDCOMCaptaGenerales = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Function


Public Function obtenerMotivoCapVoucherDeposito() As ADODB.Recordset
    Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales

    On Error GoTo ErrorMsg
    Set obtenerMotivoCapVoucherDeposito = oDCOMCaptaGenerales.devolverMotivoCapVoucherDeposito()
    Set oDCOMCaptaGenerales = Nothing
    Exit Function

ErrorMsg:
    Set oDCOMCaptaGenerales = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function guardarVoucherDeposito(ByVal pcMovNroRVD As String, _
                                       ByVal pdFecReg As Date, _
                                       ByVal psPersCodCli As String, _
                                       ByVal psIFTpo As String, _
                                       ByVal psPersCodIF As String, _
                                       ByVal psCtaIFCod As String, _
                                       ByVal pnTipMot As Integer, _
                                       ByVal pbConfirmado As Boolean, _
                                       ByVal psNroVou As String, _
                                       ByVal pnMonVou As Currency, _
                                       ByVal pdFecVou As Date, _
                                       ByVal psAgeCod As String, _
                                       ByVal psMoneda As String, _
                                       ByVal psGlosa As String, _
                                       ByVal psOpeCod As String, _
                                       ByVal psCtaContBanco As String, _
                                       ByVal psSubCtaContBanco As String, _
                                       Optional ByVal pnMovNroPen As Long = 0) As Long

Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean
Dim lnMovNroRVD As Long

Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True
clsMant.AgregaMovNoContable pcMovNroRVD, psGlosa, psOpeCod
lnMovNroRVD = clsMant.GetnMovNro(pcMovNroRVD)
clsMant.registrarCapVoucherDeposito lnMovNroRVD, Format(pdFecReg, "yyyyMMdd"), psPersCodCli, psIFTpo, psPersCodIF, psCtaIFCod, pnTipMot, IIf(pbConfirmado, 1, 0), psNroVou, pnMonVou, Format(pdFecVou, "yyyyMMdd"), psAgeCod, psMoneda, pcMovNroRVD, psCtaContBanco, psSubCtaContBanco
If pnMovNroPen > 0 Then
    clsMant.registrarMovRef lnMovNroRVD, pnMovNroPen
End If
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
guardarVoucherDeposito = lnMovNroRVD
Exit Function
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    Err.Raise Err.Number, "Error", Err.Description
    guardarVoucherDeposito = False
End Function

Public Function eliminarVoucherDeposito(ByVal pcMov As String, _
                                        ByVal pnMovNroRVD As Long, Optional ByVal pnId As Long = 0) As Long

Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean

Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True
'clsMant.eliminarCapVoucherDeposito pcMov, pnMovNroRVD
clsMant.eliminarCapVoucherDeposito pcMov, pnMovNroRVD, pnId 'EJVG20130912
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
'eliminarVoucherDeposito = pnMovNroRVD
eliminarVoucherDeposito = pnMovNroRVD + pnId 'EJVG20130912
Exit Function
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    eliminarVoucherDeposito = 0
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function obtenerVoucherSinOperacion(ByVal psAgeCod As String, _
                                           ByVal psMoneda As String) As ADODB.Recordset
    Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales

    On Error GoTo ErrorMsg
    Set obtenerVoucherSinOperacion = oDCOMCaptaGenerales.devolverVoucherSinOperacion(psAgeCod, psMoneda)
    Set oDCOMCaptaGenerales = Nothing
    Exit Function

ErrorMsg:
    Set oDCOMCaptaGenerales = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function actualizarVoucherDeposito(ByVal pnConfirmado As Integer, _
                                          ByVal psMov As String, _
                                          ByVal pnMovNroRVD As Long, Optional ByVal pnId As Long = 0) As Long

Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean
Dim lnMovNroRVD As Long

Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
On Error GoTo ErrGraba
clsMant.dbCmact.BeginTrans
bTrans = True
'clsMant.confirmarCapVoucherDeposito pnConfirmado, psMov, pnMovNroRVD
clsMant.confirmarCapVoucherDeposito pnConfirmado, psMov, pnMovNroRVD, pnId 'EJVG20130911
clsMant.dbCmact.CommitTrans
bTrans = False
Set clsMant = Nothing
'actualizarVoucherDeposito = pnMovNroRVD
actualizarVoucherDeposito = pnMovNroRVD + pnId
Exit Function
ErrGraba:
    If bTrans Then clsMant.dbCmact.RollbackTrans
    Set clsMant = Nothing
    Err.Raise Err.Number, "Error", Err.Description
    actualizarVoucherDeposito = False
End Function

Public Function obtenerVoucherIFSinOperacion(ByVal psAgeCod As String, _
                                              ByVal psMoneda As String, _
                                              ByVal pnTipMot As Integer, Optional ByVal psDetalle As String = "") As ADODB.Recordset
    Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales

    On Error GoTo ErrorMsg
    'Set obtenerVoucherIFSinOperacion = oDCOMCaptaGenerales.devolverVoucherIFSinOperacion(psAgeCod, psMoneda, pnTipMot)
    Set obtenerVoucherIFSinOperacion = oDCOMCaptaGenerales.devolverVoucherIFSinOperacion(psAgeCod, psMoneda, pnTipMot, psDetalle) 'EJVG20130914
    Set oDCOMCaptaGenerales = Nothing
    Exit Function

ErrorMsg:
    Set oDCOMCaptaGenerales = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function obtenerVoucherNroIFSinOperacion(ByVal psAgeCod As String, _
                                                ByVal psMoneda As String, _
                                                ByVal pnTipMot As Integer, _
                                                ByVal psIFTpo As String, _
                                                ByVal psPersCodIF As String, _
                                                Optional ByVal psDetalle As String = "") As ADODB.Recordset
    Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales

    On Error GoTo ErrorMsg
    'Set obtenerVoucherNroIFSinOperacion = oDCOMCaptaGenerales.devolverVoucherNroIFSinOperacion(psAgeCod, psMoneda, pnTipMot, psIFTpo, psPersCodIF)
    Set obtenerVoucherNroIFSinOperacion = oDCOMCaptaGenerales.devolverVoucherNroIFSinOperacion(psAgeCod, psMoneda, pnTipMot, psIFTpo, psPersCodIF, psDetalle) 'EJVG20130914
    Set oDCOMCaptaGenerales = Nothing
    Exit Function

ErrorMsg:
    Set oDCOMCaptaGenerales = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function obtenerVoucherClienteNroIFSinOperacion(ByVal psAgeCod As String, _
                                                       ByVal psMoneda As String, _
                                                       ByVal pnTipMot As Integer, _
                                                       ByVal psIFTpo As String, _
                                                       ByVal psPersCodIF As String, _
                                                       ByVal psCtaIFCod As String, _
                                                       ByVal psNroVou As String) As ADODB.Recordset
    Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales

    On Error GoTo ErrorMsg
    Set obtenerVoucherClienteNroIFSinOperacion = oDCOMCaptaGenerales.devolverVoucherClienteNroIFSinOperacion(psAgeCod, psMoneda, pnTipMot, psIFTpo, psPersCodIF, psCtaIFCod, psNroVou)
    Set oDCOMCaptaGenerales = Nothing
    Exit Function

ErrorMsg:
    Set oDCOMCaptaGenerales = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Function
'***Fin Agregado por ELRO******************************
'***Agregado por ELRO 20121010, según OYP-RFC101-2012
Public Function obtenerHistorialCaptacSueldosCTS(ByVal psCuenta As String, Optional ByRef psUltimaRemuneracion As Currency = 0#) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerHistorialCaptacSueldosCTS = oDCOMCaptaGenerales.devolverHistorialCaptacSueldosCTS(psCuenta, psUltimaRemuneracion)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function obtenerListadoCuentasCTS(ByVal psPersCod As String) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerListadoCuentasCTS = oDCOMCaptaGenerales.devolverListadoCuentasCTS(psPersCod)
Set oDCOMCaptaGenerales = Nothing
End Function
'***Fin Agregado por ELRO 20121010*******************
'***Agregado por ELRO 20130111, según OYP-RFC123-2012
Public Function obtenerHistoricoRenovacionesDPF(ByVal psCuenta As String) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerHistoricoRenovacionesDPF = oDCOMCaptaGenerales.devolverHistoricoRenovacionesDPF(psCuenta)
Set oDCOMCaptaGenerales = Nothing
End Function
'***Fin Agregado por ELRO 20130111*******************
'***Agregado por ELRO el 20130323, según TI-ERS011-2013****
Public Function obtenerCuentasPersona(ByVal psPersCod As String) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerCuentasPersona = oDCOMCaptaGenerales.devolverCuentasPersona(psPersCod)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function obtenerDatosCuentaPersona(ByVal psCtaCod As String) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerDatosCuentaPersona = oDCOMCaptaGenerales.devolverDatosCuentaPersona(psCtaCod)
Set oDCOMCaptaGenerales = Nothing
End Function


Public Function obtenerSubProductoDestino() As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerSubProductoDestino = oDCOMCaptaGenerales.devolverSubProductoDestino
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function obtenerTitularCuenta(ByVal psCtaCod As String) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerTitularCuenta = oDCOMCaptaGenerales.devolverTitularCuenta(psCtaCod)
Set oDCOMCaptaGenerales = Nothing
End Function

'***Fin Agregado por ELRO el 20130323, según TI-ERS011-2013
'JUEZ 20130520************************************************************************
Public Sub GuardarRegistroEnvioEstadoCta(ByVal pnTipoProd As Integer, ByVal psCtaCod As String, ByVal prsEnvEstCta As ADODB.Recordset, _
                                         ByVal pnModoEnvioEstCta As Integer, ByVal pnDebitoMismaCta As Integer, ByVal psMovNro As String)
Dim oCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oPers As COMDPersona.DCOMPersona
Set oCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
Set oPers = New COMDPersona.DCOMPersona

'If Not prsEnvEstCta.EOF Then'WIOR 20130621
Call oCapGen.EliminarRegistroEnvioEstadoCta(psCtaCod)
prsEnvEstCta.MoveFirst
Do While Not prsEnvEstCta.EOF
'COMENTADO POR APRI20180312 ERS036-2017
'    If pnModoEnvioEstCta = 2 Then
'         If prsEnvEstCta!Envio = "1" Then
'            Call oCapGen.InsertarRegistroEnvioEstadoCta(psCtaCod, prsEnvEstCta!CODIGO, pnModoEnvioEstCta, IIf(pnDebitoMismaCta = 1, psCtaCod, Trim(prsEnvEstCta!cuenta)), psMovNro)
'            Call oPers.ActualizaPersDireccion(prsEnvEstCta!CODIGO, UCase(prsEnvEstCta!Domicilio))
'        End If
'    Else
'        Call oCapGen.InsertarRegistroEnvioEstadoCta(psCtaCod, prsEnvEstCta!CODIGO, pnModoEnvioEstCta, "", psMovNro)
'    End If
       
       'APRI20180312 ERS036-2017
       If prsEnvEstCta!Envio = "1" Then
       Call oCapGen.InsertarRegistroEnvioEstadoCta(psCtaCod, prsEnvEstCta!CODIGO, pnModoEnvioEstCta, IIf(pnDebitoMismaCta = 1, psCtaCod, ""), psMovNro)
       Call oPers.ActualizaPersDireccion(prsEnvEstCta!CODIGO, IIf(pnModoEnvioEstCta = 2, UCase(prsEnvEstCta!Domicilio), UCase(prsEnvEstCta!Correo)), pnModoEnvioEstCta)
       End If
       'END APRI
       
    prsEnvEstCta.MoveNext
Loop
Set oCapGen = Nothing
End Sub
'END JUEZ ****************************************************************************
'APRI20180409 ERS036-2017
Public Function CargarProductosVigentes(ByVal psCodPers As String, ByVal pnTipo As Integer, ByVal pnTipoOpe As Integer) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set CargarProductosVigentes = oDCOMCaptaGenerales.CargarProductosVigentes(psCodPers, pnTipo, pnTipoOpe)
Set oDCOMCaptaGenerales = Nothing
End Function
Public Function AfiliarDesafiliarEnvioEstadoCta(ByVal pnTipo As Integer, ByVal pMatCtaCod As Variant, ByVal pnCantMant As Integer, Optional ByVal psPersCod As String, Optional ByVal psDireccion As String, _
                                         Optional ByVal pnModoEnvioEstCta As Integer, Optional ByVal psMovNro As String, Optional ByVal pnCanal As Integer) As String
                                         'APRI20190615 ADD pnCanal
Dim oCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim oPers As COMDPersona.DCOMPersona
Dim i As Integer
Set oCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
Set oPers = New COMDPersona.DCOMPersona
If pnTipo = 1 Then
    For i = 0 To pnCantMant - 1
    Call oCapGen.EliminarRegistroEnvioEstadoCta(pMatCtaCod(i, 0))
    Call oCapGen.InsertarRegistroEnvioEstadoCta(pMatCtaCod(i, 0), psPersCod, pnModoEnvioEstCta, IIf(Trim(Mid(pMatCtaCod(i, 0), 6, 3)) = "232" Or Trim(Mid(pMatCtaCod(i, 0), 6, 3)) = "233" Or Trim(Mid(pMatCtaCod(i, 0), 6, 3)) = "234", pMatCtaCod(i, 0), ""), psMovNro, pnTipo, pnCanal)
    Call oPers.ActualizaPersDireccion(psPersCod, psDireccion, pnModoEnvioEstCta)
    Next i
Else
    For i = 0 To pnCantMant - 1
        Call oCapGen.EliminarRegistroEnvioEstadoCta(pMatCtaCod(i, 0), psPersCod, 1, psMovNro, pnCanal)
    Next i
End If
    AfiliarDesafiliarEnvioEstadoCta = oCapGen.ObtenerNroItemEstadoCta(psPersCod, psMovNro, pnTipo)  'APRI20190615 SATI RFC1902040001
Set oCapGen = Nothing
End Function
'FIN APRI
'***Agregado por ELRO el 20130701, según RFC1306270002****
Public Function obtenerCuentasPersonaServicioPago(ByVal psPersCod As String) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerCuentasPersonaServicioPago = oDCOMCaptaGenerales.devolverCuentasPersonaServicioPago(psPersCod)
Set oDCOMCaptaGenerales = Nothing
End Function
Public Function guardarServicioPago(ByVal psPersCodSerPag As String, ByVal psNomSerPag As String, _
                               ByVal psDesSerPag As String, ByVal psCtaCod As String, _
                               ByVal pnComision As Currency, ByVal pnPenalidad As Currency, _
                               ByVal pnLimMaxOpeDia As Long, psFecReg As String, _
                               ByVal psMovNro As String, Optional ByVal pnIdSerPag As Long = 0, _
                               Optional pnComisionTransf As Currency = 0) As Long
                               'RIRO20150430 ERS146-2014 ADD pnComisionTransf
                               'RIRO20150814 Se cambió pnLimMaxOpeDia a Long

Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Dim rsServicioPago As New ADODB.Recordset
If pnIdSerPag = 0 Then

   'RIRO20150430 ERS146-2014 ADD pnComisionTransf
   guardarServicioPago = oDCOMCaptaGenerales.registrarServicioPago(psPersCodSerPag, psNomSerPag, psDesSerPag, _
                                                                  psCtaCod, pnComision, pnPenalidad, _
                                                                  pnLimMaxOpeDia, psFecReg, psMovNro, pnComisionTransf)
Else
    'RIRO20150430 ERS146-2014 ADD pnComisionTransf
    guardarServicioPago = oDCOMCaptaGenerales.modificarServicioPago(psNomSerPag, psDesSerPag, _
                                                                   psCtaCod, pnComision, pnPenalidad, _
                                                                   pnLimMaxOpeDia, psMovNro, pnIdSerPag, pnComisionTransf)
End If



Set oDCOMCaptaGenerales = Nothing
Set rsServicioPago = Nothing
End Function

Public Function obtenerConveniosPersona(ByVal psPersCod As String) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerConveniosPersona = oDCOMCaptaGenerales.devolverConveniosPersona(psPersCod)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function obtenerConvenio(ByVal pnIdSerPag As Long) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerConvenio = oDCOMCaptaGenerales.devolverConvenio(pnIdSerPag)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function confirmarNombreConvenio(ByVal psPersCodSerPag As String, ByVal psNomSerPag As String) As Long
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
confirmarNombreConvenio = oDCOMCaptaGenerales.verificarNombreConvenio(psPersCodSerPag, psNomSerPag)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function darBajaServicioPago(ByVal pnIdSerPag As Long, ByVal psMovNro As String) As Long
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
darBajaServicioPago = oDCOMCaptaGenerales.darBajaServicioPago(pnIdSerPag, psMovNro)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function obtenerConvenioVigente(ByVal psPersCodSerPag As String, ByVal pnNroSerPag As Integer) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerConvenioVigente = oDCOMCaptaGenerales.devolverConvenioVigente(psPersCodSerPag, pnNroSerPag)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function guardarDebitoConvenioServicioPago(ByVal pnIdSerPag As Long, ByVal psRefArc As String, _
                                                  ByVal psNomArc As String, pnMonto As Currency, _
                                                  ByVal psFecReg As String, psMovNro As String, _
                                                  ByVal rsBeneficiario As ADODB.Recordset) As Long

Dim oDCOMCaptaGenerales As COMDCaptaGenerales.DCOMCaptaGenerales
Set oDCOMCaptaGenerales = New COMDCaptaGenerales.DCOMCaptaGenerales
Dim bTrans As Boolean
Dim lnConfirmar1, lnConfirmar2 As Long

On Error GoTo ErrGraba

oDCOMCaptaGenerales.dbCmact.BeginTrans
bTrans = True

lnConfirmar1 = oDCOMCaptaGenerales.registrarArchivoServicioPago(pnIdSerPag, psRefArc, psNomArc, pnMonto, psFecReg, psMovNro)

If lnConfirmar1 > 0 Then
    Do While Not rsBeneficiario.EOF
        lnConfirmar2 = oDCOMCaptaGenerales.registrarDebitoServicioPago(lnConfirmar1, rsBeneficiario!cPersCod, rsBeneficiario("Monto S/."))
        If lnConfirmar2 = 0 Then
            If bTrans Then oDCOMCaptaGenerales.dbCmact.RollbackTrans
            Set oDCOMCaptaGenerales = Nothing
            guardarDebitoConvenioServicioPago = 0
            Exit Function
        End If
        rsBeneficiario.MoveNext
    Loop
Else
    If bTrans Then oDCOMCaptaGenerales.dbCmact.RollbackTrans
    Set oDCOMCaptaGenerales = Nothing
    guardarDebitoConvenioServicioPago = 0
    Exit Function
End If

oDCOMCaptaGenerales.dbCmact.CommitTrans
bTrans = False
Set oDCOMCaptaGenerales = Nothing
guardarDebitoConvenioServicioPago = lnConfirmar1
Exit Function
ErrGraba:
    If bTrans Then oDCOMCaptaGenerales.dbCmact.RollbackTrans
    Set oDCOMCaptaGenerales = Nothing
    guardarDebitoConvenioServicioPago = 0
    Err.Raise Err.Number, "Error", Err.Description
End Function

Public Function buscarConvenioServicioPago(ByVal pnTipo As Integer, ByVal psBusqueda As String) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set buscarConvenioServicioPago = oDCOMCaptaGenerales.buscarConvenioServicioPago(pnTipo, psBusqueda)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function obtenerTramaConvenioServicioPago(ByVal pnIdSerPag As Long) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerTramaConvenioServicioPago = oDCOMCaptaGenerales.devolverTramaConvenioServicioPAgo(pnIdSerPag)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function darBajaTramaConvenioServicioPago(ByVal pnIdSerPagArc As Long, ByVal psMovNro As String, _
                                                 ByVal psCtaCod As String, ByVal pnMonto As Currency) As Long
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
darBajaTramaConvenioServicioPago = oDCOMCaptaGenerales.darBajaTramaConvenioServicioPago(pnIdSerPagArc, psMovNro, psCtaCod, pnMonto)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function obtenerBeneficiariosNoCobroTramaBajaConvenioServicioPago(ByVal pnIdSerPagArc As Long) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerBeneficiariosNoCobroTramaBajaConvenioServicioPago = oDCOMCaptaGenerales.devolverBeneficiariosNoCobroTramaBajaConvenioServicioPago(pnIdSerPagArc)
Set oDCOMCaptaGenerales = Nothing
End Function

Public Function obtenerBeneficiarioConvenioServicioPago(ByVal psPersIDnro As String, ByVal pnIdSerPag As Long) As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set obtenerBeneficiarioConvenioServicioPago = oDCOMCaptaGenerales.devolverBeneficiarioConvenioServicioPago(psPersIDnro, pnIdSerPag)
Set oDCOMCaptaGenerales = Nothing
End Function
'***Fin Agregado por ELRO el 20130701, según RFC1306270002
'JUEZ 20130723 *******************************************
Public Function VerificarExisteCuentaCTS(ByVal psPersCod As String, ByVal psInstitucion As String, ByVal psMoneda As Integer, Optional ByVal psCtaCod As String = "") As Boolean
    Dim oDCOMCap As New COMDCaptaGenerales.DCOMCaptaGenerales
    VerificarExisteCuentaCTS = oDCOMCap.VerificaExisteCuentaCTS(psPersCod, psInstitucion, psMoneda, psCtaCod)
    Set oDCOMCap = Nothing
End Function
Public Function ObtenerCuentaCTS(ByVal psPersCod As String, ByVal psInstitucion As String, ByVal pnMoneda As Integer) As ADODB.Recordset
    Dim oDCOMCap As New COMDCaptaGenerales.DCOMCaptaGenerales
    Set ObtenerCuentaCTS = oDCOMCap.ObtenerCuentaCTS(psPersCod, psInstitucion, pnMoneda)
    Set oDCOMCap = Nothing
End Function
'END JUEZ ************************************************
'EJVG20130909 ***
Public Function ExisteVoucherDeposito(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String, ByVal psVoucherNro As String) As Boolean
    Dim clsMant As New COMDCaptaGenerales.DCOMCaptaGenerales
    ExisteVoucherDeposito = clsMant.ExisteVoucherDeposito(psIFTpo, psPersCod, psCtaIFCod, psVoucherNro)
    Set clsMant = Nothing
End Function
Public Function RecuperaVoucherxId(ByVal pnId As Long) As ADODB.Recordset
    Dim clsMant As New COMDCaptaGenerales.DCOMCaptaGenerales
    Set RecuperaVoucherxId = clsMant.RecuperaVoucherxId(pnId)
    Set clsMant = Nothing
End Function
Public Function RecuperaVoucherOpexId(ByVal pnId As Long) As ADODB.Recordset
    Dim clsMant As New COMDCaptaGenerales.DCOMCaptaGenerales
    Set RecuperaVoucherOpexId = clsMant.RecuperaVoucherOpexId(pnId)
    Set clsMant = Nothing
End Function
Public Function InsertarVoucherDeposito(ByVal pdFecReg As Date, ByVal psAgeCod As String, ByVal psUserCod As String, ByVal psOpeCod As String, _
                                       ByVal psPersCodCli As String, _
                                       ByVal psIFTpo As String, ByVal psPersCodIF As String, ByVal psCtaIFCod As String, _
                                       ByVal psCtaContBanco As String, ByVal psSubCtaContBanco As String, _
                                       ByVal psMoneda As String, ByVal psNroVou As String, ByVal pnMonVou As Currency, ByVal pdFecVou As Date, ByVal pbConfirmado As Boolean, _
                                       ByVal pnMovNroPen As Long, ByRef pMatOperaciones As Variant, Optional ByVal poClsMant As COMDCaptaGenerales.DCOMCaptaGenerales = Nothing, _
                                       Optional ByRef pnIdNew As Long = 0) As Boolean
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim ClsMov As COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim bTrans As Boolean
    Dim lnMovNroRVD As Long, lnId As Long
    Dim i As Integer, iSleep As Integer
    Dim lsMovNroRVD As String
    
    Dim lsVoucherNro As String
    Dim lnVoucherTpoOpe As Integer
    Dim lsVoucherMonto As Currency
    Dim lsVoucherEstado As String
    Dim lsVoucherDet As String
    Dim lsVoucherGlosa As String
    
    On Error GoTo ErrGraba
    If poClsMant Is Nothing Then
        Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        clsMant.dbCmact.BeginTrans
    Else
        Set clsMant = poClsMant
    End If
    bTrans = True
    
    'Registro Voucher
    lnId = clsMant.InsertarCapVoucherDepositoCab(Format(pdFecReg, "yyyyMMdd"), psPersCodCli, psIFTpo, psPersCodIF, psCtaIFCod, IIf(pbConfirmado, 1, 0), psNroVou, pnMonVou, Format(pdFecVou, "yyyyMMdd"), psAgeCod, psMoneda, psCtaContBanco, psSubCtaContBanco)
    If lnId = 0 Then
        If poClsMant Is Nothing Then
            clsMant.dbCmact.RollbackTrans
            Set clsMant = Nothing
        End If
    End If
    
    Set ClsMov = New COMDCaptaGenerales.DCOMCaptaMovimiento
    'Registro Detalle Voucher
    For i = 1 To UBound(pMatOperaciones, 2)
        lsVoucherNro = pMatOperaciones(1, i)
        lnVoucherTpoOpe = CInt(Trim(Right(pMatOperaciones(2, i), 2)))
        lsVoucherMonto = pMatOperaciones(3, i)
        lsVoucherEstado = pMatOperaciones(4, i)
        lsVoucherDet = pMatOperaciones(5, i)
        lsVoucherGlosa = pMatOperaciones(6, i)
        
        For iSleep = 1 To 1000
        Next
        
        lsMovNroRVD = clsMant.GeneraMovNro(pdFecReg, psAgeCod, psUserCod)
        clsMant.AgregaMovNoContable lsMovNroRVD, lsVoucherGlosa, psOpeCod
        lnMovNroRVD = clsMant.GetnMovNro(lsMovNroRVD)
        clsMant.registrarCapVoucherDeposito lnMovNroRVD, Format(pdFecReg, "yyyyMMdd"), psPersCodCli, psIFTpo, psPersCodIF, psCtaIFCod, lnVoucherTpoOpe, IIf(pbConfirmado, 1, 0), lsVoucherNro, lsVoucherMonto, Format(pdFecVou, "yyyyMMdd"), psAgeCod, psMoneda, lsMovNroRVD, psCtaContBanco, psSubCtaContBanco
                
        'RIRO20140407 ERS017 Se agregaron opciones 8,9,10,11
        Select Case lnVoucherTpoOpe
        Case 1, 3, 5 'Aperturas
            clsMant.ActualizarCapVoucherDeposito lnMovNroRVD, , , lsVoucherDet
            
        Case 2, 4, 6, 10, 11, 13, 18, 19, 20, 21 'Depósitos y Aumento de Capital'WIOR 20160413 AGREGO CODIGO 13 LIQ SEG DESG
            clsMant.ActualizarCapVoucherDeposito lnMovNroRVD, lsVoucherDet
        
        'RIRO20141013 ERS122, ADD 12
        Case 7, 8, 9, 14, 15, 16, 17, 22, 23, 24  'CTS Lote
            clsMant.ActualizarCapVoucherDeposito lnMovNroRVD, , CLng(lsVoucherDet)
            
        End Select
        If pnMovNroPen > 0 Then
            clsMant.registrarMovRef lnMovNroRVD, pnMovNroPen
        End If
        'Relación con la cabecera
        clsMant.InsertarCapVoucherDepositoRel lnId, lnMovNroRVD
    Next

    If poClsMant Is Nothing Then
        clsMant.dbCmact.CommitTrans
        Set clsMant = Nothing
    End If
    bTrans = False
    pnIdNew = lnId
    Set ClsMov = Nothing ' RIRO20140530 ERS017
    InsertarVoucherDeposito = True
    Exit Function
ErrGraba:
    InsertarVoucherDeposito = False
    If bTrans Then
        If poClsMant Is Nothing Then
            clsMant.dbCmact.RollbackTrans
            Set clsMant = Nothing
        End If
    End If
End Function
Public Function EditarVoucherDeposito(ByVal pnId As Long, _
                                       ByVal pdFecReg As Date, ByVal psAgeCod As String, ByVal psUserCod As String, ByVal psOpeCod As String, _
                                       ByVal psPersCodCli As String, _
                                       ByVal psIFTpo As String, ByVal psPersCodIF As String, ByVal psCtaIFCod As String, _
                                       ByVal psCtaContBanco As String, ByVal psSubCtaContBanco As String, _
                                       ByVal psMoneda As String, ByVal psNroVou As String, ByVal pnMonVou As Currency, ByVal pdFecVou As Date, ByVal pbConfirmado As Boolean, _
                                       ByVal pnMovNroPen As Long, ByRef pMatOperaciones As Variant) As Boolean
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim bTrans As Boolean
    Dim lnIdNew As Long
    Dim lsMovNroExt As String
    Dim bExito As Boolean

    On Error GoTo ErrEditar
    
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    clsMant.dbCmact.BeginTrans
    bTrans = True
    
    'Extorna el Voucher Anterior
    lsMovNroExt = clsMant.GeneraMovNro(pdFecReg, psAgeCod, psUserCod)
    clsMant.eliminarCapVoucherDeposito lsMovNroExt, 0, pnId
    'Inserta un Nuevo Voucher
    bExito = InsertarVoucherDeposito(pdFecReg, psAgeCod, psUserCod, psOpeCod, psPersCodCli, psIFTpo, psPersCodIF, psCtaIFCod, psCtaContBanco, psSubCtaContBanco, psMoneda, psNroVou, pnMonVou, pdFecVou, pbConfirmado, pnMovNroPen, pMatOperaciones, clsMant, lnIdNew)
    If bExito = False Or lnIdNew = 0 Then
        clsMant.dbCmact.RollbackTrans
        Set clsMant = Nothing
        Exit Function
    End If
    
    clsMant.MigrarDetCapVoucherDeposito pnId, lnIdNew
    
    clsMant.dbCmact.CommitTrans
    Set clsMant = Nothing
    bTrans = False
    EditarVoucherDeposito = True
    Exit Function
ErrEditar:
    EditarVoucherDeposito = False
    If bTrans Then
        clsMant.dbCmact.RollbackTrans
        Set clsMant = Nothing
    End If
End Function
'END EJVG *******

'*********************RECO 20131023 ERS141***************
Public Function RecuperaSubTipoProducto(ByVal psCtaCod As String) As ADODB.Recordset
    Dim objCaptaGeneral As New COMDCaptaGenerales.DCOMCaptaGenerales
    Set objCaptaGeneral = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set RecuperaSubTipoProducto = objCaptaGeneral.RecuperaSubTipoProducto(psCtaCod)
End Function

Public Function ValidaFondoGarantEcotaxi(ByVal psCtaCod As String) As Boolean
    Dim objCaptaGeneral As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oDR As New ADODB.Recordset
    Set objCaptaGeneral = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set oDR = New ADODB.Recordset
    Set oDR = objCaptaGeneral.ValidaFondoGarantEcotaxi(psCtaCod)
    
    If oDR!nCantidad > 0 Then
        ValidaFondoGarantEcotaxi = True
    Else
        ValidaFondoGarantEcotaxi = False
    End If
End Function
'********************************END RECO****************

'************ ADD BY GITU 09-12-2013 ****************'
Public Function ValidaTarjetizacion(ByVal sCuenta As String, ByRef sTieneTarj As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set ValidaTarjetizacion = clsMant.ValidaTarjetizacionDat(sCuenta)
If Not ValidaTarjetizacion.BOF() And Not ValidaTarjetizacion.EOF() Then
    sTieneTarj = ValidaTarjetizacion!Tarjeta
End If
Set clsMant = Nothing
End Function

'********************* END GITU *********************'
'FRHU 20140227 RQ14006
Public Function GetCtaAhorroPersona(ByVal sCodPer As String, ByVal nMoneda As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetCtaAhorroPersona = clsMant.GetCuentaAhorroPersona(sCodPer, nMoneda)
Set clsMant = Nothing
End Function
'FIN FRHU 20140227 RQ14006
'FRHU 20140301 RQ14008
Public Function ObtenerRelacionAhoCred(ByVal sCtaAho As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set ObtenerRelacionAhoCred = clsMant.ObtenerRelacionAhoCred(sCtaAho)
Set clsMant = Nothing
End Function
'FIN FRHU 20140301
'JUEZ 20140304 **************************************************
Public Sub CambioTasaCTSCtasVig(ByVal pnMoneda As Moneda, ByVal psMovNro As String, ByVal nTNACTSCajaSueldo As Double, ByVal nTNACTSActivo As Double, ByVal nTNACTSNoActivo As Double)
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim bTrans As Boolean
    
    bTrans = False
    On Error GoTo ErrTasa
    
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    clsCap.dbCmact.BeginTrans
    bTrans = True

    clsCap.CambioTasaCTSCtasVig pnMoneda, psMovNro, nTNACTSCajaSueldo, nTNACTSActivo, nTNACTSNoActivo

    clsCap.dbCmact.CommitTrans
    Set clsCap = Nothing
    Exit Sub
ErrTasa:
    If bTrans = True Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Sub
Public Sub CambioTasaCTS(ByVal psCtaCod As String, ByVal pnTNAAnt As Double, ByVal pnTNA As Double, ByVal psMovNro As String)
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim bTrans As Boolean
    
    bTrans = False
    On Error GoTo ErrTasa
    
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    clsCap.dbCmact.BeginTrans
    bTrans = True
    
    clsCap.CambioTasaCTS psCtaCod, pnTNAAnt, pnTNA, psMovNro

    clsCap.dbCmact.CommitTrans
    Set clsCap = Nothing
    Exit Sub
ErrTasa:
    If bTrans = True Then clsCap.dbCmact.RollbackTrans
    Set clsCap = Nothing
    Err.Raise Err.Number, "Error", Err.Description
End Sub
'END JUEZ *******************************************************
'RIRO 20140530 ERS017 *********
Public Function ObtenerSobranteVoucherPersona(ByVal psPersCod As String) As ADODB.Recordset
    Dim clsMant As New COMDCaptaGenerales.DCOMCaptaGenerales
    Set ObtenerSobranteVoucherPersona = clsMant.ObtenerSobranteVoucherPersona(psPersCod)
    Set clsMant = Nothing
End Function
Public Function RegistrarSobrante(ByVal psMovNro As String, ByVal psOpeCod As CaptacOperacion, ByVal psGlosa As String, _
                                  ByVal pnMonto As Currency, ByVal pnId As Integer, ByVal psNroVou As String, _
                                  ByVal pnMoneda As Moneda, ByVal sPersCodCli As String, ByRef pnMovNro As Long) As Boolean
        
    Dim oMov As New COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim lnMovNro As Long
    Dim bTrans As Boolean
    Dim dFecSis As Date

    On Error GoTo Error
    dFecSis = CDate(Mid(psMovNro, 7, 2) & "/" & Mid(psMovNro, 5, 2) & "/" & Left(psMovNro, 4))
    oMov.dbCmact.BeginTrans
    bTrans = True
    oMov.AgregaMov psMovNro, psOpeCod, psGlosa, gMovEstContabMovContable, gMovFlagVigente
    lnMovNro = oMov.GetnMovNro(psMovNro)
    pnMovNro = lnMovNro
    oMov.AgregaMovOpeVarias lnMovNro, pnMonto, psNroVou, sPersCodCli, pnMoneda, psOpeCod
    oMov.RegistrarSobrante pnId, lnMovNro
    oMov.dbCmact.CommitTrans
    bTrans = False
    RegistrarSobrante = True
    pnMovNro = lnMovNro
    bTrans = False
    Set oMov = Nothing
    Exit Function
Error:
    RegistrarSobrante = False
    pnMovNro = 0
    If bTrans Then
        oMov.dRollbackTrans
        Set oMov = Nothing
    End If
    Err.Raise Err.Number, "RegistrarSobrante", Err.Description
End Function
'END RIRO *********************
'****** FRHU 20140926 TI-ERS099-2014
Public Function GeneraSolicitudAperturaPersNatural(ByVal sCuenta As String, sTipCuenta As String, _
    ByVal dFecApe As Date, ByVal bOrdPag As Boolean, ByVal rsProdPers As ADODB.Recordset, _
    ByVal sAgencia As String, ByVal dFecSis As Date, ByVal sUser As String, Optional ByVal nTpoPrograma As Integer = 0, Optional nCostoMan As Currency = 0, _
    Optional ByVal psReglas As String) As String
         
    Dim sTipPer As String, sCodCli As String, sNomCli As String
    Dim sDirCli As String, sDocCli As String, sTelCli As String
    Dim sCelCli As String 'FRHU 20140906 ERS099-2014
    Dim sNomSec As String, sNomAge As String, sRelCta As String
    Dim sDesPro As String * 25, sDesTip As String * 25
    Dim i As Integer, nCarLin As Long
    Dim sMoneda As String
    Dim sCad As String
    Dim sNumPag As String
    Dim sOcupCli As String
    Dim sPromotor As String
    Dim sInstitucion As String
    'FRHU 20140927 ERS099-2014
    Dim sGrupo As String
    Dim sPlazo As String, sRetiroInteres As String
    Dim sMontoInteres As String
    Dim nFormaRetiro As Integer
    Dim lnTasa As Double, nInteres As Double
    Dim dFechaV As Date, sFechaV As String
    'FIN FRHU 20140927 ERS099-2014
    Dim nFilaPersona As Integer, nTotalFilaFirma As Integer, nFilaHoja As Integer, nFilaFirma As Integer  'FRHU 20140927 ERS099-2014
    
    'FRHU 20140927 ERS099-2014
    nFilaPersona = 1
    nTotalFilaFirma = 1
    nFilaHoja = 26
    nFilaFirma = 1
    'FIN FRHU 20140927
    sCad = ""
    
    Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
    Dim nItem As Integer
    Dim nPagina As Integer
    
    Dim sTitRp1 As String, sTitRp2 As String
    Dim sTitRp3 As String 'FRHU20140926
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oImpreCap As New NCOMCaptaImpresion
    Dim oFunLetras As New COMFunciones.FCOMCadenas
    Dim oImpre As New COMFunciones.FCOMImpresion
    Dim oFunCad As New COMFunciones.FCOMCadenas

    Dim rsPers As ADODB.Recordset
    nTamMayor = 70
    nTamMenor = 38
    nTamAdd = 22
    nItem = 12
    nPagina = 1
    sNomAge = sAgencia
    nCarLin = nTamMayor + nTamMenor + 2
    
    '------------ AVMM 18-01-2007 ---------------------
    'Procediemitno para agregar la TNA/TEA al Reporte
     Dim oConect As COMConecta.DCOMConecta
     Dim rs As New ADODB.Recordset
     Dim rs1 As New ADODB.Recordset
     Dim nConvierteTNAaTEA As Double
     Dim nTasa As Double
     Dim nPlazo As Integer
     Dim sql As String
     Dim nSaldo As Double
     Dim nTREA As Double
     Dim nSaldoIn As Currency
     
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion Then
       sql = "select nTasaInteres,nSaldo From Producto Where cCtaCod='" & sCuenta & "'"
       Set rs = oConect.CargaRecordSet(sql)
       If Not (rs.EOF And rs.BOF) Then
           nTasa = rs!ntasainteres
           nSaldoIn = rs!nSaldo
           If CLng(Mid$(sCuenta, 6, 3)) = gCapPlazoFijo Then
               'sql = "select nPlazo from CaptacPlazoFijo where cCtaCod='" & sCuenta & "'" 'FRHU 20140926
               sql = "select nPlazo,nFormaRetiro from CaptacPlazoFijo where cCtaCod='" & sCuenta & "'"
               Set rs1 = oConect.CargaRecordSet(sql)
               nPlazo = rs1!nPlazo
               'nConvierteTNAaTEA = Format$(((((((nTasa / 12) / 3000) + 1) ^ nPlazo) - 1) * 100), "#0.00")
               nConvierteTNAaTEA = Format$(((1 + nTasa / 36000) ^ 360 - 1) * 100, "#0.00")
               'ALPA 20091118***************************************
               nTREA = ObtenerTREA(Mid$(sCuenta, 6, 3), IIf(nSaldoIn = 0, 10, nSaldoIn), nConvierteTNAaTEA, nPlazo, nCostoMan)
               '****************************************************
               nFormaRetiro = rs1!nFormaRetiro 'FRHU 20140926
           Else
               nConvierteTNAaTEA = Format$(((1 + nTasa / 36000) ^ 360 - 1) * 100, "#0.00")
               'ALPA 20091118***************************************
               nTREA = ObtenerTREA(Mid$(sCuenta, 6, 3), IIf(nSaldoIn = 0, 10, nSaldoIn), nConvierteTNAaTEA, , nCostoMan)
               '****************************************************
           End If
           nSaldo = rs!nSaldo
       Else
           nTasa = 0
           nConvierteTNAaTEA = 0
           nSaldo = 0
       End If
    End If
    
    'SOLO PLAZO FIJO
    sPlazo = "-"
    sRetiroInteres = "-"
    sMontoInteres = "-"
    sFechaV = "-"
    If CLng(Mid$(sCuenta, 6, 3)) = gCapPlazoFijo Then
        Dim clsMantA As New COMNCaptaGenerales.NCOMCaptaMovimiento
        lnTasa = nConvierteTNAaTEA
        
        If nFormaRetiro = 4 Then
            sRetiroInteres = "ADELANTADO"
            Set clsMantA = New COMNCaptaGenerales.NCOMCaptaMovimiento
                'nInteres = Format$(clsMantA.GetInteres(nSaldo, lnTasa, nPlazo, TpoCalcIntAdelantado), "#,##0.00") 'FRHU 20150103
                nInteres = Format$(clsMantA.GetInteres(nSaldo, nTasa, nPlazo, TpoCalcIntAdelantado), "#,##0.00")
            Set clsMantA = Nothing
            sMontoInteres = CStr(nInteres) & "(POR SALDO DE APERTURA)"
        ElseIf nFormaRetiro = 2 Then
            sRetiroInteres = "FINAL DE PLAZO"
            'nInteres = Format$(GetInteresPF(lnTasa, nSaldo, CLng(nPlazo)), "#,##0.00") 'FRHU 20150103
            nInteres = Format$(GetInteresPF(nTasa, nSaldo, CLng(nPlazo)), "#,##0.00")
            sMontoInteres = CStr(nInteres) & "(POR SALDO DE APERTURA)"
        ElseIf nFormaRetiro = 1 Then
            sRetiroInteres = "MENSUAL"
            sMontoInteres = "SEGUN CRONOGRAMA ADJUNTO"
        ElseIf nFormaRetiro = 3 Then
            sRetiroInteres = "LIBRE"
            sMontoInteres = "SEGUN CRONOGRAMA ADJUNTO"
        End If
        sPlazo = CStr(nPlazo) & " DIAS"
        dFechaV = CDate(dFecApe) + nPlazo
        sFechaV = CStr(Format$(dFechaV, "dd mmm yyyy"))
        Set clsMant = Nothing
    End If
    'FIN SOLO PLAZO FIJO
    
    If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
        Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsPers = clsMant.GetDatosCuentaCTS(sCuenta)
        sInstitucion = rsPers("cInstitucion")
        Set clsMant = Nothing
        Set rsPers = Nothing
    End If
    
    If CLng(Mid$(sCuenta, 6, 3)) = gCapAhorros Then
        Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        If nTpoPrograma = 6 Or nTpoPrograma = 8 Then
            sInstitucion = clsMant.GetInstitucionSolicitudApertura(sCuenta, nTpoPrograma)
        Else
            sInstitucion = "-"
        End If
        Set clsMant = Nothing
    End If
    sTitRp1 = ""
    Select Case CLng(Mid$(sCuenta, 6, 3))
        Case gCapAhorros
            sDesPro = "AHORRO" '& IIf(bOrdPag, " CON ORDEN DE PAGO", "")
            If nTpoPrograma = 0 Then
            sDesPro = Trim(sDesPro) & " CORRIENTE" & IIf(bOrdPag, "-CON OP", "")
            ElseIf nTpoPrograma = 1 Then
            sDesPro = Trim(sDesPro) & " ÑAÑITO" & IIf(bOrdPag, "-CON OP", "")
            sTitRp1 = " - A NOMBRE DE MENOR DE EDAD"
            ElseIf nTpoPrograma = 4 Then
            sDesPro = Trim(sDesPro) & " DESTINO" & IIf(bOrdPag, "-CON OP", "")
            ElseIf nTpoPrograma = 5 Then
            sDesPro = " CUENTA SOÑADA"
            ElseIf nTpoPrograma = 6 Then
            sDesPro = " CAJA SUELDO"
            ElseIf nTpoPrograma = 7 Then
            sDesPro = Trim(sDesPro) & " ECOTAXI" & IIf(bOrdPag, "-CON OP", "")
            ElseIf nTpoPrograma = 8 Then
            sDesPro = Trim(sDesPro) & " CONVENIO" & IIf(bOrdPag, "-CON OP", "")
            End If
        Case gCapPlazoFijo
            If nTpoPrograma = 0 Then
            sDesPro = "PLAZO FIJO CLÁSICO"
            ElseIf nTpoPrograma = 1 Then
            sDesPro = "DEPÓSITO PLAZO PREMIUM"
            ElseIf nTpoPrograma = 2 Then
            sDesPro = "AHORRO DIARIO"
            ElseIf nTpoPrograma = 3 Then
            sDesPro = "POCO A POCO AHORRO"
            End If
        Case gCapCTS
            sDesPro = "C T S"
            If nTpoPrograma = 0 Then
                sDesPro = Trim(sDesPro) & " + CAJA SUELDO" '**JUEZ 20120216
            End If
    End Select
    
    sDesTip = sTipCuenta
    sCad = ""
    
    Set gcPEN = New COMDConstantes.DCOMConstantes
    '''sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", "NUEVOS SOLES", IIf(Mid$(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) 'marg ers044-2016
    sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL(), vbUpperCase), IIf(Mid$(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) 'marg ers044-2016
    
    sTitRp1 = "SOLICITUD DE APERTURA" & sTitRp1
    sTitRp2 = "P E R S O N A   N A T U R A L"
    sTitRp3 = "CUENTA N° " & Mid(sCuenta, 1, 3) & "-" & Mid(sCuenta, 4, 2) & "-" & Mid(sCuenta, 6, 3) & "-" & Mid(sCuenta, 9, 10)

    sCad = sCad & oImpreCap.CabeRepoCaptacSolicitud("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis, oImp.gPrnSaltoLinea, sTitRp3, sUser) & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "DATOS DEL DEPOSITO" & oImp.gPrnSaltoLinea 'FRHU20140926
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea 'FRHU20140926
    sCad = sCad & "TIPO DE CUENTA       : " & Trim(sDesTip) & Space(25 - Len(Trim(sDesTip))) & "PRODUCTO            : " & Trim(Left(sDesPro, 25)) & Space(25 - Len(Trim(Left(sDesPro, 25)))) & "MONEDA     : " & Trim(sMoneda) & Space(25 - Len(Trim(sMoneda))) & oImp.gPrnSaltoLinea
    sCad = sCad & "IMPORTE DE DEPOSITO  : " & Trim(Format(nSaldo, "#,##0.00")) & Space(25 - Len(Trim(Format(nSaldo, "#,##0.00")))) & oImp.gPrnSaltoLinea
    sCad = sCad & "PLAZO                : " & Trim(sPlazo) & Space(25 - Len(Trim(sPlazo))) & "T.E.A               : " & Trim(Format(nConvierteTNAaTEA, "0.00")) & "%" & Space(25 - Len(Trim(Format(nConvierteTNAaTEA, "0.00")) & "%")) & "T.R.E.A.   : " & Trim(Format(nTREA, "0.00")) & "%" & oImp.gPrnSaltoLinea
    sCad = sCad & "RETIRO DE INTERES    : " & Trim(sRetiroInteres) & Space(25 - Len(Trim(sRetiroInteres))) & "MONTO INTERES       : " & Trim(sMontoInteres) & oImp.gPrnSaltoLinea
    sCad = sCad & "FECHA DE APERTURA    : " & Format$(dFecApe, "dd mmm yyyy") & Space(25 - Len(Trim(Format$(dFecApe, "dd mmm yyyy")))) & "FECHA DE VENCIMIENTO: " & Trim(sFechaV) & oImp.gPrnSaltoLinea
    If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
        sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
    Else
        If nTpoPrograma = 6 Or nTpoPrograma = 8 Then
            sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & "ENTIDAD              : " & "-" & oImp.gPrnSaltoLinea
        End If
    End If
     'APRI20180406 ERS036-2017
    sql = "Select * From ProductoEnvioEstadoCta where nEstado = 1 and cCtaCod='" & sCuenta & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        sCad = sCad & "ENVIO DE ESTADO DE CUENTA: SI" & oImp.gPrnSaltoLinea
    Else
        sCad = sCad & "ENVIO DE ESTADO DE CUENTA: NO" & oImp.gPrnSaltoLinea
    End If
    Set rs = Nothing
    'END APRI
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "PERSONAS AUTORIZADOS PARA OPERAR LA CUENTA" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    
    sCad = sCad & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreFormat("DOI", 8, 0) & Space(2) & oImpre.ImpreFormat("RELACION", 18, 0) & Space(2) & oImpre.ImpreFormat("GRUPO", 5, 0) & Space(2) & oImpre.ImpreFormat("NOMBRES Y APELLIDOS", 38, 0) & Space(2) & oImpre.ImpreFormat("DIRECCION", 37, 0) & Space(2) & oImpre.ImpreFormat("TELEFONO", 9, 0) & Space(2) & oImpre.ImpreFormat("CELULAR", 10, 0) & oImp.gPrnSaltoLinea
    sCad = sCad & String(8, "-") & Space(2) & String(18, "-") & Space(2) & String(5, "-") & Space(2) & String(38, "-") & Space(2) & String(37, "-") & Space(2) & String(9, "-") & Space(2) & String(10, "-") & oImp.gPrnSaltoLinea

    sCodCli = ""

    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsPers = New ADODB.Recordset
    rsProdPers.MoveFirst
    Do While Not rsProdPers.EOF
        If CInt(Trim(Right(rsProdPers("Relacion"), 2))) <> gCapRelPersPromotor Then
            sCodCli = rsProdPers("Codigo")
            Set rsPers = clsMant.GetDatosPersona(rsProdPers("Codigo"))
            sNomCli = oImpre.ImpreCarEsp(oFunLetras.PstaNombre(rsPers("Nombre"), False))
            sRelCta = UCase(Trim(Left(rsProdPers("Relacion"), 30)))
            sDirCli = rsPers("Direccion") & ""
            sDocCli = rsPers("ID N°")
            sTelCli = rsPers("Fono")
            sNomSec = rsPers("Zona")
            sOcupCli = rsPers("cCIIUdescripcion")
            sCelCli = rsPers("Celular") 'FRHU 20140927 ERS099-2014
            sGrupo = Trim(rsProdPers("Grupo")) 'FRHU 20140927 ERS099-2014
            
            sCad = sCad & Left(Trim(sDocCli), 8) & Space(2)
            sCad = sCad & Left(Trim(sRelCta), 18) & Space(18 - Len(Left(Trim(sRelCta), 18))) & Space(2)
            sCad = sCad & Left(Trim(sGrupo), 2) & Space(5 - Len(Trim(sGrupo))) & Space(2)
            sCad = sCad & Left(Trim(oImpre.ImpreCarEsp(oFunCad.PstaNombre(sNomCli, False))), nTamMenor) & Space(38 - Len(Left(Trim(oImpre.ImpreCarEsp(oFunCad.PstaNombre(sNomCli, False))), nTamMenor))) & Space(2)
            sCad = sCad & Left(Trim(sDirCli), 37) & Space(37 - Len(Left(Trim(sDirCli), 37))) & Space(2)
            sCad = sCad & Left(Trim(sTelCli), 9) & Space(9 - Len(Left(Trim(sTelCli), 9))) & Space(2)
            sCad = sCad & Left(Trim(sCelCli), 10) & Space(10 - Len(Left(Trim(sCelCli), 10))) & oImp.gPrnSaltoLinea
            
            If nFilaPersona = 4 Or nFilaPersona = 7 Or nFilaPersona = 10 Or nFilaPersona = 13 Then
                nTotalFilaFirma = nTotalFilaFirma + 1
            End If
            nFilaHoja = nFilaHoja + 1
            nFilaPersona = nFilaPersona + 1
        End If
        rsProdPers.MoveNext
    Loop

    sCad = sCad & oImp.gPrnSaltoLinea
    Dim cadenaRegla As String, cadenaReglaAnterior As String
    Dim reglas As String, regla As String
    Dim totalCaracter As Integer, nRegla As Integer, nValor As Integer
    reglas = psReglas
    totalCaracter = Len(Trim(reglas))
    nRegla = 0: nValor = 1
    
    cadenaRegla = cadenaRegla & "REGLAS "
    If totalCaracter = 0 Then
        cadenaRegla = cadenaRegla & "¦1¦ " & Space(5) & "¦2¦ " & Space(5) & "¦3¦ " & Space(5) & "¦4¦ " & Space(5) & "¦5¦ " & Space(5) & "¦6¦ " & Space(5) & "¦7¦ " & Space(5) & "¦8¦ " & Space(5)
    Else
        Do While totalCaracter > 0
            If InStr(1, reglas, "-") = 0 Then
                regla = reglas
                totalCaracter = 0
            Else
                regla = Left(reglas, InStr(1, reglas, "-") - 1)
                reglas = Right(reglas, Len(reglas) - InStr(1, reglas, "-"))
                totalCaracter = Len(reglas)
            End If
            nRegla = nRegla + 1
            cadenaReglaAnterior = cadenaRegla
            cadenaRegla = cadenaRegla & "¦" & nRegla & "¦ " & regla & " " & Space(2)
            If ((Len(cadenaRegla) - Len(cadenaReglaAnterior)) + Len(cadenaRegla)) > (137 * nValor) Then
                cadenaRegla = cadenaRegla & oImp.gPrnSaltoLinea
                nValor = nValor + 1
                nFilaHoja = nFilaHoja + 1 'Una fila mas aumenta
            End If
        Loop
    End If
    sCad = sCad & cadenaRegla & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "DECLARACIONES Y FIRMAS" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    
    nFilaHoja = nFilaHoja + 7
    Do While nFilaFirma <= nTotalFilaFirma Or nFilaFirma = 1
        If nFilaHoja >= 64 Then
            nPagina = nPagina + 1
            sCad = sCad & oImp.gPrnSaltoPagina
            sCad = sCad & oImpreCap.CabeRepoCaptacSolicitud("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis, oImp.gPrnSaltoLinea, sTitRp3, sUser) & oImp.gPrnSaltoLinea
            sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
            sCad = sCad & "DATOS DEL DEPOSITO" & oImp.gPrnSaltoLinea 'FRHU20140926
            sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea 'FRHU20140926
            sCad = sCad & "TIPO DE CUENTA       : " & Trim(sDesTip) & Space(25 - Len(Trim(sDesTip))) & "PRODUCTO            : " & Trim(Left(sDesPro, 25)) & Space(25 - Len(Trim(Left(sDesPro, 25)))) & "MONEDA     : " & Trim(sMoneda) & Space(25 - Len(Trim(sMoneda))) & oImp.gPrnSaltoLinea
            sCad = sCad & "IMPORTE DE DEPOSITO  : " & Trim(Format(nSaldo, "#,##0.00")) & Space(25 - Len(Trim(Format(nSaldo, "#,##0.00")))) & oImp.gPrnSaltoLinea
            sCad = sCad & "PLAZO                : " & Trim(sPlazo) & Space(25 - Len(Trim(sPlazo))) & "T.E.A               : " & Trim(Format(nConvierteTNAaTEA, "0.00")) & "%" & Space(25 - Len(Trim(Format(nConvierteTNAaTEA, "0.00")) & "%")) & "T.R.E.A.   : " & Trim(Format(nTREA, "0.00")) & "%" & oImp.gPrnSaltoLinea
            sCad = sCad & "RETIRO DE INTERES    : " & Trim(sRetiroInteres) & Space(25 - Len(Trim(sRetiroInteres))) & "MONTO INTERES       : " & Trim(sMontoInteres) & oImp.gPrnSaltoLinea
            sCad = sCad & "FECHA DE APERTURA    : " & Format$(dFecApe, "dd mmm yyyy") & Space(25 - Len(Trim(Format$(dFecApe, "dd mmm yyyy")))) & "FECHA DE VENCIMIENTO: " & Trim(sFechaV) & oImp.gPrnSaltoLinea
            If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
                sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
            Else
                If nTpoPrograma = 6 Or nTpoPrograma = 8 Then
                    sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
                Else
                    sCad = sCad & "ENTIDAD              : " & "-" & oImp.gPrnSaltoLinea
                End If
            End If
            sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
            nFilaHoja = 16
        End If
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("------------------------------------------    ------------------------------------------    ------------------------------------------", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("FIRMA Y SELLO DEL REPRESENTANTE AUTORIZADO    FIRMA Y SELLO DEL REPRESENTANTE AUTORIZADO    FIRMA Y SELLO DEL REPRESENTANTE AUTORIZADO", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
        nFilaHoja = nFilaHoja + 7
        nFilaFirma = nFilaFirma + 1
    Loop
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    
    If (nFilaHoja + 11) > 64 Then
        nPagina = nPagina + 1
        sCad = sCad & oImp.gPrnSaltoPagina
        sCad = sCad & oImpreCap.CabeRepoCaptacSolicitud("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis, oImp.gPrnSaltoLinea, sTitRp3, sUser) & oImp.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & "DATOS DEL DEPOSITO" & oImp.gPrnSaltoLinea 'FRHU20140926
        sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea 'FRHU20140926
        sCad = sCad & "TIPO DE CUENTA       : " & Trim(sDesTip) & Space(25 - Len(Trim(sDesTip))) & "PRODUCTO            : " & Trim(Left(sDesPro, 25)) & Space(25 - Len(Trim(Left(sDesPro, 25)))) & "MONEDA     : " & Trim(sMoneda) & Space(25 - Len(Trim(sMoneda))) & oImp.gPrnSaltoLinea
        sCad = sCad & "IMPORTE DE DEPOSITO  : " & Trim(Format(nSaldo, "#,##0.00")) & Space(25 - Len(Trim(Format(nSaldo, "#,##0.00")))) & oImp.gPrnSaltoLinea
        sCad = sCad & "PLAZO                : " & Trim(sPlazo) & Space(25 - Len(Trim(sPlazo))) & "T.E.A               : " & Trim(Format(nConvierteTNAaTEA, "0.00")) & "%" & Space(25 - Len(Trim(Format(nConvierteTNAaTEA, "0.00")) & "%")) & "T.R.E.A.   : " & Trim(Format(nTREA, "0.00")) & "%" & oImp.gPrnSaltoLinea
        sCad = sCad & "RETIRO DE INTERES    : " & Trim(sRetiroInteres) & Space(25 - Len(Trim(sRetiroInteres))) & "MONTO INTERES       : " & Trim(sMontoInteres) & oImp.gPrnSaltoLinea
        sCad = sCad & "FECHA DE APERTURA    : " & Format$(dFecApe, "dd mmm yyyy") & Space(25 - Len(Trim(Format$(dFecApe, "dd mmm yyyy")))) & "FECHA DE VENCIMIENTO: " & Trim(sFechaV) & oImp.gPrnSaltoLinea
        If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
            sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
        Else
            If nTpoPrograma = 6 Or nTpoPrograma = 8 Then
                sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
            Else
                sCad = sCad & "ENTIDAD              : " & "-" & oImp.gPrnSaltoLinea
            End If
        End If
        sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
        nFilaHoja = 15
    End If
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("------------------------------------------             ------------------------------------------", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("            ASESOR DEL CLIENTE                                  SUPERVISOR DE OPERACIONES        ", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-")
    sCad = sCad & oImp.gPrnSaltoPagina
    
    GeneraSolicitudAperturaPersNatural = sCad
End Function
Public Function GeneraSolicitudAperturaPersJuridica(ByVal sCuenta As String, sTipCuenta As String, _
    ByVal dFecApe As Date, ByVal bOrdPag As Boolean, ByVal rsProdPers As ADODB.Recordset, _
    ByVal sAgencia As String, ByVal dFecSis As Date, ByVal sUser As String, Optional ByVal nTpoPrograma As Integer = 0, Optional nCostoMan As Currency = 0, _
    Optional ByVal psReglas As String) As String
         
    Dim sTipPer As String, sCodCli As String, sNomCli As String
    Dim sDirCli As String, sDocCli As String, sTelCli As String
    Dim sCelCli As String 'FRHU 20140906 ERS099-2014
    Dim sNomSec As String, sNomAge As String, sRelCta As String
    Dim sDesPro As String * 25, sDesTip As String * 25
    Dim i As Integer, nCarLin As Long
    Dim sMoneda As String
    Dim sCad As String
    Dim sNumPag As String
    Dim sOcupCli As String
    Dim sPromotor As String
    Dim sInstitucion As String
    'FRHU 20140927 ERS099-2014
    Dim sGrupo As String
    Dim sPlazo As String, sRetiroInteres As String
    Dim sMontoInteres As String
    Dim nFormaRetiro As Integer
    Dim lnTasa As Double, nInteres As Double
    Dim dFechaV As Date, sFechaV As String
    'FIN FRHU 20140927 ERS099-2014
    Dim nFilaPersona As Integer, nTotalFilaFirma As Integer, nFilaHoja As Integer, nFilaFirma As Integer  'FRHU 20140927 ERS099-2014
    
    'FRHU 20140927 ERS099-2014
    nFilaPersona = 1
    nTotalFilaFirma = 1
    nFilaHoja = 33
    nFilaFirma = 1
    'FIN FRHU 20140927
    sCad = ""
    
    Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
    Dim nItem As Integer
    Dim nPagina As Integer
    
    Dim sTitRp1 As String, sTitRp2 As String
    Dim sTitRp3 As String 'FRHU20140926
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oImpreCap As New NCOMCaptaImpresion
    Dim oFunLetras As New COMFunciones.FCOMCadenas
    Dim oImpre As New COMFunciones.FCOMImpresion
    Dim oFunCad As New COMFunciones.FCOMCadenas

    Dim rsPers As ADODB.Recordset
    nTamMayor = 70
    nTamMenor = 38
    nTamAdd = 22
    nItem = 12
    nPagina = 1
    sNomAge = sAgencia
    nCarLin = nTamMayor + nTamMenor + 2
    
    Dim oConect As COMConecta.DCOMConecta
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim nConvierteTNAaTEA As Double
    Dim nTasa As Double
    Dim nPlazo As Integer
    Dim sql As String
    Dim nSaldo As Double
    Dim nTREA As Double
    Dim nSaldoIn As Currency
     
    Set oConect = New COMConecta.DCOMConecta
    If oConect.AbreConexion Then
       sql = "select nTasaInteres,nSaldo From Producto Where cCtaCod='" & sCuenta & "'"
       Set rs = oConect.CargaRecordSet(sql)
       If Not (rs.EOF And rs.BOF) Then
           nTasa = rs!ntasainteres
           nSaldoIn = rs!nSaldo
           If CLng(Mid$(sCuenta, 6, 3)) = gCapPlazoFijo Then
               'sql = "select nPlazo from CaptacPlazoFijo where cCtaCod='" & sCuenta & "'" 'FRHU 20140926
               sql = "select nPlazo,nFormaRetiro from CaptacPlazoFijo where cCtaCod='" & sCuenta & "'"
               Set rs1 = oConect.CargaRecordSet(sql)
               nPlazo = rs1!nPlazo
               'nConvierteTNAaTEA = Format$(((((((nTasa / 12) / 3000) + 1) ^ nPlazo) - 1) * 100), "#0.00")
               nConvierteTNAaTEA = Format$(((1 + nTasa / 36000) ^ 360 - 1) * 100, "#0.00")
               'ALPA 20091118***************************************
               nTREA = ObtenerTREA(Mid$(sCuenta, 6, 3), IIf(nSaldoIn = 0, 10, nSaldoIn), nConvierteTNAaTEA, nPlazo, nCostoMan)
               '****************************************************
               nFormaRetiro = rs1!nFormaRetiro 'FRHU 20140926
           Else
               nConvierteTNAaTEA = Format$(((1 + nTasa / 36000) ^ 360 - 1) * 100, "#0.00")
               'ALPA 20091118***************************************
               nTREA = ObtenerTREA(Mid$(sCuenta, 6, 3), IIf(nSaldoIn = 0, 10, nSaldoIn), nConvierteTNAaTEA, , nCostoMan)
               '****************************************************
           End If
           nSaldo = rs!nSaldo
       Else
           nTasa = 0
           nConvierteTNAaTEA = 0
           nSaldo = 0
       End If
    End If
    
    'SOLO PARA PLAZO FIJO
    sPlazo = "-"
    sRetiroInteres = "-"
    sMontoInteres = "-"
    sFechaV = "-"
    If CLng(Mid$(sCuenta, 6, 3)) = gCapPlazoFijo Then
        Dim clsMantA As New COMNCaptaGenerales.NCOMCaptaMovimiento
        lnTasa = nConvierteTNAaTEA
        
        If nFormaRetiro = 4 Then
            sRetiroInteres = "ADELANTADO"
            Set clsMantA = New COMNCaptaGenerales.NCOMCaptaMovimiento
                'nInteres = Format$(clsMantA.GetInteres(nSaldo, lnTasa, nPlazo, TpoCalcIntAdelantado), "#,##0.00") 'FRHU 20150103
                nInteres = Format$(clsMantA.GetInteres(nSaldo, nTasa, nPlazo, TpoCalcIntAdelantado), "#,##0.00")
            Set clsMantA = Nothing
            sMontoInteres = CStr(nInteres) & "(POR SALDO DE APERTURA)"
        ElseIf nFormaRetiro = 2 Then
            sRetiroInteres = "FINAL DE PLAZO"
            'nInteres = Format$(GetInteresPF(lnTasa, nSaldo, CLng(nPlazo)), "#,##0.00") 'FRHU 20150103
            nInteres = Format$(GetInteresPF(nTasa, nSaldo, CLng(nPlazo)), "#,##0.00")
            sMontoInteres = CStr(nInteres) & "(POR SALDO DE APERTURA)"
        ElseIf nFormaRetiro = 1 Then
            sRetiroInteres = "MENSUAL"
            sMontoInteres = "SEGUN CRONOGRAMA ADJUNTO"
        ElseIf nFormaRetiro = 3 Then
            sRetiroInteres = "LIBRE"
            sMontoInteres = "SEGUN CRONOGRAMA ADJUNTO"
        End If
        sPlazo = CStr(nPlazo) & " DIAS"
        dFechaV = CDate(dFecApe) + nPlazo
        sFechaV = CStr(Format$(dFechaV, "dd mmm yyyy"))
        Set clsMant = Nothing
    End If
    'FIN SOLO PARA PLAZO FIJO
    If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
        Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
        Set rsPers = clsMant.GetDatosCuentaCTS(sCuenta)
        sInstitucion = rsPers("cInstitucion")
        Set clsMant = Nothing
        Set rsPers = Nothing
    End If

    Select Case CLng(Mid$(sCuenta, 6, 3))
        Case gCapAhorros
            sDesPro = "AHORRO" '& IIf(bOrdPag, " CON ORDEN DE PAGO", "")
            If nTpoPrograma = 0 Then
                sDesPro = Trim(sDesPro) & "CORRIENTE" & IIf(bOrdPag, "-CON OP", "")
            ElseIf nTpoPrograma = 1 Then
                sDesPro = Trim(sDesPro) & "ÑAÑITO" & IIf(bOrdPag, "-CON OP", "")
            ElseIf nTpoPrograma = 4 Then
                sDesPro = Trim(sDesPro) & "DESTINO" & IIf(bOrdPag, "-CON OP", "")
            ElseIf nTpoPrograma = 5 Then
                sDesPro = " CUENTA SOÑADA"
            ElseIf nTpoPrograma = 6 Then
                sDesPro = " CAJA SUELDO"
            ElseIf nTpoPrograma = 7 Then
                sDesPro = Trim(sDesPro) & "ECOTAXI" & IIf(bOrdPag, "-CON OP", "")
            ElseIf nTpoPrograma = 8 Then
                sDesPro = Trim(sDesPro) & "CONVENIO" & IIf(bOrdPag, "-CON OP", "")
            End If
        Case gCapPlazoFijo
            If nTpoPrograma = 0 Then
                sDesPro = "PLAZO FIJO CLÁSICO"
            ElseIf nTpoPrograma = 1 Then
                sDesPro = "DEPÓSITO PLAZO PREMIUM"
            ElseIf nTpoPrograma = 2 Then
                sDesPro = "AHORRO DIARIO"
            ElseIf nTpoPrograma = 3 Then
                sDesPro = "POCO A POCO AHORRO"
            End If
        Case gCapCTS
            sDesPro = "C T S"
            If nTpoPrograma = 0 Then
                sDesPro = Trim(sDesPro) & " + CAJA SUELDO" '**JUEZ 20120216
            End If
    End Select
    sDesTip = sTipCuenta
    sCad = ""
    
    '''sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", "NUEVOS SOLES", IIf(Mid$(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) 'marg ers044-2016
    sMoneda = IIf(Mid$(sCuenta, 9, 1) = "1", StrConv(gcPEN.gcPEN_PLURAL, vbUpperCase), IIf(Mid$(sCuenta, 9, 1) = "2", "DOLARES", "EUROS")) 'marg ers044-2016
    
    sTitRp1 = "SOLICITUD DE APERTURA"
    sTitRp2 = "P E R S O N A   J U R I D I C A"
    sTitRp3 = "CUENTA N° " & Mid(sCuenta, 1, 3) & "-" & Mid(sCuenta, 4, 2) & "-" & Mid(sCuenta, 6, 3) & "-" & Mid(sCuenta, 9, 10)
    sCad = sCad & oImpreCap.CabeRepoCaptacSolicitud("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis, oImp.gPrnSaltoLinea, sTitRp3, sUser) & oImp.gPrnSaltoLinea
    
    Dim sPersCodJur As String
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsPers = New ADODB.Recordset
    
    sPersCodJur = clsMant.GetPersCodJuridicaSolicitudApertura(sCuenta)
    
    Set rsPers = clsMant.GetDatosPersonaSolicitudApertura(sPersCodJur)
    Dim sDoiJur As String, sNombreJur As String, sDireccionJur As String
    Dim dFechaNacJur As Date
    Dim sDepartamJur As String, sProvinciaJur As String, sDistritoJur As String
    Dim sTelefonoJur As String, sCelularJur As String, sEmailJur As String
    
    If Not rsPers.EOF And Not rsPers.BOF Then
        sDoiJur = rsPers!DOI
        sNombreJur = rsPers!cPersNombre
        sDireccionJur = rsPers!Direccion
        dFechaNacJur = rsPers!dPersNacCreac
        sDepartamJur = rsPers!Departamento
        sProvinciaJur = rsPers!PROVINCIA
        sDistritoJur = rsPers!Distrito
        sTelefonoJur = rsPers!Telefono
        sCelularJur = rsPers!Celular
        sEmailJur = rsPers!Email
    End If
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "DATOS DE LA PERSONA JURIDICA" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "DOI / RUC            : " & Trim(sDoiJur) & Space(25 - Len(Trim(sDoiJur))) & "R.SOCIAL/NOMBRE     : " & Trim(Left(sNombreJur, 67)) & oImp.gPrnSaltoLinea
    sCad = sCad & "FECHA DE CONSTITUCION: " & Format$(dFechaNacJur, "dd mmm yyyy") & Space(25 - Len(Trim(Format$(dFechaNacJur, "dd mmm yyyy")))) & "DIRECCION           : " & Trim(Left(sDireccionJur, 67)) & oImp.gPrnSaltoLinea
    sCad = sCad & "DEPARTAMENTO         : " & Trim(sDepartamJur) & Space(25 - Len(Trim(sDepartamJur))) & "PROVINCIA           : " & Trim(sProvinciaJur) & Space(25 - Len(Trim(sProvinciaJur))) & "DISTRITO   : " & Trim(sDistritoJur) & oImp.gPrnSaltoLinea
    sCad = sCad & "TELEFONO             : " & Trim(sTelefonoJur) & Space(25 - Len(Trim(sTelefonoJur))) & "CELULAR             : " & Trim(sCelularJur) & Space(25 - Len(Trim(sCelularJur))) & "E-MAIL     : " & Trim(Left(sEmailJur, 29)) & oImp.gPrnSaltoLinea
    Set clsMant = Nothing
    Set rsPers = Nothing
    
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "DATOS DEL DEPOSITO" & oImp.gPrnSaltoLinea 'FRHU20140926
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea 'FRHU20140926
    sCad = sCad & "TIPO DE CUENTA       : " & Trim(sDesTip) & Space(25 - Len(Trim(sDesTip))) & "PRODUCTO            : " & Trim(Left(sDesPro, 25)) & Space(25 - Len(Trim(Left(sDesPro, 25)))) & "MONEDA     : " & Trim(sMoneda) & Space(25 - Len(Trim(sMoneda))) & oImp.gPrnSaltoLinea
    sCad = sCad & "IMPORTE DE DEPOSITO  : " & Trim(Format(nSaldo, "#,##0.00")) & Space(25 - Len(Trim(Format(nSaldo, "#,##0.00")))) & oImp.gPrnSaltoLinea
    sCad = sCad & "PLAZO                : " & Trim(sPlazo) & Space(25 - Len(Trim(sPlazo))) & "T.E.A               : " & Trim(Format(nConvierteTNAaTEA, "0.00")) & "%" & Space(25 - Len(Trim(Format(nConvierteTNAaTEA, "0.00")) & "%")) & "T.R.E.A.   : " & Trim(Format(nTREA, "0.00")) & "%" & oImp.gPrnSaltoLinea
    sCad = sCad & "RETIRO DE INTERES    : " & Trim(sRetiroInteres) & Space(25 - Len(Trim(sRetiroInteres))) & "MONTO INTERES       : " & Trim(sMontoInteres) & oImp.gPrnSaltoLinea
    sCad = sCad & "FECHA DE APERTURA    : " & Format$(dFecApe, "dd mmm yyyy") & Space(25 - Len(Trim(Format$(dFecApe, "dd mmm yyyy")))) & "FECHA DE VENCIMIENTO: " & Trim(sFechaV) & oImp.gPrnSaltoLinea
    If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
        sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
    Else
        If nTpoPrograma = 6 Or nTpoPrograma = 8 Then
            sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
        Else
            sCad = sCad & "ENTIDAD              : " & "-" & oImp.gPrnSaltoLinea
        End If
    End If
    'APRI20180406 ERS036-2017
    sql = "Select * From ProductoEnvioEstadoCta where nEstado = 1 and cCtaCod='" & sCuenta & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        sCad = sCad & "ENVIO DE ESTADO DE CUENTA: SI" & oImp.gPrnSaltoLinea
    Else
        sCad = sCad & "ENVIO DE ESTADO DE CUENTA: NO" & oImp.gPrnSaltoLinea
    End If
    Set rs = Nothing
    'END APRI
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "PERSONAS AUTORIZADOS PARA OPERAR LA CUENTA" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreFormat("DOI", 11, 0) & Space(2) & oImpre.ImpreFormat("RELACION", 18, 0) & Space(2) & oImpre.ImpreFormat("GRUPO", 5, 0) & Space(2) & oImpre.ImpreFormat("NOMBRES Y APELLIDOS", 37, 0) & Space(2) & oImpre.ImpreFormat("DIRECCION", 35, 0) & Space(2) & oImpre.ImpreFormat("TELEFONO", 9, 0) & Space(2) & oImpre.ImpreFormat("CELULAR", 10, 0) & oImp.gPrnSaltoLinea
    sCad = sCad & String(11, "-") & Space(2) & String(18, "-") & Space(2) & String(5, "-") & Space(2) & String(37, "-") & Space(2) & String(35, "-") & Space(2) & String(9, "-") & Space(2) & String(10, "-") & oImp.gPrnSaltoLinea

    sCodCli = ""

    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsPers = New ADODB.Recordset
    rsProdPers.MoveFirst
    Do While Not rsProdPers.EOF
        If CInt(Trim(Right(rsProdPers("Relacion"), 2))) <> gCapRelPersPromotor Then
            sCodCli = rsProdPers("Codigo")
            Set rsPers = clsMant.GetDatosPersona(rsProdPers("Codigo"))
            sNomCli = oImpre.ImpreCarEsp(oFunLetras.PstaNombre(rsPers("Nombre"), False))
            sRelCta = UCase(Trim(Left(rsProdPers("Relacion"), 30)))
            sDirCli = rsPers("Direccion")
            sDocCli = rsPers("ID N°")
            sTelCli = rsPers("Fono")
            sNomSec = rsPers("Zona")
            sOcupCli = rsPers("cCIIUdescripcion")
            sCelCli = rsPers("Celular") 'FRHU 20140927 ERS099-2014
            sGrupo = Trim(rsProdPers("Grupo")) 'FRHU 20140927 ERS099-2014
            
            sCad = sCad & Left(Trim(sDocCli), 11) & Space(11 - Len(Trim(sDocCli))) & Space(2)
            sCad = sCad & Left(Trim(sRelCta), 18) & Space(18 - Len(Trim(sRelCta))) & Space(2)
            sCad = sCad & Left(Trim(sGrupo), 5) & Space(5 - Len(Trim(sGrupo))) & Space(2)
            sCad = sCad & Left(oImpre.ImpreCarEsp(oFunCad.PstaNombre(sNomCli, False)), 37) & Space(37 - Len(Left(oImpre.ImpreCarEsp(oFunCad.PstaNombre(sNomCli, False)), 37))) & Space(2)
            sCad = sCad & Left(Trim(sDirCli), 35) & Space(35 - Len(Left(Trim(sDirCli), 35))) & Space(2)
            sCad = sCad & Left(Trim(sTelCli), 9) & Space(9 - Len(Trim(sTelCli))) & Space(2)
            sCad = sCad & Left(Trim(sCelCli), 10) & Space(10 - Len(Trim(sCelCli))) & oImp.gPrnSaltoLinea
            
            If nFilaPersona = 4 Or nFilaPersona = 7 Or nFilaPersona = 10 Or nFilaPersona = 13 Then
                nTotalFilaFirma = nTotalFilaFirma + 1
            End If
            nFilaHoja = nFilaHoja + 1
            nFilaPersona = nFilaPersona + 1
        End If
        rsProdPers.MoveNext
    Loop

    sCad = sCad & oImp.gPrnSaltoLinea
    Dim cadenaRegla As String, cadenaReglaAnterior As String
    Dim reglas As String, regla As String
    Dim totalCaracter As Integer, nRegla As Integer, nValor As Integer
    reglas = psReglas
    totalCaracter = Len(Trim(reglas))
    nRegla = 0: nValor = 1
    
    cadenaRegla = ""
    cadenaRegla = cadenaRegla & "REGLAS "
    If totalCaracter = 0 Then
        cadenaRegla = cadenaRegla & "¦1¦ " & Space(5) & "¦2¦ " & Space(5) & "¦3¦ " & Space(5) & "¦4¦ " & Space(5) & "¦5¦ " & Space(5) & "¦6¦ " & Space(5) & "¦7¦ " & Space(5) & "¦8¦ " & Space(5)
    Else
        Do While totalCaracter > 0
            If InStr(1, reglas, "-") = 0 Then
                regla = reglas
                totalCaracter = 0
            Else
                regla = Left(reglas, InStr(1, reglas, "-") - 1)
                reglas = Right(reglas, Len(reglas) - InStr(1, reglas, "-"))
                totalCaracter = Len(reglas)
            End If
            nRegla = nRegla + 1
            cadenaReglaAnterior = cadenaRegla
            cadenaRegla = cadenaRegla & "¦" & nRegla & "¦ " & regla & " " & Space(2)
            If ((Len(cadenaRegla) - Len(cadenaReglaAnterior)) + Len(cadenaRegla)) > (137 * nValor) Then
                cadenaRegla = cadenaRegla & oImp.gPrnSaltoLinea
                nValor = nValor + 1
                nFilaHoja = nFilaHoja + 1 'Una fila mas aumenta
            End If
        Loop
    End If
    sCad = sCad & cadenaRegla & oImp.gPrnSaltoLinea
    sCad = sCad & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & "DECLARACIONES Y FIRMAS" & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    
    nFilaHoja = nFilaHoja + 7
    Do While nFilaFirma <= nTotalFilaFirma Or nFilaFirma = 1
        If nFilaHoja > 64 Then
            nPagina = nPagina + 1
            sCad = sCad & oImp.gPrnSaltoPagina
            sCad = sCad & oImpreCap.CabeRepoCaptacSolicitud("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis, oImp.gPrnSaltoLinea, sTitRp3, sUser) & oImp.gPrnSaltoLinea
            sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
            sCad = sCad & "DATOS DEL DEPOSITO" & oImp.gPrnSaltoLinea
            sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
            sCad = sCad & "TIPO DE CUENTA       : " & Trim(sDesTip) & Space(25 - Len(Trim(sDesTip))) & "PRODUCTO            : " & Trim(Left(sDesPro, 25)) & Space(25 - Len(Trim(Left(sDesPro, 25)))) & "MONEDA     : " & Trim(sMoneda) & Space(25 - Len(Trim(sMoneda))) & oImp.gPrnSaltoLinea
            sCad = sCad & "IMPORTE DE DEPOSITO  : " & Trim(Format(nSaldo, "#,##0.00")) & Space(25 - Len(Trim(Format(nSaldo, "#,##0.00")))) & oImp.gPrnSaltoLinea
            sCad = sCad & "PLAZO                : " & Trim(sPlazo) & Space(25 - Len(Trim(sPlazo))) & "T.E.A               : " & Trim(Format(nConvierteTNAaTEA, "0.00")) & "%" & Space(25 - Len(Trim(Format(nConvierteTNAaTEA, "0.00")) & "%")) & "T.R.E.A.   : " & Trim(Format(nTREA, "0.00")) & "%" & oImp.gPrnSaltoLinea
            sCad = sCad & "RETIRO DE INTERES    : " & Trim(sRetiroInteres) & Space(25 - Len(Trim(sRetiroInteres))) & "MONTO INTERES       : " & Trim(sMontoInteres) & oImp.gPrnSaltoLinea
            sCad = sCad & "FECHA DE APERTURA    : " & Format$(dFecApe, "dd mmm yyyy") & Space(25 - Len(Trim(Format$(dFecApe, "dd mmm yyyy")))) & "FECHA DE VENCIMIENTO: " & Trim(sFechaV) & oImp.gPrnSaltoLinea

            If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
                sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
            Else
                If nTpoPrograma = 6 Or nTpoPrograma = 8 Then
                    sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
                Else
                    sCad = sCad & "ENTIDAD              : " & "-" & oImp.gPrnSaltoLinea
                End If
            End If
            sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
            nFilaHoja = 16
        End If
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("------------------------------------------    ------------------------------------------    ------------------------------------------", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
        sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("FIRMA Y SELLO DEL REPRESENTANTE AUTORIZADO    FIRMA Y SELLO DEL REPRESENTANTE AUTORIZADO    FIRMA Y SELLO DEL REPRESENTANTE AUTORIZADO", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
        
        nFilaHoja = nFilaHoja + 7
        nFilaFirma = nFilaFirma + 1
    Loop
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
    
    If (nFilaHoja + 11) > 64 Then
        nPagina = nPagina + 1
        sCad = sCad & oImp.gPrnSaltoPagina
        sCad = sCad & oImpreCap.CabeRepoCaptacSolicitud("", "", nCarLin, "SECCION AHORROS", sTitRp1, sTitRp2, sMoneda, Trim(CStr(nPagina)), sAgencia, dFecSis, oImp.gPrnSaltoLinea, sTitRp3, sUser) & oImp.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & "DATOS DEL DEPOSITO" & oImp.gPrnSaltoLinea
        sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
        sCad = sCad & "TIPO DE CUENTA       : " & Trim(sDesTip) & Space(25 - Len(Trim(sDesTip))) & "PRODUCTO            : " & Trim(Left(sDesPro, 25)) & Space(25 - Len(Trim(Left(sDesPro, 25)))) & "MONEDA     : " & Trim(sMoneda) & Space(25 - Len(Trim(sMoneda))) & oImp.gPrnSaltoLinea
        sCad = sCad & "IMPORTE DE DEPOSITO  : " & Trim(Format(nSaldo, "#,##0.00")) & Space(25 - Len(Trim(Format(nSaldo, "#,##0.00")))) & oImp.gPrnSaltoLinea
        sCad = sCad & "PLAZO                : " & Trim(sPlazo) & Space(25 - Len(Trim(sPlazo))) & "T.E.A               : " & Trim(Format(nConvierteTNAaTEA, "0.00")) & "%" & Space(25 - Len(Trim(Format(nConvierteTNAaTEA, "0.00")) & "%")) & "T.R.E.A.   : " & Trim(Format(nTREA, "0.00")) & "%" & oImp.gPrnSaltoLinea
        sCad = sCad & "RETIRO DE INTERES    : " & Trim(sRetiroInteres) & Space(25 - Len(Trim(sRetiroInteres))) & "MONTO INTERES       : " & Trim(sMontoInteres) & oImp.gPrnSaltoLinea
        sCad = sCad & "FECHA DE APERTURA    : " & Format$(dFecApe, "dd mmm yyyy") & Space(25 - Len(Trim(Format$(dFecApe, "dd mmm yyyy")))) & "FECHA DE VENCIMIENTO: " & Trim(sFechaV) & oImp.gPrnSaltoLinea
        If CLng(Mid$(sCuenta, 6, 3)) = gCapCTS Then
            sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
        Else
            If nTpoPrograma = 6 Or nTpoPrograma = 8 Then
                sCad = sCad & "ENTIDAD              : " & Trim(sInstitucion) & oImp.gPrnSaltoLinea
            Else
                sCad = sCad & "ENTIDAD              : " & "-" & oImp.gPrnSaltoLinea
            End If
        End If
        sCad = sCad & String(nTamMayor + nTamAdd + 45, "-") & oImp.gPrnSaltoLinea
        nFilaHoja = 16
    End If
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("------------------------------------------             ------------------------------------------", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena("            ASESOR DEL CLIENTE                                  SUPERVISOR DE OPERACIONES        ", nTamMenor + nTamMayor + nTamAdd + 4) & "¦") & oImp.gPrnSaltoLinea
    sCad = sCad & oImpre.ImpreCarEsp("¦" & oImpre.CentrarCadena(" ", nTamMenor + nTamMayor + nTamAdd + 4) & " ¦") & oImp.gPrnSaltoLinea
    sCad = sCad & String(nTamMayor + nTamAdd + 45, "-")
    sCad = sCad & oImp.gPrnSaltoPagina
    
    GeneraSolicitudAperturaPersJuridica = sCad
End Function
Public Function GetDatosParaCronogramaPlazoFijo(ByVal psCuenta As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetDatosParaCronogramaPlazoFijo = clsMant.GetDatosParaCronogramaPlazoFijo(psCuenta)
Set clsMant = Nothing
End Function
'FIN FRHU 20140926
'JUEZ 20141014 *******************************************************
Public Function ObtenerCantidadOperaciones(ByVal psCtaCod As String, ByVal pnTipoOpe As CaptacMovTipo, ByVal psPeriodo As String) As Integer
Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rs As ADODB.Recordset
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rs = clsCap.ObtenerCantidadOperaciones(psCtaCod, pnTipoOpe, Format(psPeriodo, IIf(Mid(psCtaCod, 6, 3) = gCapCTS, "yyyy", "yyyymm")))
    ObtenerCantidadOperaciones = rs!nCantOpe
    Set clsCap = Nothing
End Function
'END JUEZ ************************************************************
'FRHU 20141201 ERS048-2014 ******************************************
Public Function GetProductosAhorros(ByVal psPersCod As String) As ADODB.Recordset
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
Set GetProductosAhorros = clsMant.GetProductosAhorros(psPersCod)
Set clsMant = Nothing
End Function
'FIN FRHU 20141201 ***************************************************
'JUEZ 20150205 **********************************************************
Public Function ValidaServCobDebitoAuto(ByVal psCtaCod As String) As ADODB.Recordset
Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set ValidaServCobDebitoAuto = clsCap.ValidaServCobDebitoAuto(psCtaCod)
    Set clsCap = Nothing
End Function
Public Function ValidaServCobDebitoAutoDet(ByVal pRs As ADODB.Recordset, ByVal psCtaCod As String, ByVal pnTipoDebito As CaptacServDebitoAutomaticoTipo) As String
Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
Dim i As Integer
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
        For i = 0 To pRs.RecordCount - 1
            If clsCap.ValidaServCobDebitoAutoDet(psCtaCod, pnTipoDebito, IIf(pnTipoDebito = gServConvenio, pRs(1), pRs(0)), pRs(2)) Then
                If pnTipoDebito = gServConvenio Then
                    ValidaServCobDebitoAutoDet = "El Código de Usuario " & pRs(2) & " ya se encuentra registrado en el pago por débito automático para el Convenio " & pRs(1)
                Else
                    ValidaServCobDebitoAutoDet = "El crédito " & pRs(0) & " ya se encuentra registrado en el pago por débito automático"
                End If
                Exit Function
            End If
        Next i
    Set clsCap = Nothing
End Function
Public Function ObtenerServCobDebitoAuto(ByVal pnMovNro As Long, ByVal psUsuario As String, ByVal psFecha As String, ByVal psAgeCod As String) As ADODB.Recordset
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set ObtenerServCobDebitoAuto = oDCapGen.ObtenerServCobDebitoAuto(pnMovNro, psUsuario, psFecha, psAgeCod)
Set oDCapGen = Nothing
End Function
'END JUEZ ***************************************************************
'JUEZ 20151229 *****************************************************************
Public Sub RegistrarPendComisionMantPoderes(ByVal psCuenta As String, ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal pnComiMantPoderes As Double)
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    oDCapGen.RegistrarPendComisionMantPoderes psCuenta, pnMovNro, psMovNro, pnComiMantPoderes
Set oDCapGen = Nothing
End Sub
'ANDE SE AGREGÓ EL PARAMETRO pdFecSis 20171108
Public Sub VerificarCobroComisionRetencionFondos(ByVal pRs As ADODB.Recordset, ByVal psCuenta As String, ByVal psMovNro As String, ByVal pdFecSis As Date, Optional ByVal pbBloqueoParcial As Boolean = False, Optional ByVal pnMontoBloq As Double = 0)
Dim nMotRet As Long 'JUEZ 20160226
Dim nMontoBloqParc As Double 'JUEZ 20160226
    If Not pRs Is Nothing Then
        pRs.MoveFirst
        Do While Not pRs.EOF
            nMotRet = CLng(Trim(Right(pRs("Motivo"), 3))) 'JUEZ 20160226
            If pbBloqueoParcial Then nMontoBloqParc = CDbl(pRs("Monto")) 'JUEZ 20160226
            If pRs("Flag") = "N" And IIf(pRs("Est") = "0", False, True) Then
                If nMotRet = 1 Or nMotRet = 19 Then 'JUEZ 20160226
                    'ANDE 20170726 AGREGADO DE pdFecSis
                    If Mid(psCuenta, 6, 3) = gCapAhorros Then CobrarComisionRetencionFondos psCuenta, psMovNro, pbBloqueoParcial, nMontoBloqParc, pdFecSis
                    'END ANDE
                    Exit Do
                End If
            End If
            pRs.MoveNext
        Loop
    End If
End Sub
'ANDE SE AGREGÓ EL PARAMETRO pdFecSis 20171108
Public Sub CobrarComisionRetencionFondos(ByVal psCuenta As String, ByVal psMovNro As String, Optional ByVal pbBloqueoParcial As Boolean = False, Optional ByVal pnMontoBloq As Double = 0, Optional ByVal pdFecSis As Date)
Dim clsCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim rs As New ADODB.Recordset
Dim lnSaldo As Double
Dim pnMovNro As Long
'ANDE 20171108
Dim nMontoACobrar, nTC As Double
Dim nTipoMoneda As Integer
Dim clsTC As New COMDConstSistema.NCOMTipoCambio

nTC = clsTC.EmiteTipoCambio(pdFecSis, TCFijoDia)
Set clsTC = Nothing
nTipoMoneda = IIf(Mid(psCuenta, 9, 1) = "1", 1, 2)

Set clsCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
Set rs = clsCapDef.GetParametrosComision("CGA2002", , "A")
Set clsCapDef = Nothing

If nTipoMoneda = 1 Then
    nMontoACobrar = CDbl(rs!nParMonto)
Else
    nMontoACobrar = CDbl(rs!nParMonto) / nTC
End If
nMontoACobrar = Format(nMontoACobrar, "#.##")
'se agrego cambio CDbl(rs!nParMonto) por nMontoACobrar
If pbBloqueoParcial Then
    Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    If Not oNCapMov.ValidaSaldoCuenta(psCuenta, pnMontoBloq + nMontoACobrar) Then Exit Sub
    Set oNCapMov = Nothing
End If
    
Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
lnSaldo = oNCapMov.CapCargoCuentaAho(psCuenta, nMontoACobrar, gAhoCargoCobroComDiversasAho, psMovNro, "Cuenta = " & psCuenta, , , , , , , , , , , , , , , "109", , Mid(psMovNro, 18, 2), , , 0, , , gAhoCargoCobroComDiversasAho, , , , , , , , , , , , , pnMovNro, , , , , , , , , , 12002)

'end ande
    If pnMovNro > 0 Then
        Call oNCapMov.InsertaCapComisiones(pnMovNro, "CGA2002", "1090100012521", 2, psCuenta)
    End If
    Set oNCapMov = Nothing
End Sub
'END JUEZ **********************************************************************
'RIRO20160311 ********************************
Public Function GetDoiCta(ByVal psCtaCod As String) As String
Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
GetDoiCta = clsCap.GetDoiCta(psCtaCod)
Set clsCap = Nothing
End Function
'END RIRO ************************************

'************** ANDE 20170615 ARQUEO DE EXPEDIENTES DE AHORRO ***************

Public Function ObtenerDatosParticipantes(ByVal pcUsuario As String) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    
    Set R = oCaptaAD.ObtenerDatosParticipantesArqueoExpAho(pcUsuario)
    Set ObtenerDatosParticipantes = R
    Set oCaptaAD = Nothing
    Set R = Nothing
End Function

Public Function ObtenerCuentasAperturadas(ByVal pcDel As String, ByVal pcAl As String, ByVal pcAgeCod As String) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.ObtenerCuentasAperturadas(pcDel, pcAl, pcAgeCod)
    Set ObtenerCuentasAperturadas = R
    Set oCaptaAD = Nothing
    Set R = Nothing
End Function

Public Function ObtenerTotalCuentasAperturadas(ByVal pcDel As String, ByVal pcAl As String, ByVal pcAgeCod As String)
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.ObtenerTotalCuentasAperturadas(pcDel, pcAl, pcAgeCod)
    Set ObtenerTotalCuentasAperturadas = R
    Set oCaptaAD = Nothing
    Set R = Nothing
End Function

Public Sub GuardarPistaArqueoExpAho(ByVal pcCodArqueo As String, ByVal pcPersCod As String, ByVal pnTipoArqueo As Integer, ByVal pcDesTipoArqueo As String, ByVal pnPerfil As Integer, ByVal pcAgeCod As String, ByVal pcFecArqueo As String)
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Call oCaptaAD.GuardarPistaArqueoExpAho(pcCodArqueo, pcPersCod, pnTipoArqueo, pcDesTipoArqueo, pnPerfil, pcAgeCod, pcFecArqueo)
    Set oCaptaAD = Nothing
End Sub

Public Function ObtenerArqueoDelDia(ByVal pdDia As Date) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    
    Set R = oCaptaAD.ObtenerArqueosDeldia(pdDia)
    Set ObtenerArqueoDelDia = R
    Set R = Nothing
    Set oCaptaAD = Nothing
    
End Function

Public Sub ExtornarArqueoExpAho(ByVal cIdArqueo As String)
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Call oCaptaAD.ExtornarArqueoExpAho(cIdArqueo)
    Set oCaptaAD = Nothing
End Sub

Public Function ObtenerCiudad(ByVal pcAgeCod As String) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.ObtenerCiudad(pcAgeCod)
    Set ObtenerCiudad = R
    Set R = Nothing
End Function

Public Sub RegistrarDF(ByVal pcCodArqueo As String, ByVal pcCuenta As String, ByVal pdFecha As Date, ByVal pcComentario As String, ByVal pcMotivo As String)
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Call oCaptaAD.RegistrarDeficientesYFaltantes(pcCodArqueo, pcCuenta, pdFecha, pcComentario, pcMotivo)
    Set oCaptaAD = Nothing
End Sub

Public Function ObtenerCodigoArqueo(ByVal pcFecha As String) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.ObtenerCodigoArqueo(pcFecha)
    Set ObtenerCodigoArqueo = R
    Set R = Nothing
End Function

'********************** END ANDE ****************************
'******************* ANDE 16/04/2018 verificación si participa de campaña
Public Sub ParticipaDeCampania(ByVal pcPersCod As String, ByVal pcCtaCod As String, ByVal pnMoneda As Integer, ByVal pnMonto As Currency, ByRef pbParticipa As Boolean, ByVal pcMovNro As String, _
Optional ByVal pnTipoCambio As Currency, Optional ByVal pnTipoApertura As Integer = 1, Optional ByVal pnTipoProceso As Integer = 1, Optional ByVal pnDias As Integer = 0, _
Optional ByVal pnTipoPersona As Integer = 1, Optional ByRef pnPuntosRef As Integer, Optional ByRef pnCondicion As Integer, Optional pcNomCliente As String, Optional ByRef pnTotalAcumulado As Integer)
    'pnTipoProceso 1 Apertura: 2 Deposito: 3 Retiro : 4 Cancelación
    Dim nPuntosGanados As Integer, nMontoPunto As Currency, num_entero As Integer, num_str As String
    Dim nMontoSumatProdCS As Currency, nMontoDifCS As Currency, nSaldoAcum As Currency 'verificando saldos
    Dim cProductos As String, nPuntoAcumulado As Integer, nPuntoAcumuladoAH As Integer, nPuntoAcumuladoPF As Integer
    Dim rInfoParticipante As ADODB.Recordset
    
    pbParticipa = False
    If pnTipoPersona = 1 Then
        If pnTipoProceso = 1 Or pnTipoProceso = 2 Then
            If pnTipoProceso = 1 Then 'apertura
            
                'proceso '1:apertura y 2:deposito
                Dim bTienePunto As Boolean, oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales, bHayCuentasCanceladas As Boolean
                'apertura si es plazo fijo
                If pnTipoApertura = 2 Then
                    'If (pnMonto >= 1000 And pnMonto <= 350000) And pnMoneda = 1 Then
                    If (pnMonto >= 500 And pnMonto <= 350000) And pnMoneda = 1 Then 'APRI20191128
                        'If pnDias >= 90 And pnDias <= 359 Then
                        If pnDias >= 180 And pnDias <= 2000 Then 'APRI20191128
                        
                            'extraer cuentas de ahorro canceladas y ordenadas por fecha de cancelación
                            Dim rCuentas As New ADODB.Recordset
                            Dim nTpoPrograma As Integer, nProducto As Integer, nMontoRetiro As Currency, nMoneda As Integer
                            'nMontoPunto = 1000 'valor en soles por punto
                            nMontoPunto = 500 'valor en soles por punto 'APRI20191128
                            
                            Set rCuentas = oCaptaAD.ExtraerCuentasCanceladas(pcPersCod)
                            bHayCuentasCanceladas = False
                            
                            If Not (rCuentas.BOF And rCuentas.EOF) Then
                                bHayCuentasCanceladas = True
                                Do While Not rCuentas.EOF
                                    nTpoPrograma = rCuentas!nTpoPrograma
                                    nProducto = rCuentas!nProducto
                                    nMontoRetiro = rCuentas!nMontoRetiro
                                    nMoneda = rCuentas!nMoneda
                                    
                                    
                                    If nTpoPrograma = 5 And nProducto = 232 Then 'si es ahorro soñada
                                        'verificando moneda para considerar tipo de cambio
                                        'If nMoneda = 2 Then
                                        '    nMonto = nMonto * pnTipoCambio
                                        'End If
                                        If nMoneda <> 2 Then
                                            'si monto retiro anterior es menor que el de apertura, consigue puntos, claro que se debe aplicar la regla de punto
                                            If nMontoRetiro < pnMonto Then
                                                If (nMontoRetiro / nMontoPunto) >= 1 Then
                                                    'tiene punto:  S/ 1000 = 1 punto
                                                    bTienePunto = True
                                                    num_str = Trim(CStr(pnMonto / nMontoPunto))
                                                    If InStr(num_str, ".") > 0 Then
                                                        num_str = Left(num_str, InStr(num_str, ".") - 1)
                                                    End If
                                                    num_entero = CInt(num_str)
                                                    nPuntosGanados = num_entero
                                                   GoTo continuarAqui
                                                End If
                                            End If
                                        End If
                                    End If
                                    rCuentas.MoveNext
                                Loop
                                If Not bTienePunto Then
                                    If Round(pnMonto / nMontoPunto) >= 1 Then
                                        'tiene punto:  S/ 1000 = 1 punto
                                        bTienePunto = True
                                        num_str = Trim(CStr(pnMonto / nMontoPunto))
                                        If InStr(num_str, ".") > 0 Then
                                            num_str = Left(num_str, InStr(num_str, ".") - 1)
                                        End If
                                        num_entero = CInt(num_str)
                                        nPuntosGanados = num_entero
                                    End If
                               End If
continuarAqui:
                            Else
                                'no tiene cuentas canceladas pero si debe participar
                                    If (pnMonto / nMontoPunto) >= 1 Then
                                        'tiene punto:  S/ 1000 = 1 punto
                                        bTienePunto = True
                                        num_str = Trim(CStr(pnMonto / nMontoPunto))
                                        If InStr(num_str, ".") > 0 Then
                                            num_str = Left(num_str, InStr(num_str, ".") - 1)
                                        End If
                                        num_entero = CInt(num_str)
                                        nPuntosGanados = num_entero
                                    End If
                            End If
                        End If
                    End If
                Else
                    'apertura si es ahorro
                    nMontoPunto = 100 'valor en soles por punto
                    
                    If pnMoneda = 1 And pnMonto >= 100 Then
                        If (pnMonto / nMontoPunto) >= 1 Then
                            bTienePunto = True
                            num_str = Trim(CStr(pnMonto / nMontoPunto))
                            If InStr(num_str, ".") > 0 Then
                                num_str = Left(num_str, InStr(num_str, ".") - 1)
                            End If
                            num_entero = CInt(num_str)
                            nPuntosGanados = num_entero
                        End If
                    End If
                End If
            Else 'depósito ahorro
                Set rInfoParticipante = oCaptaAD.ExtraerInfoParticipantesCamp(pcPersCod)
                If Not rInfoParticipante.EOF Then
                    cProductos = rInfoParticipante!TipoProductos
                    'nPuntoAcumulado = rInfoParticipante!nPuntoAcumulado
                    nPuntoAcumuladoAH = rInfoParticipante!nPuntoAcumuladoAH
                    nPuntoAcumuladoPF = rInfoParticipante!nPuntoAcumuladoPF
                    
                    'If InStr(cProductos, "232") > 0 Then
                    If (Mid(pcCtaCod, 6, 3) = "232") Then 'APRI20191128
                        nMontoSumatProdCS = rInfoParticipante!nMontoSumatProdCS
                        nMontoDifCS = rInfoParticipante!nMontoDifCS
                        nSaldoAcum = nMontoDifCS + pnMonto
                        If (nSaldoAcum > 0 And nSaldoAcum > nMontoSumatProdCS) Then
                            nSaldoAcum = nSaldoAcum - nMontoSumatProdCS
                            nMontoPunto = 100
                            If (nSaldoAcum / nMontoPunto) >= 1 Then
                                num_str = Trim(CStr(nSaldoAcum / nMontoPunto))
                                If InStr(num_str, ".") > 0 Then
                                    num_str = Left(num_str, InStr(num_str, ".") - 1)
                                End If
                                num_entero = CInt(num_str)
                                bTienePunto = True
                            End If
                        End If
                    'APRI20191128
                    Else
                        nMontoPunto = 500
                        If (pnMonto / nMontoPunto) >= 1 Then
                                'tiene punto:  S/ 1000 = 1 punto
                                bTienePunto = True
                                num_str = Trim(CStr(pnMonto / nMontoPunto))
                                If InStr(num_str, ".") > 0 Then
                                    num_str = Left(num_str, InStr(num_str, ".") - 1)
                                End If
                                num_entero = CInt(num_str)
                                nPuntosGanados = num_entero
                        End If
                    'END APRI
                    End If
                Else
                    'sino estaba participando en la campaña y debe hacerlo
                    nMontoPunto = 100 'valor en soles por punto
                    
                    If pnMoneda = 1 And pnMonto >= 100 Then
                        If (pnMonto / nMontoPunto) >= 1 Then
                            bTienePunto = True
                            num_str = Trim(CStr(pnMonto / nMontoPunto))
                            If InStr(num_str, ".") > 0 Then
                                num_str = Left(num_str, InStr(num_str, ".") - 1)
                            End If
                            num_entero = CInt(num_str)
                        End If
                    End If
                End If
                nPuntosGanados = num_entero
                nPuntoAcumulado = nPuntosGanados
                
            End If
            Dim nPuntoMuestra As Integer
            If bTienePunto And pnTipoProceso = 1 Then
                nPuntoMuestra = nPuntosGanados
                If nPuntoMuestra > 0 Then
                    Call oCaptaAD.RegistrarParticipanteCampanaAho(pcCtaCod, pnMonto, pnTipoProceso, nPuntosGanados, pcPersCod, pcMovNro)
                End If
            ElseIf pnTipoProceso = 2 Then
                nPuntoMuestra = nPuntoAcumulado - nPuntoAcumuladoAH
                If nPuntoMuestra > 0 Then
                    Call oCaptaAD.RegistrarParticipanteCampanaAho(pcCtaCod, pnMonto, pnTipoProceso, nPuntoAcumulado, pcPersCod, pcMovNro)
                End If
            End If
            
            If bTienePunto Then
                'Call oCaptaAD.RegistrarParticipanteCampanaAho(pcCtaCod, pnMonto, pnTipoProceso, nPuntosGanados, pcPersCod, pnMovNro)
                'extrayendo total acumulado
                Set rInfoParticipante = oCaptaAD.ExtraerInfoParticipantesCamp(pcPersCod)
                pnTotalAcumulado = rInfoParticipante!nPuntoAcumulado
                'nPuntoMuestra = pnTotalAcumulado - nPuntoMuestra
                'fin total acumulado
                If pnTotalAcumulado >= 0 And nPuntoMuestra > 0 Then
                    'MsgBox "El cliente " & pcNomCliente & "se ha ganado " & CStr(nPuntoMuestra) & " punto(s) por su operación en nuestra campaña mundial del ahorro 2018. ", vbInformation + vbOKOnly, "Aviso"
                    MsgBox "El cliente " & pcNomCliente & "se ha ganado " & CStr(nPuntoMuestra) & " punto(s) por su operación en nuestra campaña Ahorra Más. ", vbInformation + vbOKOnly, "Aviso" 'APRI20191128
                    pbParticipa = True
                    pnPuntosRef = nPuntoMuestra
                    pnCondicion = 0 'gana puntos o gol
                End If
            End If
            
        Else
            '3:retiro y 4:cancelación
            Set rInfoParticipante = oCaptaAD.ExtraerInfoParticipantesCamp(pcPersCod)
            If Not rInfoParticipante.EOF Then
                cProductos = rInfoParticipante!TipoProductos
                nPuntoAcumulado = rInfoParticipante!nPuntoAcumulado
                nPuntoAcumuladoAH = rInfoParticipante!nPuntoAcumuladoAH
                nPuntoAcumuladoPF = rInfoParticipante!nPuntoAcumuladoPF
                
                If pnTipoProceso = 3 Then
                    Dim bQuitaPunto As Boolean, nPuntoAnterior As Integer
                    If (Mid(pcCtaCod, 6, 3) = "232") Then
                        If InStr(cProductos, "232") > 0 Then
                            nMontoSumatProdCS = rInfoParticipante!nMontoSumatProdCS
                            nMontoDifCS = rInfoParticipante!nMontoDifCS
                            nSaldoAcum = nMontoDifCS - pnMonto
                    
                            nMontoPunto = 100
                            If (nSaldoAcum > 0 And nSaldoAcum > nMontoSumatProdCS) Then
                                If (nSaldoAcum / nMontoPunto) >= 1 Then
                                    nSaldoAcum = nSaldoAcum - nMontoSumatProdCS
                                    num_str = Trim(CStr(nSaldoAcum / nMontoPunto))
                                    If InStr(num_str, ".") > 0 Then
                                        num_str = Left(num_str, InStr(num_str, ".") - 1)
                                    End If
                                    num_entero = CInt(num_str)
                                    nPuntosGanados = num_entero
                                    nPuntoAnterior = nPuntoAcumuladoAH
                                    nPuntoAcumulado = IIf((nPuntoAcumuladoAH - nPuntosGanados) < 0, nPuntoAnterior, (nPuntoAcumuladoAH - nPuntosGanados))
                                    'nPuntoAcumulado = nPuntoAnterior
                                    bQuitaPunto = True
'                                    num_str = Trim(CStr(pnMonto / nMontoPunto))
'                                    If InStr(num_str, ".") > 0 Then
'                                        num_str = Left(num_str, InStr(num_str, ".") - 1)
'                                    End If
'                                    num_entero = CInt(num_str)
'                                    nPuntosGanados = num_entero
'                                    nPuntoAnterior = nPuntoAcumuladoAH
'                                    nPuntoAcumulado = IIf((nPuntoAcumuladoAH - nPuntosGanados) < 0, nPuntoAnterior, (nPuntoAcumuladoAH - nPuntosGanados))
'                                    bQuitaPunto = True
                                End If
                            Else
                                 If (pnMonto / nMontoPunto) >= 1 Then
                                    num_str = Trim(CStr(pnMonto / nMontoPunto))
                                    If InStr(num_str, ".") > 0 Then
                                        num_str = Left(num_str, InStr(num_str, ".") - 1)
                                    End If
                                    num_entero = CInt(num_str)
                                    nPuntosGanados = num_entero
                                    nPuntoAnterior = nPuntoAcumuladoAH
                                    nPuntoAcumulado = nPuntosGanados
                                End If
                                bQuitaPunto = True
                            End If
                        End If
                    Else
                        If InStr(cProductos, "233") > 0 Then
                            nMontoPunto = 1000
                            If (pnMonto / nMontoPunto) >= 1 Then
                                num_str = Trim(CStr(pnMonto / nMontoPunto))
                                If InStr(num_str, ".") > 0 Then
                                    num_str = Left(num_str, InStr(num_str, ".") - 1)
                                End If
                                num_entero = CInt(num_str)
                                nPuntosGanados = num_entero
                                nPuntoAnterior = nPuntoAcumulado
                                nPuntoAcumulado = IIf((nPuntoAcumuladoPF - nPuntosGanados) < 0, nPuntoAnterior, (nPuntoAcumuladoPF - nPuntosGanados))
                                bQuitaPunto = True
                            End If
                        End If
                    End If
                        
                    If bQuitaPunto Then
                        Set rInfoParticipante = oCaptaAD.ExtraerInfoParticipantesCamp(pcPersCod)
                        pnTotalAcumulado = rInfoParticipante!nPuntoAcumulado
                        pnTotalAcumulado = pnTotalAcumulado - nPuntoAcumulado
                        
                        If pnTotalAcumulado >= 0 And nPuntoAcumulado > 0 Then
                            pbParticipa = True
                            Call oCaptaAD.RegistrarParticipanteCampanaAho(pcCtaCod, pnMonto, pnTipoProceso, nPuntoAcumulado, pcPersCod, pcMovNro)
                            'MsgBox "El cliente " & pcNomCliente & IIf(pcNomCliente <> "", " ", "") & "ha perdido " & CStr(nPuntoAcumulado) & " punto(s) por su operación, en nuestra campaña mundial del ahorro 2018. ", vbInformation + vbOKOnly, "Aviso"
                            MsgBox "El cliente " & pcNomCliente & IIf(pcNomCliente <> "", " ", "") & "ha perdido " & CStr(nPuntoAcumulado) & " punto(s) por su operación, en nuestra campaña Ahorra Más. ", vbInformation + vbOKOnly, "Aviso" 'APRI20191128
                            pnPuntosRef = nPuntoAcumulado
                            pnCondicion = 1 'pierde puntos o autogol
                        End If
                    End If
                Else
                    Dim nPuntoExtraido As Integer
                    nPuntoExtraido = ExtraePuntoCuenta(pcCtaCod)
                    nPuntoAcumulado = nPuntoAcumulado - nPuntoExtraido
                    pnTotalAcumulado = nPuntoAcumulado
                    
                    If pnTotalAcumulado >= 0 And nPuntoExtraido > 0 Then
                        pbParticipa = True
                        Call oCaptaAD.RegistrarParticipanteCampanaAho(pcCtaCod, pnMonto, pnTipoProceso, nPuntoExtraido, pcPersCod, pcMovNro)
                        'MsgBox "El cliente " & pcNomCliente & IIf(pcNomCliente <> "", " ", "") & "ha perdido " & CStr(nPuntoExtraido) & " punto(s) por su operación, en nuestra campaña mundial del ahorro 2018. ", vbInformation + vbOKOnly, "Aviso"
                        MsgBox "El cliente " & pcNomCliente & IIf(pcNomCliente <> "", " ", "") & "ha perdido " & CStr(nPuntoExtraido) & " punto(s) por su operación, en nuestra campaña Ahorra Más. ", vbInformation + vbOKOnly, "Aviso" 'APRI20191128
                        pnPuntosRef = nPuntoExtraido
                        pnCondicion = 1 'pierde puntos o autogol
                    End If
                End If
            End If
            
            
        End If
    End If
End Sub
Public Function getVerificarPersonaNatJur(ByVal cPersCod As String) As Integer
    Dim R As ADODB.Recordset, oCampanha As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCampanha = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set R = oCampanha.getVerificarPersonaNatJur(cPersCod)
    getVerificarPersonaNatJur = R!nPersPersoneria
    Set oCampanha = Nothing
End Function
Public Sub VerificarParticipacionCampMundial(ByVal pcPersCod As String, ByVal pcCtaCod As String, ByVal pnOperacion As COMDConstantes.CaptacOperacion, ByVal pnOpeTipo As Integer, ByVal pnMoneda As Integer, ByVal pnMonto As Currency, ByVal pnTipoPersona As Integer, ByRef pbParticipa As Boolean, ByVal pcMovNro As String, _
Optional ByVal pnDias As Integer, Optional ByVal pdSistema As Date, Optional ByVal pnSubProducto As Integer, Optional ByVal pcAgencia As String, Optional ByRef nPuntosRef As Integer, Optional pnCondicion As Integer, Optional pcNomCliente As String, Optional ByRef nPTotalCumulado As Integer)
    Dim rCampVigentes As ADODB.Recordset
    Set rCampVigentes = getCampVigentes(pdSistema)
    pbParticipa = False
    If Not (rCampVigentes.BOF And rCampVigentes.EOF) Then
        'nCampanaCod
'        cCampanaDesc
'        nMoneda
'        nProducto
'        nSubProducto
'        nMontoMin
'        nPeriodoVigMeses
'        dFechaIni
'        dFechaFin
'        bPersJur,
'        cAgencias
        Dim bCamp As Boolean
        bCamp = VerificarCtaCodCamp(pcCtaCod)
        
        If pnTipoPersona = 1 Then
            If pnOperacion = gAhoApeEfec Or pnOperacion = gAhoApeTransf Or pnOperacion = gAhoApeLoteEfec Or pnOperacion = gAhoApeLoteTransfBanco _
                Or pnOperacion = gPFApeEfec Or pnOperacion = gPFApeChq Or pnOperacion = gPFApeTransf Or pnOperacion = gPFApeLoteEfec Then
                
                Dim nTipoApertura As Integer
                If pnOperacion = gAhoApeEfec Or pnOperacion = gAhoApeChq Or pnOperacion = gAhoApeTransf Or pnOperacion = gAhoApeLoteEfec Or pnOperacion = gAhoApeLoteTransfBanco _
                Then
                    nTipoApertura = 1 'ahorro
                Else
                    nTipoApertura = 2 'plazo fijo
                End If
                
                Do While Not rCampVigentes.EOF ' poes esta ves solo se extrae una campaña , es decir solo debe haber una campaña mundialito o que funcione como tal
                    If pnMoneda = rCampVigentes!nMoneda And Mid(pcCtaCod, 6, 3) = Trim(rCampVigentes!nProducto) _
                    And pnMonto >= rCampVigentes!nMontoMin And rCampVigentes!bPersJur = False _
                    And InStr(rCampVigentes!cAgencias, pcAgencia) > 0 _
                    Then
                        If Trim(rCampVigentes!nSubProducto) = pnSubProducto Then
                            Call ParticipaDeCampania(pcPersCod, pcCtaCod, pnMoneda, pnMonto, pbParticipa, pcMovNro, , nTipoApertura, pnOpeTipo, pnDias, , nPuntosRef, pnCondicion, pcNomCliente, nPTotalCumulado)
                        End If
                    End If
                    rCampVigentes.MoveNext
                Loop
            Else
                Do While Not rCampVigentes.EOF
                    If pnMoneda = rCampVigentes!nMoneda And Mid(pcCtaCod, 6, 3) = Trim(rCampVigentes!nProducto) _
                        And rCampVigentes!bPersJur = False _
                        And InStr(rCampVigentes!cAgencias, pcAgencia) > 0 _
                    Then
                        If pnOperacion = gAhoDepEfec Or pnOperacion = gAhoDepChq Or pnOperacion = gAhoDepTransf Then
                            If Trim(rCampVigentes!nSubProducto) = pnSubProducto Then
                               Call ParticipaDeCampania(pcPersCod, pcCtaCod, pnMoneda, pnMonto, pbParticipa, pcMovNro, , , pnOpeTipo, , , nPuntosRef, pnCondicion, pcNomCliente, nPTotalCumulado)
                            End If
                        Else
                            If Trim(rCampVigentes!nSubProducto) = pnSubProducto And bCamp Then
                                Call ParticipaDeCampania(pcPersCod, pcCtaCod, pnMoneda, pnMonto, pbParticipa, pcMovNro, , , pnOpeTipo, , , nPuntosRef, pnCondicion, pcNomCliente, nPTotalCumulado)
                            End If
                        End If
                    End If
                    rCampVigentes.MoveNext
                Loop
            End If
        End If
    End If
End Sub

Public Function getCampVigentes(ByVal pdFecSis As Date) As ADODB.Recordset
    Dim oCaptaLN As New COMDCaptaGenerales.DCOMCaptaGenerales
    Set getCampVigentes = oCaptaLN.CampaniaVigentes(pdFecSis)
    Set oCaptaLN = Nothing
End Function

Public Function VerificarCtaCodCamp(ByVal pcCtaCod As String) As Boolean
    Dim oCaptaLN As New COMDCaptaGenerales.DCOMCaptaGenerales, R As ADODB.Recordset
    On Error GoTo ErrorVerificarCtaCodCamp
    Set R = oCaptaLN.VerificarCtaCodCamp(pcCtaCod)
    If Not (R.BOF And R.EOF) Then
        VerificarCtaCodCamp = IIf(R!nExiste = 0, True, False)
    Else
        VerificarCtaCodCamp = False
    End If
    Set R = Nothing
    Set oCaptaLN = Nothing
    Exit Function
ErrorVerificarCtaCodCamp:
    VerificarCtaCodCamp = False
End Function

Public Function ExtraePuntoCuenta(ByVal pcCtaCod As String) As Integer
    Dim oCapta As New COMDCaptaGenerales.DCOMCaptaGenerales, R As ADODB.Recordset
    Set R = oCapta.ExtraerPuntoCuenta(pcCtaCod)
    Dim nPuntoExtraido As Integer
    If Not (R.BOF And R.EOF) Then
        nPuntoExtraido = R!nPunto
        ExtraePuntoCuenta = nPuntoExtraido
    End If
End Function


'********************** END ANDE ****************************
'APRI20180405 ERS036-2017
Public Function GeneraSolicitudAfiliacionDesafiliacionEnvioEstadoCta(ByVal nTipo As Integer, ByVal sNombre As String, sDoi As String, _
    ByVal dFecAfiliacion As String, ByVal sTpoProducto As String, ByVal rsProd As Variant, ByVal pnCantProd As Integer, _
    ByVal sTpoEnvio As String, ByVal sCorreo As String, ByVal sDireccion As String, ByVal sAgecod As String, ByVal sCmacm As String, ByVal sAgencia As String, ByVal sUser As String, _
    Optional pcNroItem As String) As String
    'APRI20190615 ADD pcNroItem
    Dim i As Integer, nCarLin As Long
    Dim sMoneda As String
    Dim sCad As String
    Dim sNumPag As String
    Dim lsTitulo As String
    Dim lsTpo As String
    Dim lsTab As String
    sCad = ""
    
    Dim nTamMayor As Long, nTamMenor As Long, nTamAdd As Long
    Dim nItem As Integer
    Dim nPagina As Integer
    
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oImpreCap As New NCOMCaptaImpresion
    Dim oFunLetras As New COMFunciones.FCOMCadenas
    Dim oImpre As New COMFunciones.FCOMImpresion
    Dim oFunCad As New COMFunciones.FCOMCadenas
    
    
    Dim psTitulo As String
    Dim psSubTitulo As String
    Dim rsUbi As ADODB.Recordset
    Dim rs As ADODB.Recordset
    
    If nTipo = 1 Then
        psTitulo = "AFILIACIÓN A ENVIO DE ESTADOS DE CUENTAS DE " & UCase(sTpoProducto)
        psSubTitulo = "Solicitud de Afiliación"
    Else
        psTitulo = "DESAFILIACIÓN A ENVIO DE ESTADOS DE CUENTAS DE " & UCase(sTpoProducto)
        psSubTitulo = "Solicitud de Desafiliación"
    End If
    '*****CABECERA
    sCad = sCad & lsTab & sCmacm & Space(95) & "FECHA   :" & dFecAfiliacion & Chr$(10)
    
    sCad = sCad & lsTab & UCase(sAgencia) & Space(124 - Len(sAgencia)) & "USUARIO : " & sUser & Chr$(10)
    
    sCad = sCad & lsTab & Space((170 - Len(psTitulo)) / 2 - 18) & psTitulo & Chr$(10)
    sCad = sCad & lsTab & Space((170 - Len(psSubTitulo)) / 2 - 18) & psSubTitulo & Chr$(10) 'DAOR 20070816
    '*****END CABECERA
    
    lsTpo = IIf(nTipo = 1, "Solicitud de Afiliación", "Solicitud de Desafiliación")

    sCad = sCad & lsTab & String(138, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & lsTab & "DOI                      : " & sDoi & oImp.gPrnSaltoLinea
    sCad = sCad & lsTab & "NOMBRE CLIENTE           : " & sNombre & oImp.gPrnSaltoLinea
    sCad = sCad & lsTab & "CORREO                   : " & sCorreo & oImp.gPrnSaltoLinea
    sCad = sCad & lsTab & "DIRECCIÓN                : " & sDireccion & oImp.gPrnSaltoLinea
    If nTipo = 1 Then
    sCad = sCad & lsTab & "FECHA DE AFILIACIÓN      : " & dFecAfiliacion & oImp.gPrnSaltoLinea
    sCad = sCad & lsTab & "MEDIO DE ENVÍO           : " & sTpoEnvio & oImp.gPrnSaltoLinea
    End If
    sCad = sCad & lsTab & "TIPO DE PRODUCTO         : " & sTpoProducto & oImp.gPrnSaltoLinea
    If nTipo = 1 Then
    sCad = sCad & lsTab & "COD. AFILIACIÓN          : " & pcNroItem & oImp.gPrnSaltoLinea 'APRI20190615 SATI
    sCad = sCad & lsTab & "CUENTA(S) A AFILIAR      : " & pnCantProd & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    sCad = sCad & Space(35) & "PRODUCTO" & Space(25) & "CUENTA" & oImp.gPrnSaltoLinea
    sCad = sCad & Space(35) & String(50, "-") & oImp.gPrnSaltoLinea
    Else
    sCad = sCad & lsTab & "COD. DESAFILIACIÓN       : " & pcNroItem & oImp.gPrnSaltoLinea 'APRI20190615 SATI
    sCad = sCad & lsTab & "CUENTA(S) A DESAFILIAR   : " & pnCantProd & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    sCad = sCad & Space(35) & "PRODUCTO" & Space(25) & "CUENTA" & Space(17) & "MEDIO ENVÍO" & oImp.gPrnSaltoLinea
    sCad = sCad & Space(35) & String(80, "-") & oImp.gPrnSaltoLinea
    End If
    
    
   
    For i = 0 To pnCantProd - 1
      If nTipo = 1 Then
        sCad = sCad & Space(35) & rsProd(i, 1) & Space(33 - Len(rsProd(i, 1))) & rsProd(i, 0) & oImp.gPrnSaltoLinea
      Else
        sCad = sCad & Space(35) & rsProd(i, 1) & Space(33 - Len(rsProd(i, 1))) & rsProd(i, 0) & Space(6) & UCase(rsProd(i, 2)) & oImp.gPrnSaltoLinea
      End If
    Next i
    
    If nTipo = 1 Then
    sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    sCad = sCad & "(*)Si la seleccion es Envío de Estado de Cuenta por medio Físico, se le aplicará cargo a su cuenta de Ahorros y/o Créditos, según comisión " & oImp.gPrnSaltoLinea
    'sCad = sCad & "su cuenta de Ahorros y/o Créditos, según comisión establecida en el tarifario  de vigente" & oImp.gPrnSaltoLinea
    sCad = sCad & "establecida en el tarifario  de vigente de servicios tranversales. Dicha comisión se considerá por cuenta de Ahorro o Crédito y por cliente." & oImp.gPrnSaltoLinea
    'sCad = sCad & "por cliente" & oImp.gPrnSaltoLinea
    End If
    
    sCad = sCad & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    
    'ciudad
    Dim sEmision As String
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsUbi = clsMant.RecuperaCiudadAgencia(sAgecod)
    sEmision = rsUbi!cUbiGeoDescripcion
    Dim liPosicion As Integer
    Dim lsCiudad As String
    lsCiudad = Trim(sEmision)
    liPosicion = InStr(lsCiudad, "(")
    If liPosicion > 0 Then
    lsCiudad = Left(lsCiudad, liPosicion - 1)
    End If
    Set rsUbi = Nothing
    
    'fecha
    Dim sql As String
    Dim fecha As String
    Dim oConect As COMConecta.DCOMConecta
    Set oConect = New COMConecta.DCOMConecta
    sql = "select dbo.fnc_GetFechaCompleta('" & Format(dFecAfiliacion, "YYYY-MM-DD") & "') dFecha"
    oConect.AbreConexion
    Set rs = oConect.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        fecha = rs!dFecha
    End If
    oConect.CierraConexion
    Set rs = Nothing
    
    sCad = sCad & lsCiudad & "," & UCase(fecha) & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea & oImp.gPrnSaltoLinea
    
    sCad = sCad & Space(35) & String(Len(sNombre) + 5, "-") & oImp.gPrnSaltoLinea
    sCad = sCad & Space(35) & sNombre & oImp.gPrnSaltoLinea
    sCad = sCad & Space(35) & "DOI: " & sDoi & oImp.gPrnSaltoLinea
    
    GeneraSolicitudAfiliacionDesafiliacionEnvioEstadoCta = sCad
End Function
'END APRI
'INICIO EAAS20181122 SEGUN ACTA 90-2018 ERS-006-2018
Public Function ObtieneCCI(ByVal pcCtaCod As String) As ADODB.Recordset
Dim oCCI As New COMDCaptaGenerales.DCOMCaptaGenerales
    Set ObtieneCCI = oCCI.ObtieneCCI(pcCtaCod)
    Set oCCI = Nothing
End Function

'FIN EAAS EAAS20181122 SEGUN ACTA 90-2018 ERS-006-2018
'*************** ANDE 20171107 habilitacion de operaciones  -- iniciativa TI-Desarrollo ****************************
Public Sub RegistrarLogHabilitacion(ByVal pnId As Integer, ByVal pcUserH As String, ByVal pdFechaOpe As Date, ByVal pnCondicion As Integer)
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Call oCaptaAD.RegistrarLogHabilitacion(pnId, pcUserH, pdFechaOpe, pnCondicion)
End Sub
Public Function ObtenerOperacionesH() As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.ObtenerOperacionesHabilitar
    Set ObtenerOperacionesH = R
    Set R = Nothing
End Function

Public Function ObtenerUsuarios_RF_AS_SUP(ByVal pcUser As String) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.ObtenerUsuarios_RF_AS_SUP(pcUser)
    Set ObtenerUsuarios_RF_AS_SUP = R
    Set R = Nothing
End Function

Public Function ObtenerOperacionesHabilitadasxUsuario(ByVal pcUser As String, ByVal dFechaSist As Date) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.ObtenerOperacionesHabilitadasxUsuario(pcUser, dFechaSist)
    Set ObtenerOperacionesHabilitadasxUsuario = R
    Set R = Nothing
End Function

Public Function RegistrarHabilitacionDeOperacion(ByVal pcUserH As String, ByVal pcOpeCod As String, ByVal pcHabilitacionSist As String, Optional ByVal pcUserQH As String = "", Optional ByVal pnCantOpe As Integer = 0, Optional ByVal pnEstado As Integer = 0) As Integer
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    On Error GoTo RegistrarHabError
    Set R = oCaptaAD.RegistrarHabilitacionDeOperacion(pcUserH, pcOpeCod, pcHabilitacionSist, pcUserQH, pnCantOpe, pnEstado)
    RegistrarHabilitacionDeOperacion = 1
    If Not (R.BOF And R.EOF) Then
        RegistrarHabilitacionDeOperacion = R!Resultado
    End If
    Exit Function
RegistrarHabError:
    RegistrarHabilitacionDeOperacion = 1
End Function

Public Function RegistrarActividad(ByVal pcOpeCod As String, ByVal pcUser As String, ByVal pdFecha As Date) As Integer
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim nEstadoActividad As Integer
        nEstadoActividad = oCaptaAD.RegistrarActividad(pcOpeCod, pcUser, pdFecha)
'        If nEstadoActividad = 1 Then
'            MsgBox "He detectado un problema; su operación no fue afectada, pero por favor comunciar a TI-Desarrollo.", vbError, "Error"
'        ElseIf nEstadoActividad = 2 Then
'            MsgBox "Ha usado el total de operaciones permitidas para el día de hoy. Si desea realizar más operaciones, comuníquese con el área de Operaciones.", vbInformation + vbOKOnly, "Aviso"
'        End If
        RegistrarActividad = nEstadoActividad
End Function

Public Function VerificarOperaciones(ByVal pcUser As String, ByVal pdFecha As Date, ByVal pcOpeCod As String) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.VerificarOperaciones(pcUser, pdFecha, pcOpeCod)
    Set VerificarOperaciones = R
    Set R = Nothing
End Function

Public Function OperacionPermitida(ByVal pcUser As String, ByVal pdFecha As Date, ByVal pnOperacion As COMDConstantes.CaptacOperacion, ByRef cMsg As String) As Boolean
      If InStr(1, "122900,200221,200243,200244,200245,200266,200267,300401,300419,300501,300593", pnOperacion) Then
        Dim R As ADODB.Recordset, nOpeDisponible As Integer
        Set R = VerificarOperaciones(pcUser, pdFecha, CStr(pnOperacion))
        If Not (R.BOF And R.EOF) Then
            R.MoveFirst
            Do While Not R.EOF
                If pnOperacion = R!copecod Then
                    nOpeDisponible = R!nCantOpe - R!nCantUsadas
                    If nOpeDisponible <= 0 Then 'or =0
                        cMsg = "Se ha superado el total de operaciones permitidas para el dia de hoy, si desea realizar ésta operación, por favor comuníquese con el área de Operaciones."
                        OperacionPermitida = False
                    Else
                        OperacionPermitida = True
                    End If
                    Exit Function
                End If
                R.MoveNext
            Loop
        Else
            cMsg = "No cuentas con permiso para realizar ésta operación. Si desea realizar ésta operación, por favor comuníquese con el área de Operaciones."
            OperacionPermitida = False
        End If
    Else
        OperacionPermitida = True
    End If
End Function
'*******CTI3 TI-ERS082-2018 *****
Public Function GetTpoCta(ByVal sCuenta As String) As String
Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
Dim R As ADODB.Recordset
Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales

Set R = clsMant.ConsultaTpoCta(sCuenta)

If Not R Is Nothing Then
    If Not R.EOF Then
        If R.RecordCount > 0 Then
           GetTpoCta = R!codLeyenda
        End If
    End If
End If
Set clsMant = Nothing
End Function

'APRI20190601 RFC1902040001
Public Sub GuardarMotivoCancelacion(ByVal pcCtaCod As String, ByVal pnMotivo As Integer, ByVal pcGlosa As String)
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    oCaptaAD.GuardarMotivoCancelacion pcCtaCod, pnMotivo, pcGlosa
    Set oCaptaAD = Nothing
End Sub
'APRI20190109 ERS077-2018
Public Function ObtenerSaldoEquilibrio(ByVal nComision As Currency, ByVal nTEA As Currency) As Currency
    Dim nTEM As Double
    Dim nSaldoEqui As Double
    If nTEA > 0 Then
        nTEM = Format$(((1 + nTEA / 100) ^ (1 / 12) - 1), "#0.000000000000")
        nSaldoEqui = Format$((nComision / nTEM), "#0.00")
    End If
    ObtenerSaldoEquilibrio = nSaldoEqui
End Function
Public Function ObtenerCapValorComision(ByVal pcCtaCod As String, ByVal pnCantOpe As Integer, ByVal pnTpoOpe As Integer, ByVal pnTipoCanal As Integer) As Currency
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim nValor As Currency
    nValor = oCaptaAD.ObtenerCapValorComision(pcCtaCod, pnCantOpe, pnTpoOpe, pnTipoCanal)
    ObtenerCapValorComision = nValor
End Function
Public Function CargarCuentasMantCuentaPendiente(Optional ByVal psCodPers As String = "", Optional psCtaCod As String = "") As ADODB.Recordset
Dim oDCOMCaptaGenerales As New COMDCaptaGenerales.DCOMCaptaGenerales
Set CargarCuentasMantCuentaPendiente = oDCOMCaptaGenerales.CargarCuentasMantCuentaPendiente(psCodPers, psCtaCod)
Set oDCOMCaptaGenerales = Nothing
End Function
'END APRI


'JIPR20190916 INICIO
Public Function ObtieneDatosCuentas(ByVal pcPersCod As String) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.ObtieneDatosCuentas(pcPersCod)
    Set ObtieneDatosCuentas = R
    Set R = Nothing
End Function

Public Function DatosPersCuenta(ByVal pcPersCod As String) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.DatosPersCuenta(pcPersCod)
    Set DatosPersCuenta = R
    Set R = Nothing
End Function

Public Function RegistraTarjetaCuenta(ByVal lsNumTarj As String, ByVal lsCuenta As String, ByVal lsFecha As Date, ByVal lnPrior As Integer, ByVal lnRelac As Integer, ByVal lnConsulta As Integer, ByVal lnRetiro As Integer)
    Dim oCaptaAD As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCaptaAD = New COMDCaptaGenerales.DCOMCaptaGenerales
    Call oCaptaAD.RegistraTarjetaCuenta(lsNumTarj, lsCuenta, lsFecha, lnPrior, lnRelac, lnConsulta, lnRetiro)
  Set oCaptaAD = Nothing
End Function

Public Function RegistraHistorPrioridad(ByVal lsNumTarj As String, ByVal lsCuenta As String, ByVal lsFecha As Date, ByVal lnPrior As Integer)
    Dim oCaptaAD As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCaptaAD = New COMDCaptaGenerales.DCOMCaptaGenerales
    Call oCaptaAD.RegistraHistorPrioridad(lsNumTarj, lsCuenta, lsFecha, lnPrior)
  Set oCaptaAD = Nothing
End Function

Public Function ActualizaTarjetaCuenta(ByVal lsNumTarj As String, ByVal lsCuenta As String, ByVal lnRelac As Integer, ByVal lnPrior As Integer, ByVal lnConsulta As Integer, ByVal lnRetiro As Integer)
    Dim oCaptaAD As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCaptaAD = New COMDCaptaGenerales.DCOMCaptaGenerales
    Call oCaptaAD.ActualizaTarjetaCuenta(lsNumTarj, lsCuenta, lnRelac, lnPrior, lnConsulta, lnRetiro)
  Set oCaptaAD = Nothing
End Function

Public Function DesafiliarTarjetaCuenta(ByVal lsNumTarj As String, ByVal lsCuenta As String)
    Dim oCaptaAD As COMDCaptaGenerales.DCOMCaptaGenerales
    Set oCaptaAD = New COMDCaptaGenerales.DCOMCaptaGenerales
    Call oCaptaAD.DesafiliarTarjetaCuenta(lsNumTarj, lsCuenta)
  Set oCaptaAD = Nothing
End Function

Public Function DatosPDF(ByVal lcNumTarjeta As String) As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.DatosPDF(lcNumTarjeta)
    Set DatosPDF = R
    Set R = Nothing
End Function

Public Function DocumentoPDF() As ADODB.Recordset
    Dim oCaptaAD As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim R As ADODB.Recordset
    Set R = oCaptaAD.DocumentoPDF()
    Set DocumentoPDF = R
    Set R = Nothing
End Function
'JIPR20190916 FIN
'CTI4 ERS0112020
Public Function ObtieneDOI_Titular(ByVal psCuenta As String) As String
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim rsDoi As ADODB.Recordset
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsDoi = clsMant.ObtieneDOI_Titular(psCuenta)
    If Not (rsDoi.EOF And rsDoi.BOF) Then
        ObtieneDOI_Titular = rsDoi!cPersIDnro
        Exit Function
    End If
    ObtieneDOI_Titular = ""
End Function
Public Function ObtieneDatosTrBanco(ByVal nMovNro As Long) As ADODB.Recordset
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim rsDoi As ADODB.Recordset
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set ObtieneDatosTrBanco = clsMant.ObtieneDatosTrBanco(nMovNro)
End Function
'END CTI4
'APRI20210621 ERS031-2021
Public Function AplicaDineroNuevoCaptaciones(ByVal psPersCod As String, ByVal psMonto As Currency) As Boolean
    Dim oDCOMCap As New COMDCaptaGenerales.DCOMCaptaGenerales
    AplicaDineroNuevoCaptaciones = oDCOMCap.AplicaDineroNuevoCaptaciones(psPersCod, psMonto)
    Set oDCOMCap = Nothing
End Function
'END APRI


