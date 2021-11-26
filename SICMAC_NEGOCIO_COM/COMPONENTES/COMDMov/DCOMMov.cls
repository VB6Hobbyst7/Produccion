VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Base 0
Option Explicit
Dim vsServerAdmin As String
Dim vsServerComunes As String
Dim vsServerImagenes As Variant
Dim vsServerNegocio As String
Dim vsServerPersonas As String
Dim oConect As COMConecta.DCOMConecta
Dim oError As New COMConecta.COMErrorHandling
Dim oVarPublicas As New COMFunciones.FCOMVarPublicas
Public bTransaccion As Boolean 'RIRO20170614
Dim sSql As String
Dim rs As ADODB.Recordset

Public Sub RollbackTrans()
    oConect.RollbackTrans
End Sub

Public Sub CommitTrans()
    oConect.CommitTrans
End Sub
Public Sub BeginTrans()
    oConect.BeginTrans
End Sub
'RIRO20170322
Public Sub SetConexion(ByVal pCon As COMConecta.DCOMConecta)
    Set oConect = pCon
End Sub
'END RIRO
Public Function BuscarMov(ByVal psMovNro As String, Optional psFiltro As String = "") As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    
    On Error GoTo BuscarMovErr
    BuscarMov = False
    
    sql = "SELECT cMovnro from Mov WHere cMovNro LIKE '" & psMovNro & "%'" & IIf(psFiltro <> "", " AND ", "") & psFiltro
    Set rs = oConect.CargaRecordSet(sql)
    BuscarMov = Not rs.EOF
    rs.Close
    Set rs = Nothing
    Exit Function
BuscarMovErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:BuscarMov Method")
End Function

Public Function InsertaMovCont(ByVal pnMovNro As Long, ByVal pnMovMonto As Double, ByVal pnMovCorrela As Integer, ByVal psMovParalelo As String) As Integer
    
    On Error GoTo InsertaMovContErr
    
    Dim sql As String
    InsertaMovCont = 1
    
    sql = "INSERT INTO MovCont (nMovNro, nMovMonto, nMovCorrela, cMovParalelo ) " _
         & "VALUES (" & pnMovNro & "," & pnMovMonto & "," & pnMovCorrela & ",'" & psMovParalelo & "')"
     
    oConect.Ejecutar sql
    InsertaMovCont = 0
    Exit Function
InsertaMovContErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovCont Method")
End Function

Public Function InsertaMovPendientesRend(ByVal pnMovNro As Long, ByVal psCtaContCod As String, ByVal pnSaldo As Currency) As Integer
    
    On Error GoTo InsertaMovPendientesRendErr
    Dim sql As String
    InsertaMovPendientesRend = 1
    sql = "SELECT nMovNro FROM MovPendientesRend Where nMovNro = " & pnMovNro & " and cCtaContCod = '" & psCtaContCod & "' "
    Set rs = oConect.CargaRecordSet(sql)
    If rs.EOF Then
        sql = "INSERT INTO MovPendientesRend (nMovNro, nMovMonto, nMovCorrela, cMovParalelo ) " _
             & "VALUES (" & pnMovNro & ", '" & psCtaContCod & "'," & pnSaldo & ") "
    Else
        sql = "UPDATE MovPendientesRend SET nSaldo = " & pnSaldo & " WHERE nMovNro = " & pnMovNro & " "
    End If
    oConect.Ejecutar sql
    InsertaMovPendientesRend = 0
    Exit Function
InsertaMovPendientesRendErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovPendientesRend Method")
End Function


Public Function InsertaMovARendir(ByVal pnMovNro As Long, ByVal psTpoArendir As ArendirTipo, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal psPersCod As String, ByVal pnMovSaldo As Currency, ByVal psCategCod As String) As Integer
    
    
    On Error GoTo InsertaMovARendirErr
    Dim sql As String
    InsertaMovARendir = 1
    
    sql = "INSERT INTO MovARendir(nMovNro,cTpoArendir, cAgeCod, cAreaCod, cPersCod, nMovSaldo , cCategCod ) " _
     & "VALUES (" & pnMovNro & ", '" & psTpoArendir & "','" & psAgeCod & "','" & psAreaCod & "','" & psPersCod & "'," & pnMovSaldo & ",'" & psCategCod & "')"
     
    oConect.Ejecutar sql
    InsertaMovARendir = 0
    Exit Function
InsertaMovARendirErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovARendir Method")
End Function
Public Sub ActualizaMov(ByVal pnMovNro As Long, Optional ByVal psMovDesc As String = "", Optional ByVal pnMovEstado As MovEstado = -1, Optional ByVal psMovFlag As MovFlag = gMovFlagVigente)
    
    On Error GoTo ActualizaMovErr
    Dim sql As String
    Dim lsActualiza As String
    lsActualiza = ""
    If psMovDesc <> "" Then
        lsActualiza = lsActualiza & " cMovDesc='" & Replace(psMovDesc, "'", "''") & "',"
    End If
    If pnMovEstado > -1 Then
      lsActualiza = lsActualiza & " nMovEstado=" & pnMovEstado & ","
    End If
    lsActualiza = lsActualiza & " nMovFlag=" & psMovFlag & ","
    lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    If lsActualiza = "" Then Exit Sub
    
    sql = " Update Mov set " & lsActualiza & " Where nMovNro =" & pnMovNro & ""
    oConect.Ejecutar sql
    Exit Sub
ActualizaMovErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaMov Method")
End Sub
'MIOL 20120608, SEGUN RQ12093 AGREGAR PARAMETRO psMovDesc **************************************
Public Sub EliminaMov(ByVal psMovNro As String, Optional psMovNroActualiza As String, Optional psMovDesc As String)

Dim sql As String
   On Error GoTo EliminaMovErr
    sql = "UPDATE Mov SET nMovFlag = '" & gMovFlagEliminado & "' WHERE cMovNro = '" & psMovNro & "'"
    oConect.Ejecutar sql
    If Not psMovNroActualiza = "" Then
        InsertaMovModifica psMovNroActualiza, psMovNro, , psMovDesc
    End If
    ActualizaSaldoMovimiento psMovNro, "-"
Exit Sub
EliminaMovErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:EliminaMov Method")
End Sub

Public Sub InsertaMovModifica(ByVal psMovNro As String, ByVal psMovNroAnt As String, Optional psMovNroNew As String, Optional psMovDesc As String)
Dim sql As String


    On Error GoTo InsertaMovModificaErr
    sql = "INSERT MovModifica (cMovNro, cMovNroAnt,cMovNroNew,cMovGlosa) VALUES ('" & psMovNro & "','" & psMovNroAnt & "', '" & psMovNroNew & "', '" & psMovDesc & "')"
    oConect.Ejecutar sql
Exit Sub
InsertaMovModificaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovModifica Method")
End Sub
'END MIOL **************************************************************************************
'MADM 20120102 - ##ModelId=3A848F08038A
Public Function InsertaMovCta(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency, Optional pnTipoPago As Integer = 0) As Integer
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    InsertaMovCta = 1
    
    sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte,nTipoPago )" _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & "," & pnTipoPago & ") "
      
    oConect.Ejecutar sql
    InsertaMovCta = 0
    Exit Function
InsertaMovCtaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovCta Method")
End Function

Public Function InsertaMovCartaFianza(ByVal pnMovNro As Long, psPersCod As String, psIFPersCod As String, psCartaVenc As String) As Integer
    On Error GoTo InsertaMovCartaFianzaErr
    Dim sql As String
    InsertaMovCartaFianza = 1
    
    sql = " INSERT INTO MOVCartaFianza (nMovNro, cPersCod, cIFPersCod, dCartaVenc )" _
        & " VALUES(" & pnMovNro & ",'" & psPersCod & "','" & psIFPersCod & "','" & psCartaVenc & "') "
    oConect.Ejecutar sql
    
    InsertaMovCartaFianza = 0
    Exit Function
InsertaMovCartaFianzaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovCartaFianza Method")
End Function

Public Sub ActualizaMovCta(ByVal pnMovNro As Long, ByVal pnMovItem As Long, Optional psCtaCod As String = "", Optional pnMovImporte As Currency = 0)
    On Error GoTo ActualizaMovCtaErr
    Dim sql As String
    Dim lsActualiza As String
    lsActualiza = ""
    If psCtaCod <> "" Then
        lsActualiza = lsActualiza & " cCtaContCod ='" & psCtaCod & "',"
    End If
    If pnMovImporte <> 0 Then
       lsActualiza = lsActualiza & " nMovImporte = " & pnMovImporte & ","
    End If
    If lsActualiza <> "" Then
      lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    End If
    If lsActualiza = "" Then Exit Sub
    
    sql = " UPDATE MovCta SET " & lsActualiza & " WHERE nMovNro =" & pnMovNro & " and nMovItem = " & pnMovItem & ""
    oConect.Ejecutar sql
    Exit Sub
ActualizaMovCtaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaMovCta Method")
End Sub
Public Function InsertaMovMe(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovMeImporte As Currency) As Integer
    On Error GoTo InsertaMovMeErr
    Dim sql As String
    InsertaMovMe = 1
    
    sql = " INSERT INTO MOVME (nMovNro, nMovItem, nMovMEImporte )" _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ", " & pnMovMeImporte & ") "
      
    oConect.Ejecutar sql
    InsertaMovMe = 0
    Exit Function
InsertaMovMeErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovMe Method")
End Function

'##ModelId=3A848F3301A5
Public Function InsertaMovObj(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psObjetoCod As String) As Integer
    On Error GoTo InsertaMovObjErr
    Dim sql As String
    InsertaMovObj = 1
     
    sql = "INSERT INTO MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & Format(psObjetoCod, "00") & "')"

    oConect.Ejecutar sql
    InsertaMovObj = 0
    Exit Function
InsertaMovObjErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovObj Method")
End Function
Public Function InsertaMovObjViaticos(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psObjetoCod As String, ByVal pnImporte As Currency) As Integer
    On Error GoTo InsertaMovObjViaticosErr
    Dim sql As String
    InsertaMovObjViaticos = 1
     
    sql = "INSERT INTO MovObjViaticos (nMovNro, nMovItem, cObjetoCod, nMovImporte) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psObjetoCod & "'," & pnImporte & ")"

    oConect.Ejecutar sql
    InsertaMovObjViaticos = 0
    Exit Function
InsertaMovObjViaticosErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovObjViaticos Method")
End Function

Public Function InsertaMovObjIF(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As String, ByVal psPersCod As String, ByVal psTipoIF As String, ByVal psCtaIfCod As String) As Integer
    On Error GoTo InsertaMovObjIFErr
    Dim sql As String
    InsertaMovObjIF = 1
     
    sql = "INSERT INTO MOVOBJIF(nMovNro, nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod ) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psPersCod & "','" & Format(psTipoIF, "00") & "','" & psCtaIfCod & "')"
    
    oConect.Ejecutar sql
    InsertaMovObjIF = 0
    Exit Function
InsertaMovObjIFErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovObjIF Method")
End Function

'##ModelId=3A848F4F002E
Public Function InsertaMovDoc(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psDocFecha As String) As Integer
    On Error GoTo InsertaMovDocErr
    Dim sql As String
    InsertaMovDoc = 1
    
    sql = "INSERT INTO MOVDOC (nMovNro, nDocTpo, cDocNro, dDocFecha ) " _
        & "VALUES (" & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & psDocFecha & "')"
        
    oConect.Ejecutar sql
    InsertaMovDoc = 0
    Exit Function
InsertaMovDocErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovDoc Method")
End Function
'##ModelId=3A84A3FE038A
Public Function InsertaMovCajaChica(ByVal pnMovNro As Long, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnProcNro As Integer, ByVal pnProcTpo As CHTipoProc) As Integer
    On Error GoTo InsertaMovCajaChicaErr
    Dim sql As String
    InsertaMovCajaChica = 1
    sql = " INSERT INTO MOVCAJACHICA(nMovNro, cAreaCod, cAgeCod, nProcNro, cProcTpo ) " _
        & " VALUES(" & pnMovNro & ",'" & psAreaCod & "','" & psAgeCod & "'," & pnProcNro & ",'" & pnProcTpo & "')"
    
    oConect.Ejecutar sql
    
    InsertaMovCajaChica = 0
    Exit Function
InsertaMovCajaChicaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovCajaChica Method")
End Function
Public Function ActualizaMovCajaChicaNroProc(ByVal pnMovNro As Long, Optional ByVal pnProcNro As Integer = 0) As Integer
    On Error GoTo ActualizaMovCajaChicaNroProcErr
    Dim sql As String
    ActualizaMovCajaChicaNroProc = 1
    sql = " UPDATE MOVCAJACHICA SET nProcNro = " & pnProcNro & " WHERE nMovNro = " & pnMovNro & " "
    oConect.Ejecutar sql
    
    ActualizaMovCajaChicaNroProc = 0
    Exit Function
ActualizaMovCajaChicaNroProcErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaMovCajaChicaNroProc Method")
End Function
Public Function InsertaCajaChica(ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                ByVal pnProcNro As Integer, ByVal psPersCod As String, _
                                ByVal pnMontoAsig As Currency, ByVal pnTopeEgresos As Currency, _
                                ByVal pnMontoDesem As Currency, ByVal pnSaldoAnt As Currency, _
                                ByVal pnSaldo As Currency) As Integer
    On Error GoTo InsertaCajaChicaErr
    Dim sql As String
    InsertaCajaChica = 1
    sql = " INSERT INTO CAJACHICA(cAreaCod ,cAgeCod,nProcNro,cPersCod,nMontoAsig," _
        & "                         nTopeEgresos,nMontoDesem,nSaldoAnt,nSaldo) " _
        & " VALUES('" & psAreaCod & "','" & psAgeCod & "'," & pnProcNro & ",'" & psPersCod & "'," _
        & pnMontoAsig & "," & pnTopeEgresos & "," & pnMontoDesem & "," & pnSaldoAnt & "," _
        & pnSaldo & ")"
    
    oConect.Ejecutar sql
    InsertaCajaChica = 0
    Exit Function
InsertaCajaChicaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaCajaChica Method")
End Function

Public Function InsertaMovBS(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovBSOrden As Long, ByVal psBSCod As String) As Integer
    On Error GoTo InsertaMovBSErr
    Dim sql As String
    InsertaMovBS = 1
     
    sql = "INSERT INTO MOVBS (nMovNro, nMovItem, nMovBSOrden, cBSCod) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovBSOrden & ",'" & psBSCod & "')"
    oConect.Ejecutar sql
    InsertaMovBS = 0
    Exit Function
InsertaMovBSErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovBS Method")
End Function

Private Sub Class_Initialize()
Set oConect = New COMConecta.DCOMConecta
If oConect.AbreConexion = False Then
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos")
End If

End Sub
'RIRO20170614 añadio filtro btransaccion
Private Sub Class_Terminate()
    If bTransaccion = True Then
    Else
        oConect.CierraConexion
        Set oConect = Nothing
    End If
End Sub
'MADM 20120201 nTipoPago
Public Function GeneraMovME(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional nTipoPago As Integer = 0) As Boolean
Dim nTCFijo As Currency
Dim nTCVent As Currency
Dim nTCComp As Currency
Dim sSql As String
Dim rs   As New ADODB.Recordset
Dim dFecha As Date
Dim nImporteD As Currency
Dim nImporteS As Currency
Dim oGen As New COMDConstSistema.DCOMGeneral

   On Error GoTo GeneraMovMEErr
   GeneraMovME = False
   dFecha = CDate(oGen.GetFechaMov(psMovNro, True))
   
   Dim oTC As New COMDConstSistema.NCOMTipoCambio
   nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
   If nTCMerc = 0 Then
        nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
        nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)
   Else
        nTCVent = nTCMerc
        nTCComp = nTCMerc
   End If

   Set oTC = Nothing
   oConect.Ejecutar " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent & "," & nTipoPago
   
   GeneraMovME = True
Exit Function
GeneraMovMEErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:GeneraMovME Method")
End Function

'Private Sub GetClaseCtaCont(ByRef pnBalaMoneda As Boolean, ByRef psCtaClase, ByVal psCtaCod As String)
'Dim oCta As New COMDContabilidad.DCOMCtaCont
'Dim prs  As ADODB.Recordset
'Set prs = oCta.CargaCtaContClase(psCtaCod)
'If Not prs.EOF Then
'   pnBalaMoneda = prs!nCtaMoneda
'   psCtaClase = prs!cCtaCaracter
'Else
'   pnBalaMoneda = 0
'   psCtaClase = ""
'End If
''RSClose prs
'prs.Close
'Set oCta = Nothing
'End Sub

Public Function InsertaMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long) As Integer
    On Error GoTo InsertaMovRefErr
    Dim sql As String
    InsertaMovRef = 1
    
    sql = "INSERT INTO MovRef(nMovNro, nMovNroRef) " _
         & "VALUES (" & pnMovNro & "," & pnMovNroRef & ")"
     
    oConect.Ejecutar sql
    InsertaMovRef = 0
    Exit Function
InsertaMovRefErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovContErr Method")
End Function

Public Function InsertaMovCajaChicaRend(ByVal pnMovNroRend As Long, ByVal pnMovNroAnt As Long, ByVal pnMovNroNew As Long) As Integer
    On Error GoTo InsertaMovCajaChicaRendErr
    Dim sql As String
    InsertaMovCajaChicaRend = 1
    
    sql = "INSERT INTO MovCajaChicaRend(nMovNroRend, nMovNroAnt, nMovNroNew) " _
         & "VALUES (" & pnMovNroRend & "," & pnMovNroAnt & "," & pnMovNroNew & ")"
     
    oConect.Ejecutar sql
    InsertaMovCajaChicaRend = 0
    Exit Function
InsertaMovCajaChicaRendErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovCajaChicaRend Method")
End Function

Public Function ActualizaMovArendir(ByVal pnMovNro As Long, Optional ByVal psTpoArendir As ArendirTipo = gArendirTipoCajaGeneral, Optional ByVal psAreaCod As String, Optional ByVal psAgeCod As String, Optional ByVal psPersCod As String, Optional ByVal pnMovSaldo As Currency = -9999) As Integer
    On Error GoTo ActualizaMovARendirErr
    Dim sql As String
    Dim lsActualiza As String
    lsActualiza = ""
    ActualizaMovArendir = 1
    If psAgeCod <> "" Then
        lsActualiza = lsActualiza & " cAgeCod='" & psAgeCod & "',"
    End If
    If psAreaCod <> "" Then
        lsActualiza = lsActualiza & " cAreaCod='" & psAreaCod & "',"
    End If
    If psPersCod <> "" Then
        lsActualiza = lsActualiza & " cPersCod='" & psPersCod & "',"
    End If
    If pnMovSaldo <> -9999 Then
        lsActualiza = lsActualiza & " nMovSaldo=" & pnMovSaldo & ","
    End If
    'nMovSaldo
    lsActualiza = lsActualiza & " cTpoArendir='" & psTpoArendir & "',"
    lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    
    If lsActualiza = "" Then Exit Function
    sql = " Update MovARendir set " & lsActualiza & " Where nMovNro =" & pnMovNro & ""
    
    oConect.Ejecutar sql
    ActualizaMovArendir = 0
    Exit Function
ActualizaMovARendirErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaMovARendir Method")
End Function
Public Function InsertaMovObjPers(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psPersCod As String) As Integer
    On Error GoTo InsertaMovObjPersErr
    Dim sql As String
    InsertaMovObjPers = 1
    
    sql = "INSERT INTO MovObjPers(nMovNro, nMovItem, nMovObjOrden, cPersCod ) " _
         & "VALUES (" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psPersCod & "')"
     
    oConect.Ejecutar sql
    InsertaMovObjPers = 0
    Exit Function
InsertaMovObjPersErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovObjPers Method")
End Function
Public Function InsertaMovObjAgenciaArea(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psAgeCod As String, ByVal psAreaCod As String) As Integer
    On Error GoTo InsertaMovObjAgenciaAreaErr
    Dim sql As String
    InsertaMovObjAgenciaArea = 1
    
    sql = "INSERT INTO MovObjAreaAgencia(nMovNro, nMovItem, nMovObjOrden, cAreaCod, cAgeCod) " _
         & "VALUES (" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psAreaCod & "','" & psAgeCod & "')"
     
    oConect.Ejecutar sql
    InsertaMovObjAgenciaArea = 0
    Exit Function
InsertaMovObjAgenciaAreaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovObjAgenciaArea Method")
End Function
Public Function InsertaMovObjEfectivo(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psEfectivoCod As String) As Integer
    On Error GoTo InsertaMovObjEfectivoErr
    Dim sql As String
    InsertaMovObjEfectivo = 1
     
    sql = "INSERT INTO MovObjEfectivo(nMovNro, nMovItem, nMovObjOrden, cEfectivoCod ) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psEfectivoCod & "')"

    oConect.Ejecutar sql
    InsertaMovObjEfectivo = 0
    Exit Function
InsertaMovObjEfectivoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovObjEfectivo Method")
End Function
Public Function InsertProv(ByVal psPersCod As String, psProvEstado As String, ByVal psUltimaActualizacion As String) As Integer
    On Error GoTo InsertProvErr
    Dim sql As String
    InsertProv = 1
     
    sql = "INSERT INTO Proveedor( cPersCod, nProvEstado, cUltimaActualizacion  ) " _
        & " VALUES('" & psPersCod & "','" & psProvEstado & "','" & psUltimaActualizacion & "')"

    oConect.Ejecutar sql
    InsertProv = 0
    Exit Function
InsertProvErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertProv Method")
End Function

Public Function InsertaMovGasto(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal psPersCodDest As String) As Integer
    On Error GoTo InsertaMovGastoErr
    Dim sql As String
    InsertaMovGasto = 1
    
    sql = "INSERT INTO MovGasto (nMovNro, cPersCod, cPersCodDest) " _
        & "VALUES (" & pnMovNro & ", '" & psPersCod & "','" & psPersCodDest & "')"
    oConect.Ejecutar sql
    InsertaMovGasto = 0
    Exit Function
InsertaMovGastoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovGasto Method")
End Function

Public Function InsertaMovOtrosItem(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psMovOtrasVar As String, ByVal pnMovOtroImporte As Currency, ByVal psMovOtroTexto As String) As Integer
    On Error GoTo InsertaMovOtrosItemErr
    Dim sql As String
    InsertaMovOtrosItem = 1
    
    sql = "INSERT INTO MovOtrosItem (nMovNro,nMovItem,cMovOtroVariable,nMovOtroImporte,cMovOtroTexto) " _
     & "VALUES (" & pnMovNro & "," & pnMovItem & ",'" & psMovOtrasVar & "'," & pnMovOtroImporte & ",'" & psMovOtroTexto & "')"
     
    oConect.Ejecutar sql
    InsertaMovOtrosItem = 0
    Exit Function
InsertaMovOtrosItemErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovOtrosItem Method")
End Function
Public Function InsertaMovArendirSust(ByVal pnMovNroSol As Long, ByVal pnMovNroSust As Long, Optional ByVal pnMovNroRend As Long = -1) As Integer
    On Error GoTo InsertaMovArendirSustErr
    Dim sql As String
    InsertaMovArendirSust = 1
    
    sql = "INSERT INTO MovArendirSust (nMovNro, nMovNroSust, nMovNroRend) " _
     & "VALUES (" & pnMovNroSol & "," & pnMovNroSust & "," & IIf(pnMovNroRend = -1, "Null", pnMovNroRend) & ")"
     
    oConect.Ejecutar sql
    InsertaMovArendirSust = 0
    Exit Function
InsertaMovArendirSustErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovArendirSust Method")
End Function
Public Function InsertaCheque(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psTipoIF As String, _
            ByVal pnPlaza As ChequePlaza, ByVal psIFCta As String, ByVal pnMonto As Currency, _
            ByVal pdValorizaRef As Date, ByVal pdValorizacion As Date, ByVal psDepIF As CGEstadosChq, _
            ByVal pnConfCaja As CGEstadoConfCheque, ByVal pnMonedaCheque As Moneda, _
            Optional ByVal pnEstado As ChequeEstado = gChqEstEnValorizacion, Optional psAreaCod As String, _
            Optional psAgeCod As String, Optional pnProducto As Producto = 0, Optional ByVal psPersCodIF1 As String = "", Optional ByVal psNumChequeComp As String = "", Optional ByVal pnMovNro As Long = 0, Optional ByVal pbGerencia As Boolean = False)
'MADM 20110228 - se agrego psPersCodIF1 = girador - cambio store
    On Error GoTo InsertaChequeErr
    Dim sql As String
    InsertaCheque = 1
    
'''    sql = " INSERT INTO DOCREC(nTpoDoc, cNroDoc, cPersCod , cIFTpo, bPlaza, cIFCta, nMonto, dValorizaRef, dValorizacion, cDepIF, nConfCaja ,  nMoneda , nEstado, cAreaCod, cAgeCod, nProducto ,cPersCodGirador) " _
'''     & "    VALUES (" & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "'," & pnPlaza & ",'" & psIFCta & "'," & pnMonto & ",'" _
'''     & Format(pdValorizaRef, oVarPublicas.gsFormatoFecha) & "','" & Format(pdValorizacion, oVarPublicas.gsFormatoFecha) & "','" & psDepIF & "'," & pnConfCaja & "," & pnMonedaCheque & "," & pnEstado & ",'" & psAreaCod & "','" & psAgeCod & "'," & pnProducto & ", '" & psPersCodIF1 & "')"

    'sql = "exec stp_ins_Cheque " & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "'," & pnPlaza & ",'" & psIFCta & "'," & pnMonto & ",'" & Format(pdValorizaRef, oVarPublicas.gsFormatoFecha) & "','" & Format(pdValorizacion, oVarPublicas.gsFormatoFecha) & "','" & psDepIF & "'," & pnConfCaja & "," & pnMonedaCheque & "," & pnEstado & ",'" & psAreaCod & "','" & psAgeCod & "'," & pnProducto & ", '" & psPersCodIF1 & "','" & psNumChequeComp & "'"
    sql = "exec stp_ins_Cheque " & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "'," & pnPlaza & ",'" & psIFCta & "'," & pnMonto & ",'" & Format(pdValorizaRef, oVarPublicas.gsFormatoFecha) & "','" & Format(pdValorizacion, oVarPublicas.gsFormatoFecha) & "','" & psDepIF & "'," & pnConfCaja & "," & pnMonedaCheque & "," & pnEstado & ",'" & psAreaCod & "','" & psAgeCod & "'," & pnProducto & ", '" & psPersCodIF1 & "','" & psNumChequeComp & "'," & pnMovNro & "," & IIf(pbGerencia, 1, 0) 'EJVG20140408
    oConect.Ejecutar sql
    
    InsertaCheque = 0
    Exit Function
InsertaChequeErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaCheque Method")
End Function

Public Function InsertaRegDocCuenta(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal cCtaCod As String) As Integer
    On Error GoTo InsertaRegDocCuentaErr
    Dim sql As String
    InsertaRegDocCuenta = 1
    
    sql = " INSERT INTO RegDocCuenta(nDocTpo, cDocNro, cCtaCod)" _
        & " VALUES(" & pnTpoDoc & ",'" & psNroDoc & "','" & cCtaCod & "') "
    
    oConect.Ejecutar sql
    InsertaRegDocCuenta = 0
    Exit Function
InsertaRegDocCuentaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaNotaAbonoCargoEst Method")
End Function

Public Function InsertaDocRecCapta(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psTipoIF As String, _
                              ByVal psCodCta As String, ByVal pnMonto As Currency)
    On Error GoTo InsertaChequeErr
    Dim sql As String
    InsertaDocRecCapta = 1
    
    sql = "INSERT INTO DocRecCapta (nTpoDoc, cNroDoc, cPersCod, cIFTpo, cCtaCod, nMonto) VALUES " _
        & "(" & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "', '" & psCodCta & "', " & pnMonto & ") "
    oConect.Ejecutar sql
    
    InsertaDocRecCapta = 0
    Exit Function
InsertaChequeErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaDocRecCapta Method")
End Function

Public Function ActualizaCheque(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psTipoIF As String, _
                                Optional ByVal pdValorizaRef As String = "", Optional ByVal pdValorizacion As String = "", _
                                Optional ByVal pnPlaza As ChequePlaza = -1, Optional ByVal psIFCta As String = "", _
                                Optional ByVal pnMonto As Currency = -999999, _
                                Optional ByVal pnDepIF As CGEstadosChq = -1, _
                                Optional ByVal pnConfCaja As CGEstadoConfCheque = -1, _
                                Optional ByVal pnMonedaCheque As Moneda = -1, _
                                Optional ByVal pnEstado As ChequeEstado = -1)
    On Error GoTo ActualizaChequeErr
    Dim sql As String
    Dim lsFiltro As String
    
    ActualizaCheque = 1
    lsFiltro = ""
    If pdValorizaRef <> "" Then
        lsFiltro = lsFiltro + " dValorizaRef  ='" & Format(pdValorizaRef, oVarPublicas.gsFormatoFecha) & "', "
    End If
    If pdValorizacion <> "" Then
        lsFiltro = lsFiltro + " dValorizacion ='" & Format(pdValorizacion, oVarPublicas.gsFormatoFecha) & "', "
    End If
    If pnPlaza <> -1 Then
        lsFiltro = lsFiltro + " bPlaza = " & pnPlaza & ","
    End If
    If psIFCta <> "" Then
        lsFiltro = lsFiltro + " cIFCta = '" & psIFCta & "',"
    End If
    If pnMonto <> -999999 Then
        lsFiltro = lsFiltro + " nMonto =" & pnMonto & ","
    End If
    If pnDepIF <> -1 Then
        lsFiltro = lsFiltro + " cDepIF='" & pnDepIF & "',"
    End If
    If pnConfCaja <> -1 Then
        lsFiltro = lsFiltro + " nConfCaja=" & pnConfCaja & ","
    End If
    If pnMonedaCheque <> -1 Then
        lsFiltro = lsFiltro + " nMoneda=" & pnMonedaCheque & ","
    End If
    If pnEstado <> -1 Then
        lsFiltro = lsFiltro + " nEstado=" & pnEstado & ","
    End If
    lsFiltro = Mid(lsFiltro, 1, Len(lsFiltro) - 1)
        
    sql = " UPDATE  DOCREC SET " & lsFiltro & " " _
     & "    WHERE   nTpoDoc =" & pnTpoDoc & " AND cNroDoc='" & psNroDoc & "' AND " _
     & "            cPersCod='" & psPersCod & "' AND cIFTpo ='" & psTipoIF & "'"
     
    oConect.Ejecutar sql
    ActualizaCheque = 0
    Exit Function
ActualizaChequeErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaCheque Method")
End Function

Public Function InsertaChequeEstado(ByVal pstpodoc As TpoDoc, ByVal psNroDoc As String, psPersCod As String, ByVal psTipoIF As String, ByVal pdFechaEstado As Date, ByVal psEstado As ChequeEstado, ByVal psMovNro As String, Optional ByVal psCtaCheque As String) As Integer
    On Error GoTo InsertaChequeEstadoErr
    Dim sql As String
    InsertaChequeEstado = 1
    
    sql = " INSERT INTO DocRecEst(nTpoDoc, cNroDoc, cPersCod, cIFTpo , nEstado , cMovNro,CIFCTA ) " _
     & "    VALUES (" & pstpodoc & ", '" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "'," _
     & psEstado & ",'" & psMovNro & "','" & psCtaCheque & "')"
     
    oConect.Ejecutar sql
    InsertaChequeEstado = 0
    Exit Function
InsertaChequeEstadoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaChequeEstado Method")
End Function
Public Function ActualizaMovimiento(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long, Optional pnMovEstado As MovEstado = gMovEstContabMovContable, _
                                Optional plActualizaAsiento As Boolean = True, Optional pbInsertaMov As Boolean = True, Optional pbDeleteMovAnt As Boolean = True, _
                                Optional ByVal psOpeCod As String = "", Optional ByVal psMovDesc As String = "", Optional ByVal lbUpdaMovRef As Boolean = True) As Integer
On Error GoTo ActualizaMovimientoErr
   Dim sql As String
   Dim lnMovNronew As Long
   Dim lsMovNroAnt As String
   ActualizaMovimiento = 1
   
   lsMovNroAnt = GetcMovNro(pnMovNroAnt)
   If lsMovNroAnt = psMovNroNew Then
   End If
   If plActualizaAsiento Then
      If pbInsertaMov Then
        sql = " INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag )" _
              & " SELECT '" & psMovNroNew & "'," & IIf(psOpeCod = "", "cOpeCod", "'" & psOpeCod & "'") & "," _
            & IIf(psMovDesc = "", "cMovDesc", "'" & psMovDesc & "'") & ",'" & pnMovEstado & "', nMovFlag  " _
            & " FROM Mov WHERE nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        lnMovNronew = GetnMovNro(psMovNroNew)
      Else
        lnMovNronew = GetnMovNro(psMovNroNew)
      End If
      sql = " UPDATE MOVCONT SET nMovNro=" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
      oConect.Ejecutar sql
   
      If pbDeleteMovAnt Then
          sql = "UPDATE MOV set nMovFlag=" & gMovFlagEliminado & " where nMovNro =" & pnMovNroAnt & " "
          oConect.Ejecutar sql
          ActualizaSaldoMovimiento lsMovNroAnt, "-"
      End If
   Else
      lnMovNronew = GetnMovNro(psMovNroNew)
   End If
   
   'Cambio de Mov de Arqueos
   sql = "INSERT INTO MOVARQUEO(nMovNro, cAreaCod, cPersCod, cHoraIni, cHoraFin, nEfectivo, nOPImporte, nTotDoc, nTotDocRend, nArendirPend, nDiferencia )      " _
    & " SELECT " & lnMovNronew & ",cAreaCod, cPersCod, cHoraIni, cHoraFin, nEfectivo, nOPImporte, nTotDoc, nTotDocRend, nArendirPend, nDiferencia FROM MOVARQUEO Where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   
   sql = "UPDATE MOVARQUEOOP SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
'   oConect.Ejecutar sql
   
   sql = "UPDATE MOVARQUEOEFECTIVO      SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   
   sql = "DELETE MOVARQUEO Where nMovNro =" & pnMovNroAnt & "; "
   oConect.Ejecutar sql
   ' Fin de Cambio de Mov de Arqueos
   
   If lbUpdaMovRef Then
        sql = "UPDATE MOVREF         SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
'        oConect.Ejecutar sql
        sql = sql & "UPDATE MOVREF         SET nMovNroRef =" & lnMovNronew & " Where nMovNroRef =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
   End If
   
   sql = "UPDATE MOVARENDIR     SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'oConect.Ejecutar sql
   sql = sql & "UPDATE MOVARENDIRSUST SET nMovNroRend=" & lnMovNronew & " where nMovNroSust =" & pnMovNroAnt & "; "
   'oConect.Ejecutar sql
   sql = sql & "UPDATE MOVOTROSITEM   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'oConect.Ejecutar sql
   sql = sql & "UPDATE MOVCAJACHICA   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'oConect.Ejecutar sql
   sql = sql & "UPDATE MOVCARTAFIANZA SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'oConect.Ejecutar sql
   sql = sql & "UPDATE MOVTPOCAMBIO   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & "; "
   'oConect.Ejecutar sql
   sql = sql & "UPDATE MOVVIATICOS    SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
    
   If plActualizaAsiento Then
        sql = " INSERT MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, cCtaContCod, nMovImporte FROM MOVCTA Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cObjetoCod FROM MOVOBJ Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "UPDATE MOVOBJIF          SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
        'oConect.Ejecutar sql
        sql = sql & "UPDATE MOVOBJEfectivo    SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
        'oConect.Ejecutar sql
        
        sql = sql & "UPDATE MOVOBJAreaAgencia SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & "; "
        'oConect.Ejecutar sql
        
        sql = sql & "UPDATE MOVME             SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql

        sql = "DELETE MOVOBJ            Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "DELETE MOVCTA            Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "UPDATE MOVDOC          SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = " UPDATE MOVGASTO       SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        ActualizaSaldoMovimiento psMovNroNew, "+"
   End If
   ActualizaMovimiento = 0
   Exit Function
ActualizaMovimientoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaMovimiento Method")
End Function
'**************************
'** ULTIMAS FUNCIONES
'**************************

Public Function CargaMovDocAsiento(ByVal pnMovNro As Long) As ADODB.Recordset
Dim psSql As String
   On Error GoTo CargaMovDocAsientoErr
   psSql = "SELECT md.nDocTpo, d.cDocDesc, cDocNro, Convert(varchar(10),md.dDocFecha,103) dDocFecha, m.nMovNro, m.cMovNro " _
         & "FROM Mov m JOIN MovDoc md ON md.nMovNro = m.nMovNro " _
         & "           JOIN Documento d ON d.nDocTpo = md.nDocTpo " _
         & "WHERE m.nMovNro = " & pnMovNro
   Set CargaMovDocAsiento = oConect.CargaRecordSet(psSql)
   Exit Function
CargaMovDocAsientoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DMov:CargaMovDocAsiento Method")
End Function

Public Function CargaMovOpeAsiento(ByVal pnMovNro As Long, Optional ByVal psMovNro As String = "") As ADODB.Recordset
Dim psSql As String
   On Error GoTo CargaOpeMovAsientoErr
      psSql = "SELECT m.cMovNro, m.cOpeCod, o.cOpeDesc, m.cMovDesc, m.nMovEstado, m.nMovFlag, ISNULL(mg.cPersCod,'') cPersCod, ISNULL(p.cPersNombre,'') cPersNombre, ISNULL(pid.cPersIDnro, '00000000') cRuc  " _
            & "FROM Mov M JOIN " & vsServerComunes & "OpeTpo O ON O.cOpeCod = M.cOpeCod " _
            & "      LEFT JOIN MovGasto mg ON mg.nMovNro = m.nMovNro " _
            & "      LEFT JOIN " & vsServerPersonas & "Persona p ON p.cPersCod = mg.cPersCod LEFT JOIN  PersID pid ON pid.cPersCod = p.cPersCod and pid.cPersIDTpo = " & gPersIdRUC _
            & "WHERE " & IIf(psMovNro <> "", " M.cMovNro = '" & psMovNro & "' ", " M.nMovNro = '" & pnMovNro & "'")
      Set CargaMovOpeAsiento = oConect.CargaRecordSet(psSql)
   Exit Function
CargaOpeMovAsientoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DMov:CargaMovOpeAsiento Method")
End Function

Public Function CargaMovCtaAsiento(ByVal pnMovNro As Long) As ADODB.Recordset
Dim psSql As String
   On Error GoTo CargaMovCtaAsientoErr
      psSql = "SELECT cMovNro, a.nMovItem, a.cCtaContCod, isnull(b.cCtaContDesc,'') AS cCtaContDesc, " _
         & "       CONVERT(varchar(20),ISNULL(SUM(CASE WHEN a.nMovImporte > 0 THEN a.nMovImporte END),0),105) as nDebe, " _
         & "       CONVERT(varchar(20),ISNULL(SUM(CASE WHEN a.nMovImporte < 0 THEN a.nMovImporte * -1 END),0),105) as nHaber, " _
         & "       CONVERT(varchar(20),ISNULL(SUM(CASE WHEN me.nMovMeImporte > 0 THEN me.nMovMeImporte END),0),105) as nDebeME, " _
         & "       CONVERT(varchar(20),ISNULL(SUM(CASE WHEN me.nMovMeImporte < 0 THEN me.nMovMeImporte * -1 END),0),105) nHaberME " _
         & " FROM  MovCta a      JOIN Mov M ON M.nMovNro = a.nMovNro " _
         & "                LEFT JOIN MovMe Me ON Me.nMovNro = a.nMovNro and Me.nMovItem = a.nMovItem " _
         & "                LEFT JOIN " & vsServerComunes & "CtaCont b ON a.cCtaContCod = b.cCtaContCod " _
         & "WHERE  M.nMovNro = '" & pnMovNro & "'" _
         & "GROUP BY cMovNro, a.nMovItem, a.cCtaContCod, b.cCtaContDesc ORDER BY a.nMovItem "
      Set CargaMovCtaAsiento = oConect.CargaRecordSet(psSql)
   Exit Function
CargaMovCtaAsientoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DMov:CargaMovCtaAsiento Method")
End Function

Public Function CargaMovObjAsiento(ByVal pnMovNro As Long) As ADODB.Recordset
Dim psSql As String
Dim lnMovNro As Long
Dim prs      As ADODB.Recordset
Dim prsObj   As ADODB.Recordset
Dim prsRes   As ADODB.Recordset

Dim oGen As New COMFunciones.FCOMObjetos

   On Error GoTo CargaMovObjAsientoErr
    psSql = "SELECT cMovNro, m.nMovNro, mo.nMovItem, mo.nMovObjOrden, mo.cObjetoCod " _
          & "FROM Mov m JOIN MovObj mo ON m.nMovNro = mo.nMovNro " _
          & "WHERE M.nMovNro = '" & pnMovNro & "' "
    Set prs = oConect.CargaRecordSet(psSql)
    Do While Not prs.EOF
        Select Case prs!cObjetoCod
           Case ObjCMACAgencias, ObjCMACAgenciaArea, ObjCMACArea
              psSql = "SELECT mo.nMovNro, mo.nMovItem, nMovObjOrden, " & prs!cObjetoCod & " ObjPadre, mo.cAreacod+mo.cAgeCod cObjetoCod, ISNULL(ag.cAgeDescripcion,'')+' '+ISNULL(a.cAreaDescripcion,'') cObjetoDesc " _
                    & "FROM MovObjAreaAgencia mo JOIN AreaAgencia aa ON aa.cAreaCod = mo.cAreaCod and aa.cAgeCod = mo.cAgeCod " _
                    & "     LEFT JOIN Areas a ON a.cAreaCod = aa.cAreaCod " _
                    & "     LEFT JOIN Agencias ag ON ag.cAgeCod = aa.cAgeCod " _
                    & "WHERE mo.nMovNro = '" & prs!nMovNro & "' and mo.nMovItem = '" & prs!nMovItem & "' and mo.nMovObjOrden = " & prs!nMovObjOrden & ""
                    
           Case ObjEntidadesFinancieras
              psSql = "SELECT mif.nMovNro, mif.nMovItem, mif.nMovObjOrden, " & prs!cObjetoCod & " ObjPadre, cif.cPersCod + cif.cCtaIfCod cObjetoCod, RTRIM(p.cPersNombre) + ' ' + cCtaIFDesc cObjetoDesc " _
                    & "FROM MovObjIF mif " _
                    & "     JOIN Persona p ON p.cPersCod = mif.cPersCod " _
                    & "     JOIN CtaIF cif ON cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod " _
                    & "WHERE mif.nMovNro = '" & prs!nMovNro & "' and mif.nMovItem = " & prs!nMovItem & " and mif.nMovObjOrden = " & prs!nMovObjOrden
           
           Case ObjDescomEfectivo
              psSql = "SELECT mo.nMovNro, mo.nMovItem, nMovObjOrden, " & prs!cObjetoCod & " ObjPadre, aa.cEfectivoCod cObjetoCod, aa.nEfectivoValor cObjetoDesc " _
                   & "FROM MovObjEfectivo mo JOIN Efectivo aa ON mo.cEfectivoCod = aa.cEfectivoCod " _
                   & "    LEFT JOIN Areas a ON a.cAreaCod = aa.cAreaCod LEFT JOIN Agencias ag ON ag.cAgeCod = aa.cAgeCod " _
                   & "WHERE  mo.nMovNro = " & prs!nMovNro & " and mo.nMovItem = " & prs!nMovItem & " and mo.nMovObjOrden = " & prs!nMovObjOrden

           Case ObjBienesServicios
              psSql = "SELECT bs.nMovNro, bs.nMovItem, bs.nMovBsOrden nMovObjOrden, " & prs!cObjetoCod & " ObjPadre, bs.cBSCod cObjetoCod, b.cBSDescripcion cObjetoDesc " _
                   & "FROM MovBS bs JOIN BienesServicios b ON b.cBSCod = bs.cBsCod " _
                   & "WHERE  bs.nMovNro = " & prs!nMovNro & " and bs.nMovItem = " & prs!nMovItem
           Case Else
              psSql = "SELECT mo.nMovNro, mo.nMovItem, nMovObjOrden, '" & prs!cObjetoCod & "' ObjPadre, mo.cObjetoCod, o.cObjetoDesc " _
                   & "FROM MovObj mo JOIN  Objeto o ON o.cObjetoCod = mo.cObjetoCod " _
                   & "WHERE  mo.nMovNro = " & prs!nMovNro & " and mo.nMovItem = " & prs!nMovItem & " and not mo.cObjetoCod IN ('" & objPersona & "', '" & ObjEntidadesFinancieras & "', '" & ObjCMACAgenciaArea & "', '" & ObjCMACArea & "','" & ObjCMACAgencias & "','" & ObjBienesServicios & "') "
        End Select
        Set prsObj = oConect.CargaRecordSet(psSql)
        oGen.RecordSetAdiciona prsRes, prsObj
        prs.MoveNext
    Loop
    Set CargaMovObjAsiento = prsRes
Exit Function

CargaMovObjAsientoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DMov:CargaMovObjAsiento Method")
End Function

Public Function CargaSumaMovAsiento(psMovNro As String) As ADODB.Recordset
Dim psSql As String
   On Error GoTo CargaSumaMovAsientoErr
      psSql = "SELECT SUM(CASE WHEN mc.nMovImporte > 0 THEN mc.nMovImporte END) as nDebe, " _
            & "       SUM(CASE WHEN mc.nMovImporte < 0 THEN mc.nMovImporte * -1 END) as nHaber, " _
            & "       ISNULL(SUM(CASE WHEN me.nMovMeImporte > 0 THEN me.nMovMeImporte END),0) as nDebeME, " _
            & "       ISNULL(SUM(CASE WHEN me.nMovMeImporte < 0 THEN me.nMovMeImporte * -1 END),0) as nHaberME " _
            & "FROM   Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
            & "WHERE  m.cMovNro = '" & psMovNro & "'"
      Set CargaSumaMovAsiento = oConect.CargaRecordSet(psSql)
   Exit Function
CargaSumaMovAsientoErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DMov:CargaSumaMovAsiento Method")
End Function
Public Function ActualizaArendirSust(ByVal pnMovNroSol As Long) As Integer
    On Error GoTo ActualizaArendirSustErr
    Dim sql As String
    ActualizaArendirSust = 1
    
    sql = "UPDATE MovArendirSust SET nMovNroRend=NULL WHERE nMovNro =" & pnMovNroSol & ""
    oConect.Ejecutar sql
    ActualizaArendirSust = 0
    Exit Function
ActualizaArendirSustErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaArendirSust Method")
End Function
Public Function ExtornaMovimiento(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long, _
            Optional ByVal psOpeCod As String = "", Optional ByVal psMovDesc As String = "", _
            Optional ByVal pbEliminaMov As Boolean = False, Optional psMovActualiza As String = "", _
            Optional ByVal psMotExtorno As Variant, Optional psMovNroCap As String = "", _
            Optional ByRef psOpeCodExtornoCap As String = "", Optional ByRef psCtaAhoExt As String = "", _
            Optional ByRef pnMontoAhorroExt As Double = 0#, Optional ByRef psOperacionDescExt As String, _
            Optional ByRef pnITFAhoExt As Double, Optional ByRef psMovNroAExt As String = "", Optional ByRef psClienteExt As String = "") As Integer

On Error GoTo ExtornaMovimientoErr
    Dim sql As String
    Dim lnMovNronew As Long
    Dim lsMovNroAnt As String
    Dim ldFechaAnt As Date
    Dim ldFechaExt As Date
    
    Dim oGen As New COMDConstSistema.DCOMGeneral
    
    'cti3
    Dim psCuenta As String
    psCuenta = ""
    
    lsMovNroAnt = GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(oGen.GetFechaMov(lsMovNroAnt, True))
    ldFechaExt = CDate(oGen.GetFechaMov(psMovNroNew, True))
    
    ExtornaMovimiento = 1
    If ldFechaAnt = ldFechaExt Or pbEliminaMov = True Then
        'MIOL 20120608, SEGUN RQ12093 SE AGREGO EL PARAMETRO psMovDesc**********************
        EliminaMov lsMovNroAnt, , psMovDesc
        InsertaMovModifica psMovNroNew, lsMovNroAnt, , psMovDesc
        'END MIOL **************************************************************************
        'CTI3 (ferimoro) 19102018
        InsertaDetExtorno pnMovNroAnt, psOpeCod, , psMotExtorno

    ElseIf ldFechaAnt <> ldFechaExt Or pbEliminaMov = False Then
        
        sql = " INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag )" _
            & " SELECT '" & psMovNroNew & "'," & IIf(psOpeCod = "", "cOpeCod", "'" & psOpeCod & "'") & ", " _
            & IIf(psMovDesc = "", "cMovDesc", "'" & Replace(psMovDesc, "''", "'") & "'") & ",nMovEstado, '" & gMovFlagDeExtorno & "'  " _
            & " FROM Mov WHERE nMovNro ='" & pnMovNroAnt & "' "
        oConect.Ejecutar sql
        lnMovNronew = GetnMovNro(psMovNroNew)
        
        sql = "INSERT MovDetExtornos (nMovNro,cOpedCod,cCtaCod,cMotExtorno,cDetMotExto) " _
            & "VALUES (" & lnMovNronew & ",'" & psOpeCod & "','" & psCuenta & "','" & psMotExtorno(0) & "','" & UCase(psMotExtorno(1)) & "')"
        oConect.Ejecutar sql
        
        InsertaMovRef lnMovNronew, pnMovNroAnt
        ActualizaSaldoMovimiento lsMovNroAnt, "-"
        
        sql = "INSERT MOVCONT (nMovNro, nMovMonto, nMovCorrela, cMovParalelo) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nMovMonto, nMovCorrela, cMovParalelo FROM MovCont Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
       
        sql = "INSERT MOVDOC (nMovNro, nDocTpo, cDocNro, dDocFecha) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nDocTpo, cDocNro, dDocFecha FROM MovDoc Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVGASTO (nMovNro, cPersCod, cPersCodDest) " _
            & " SELECT nMovNro = " & lnMovNronew & ",cPersCod, cPersCodDest FROM MovGasto Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, cCtaContCod, nMovImporte * -1  FROM MOVCTA Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVME (nMovNro, nMovItem, nMovMEImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovMEImporte*-1 FROM MOVME Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cObjetoCod FROM MOVOBJ Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJIF (nMovNro, nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod FROM MOVOBJIF Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
'        sql = "INSERT MOVCTAIF (nMovNro, cPersCod, cIFTpo, cCtaIFCod, nConceptoCod, nMovImporte) " _
'            & " SELECT nMovNro =" & lnMovNronew & ", cPersCod, cIFTpo, cCtaIFCod, nConceptoCod, nMovImporte * -1 FROM MOVCTAIF Where nMovNro =" & pnMovNroAnt & " "
'        oConect.Ejecutar sql
                
        sql = "INSERT MOVOBJEfectivo (nMovNro, nMovItem, nMovObjOrden, cEfectivoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cEfectivoCod FROM MOVOBJEfectivo Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJAreaAgencia (nMovNro, nMovItem, nMovObjOrden, cAreaCod, cAgeCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cAreaCod, cAgeCod FROM MOVOBJAreaAgencia Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT INTO MOVHABILITACION(nMovNro, cTranspCod, nMontoTrans) " _
            & " SELECT nMovNro=" & lnMovNronew & ", cTranspCod, nMontoTrans * -1 FROM MOVHABILITACION Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT INTO MOVHABDEVCAJERO(nMovNro, cUser, nMovImporte)  " _
            & " SELECT nMovNro=" & lnMovNronew & ",cUser, nMovImporte *-1 FROM MOVHABDEVCAJERO Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "UPDATE MOV SET nMovFlag = '" & gMovFlagExtornado & "' where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        'ExtornaMovCtaIF pnMovNroAnt
    End If
    'CTI6 ERS0112020
    If psOpeCod = "909002" Or psOpeCod = "909003" Then
           Dim RCapMov As ADODB.Recordset
           Set RCapMov = RecuperaMovimientoCapataciones(pnMovNroAnt)
            pnITFAhoExt = 0#
            psMovNroAExt = psMovNroNew
            Do While Not RCapMov.EOF
                    If RCapMov!nMonto > 0 Then
                        If Mid(RCapMov!cOpeCod, 1, 2) <> "99" Then
                            If RCapMov!cOpeCod = gAhoCargoCompra Or RCapMov!cOpeCod = gAhoCargoVenta Then
                                psOpeCodExtornoCap = RCapMov!cOpeCod
                                psCtaAhoExt = RCapMov!cCtaCod
                                pnMontoAhorroExt = RCapMov!nMonto
                                psOperacionDescExt = IIf(Mid(RCapMov!cCtaCod, 9, 1) = gMonedaNacional, "Extorno Cargo a Cuenta por Compra ME", "Extorno Cargo a Cuenta por Venta ME")
                                Call CapExtornoCargoAho(pnMovNroAnt, 0, IIf(Mid(RCapMov!cCtaCod, 9, 1) = gMonedaNacional, gAhoExtCargoCtaCompraME, gAhoExtCargoCtaVentaME), RCapMov!cCtaCod, psMovNroCap, psOperacionDescExt, RCapMov!nMonto)
                            End If
                        Else
                            Call CapExtornoCargoAho(pnMovNroAnt, 0, gITFCobroCargoExt, RCapMov!cCtaCod, psMovNroCap, "Extorno Cargo a Cuenta por Compra/Venta ME", RCapMov!nMonto)
                            pnITFAhoExt = RCapMov!nMonto
                        End If
                    End If
                    RCapMov.MoveNext
                Loop
            'CTI7 ERSOpeV2
            actualizarRelacionExtornoVoucherCaptacion (pnMovNroAnt)
            '************************
    End If
    'CTI6 ERS0112020
    ActualizaSaldoMovimiento psMovNroNew, "+"
    ExtornaMovimiento = 0
    Exit Function
ExtornaMovimientoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ExtornaMovimiento Method")
End Function
'CTI6 RS0112020
Public Function CapExtornoCargoAho(ByVal pnMovNro As Long, ByVal nMovNroBus As Long, ByVal nOperacion As CaptacOperacion, _
            ByVal sCuenta As String, ByVal sMovNro As String, ByVal sGlosa As String, ByVal nMonto As Double, _
            Optional nTipoDoc As TpoDoc = TpoDocCheque, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional sNomAge As String = "", Optional sLpt As String = "") As Double

Dim rsCta As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
Dim nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo
Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 1, True, False)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

'Inicia la transaccion

bTrans = True
If sGlosa = "" Then sGlosa = "Extorno Cargo Cuenta = " & sCuenta
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt + nMonto
nSaldoDisp = nSaldoDisp + nMonto
ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
'CTI6 ERS0112020
If nOperacion = gAhoExtCargoCtaCompraME Or nOperacion = gAhoExtCargoCtaVentaME Then
    sMsgOpe = sGlosa
Else
    If sNroDoc = "" Then
        sMsgOpe = "Ext. Retiro Efectivo"
    Else
        If nTipoDoc = TpoDocOrdenPago Then
            If nOperacion = gAhoExtRetOPCanje Then
                sMsgOpe = "Ext. Retiro OP Canje"
            Else
                sMsgOpe = "Ext. Retiro OP"
            End If
            AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstExtornada
        ElseIf nTipoDoc = TpoDocNotaCargo Then
            sMsgOpe = " Ext. Retiro NC"
        End If
    End If
End If
'end CTI6
ActualizaEstadoMov nMovNroBus, gMovFlagExtornado
AgregaMov sMovNro, nOperacion, sGlosa, , gMovFlagDeExtorno
nMovNro = GetnMovNro(sMovNro)

AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
UltimaActualizacionCuenta sCuenta, sMovNro

CapExtornoCargoAho = nSaldoCnt

End Function
Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
Dim sSql As String
sSql = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
oConect.ConexionActiva.Execute sSql
End Sub
Public Function RecuperaMovimientoCapataciones(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sSql As String

    sSql = "Select M.cMovNro, MC.cOpeCod, MC.cCtaCod, MC.nMonto from MovCap MC "
    sSql = sSql & " Inner Join Mov M ON M.nMovNro = MC.nMovNro "
    sSql = sSql & " Where MC.nMovNro = " & pnMovNro
    
    Set RecuperaMovimientoCapataciones = New ADODB.Recordset
    
    RecuperaMovimientoCapataciones.CursorLocation = adUseClient
    RecuperaMovimientoCapataciones.Open sSql, oConect.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
    RecuperaMovimientoCapataciones.ActiveConnection = Nothing

End Function
'CTI6 END
Public Function InsertaMovViaticos(ByVal pnViaticoMovNro As Long, ByVal pnMovNro As Long, _
                                   ByVal psDestinoCod As ViaticosDestino, _
                                   ByVal psTranspIda As ViaticosTransporte, _
                                   ByVal psTranspVuelta As ViaticosTransporte, ByVal pdPartida As Date, _
                                   ByVal pnMovViaticosDias As String, ByVal psDestinoDesc As String) As Integer
                                   
    On Error GoTo InsertaMovViaticosErr
    Dim sql As String
    InsertaMovViaticos = 1
    
    sql = "INSERT INTO MOVVIATICOS(nViaticoMovNro, nMovNro, " _
        & "cDestinoCod, cTranspIda, cTranspVuelta,dPartida, nMovViaticosDias, cDestinoDesc  ) " _
        & " VALUES (" & pnViaticoMovNro & "," & pnMovNro & ",'" _
        & psDestinoCod & "','" & psTranspIda & "','" & psTranspVuelta & "','" _
        & Format(pdPartida, oVarPublicas.gsFormatoFecha) & "'," & pnMovViaticosDias & ",'" & psDestinoDesc & "')"
     
    oConect.Ejecutar sql
    InsertaMovViaticos = 0
    Exit Function
InsertaMovViaticosErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovViaticos Method")
End Function

''Public Sub InsertaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, Optional SPersOrd As String = "", _
''                                Optional SPersRea As String = "", Optional SPersBen As String = "", _
''                                Optional SPersVis As String = "")
Public Sub InsertaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String, Optional SPersOrd As String = "", _
                                Optional SPersRea As String = "", Optional SPersBen As String = "", _
                                Optional SPersVis As String = "", _
                                Optional nTipoREU As Integer = 1, Optional nMontoAcu As Double = 0, Optional sOrigen As String = "")
                                'By Capi 28012008 se agrego parametros
                                'ALPA 20081009 *******************************
Dim sSql As String
If SPersBen = "" Then SPersBen = sPersCod
'ALPA 20081009 *******************************
''sSQL = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodOrd,cPersCodRea,cPersCodBen,cPersCodVisto) " _
''    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & SPersOrd & "','" & SPersRea & "','" & SPersBen & "','" & SPersVis & "')"
sSql = "INSERT MovLavDinero (nMovNro,cPersCod,cPersCodOrd,cPersCodRea,cPersCodBen,cPersCodVisto,nTipoReu,nMontoAcum,cOrigenEfectivo) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & SPersOrd & "','" & SPersRea & "','" & SPersBen & "','" & SPersVis & "'," & nTipoREU & ", " & nMontoAcu & ",'" & sOrigen & "')"

oConect.Ejecutar sSql
End Sub

Public Function ActualizaCajaChica(ByVal psAreaCod As String, ByVal psAgeCod As String, _
                                    ByVal pnProcNro As Integer, _
                                    Optional ByVal psPersCod As String = "", _
                                    Optional ByVal pnMontoAsig As Currency = -99999, _
                                    Optional ByVal pnMontoDesem As Currency = -99999, _
                                    Optional ByVal pnMontoSaldAnt As Currency = -99999, _
                                    Optional ByVal pnMontoSaldo As Currency = -99999) As Integer
                                    
    On Error GoTo ActualizaCajaChicaErr
    Dim sql As String
    Dim lsFiltro As String
    lsFiltro = ""
    If psPersCod <> "" Then
        lsFiltro = " cPersCod= '" & psPersCod & "',"
    End If
    If pnMontoAsig <> -99999 Then
        lsFiltro = lsFiltro & " nMontoAsig = " & pnMontoAsig & ","
    End If
    If pnMontoDesem <> -99999 Then
        lsFiltro = lsFiltro & " nMontoDesem = " & pnMontoDesem & ","
    End If
    If pnMontoSaldAnt <> -99999 Then
        lsFiltro = lsFiltro & " nMontoSaldAnt  = " & pnMontoSaldAnt & ","
    End If
    If pnMontoSaldo <> -99999 Then
        lsFiltro = lsFiltro & " nSaldo = " & pnMontoSaldo & ","
    End If
    lsFiltro = Mid(lsFiltro, 1, Len(lsFiltro) - 1)
    
    
    sql = "UPDATE CAJACHICA SET " & lsFiltro _
    & "     WHERE  cAreaCod ='" & psAreaCod & "' AND  cAgeCod='" & psAgeCod & "' AND nProcNro=" & pnProcNro
    
    oConect.Ejecutar sql
    
    Exit Function
ActualizaCajaChicaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaCajaChica Method")
End Function

Public Function CargaBilletaje(lsFecha As String, lsMoneda As String, lsCodAge As String, lnMoneda As String, Optional lsEfectivoCod As String = "") As ADODB.Recordset
Dim psSql As String

psSql = "SELECT m.cMovNro, mue.* " _
    & "FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro " _
    & "WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and " _
    & "       m.cMovNro LIKE '" & lsFecha & "%' and " _
    & "       m.cOpeCod = '" & 901016 & "' and " _
    & "       SubString(m.cMovNro,18,2) = '" & lsCodAge & "' " & IIf(lsEfectivoCod = "", "", " and cEfectivoCod = '" & lsEfectivoCod & "'")

' OJO gOpeHabCajRegEfect 901016
Set CargaBilletaje = oConect.CargaRecordSet(psSql)
End Function

Public Function CargaMovCta(Optional psMovNro As String = "", Optional psFiltro As String = "") As ADODB.Recordset
Dim psSql As String
   On Error GoTo CargaMovCtaErr
      psSql = "SELECT cMovNro, nMovItem, mc.cCtaContCod, isnull(b.cCtaContDesc,'') AS cCtaContDesc, " _
         & "       CONVERT(varchar(20),ISNULL(CASE WHEN nMovImporte > 0 THEN nMovImporte END,0),105) as nDebe, " _
         & "       CONVERT(varchar(20),ISNULL(CASE WHEN nMovImporte < 0 THEN nMovImporte * -1 END,0),105) as nHaber " _
         & " FROM  Mov M JOIN MovCta mc ON mc.nMovNro = m.nMovNro LEFT JOIN " & vsServerComunes & "CtaCont b ON mc.cCtaContCod = b.cCtaContCod " _
         & IIf(psMovNro = "", "", "WHERE cMovNro = '" & psMovNro & "' ") _
         & IIf(psFiltro = "", "", IIf(psMovNro = "", "WHERE ", "AND ")) & psFiltro
      Set CargaMovCta = oConect.CargaRecordSet(psSql)
   Exit Function
CargaMovCtaErr:
   Call oError.RaiseError(oError.MyUnhandledError, "DMov:CargaMovCta Method")
End Function
Public Sub ActualizaMovCta_Cuenta(psCtaCodant As String, psCtaCodNew As String, psFecha1 As String, psFecha2 As String)
    Dim sql As String
    On Error GoTo ActualizaMovCtaErr
    sql = "UPDATE mc SET cCtaContCod = '" & psCtaCodNew & "' FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro WHERE mc.cCtaContCod = '" & psCtaCodant & "' and LEFT(m.cMovNro,8) BETWEEN '" & psFecha1 & "' and '" & psFecha2 & "' "
    oConect.Ejecutar sql
    Exit Sub
ActualizaMovCtaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaMovCta_Cuenta Method")

End Sub
'ALPA 20140316******************************************
'Public Sub InsertaMovTpoCambio(ByVal pnMovNro As Long, ByVal pnMovTpoCambio As Currency)
'    On Error GoTo InsertaMovTipCambioErr
'    Dim sql As String
'
'    sql = "INSERT INTO MovTpoCambio (nMovNro, nMovTpoCambio) " _
'     & "VALUES (" & pnMovNro & ", " & pnMovTpoCambio & ")"
'
'    oConect.Ejecutar sql
'    Exit Sub
'InsertaMovTipCambioErr:
'    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovTipCambio Method")
'End Sub
Public Sub InsertaMovTpoCambio(ByVal pnMovNro As Long, ByVal pnMovTpoCambio As Currency, _
Optional ByVal pnValVent As Currency = 0, _
Optional ByVal pnValComp As Currency = 0, _
Optional ByVal pnValFijo As Currency = 0, _
Optional ByVal pnValFijoDia As Currency = 0, _
Optional ByVal pnValVentEsp As Currency = 0, _
Optional ByVal pnValCompEsp As Currency = 0, _
Optional ByVal pnValPond As Currency = 0, _
Optional ByVal pnValPondVenta As Currency = 0, _
Optional ByVal pnValPondREU As Currency = 0, _
Optional ByVal pnValSBSDia As Currency = 0, _
Optional ByVal pnValCompTr As Currency = 0, _
Optional ByVal pnValVentTr As Currency = 0)
    On Error GoTo InsertaMovTipCambioErr
    Dim sql As String
    
    sql = "INSERT INTO MovTpoCambio (nMovNro, nMovTpoCambio,nValVent,nValComp,nValFijo,nValFijoDia,nValVentEsp,nValCompEsp,nValPond,nValPondVenta,nValPondREU,nValSBSDia,nValCompTr,nValVentTr) " _
     & "VALUES (" & pnMovNro & ", " & pnMovTpoCambio & ", " & pnValVent & ", " & pnValComp & "," & pnValFijo & ", " & pnValFijoDia & ", " & pnValVentEsp & ", " & pnValCompEsp & ", " & pnValPond & ", " & pnValPondVenta & ", " & pnValPondREU & ", " & pnValSBSDia & "," & pnValCompTr & ", " & pnValVentTr & ")"
     
    oConect.Ejecutar sql
    Exit Sub
InsertaMovTipCambioErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovTipCambio Method")
End Sub
'CTI7 OpeV2********************************************************************
Public Sub AgregaMovRef(ByVal nMovNro As Long, ByVal nMovRef As Long)
sSql = "exec stp_ins_RegistrarMovRef " & nMovNro & "," & nMovRef & ""
oConect.Ejecutar sSql
End Sub
Public Sub ActualizaMovPendientesRend(ByVal nMovNro As Long, ByVal pnMonto As Currency)
    sSql = " UPDATE MovPendientesRend " _
         & " Set nSaldo = nSaldo - " & pnMonto & "" _
         & " Where nMovNro = " & nMovNro & ""
    oConect.Ejecutar sSql
End Sub

Public Sub actualizarRelacionVoucherCaptacion(ByVal pnMovNroOpe As Long, _
                                              ByVal pnMovNroRVD As Long)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_2  " & pnMovNroOpe & ", " & pnMovNroRVD & ""
    oConect.Ejecutar lsSQL
End Sub

Public Sub actualizarRelacionExtornoVoucherCaptacion(ByVal pnMovNroOpe As Long)
    Dim lsSQL As String
    lsSQL = "exec stp_upd_ActualizarCapVoucherDeposito_3  " & pnMovNroOpe & ""
    oConect.Ejecutar lsSQL
End Sub
'******************************************************************************
'CTI6 ERS0112020
Public Function GetDatosCuentaPF(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nPlazo, A.nIntPag, A.dRenovacion, A.nApertura, " _
    & "A.nFormaRetiro, A.dAuxiliar, A.nDuplicado, T.cConsDescripcion cEstado, A.bBusCredPend, " _
    & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, T3.cConsDescripcion cRetiro, " _
    & "A.dUltCierreAnt, A.dAuxiliarAnt, ISNULL(CI.cCtaCodAbono,'') cCtaCodAbono,a.nformaretiro,c.calias,c.nfirmasmin FROM Producto P " _
    & "INNER JOIN Captaciones C INNER JOIN CaptacPlazoFijo A LEFT JOIN CaptacCtaAboIntPF CI " _
    & "ON A.cCtaCod = CI.cCtaCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Constante T3 ON A.nFormaRetiro = T3.nConsValor " _
    & "INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod WHERE P.cCtaCod = '" & sCuenta & "' AND " _
    & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa & " AND T3.nConsCod = " & gCaptacPFFormaRetiro


Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, oConect.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaPF = rsCta
Set rsCta = Nothing

End Function
Public Function GetDatosCuentaCTS(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, P.dPrdEstado, " _
    & "C.nSaldoDisp, C.nPersoneria, C.nFirmas,c.nfirmasmin ,C.nIntAcum, dUltCierre=( case when dbo.ultimodiames(dUltCierre)=1 then dUltCierre else dateadd(dd,1, dUltCierre) end )  , C.dApertura, AG.cAgeDescripcion, " _
    & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.nSaldRetiro, A.nIntSaldo, PE.cPersNombre cInstitucion, " _
    & "T.cConsDescripcion cEstado, T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.cCodInst,C.cAlias " _
    & "FROM Producto P INNER JOIN Captaciones C INNER JOIN CaptacCTS A INNER JOIN " _
    & "Persona PE ON A.cCodInst = PE.cPersCod ON C.cCtaCod = A.cCtaCod ON " _
    & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
    & "INNER JOIN Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN Constante T2 " _
    & "ON C.nPrdTasaInteres = T2.nConsValor INNER JOIN Agencias AG ON SUBSTRING(P.cCtaCod,4,2) = AG.cAgeCod " _
    & "WHERE P.cCtaCod = '" & sCuenta & "' AND T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
    & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, oConect.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaCTS = rsCta
Set rsCta = Nothing

End Function
Public Function GetDatosCuenta(ByVal sCuenta As String) As ADODB.Recordset
Dim rsCta As ADODB.Recordset
Dim nProd As Producto
nProd = CInt(Mid(sCuenta, 6, 3))

Select Case nProd
    Case gCapAhorros
        Set rsCta = GetDatosCuentaAho(sCuenta)
        
    Case gCapPlazoFijo
        Set rsCta = GetDatosCuentaPF(sCuenta)
        
    Case gCapCTS
        Set rsCta = GetDatosCuentaCTS(sCuenta)
        
End Select
Set GetDatosCuenta = rsCta
Set rsCta = Nothing

End Function
Public Function ValidaSaldoCuenta(ByVal sCuenta As String, ByVal nMonto As Double) As Boolean
Dim nMoneda As Moneda
Dim rsCta As ADODB.Recordset
Dim nSaldo As Double, nMontoMinimo As Double

Set rsCta = GetDatosCuenta(sCuenta)
nSaldo = rsCta("nSaldoDisp")
rsCta.Close
Set rsCta = Nothing

nMoneda = CLng(Mid(sCuenta, 9, 1))
If nMoneda = gMonedaNacional Then
    nMontoMinimo = GetCapParametro(gSaldMinAhoMN)
Else
    nMontoMinimo = GetCapParametro(gSaldMinAhoME)
End If
If nSaldo - nMontoMinimo - nMonto >= 0 Then
    ValidaSaldoCuenta = True
Else
    ValidaSaldoCuenta = False
End If
End Function
Public Function GetCapParametro(ByVal nParametro As CaptacParametro) As Double
Dim rsVar As ADODB.Recordset
Dim sSql As String
sSql = "SELECT nParValor FROM Parametro WHERE nParCod = " & nParametro & " " _
    & "And nParProd = " & gPrdParamCaptac
Set rsVar = New ADODB.Recordset
rsVar.CursorLocation = adUseClient
rsVar.Open sSql, oConect.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsVar.ActiveConnection = Nothing
If rsVar.EOF And rsVar.BOF Then
    GetCapParametro = 0
Else
    GetCapParametro = rsVar("nParValor")
End If
rsVar.Close
Set rsVar = Nothing
End Function
Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function
Public Function GetDatosCuentaAho(ByVal sCuenta As String) As ADODB.Recordset
Dim sSql As String
Dim rsCta As ADODB.Recordset

sSql = "Exec sp_sel_DatosCuentaAhorro '" & sCuenta & "' , " & gCaptacEstado & " , " & gProductoCuentaTipo & " , " & gCaptacTipoTasa

Set rsCta = New ADODB.Recordset
rsCta.CursorLocation = adUseClient
rsCta.Open sSql, oConect.ConexionActiva, adOpenStatic, adLockReadOnly, adCmdText
Set rsCta.ActiveConnection = Nothing
Set GetDatosCuentaAho = rsCta
Set rsCta = Nothing
End Function
Public Function CapCargoCuentaAho(ByRef pMatAho As Variant, ByVal sCuenta As String, ByVal nMonto As Double, _
            ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, ByVal sGlosa As String, ByVal pdFecSis As Date, _
            Optional nTipoDoc As TpoDoc = TpoDocNotaCargo, Optional sNroDoc As String = "", _
            Optional sCodIF As String = "", Optional bOPExiste As Boolean = False, _
            Optional bActivaCta As Boolean = True, Optional bNumExtracto As Boolean = True, _
            Optional sCodCMAC As String = "", Optional sNomCmac As String = "", _
            Optional sNomAge As String = "", Optional sLpt As String = "", _
            Optional sPersLavDinero As String = "", Optional bCancelaciones As Boolean = False, _
            Optional pnMonedaCred As Moneda = 0, _
            Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0, Optional pbITFAsumido As Boolean = False, _
            Optional psItfOperacion As String = "", Optional pbRetiroEfectivo As Boolean = False, _
            Optional ByVal pnTipoPago As Integer = 1) As Double
            'FRHU RQ14008 Se agrego: pbRetiroEfectivo

Dim rsCta As ADODB.Recordset
Dim nEstado As CaptacEstado
Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
Dim nSaldoInac As Double
Dim dUltMov As Date
Dim bInactiva As Boolean, bTrans As Boolean
Dim nNumExtracto As Long, nDiasTranscurridos As Long, i As Long
Dim nExtracto As Long, nMovNro As Long
Dim sMsgOpe As String, sCodUser As String
Dim dFecSis As Date
Dim lnMontoOri As Double
Dim oGen As COMDConstSistema.DCOMGeneral
Dim lnTipCambioV As Currency
Dim lnTipCambioC As Currency
Dim lnTipCambio As Currency

'Dim clsMov As COMNContabilidad.NCOMContFunciones

dFecSis = CDate(Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4))
sCodUser = Right(sMovNro, 4)
bTrans = False
'Obtiene los datos para el calculo

Set rsCta = GetDatosCuentaAho(sCuenta)

nEstado = rsCta("nPrdEstado")
nSaldoDisp = rsCta("nSaldoDisp")
nSaldoCnt = rsCta("nSaldo")
dUltMov = rsCta("dUltCierre")
nTasa = rsCta("nTasaInteres")
nIntAcum = rsCta("nIntAcum")
bInactiva = IIf(rsCta("bInactiva") = 0, False, True)
nExtracto = rsCta("nTransacc") + 1
rsCta.Close
Set rsCta = Nothing

Set oGen = New COMDConstSistema.DCOMGeneral

'GetTipCambio pdFecSis
lnTipCambioC = oGen.EmiteTipoCambio(pdFecSis, TCCompra)
lnTipCambioV = oGen.EmiteTipoCambio(pdFecSis, TCVenta)
lnTipCambio = oGen.EmiteTipoCambio(pdFecSis, TCFijoDia)

'RIRO 20200706 ****************
Dim sMensajeReactiva As String
sMensajeReactiva = oGen.LeeConstSistema("725")
If Not Len(Trim(sMensajeReactiva)) > 0 Then
    sMensajeReactiva = "Retiro Comisión Reactiva"
End If
'END RIRO *********************

If pnMonedaCred = 0 Then pnMonedaCred = Mid(sCuenta, 9, 1)

If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    lnMontoOri = nMonto
    If pnMonedaCred = gMonedaExtranjera And Mid(sCuenta, 9, 1) = gMonedaNacional Then
        nMonto = Round(nMonto * lnTipCambioV, 2)
    Else
        nMonto = Round(nMonto / lnTipCambioC, 2)
    End If
Else
    lnMontoOri = 0
End If

bTrans = True
Randomize
For i = 0 To Rnd(2000) * 1000
Next i

If Not ValidaSaldoCuenta(sCuenta, nMonto) Then 'Valida el saldo de la cuenta nuevamente
    Err.Raise 1000, "CapCargoCuentaAho", "Cuenta NO Posee Saldo Suficiente"
    CapCargoCuentaAho = 0
    Exit Function
End If
If bActivaCta Then
    If bInactiva Then
        nSaldoInac = nSaldoCnt
    End If
End If
'Calcula los intereses
nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
If Not bCancelaciones Then
    pMatAho(3) = Format(nIntGanado, "#0.00")
Else
    pMatAho(4) = Format(nIntGanado, "#0.00")
End If

dUltMov = DateAdd("d", -1, dFecSis)
nSaldoCnt = nSaldoCnt - nMonto

ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp

nSaldoDisp = nSaldoDisp - nMonto
If Not bCancelaciones Then
    pMatAho(10) = nSaldoDisp
    pMatAho(11) = nSaldoCnt
Else
    pMatAho(12) = nSaldoDisp
    pMatAho(13) = nSaldoCnt
End If

'FRHU20140228 RQ14006
If pbRetiroEfectivo Then
    pMatAho(15) = nSaldoDisp
    pMatAho(16) = nSaldoCnt
    pMatAho(17) = Format(nIntGanado, "#0.00")
End If
'FIN FRHU20140228 RQ14006
ActualizaCargoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True

If sNroDoc <> "" Then
    If nTipoDoc = TpoDocOrdenPago Then 'Si es con orden de pago
        If bOPExiste Then
            AgregaOrdenPagoEstado sCuenta, sNroDoc, sMovNro, nMonto, gCapOPEstCobrada
        Else
            AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMonto, gCapOPEstCobrada
        End If
        If nOperacion = gAhoRetOPCanje Then
            sMsgOpe = "Retiro OP Canje"
        ElseIf nOperacion = gAhoRetOP Then
            sMsgOpe = "Retiro OP"
        ElseIf nOperacion = gAhoOPCertificacion Then
            sMsgOpe = "Certificacion OP"
        
        'RIRO 20200630 ******
        ElseIf nOperacion = "100949" Then
            sMsgOpe = sMensajeReactiva
        'END RIRO ***********
            
        End If
    ElseIf nTipoDoc = TpoDocNotaCargo Then
        sMsgOpe = "Retiro NC"
    End If
Else
        'CTI6 ERS0112020
        Select Case pnTipoPago
         Case gCVTipoPagoEfectivo   ' Efectivo
                'RIRO 20200630 ******
                If nOperacion = "100949" Then
                    sMsgOpe = sMensajeReactiva
                'END RIRO ***********
                Else
                    sMsgOpe = "Retiro Efectivo"
                End If
           Case gCVTipoPagoCargoCta
           If Mid(sCuenta, 9, 1) = gMonedaNacional Then
                      sMsgOpe = "Retiro OP Compra Cargo a Cuenta"
                 Else
                   sMsgOpe = "Retiro OP Venta Cargo a Cuenta"
                 End If
        End Select
        'END CTI6
        
'        'RIRO 20200630 ******
'        If nOperacion = "100949" Then
'            sMsgOpe = sMensajeReactiva
'        'END RIRO ***********
'        Else
'            sMsgOpe = "Retiro Efectivo"
'        End If
End If

'AgregaMov sMovNro, nOperacion, sGlosa

nMovNro = GetnMovNro(sMovNro)
AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
If bInactiva Then
    AgregaMovCap nMovNro, gAhoEstInacAct, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
    AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
    ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
    ActualizaInactivaAct (sCuenta) 'JUEZ 20150127
End If

'RIRO 20200611 Comisión Reactiva ****
If nOperacion = "100949" Then
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, 1251, nMonto
Else
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto
End If
'End RIRO 20200611 ******************

'AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto RIRO20200611 Comentado
'si el pago de credito es en diferente moneda a la cuenta de ahorros
Dim lnMontoCV As Double
If pnMonedaCred <> Mid(sCuenta, 9, 1) Then
    If Mid(sCuenta, 9, 1) = gMonedaExtranjera Then
        lnMontoCV = Round((lnMontoOri / lnTipCambioC) * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioC
        
        If lnTipCambioC < lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
        End If
    Else
        lnMontoCV = Round(lnMontoOri * lnTipCambio, 2)
        AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
        AgregaMovTipoCambio nMovNro, lnTipCambioV
        
        If lnTipCambioV > lnTipCambio Then
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        Else
            AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(lnMontoOri * lnTipCambioV, 2))
        End If
    End If
End If
CapCargoCuentaAho = nSaldoCnt

If pbITFAplica And pnITFValor > 0 Then
    If pbITFAsumido Then
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp, nSaldoCnt
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFAsumido, pnITFValor
    Else
        ActualizaCargoCaptacion sCuenta, pnITFValor, pnITFValor, nIntGanado, dUltMov, sMovNro, bNumExtracto
        AgregaMovCap nMovNro, psItfOperacion, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
        AgregaMovCapDet nMovNro, psItfOperacion, sCuenta, gConcITFCliente, pnITFValor
    End If
    
    nSaldoDisp = nSaldoDisp - pnITFValor
    If Not bCancelaciones Then
        pMatAho(10) = nSaldoDisp
        pMatAho(11) = nSaldoCnt - pnITFValor
    Else
        pMatAho(12) = nSaldoDisp
        pMatAho(13) = nSaldoCnt - pnITFValor
    End If
    'FRHU20140429 INCIDENTE: Se agrego para que en el vouche muestre el saldo correcto. cuando hace el retiro
    If pbRetiroEfectivo Then
        pMatAho(15) = nSaldoDisp
        pMatAho(16) = nSaldoCnt - pnITFValor
        pMatAho(17) = Format(nIntGanado, "#0.00")
    End If
    'FIN FRHU20140429
    CapCargoCuentaAho = nSaldoCnt - pnITFValor
End If
UltimaActualizacionCuenta sCuenta, sMovNro

bTrans = False
End Function
Public Sub AgregaMovTipoCambio(ByVal nMovNro As Long, ByVal pnTipoCambio As Currency)
Dim sSql As String

sSql = " INSERT MovTpoCambio (nMovNro, nMovTpoCambio) " _
     & " VALUES (" & nMovNro & "," & pnTipoCambio & ")"

oConect.ConexionActiva.Execute sSql
End Sub
Public Sub ActualizaInactivaAct(ByVal sCuenta As String)
Dim sSql As String
    sSql = "Update CaptacAhorros Set bInactiva = 0 WHERE cCtaCod = '" & sCuenta & "'"
    oConect.ConexionActiva.Execute sSql
End Sub
Public Function InsertaMov(ByVal psMovNro As String, _
                            ByVal psOpeCod As String, ByVal psMovDesc As String, _
                            Optional ByVal pnMovEstado As MovEstado = gMovEstContabMovContable, _
                            Optional ByVal pnMovFlag As MovFlag = gMovFlagVigente) As Integer
On Error GoTo InsertaMovErr
    Dim sql As String
    
    sql = "INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag)" _
        & " VALUES('" & psMovNro & "','" & psOpeCod & "','" & Replace(psMovDesc, "'", "''") & "'," _
        & pnMovEstado & "," & pnMovFlag & ")"
     
    oConect.Ejecutar sql
    
    Exit Function
InsertaMovErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMov Method")
End Function
Public Function GetnMovNro(ByVal psMovNro As String) As Long
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

sql = "Select Max(nMovNro) nMovNro From Mov where cMovNro ='" & psMovNro & "'"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetnMovNro = rs!nMovNro
End If
rs.Close
Set rs = Nothing
End Function
Public Function GetcMovNro(ByVal pnMovNro As Long) As String
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

sql = "Select cMovNro From Mov where nMovNro ='" & pnMovNro & "'"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetcMovNro = rs!cMovNro
End If
rs.Close
Set rs = Nothing
End Function
Public Sub EliminaMovRef(ByVal pnMovNro As Long)
    On Error GoTo EliminaMovRefErr
    Dim sql As String
    sql = "DELETE MOVREF WHERE nMOVNRO =" & pnMovNro & ""
        
    oConect.Ejecutar sql
    Exit Sub
EliminaMovRefErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:EliminaMovRef Method")
End Sub
Public Function InsertaMovHabilitaciones(ByVal pnMovNro As Long, ByVal psCodTransp As String, _
    ByVal psAreaOrig As String, ByVal psAgeOrig As String, ByVal psAreaDest As String, _
    ByVal psAgeDest As String, ByVal pnMontoTrans As Double, ByVal pnMovImporte As Double, _
    ByVal nMoneda As Moneda) As Integer
    
On Error GoTo InsertaMovHabilitacionesErr
    Dim sql As String
    InsertaMovHabilitaciones = 1
    sql = "INSERT INTO MovHabilitacion (nMovNro, cTranspCod, cAreaOrig, cAgeOrig, cAreaDest, cAgeDest, nMontoTrans, nMovImporte, nMoneda)" _
        & " VALUES(" & pnMovNro & ",'" & psCodTransp & "','" & psAreaOrig & "','" & psAgeOrig & "','" _
        & psAreaDest & "','" & psAgeDest & "'," & pnMontoTrans & "," & pnMovImporte & "," & nMoneda & ")"
     
    oConect.Ejecutar sql
    InsertaMovHabilitaciones = 0
    Exit Function
InsertaMovHabilitacionesErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovTipCambio Method")
End Function

Public Function InsertaMovHabDev(ByVal pnMovNro As Long, ByVal pnImporte As Double, _
        ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal nMoneda As Moneda, _
        ByVal sUsuOrig As String, ByVal sUsuDest As String) As Integer
    On Error GoTo InsertaMovHabDevErr
    Dim sql As String
    InsertaMovHabDev = 1
    sql = "Insert MovHabDevCajero(nMovNro,nMovImporte,cAgeCod,cAreaCod,nMoneda,cUsuOrig,cUsuDest) " _
        & "Values(" & pnMovNro & "," & pnImporte & ",'" & psAgeCod & "','" & psAreaCod & "'," & nMoneda & ",'" & sUsuOrig & "','" & sUsuDest & "')"
    oConect.Ejecutar sql
    InsertaMovHabDev = 0
    Exit Function
InsertaMovHabDevErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovHabDev Method")
End Function

'MADM 20110128
Public Function InsertaMovHabDevPreCuadre(ByVal pnMovNro As Long, ByVal pnImporte As Double, _
        ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal nMoneda As Moneda, _
        ByVal sUsuOrig As String, ByVal sUsuDest As String) As Integer
    On Error GoTo InsertaMovHabDevPreCuadreErr
    Dim sql As String
    InsertaMovHabDevPreCuadre = 1
    sql = "Insert MovHabDevCajeroPreCuadre(nMovNro,nMovImporte,cAgeCod,cAreaCod,nMoneda,cUsuOrig,cUsuDest) " _
        & "Values(" & pnMovNro & "," & pnImporte & ",'" & psAgeCod & "','" & psAreaCod & "'," & nMoneda & ",'" & sUsuOrig & "','" & sUsuDest & "')"
    oConect.Ejecutar sql
    InsertaMovHabDevPreCuadre = 0
    Exit Function
InsertaMovHabDevPreCuadreErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovHabDev Method")
End Function
'END MADM

Public Function InsertaMovUserEfectivo(ByVal pnMovNro As Long, ByVal psCodUser As String, ByVal psEfectivoCod As String, _
                                        ByVal pnMonto As Currency) As Integer
    On Error GoTo InsertaMovUserEfectivoErr
    Dim sql As String
    InsertaMovUserEfectivo = 1
     
    sql = "INSERT INTO MovUserEfectivo(nMovNro, cUser,cEfectivoCod, nMonto) " _
        & " VALUES(" & pnMovNro & ",'" & psCodUser & "','" & psEfectivoCod & "'," & pnMonto & ")"

    oConect.Ejecutar sql
    InsertaMovUserEfectivo = 0
    Exit Function
InsertaMovUserEfectivoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovUserEfectivo Method")
End Function

'MADM 20110121
Public Function InsertaMovPreCuadre(ByVal pnMovNro As Long, ByVal psCodUser As String, ByVal pnMontoSol As Currency, _
                                        ByVal pnMontoDol As Currency) As Integer
    On Error GoTo InsertaMovPreCuadreErr
    Dim sql As String
    InsertaMovPreCuadre = 1
     
    sql = "INSERT INTO MovPreCuadre(nMovNro, cCodUser, nMontoSol, nMontoDol, nFlag) " _
        & " VALUES(" & pnMovNro & ",'" & psCodUser & "'," & pnMontoSol & "," & pnMontoDol & ", 0)"

    oConect.Ejecutar sql
    InsertaMovPreCuadre = 0
    Exit Function
InsertaMovPreCuadreErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovUserEfectivo Method")
End Function
'END MADM

Public Function ActualizaMovUserEfectivo(ByVal pnMovNro As Long, ByVal psCodUser As String, ByVal psEfectivoCod As String, _
                                         ByVal pnMonto As Currency) As Integer
    On Error GoTo ActualizaMovUserEfectivoErr
    Dim sql As String
    ActualizaMovUserEfectivo = 1
     
    sql = "UPDATE MovUserEfectivo SET nMonto =" & pnMonto & " " _
        & " WHERE nMovNro =" & pnMovNro & " AND cUser='" & psCodUser & "' AND cEfectivoCod='" & psEfectivoCod & "'"

    oConect.Ejecutar sql
    ActualizaMovUserEfectivo = 0
    Exit Function
ActualizaMovUserEfectivoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaMovUserEfectivo Method")
End Function
Public Sub ActualizaMovCont(ByVal pnMovNro As Long, Optional ByVal pnMovMonto As Currency, Optional ByVal pnMovCorrela As Integer, Optional ByVal psMovParalelo As String = "")
Dim sModif As String
Dim sql    As String
    On Error GoTo ActualizaMovContErr
    If pnMovMonto <> 0 Then
      sModif = "nMovMonto = " & pnMovMonto
    End If
    If pnMovCorrela <> 0 Then
      sModif = sModif & IIf(sModif = "", "", ",") & "nMovCorrela = " & pnMovCorrela
    End If
    If Not psMovParalelo = "" Then
      sModif = sModif & IIf(sModif = "", "", ",") & "cMovParalelo = '" & psMovParalelo
    End If
    If sModif = "" Then
       Err.Raise "50001", "DMov:ActualizaMovCont", "Indicar campos de Mov a Actualizar"
    Else
        sql = "UPDATE MovCont SET " & sModif & " WHERE nMovNro = " & pnMovNro & ""
          
        oConect.Ejecutar sql
    End If
Exit Sub
ActualizaMovContErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaMovCont Method")
End Sub
Public Function GeneraMovNro(Optional ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional ByVal psMovNro As String = "", Optional pnCont As Integer = 0) As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    If pnCont = 0 And psMovNro <> "" Then
        psMovNro = Left(psMovNro, 19) & Format(Val(Mid(psMovNro, 20, 2)) + 1, "00") & Right(psMovNro, 4)
    End If
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConect.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "','" & psCodAge & "','" & psUser & "'"
    Else
       sql = "sp_GeneraMovNro '','','','" & psMovNro & "'"
    End If
    Set rs = oConect.Ejecutar(sql)
    If Not rs.EOF Then
        GeneraMovNro = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
GeneraMovNroErr:
    Call oError.RaiseError(oError.MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function
Public Function CargaRecordSet(ByVal sql As String) As ADODB.Recordset
Set CargaRecordSet = oConect.CargaRecordSet(sql)
End Function
Public Function InsertaMovCompraVenta(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal pnImporte As Currency) As Integer
    On Error GoTo MovCompraVentaErr
    Dim sql As String
    InsertaMovCompraVenta = 1
    
    sql = " INSERT INTO MovCompraVenta(nMovNro, cPersCod, nMovImporte )" _
        & " VALUES(" & pnMovNro & ",'" & psPersCod & "'," & pnImporte & ") "
      
    oConect.Ejecutar sql
    InsertaMovCompraVenta = 0
    Exit Function
MovCompraVentaErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:MovCompraVenta Method")
End Function

Public Function InsertaMovOpeVarias(ByVal pnMovNro As Long, ByVal psNroDoc As String, _
            ByVal psReferencia As String, ByVal pnImporte As Double, _
            Optional nMoneda As Moneda = gMonedaNacional) As Integer
    On Error GoTo InsertaMovOpeVariasErr
    Dim sql As String
    InsertaMovOpeVarias = 1
    
    sql = " INSERT INTO MovOpeVarias(nMovNro, cNroDoc, cReferencia, nMovImporte,nMoneda)" _
        & " VALUES(" & pnMovNro & ",'" & psNroDoc & "','" & psReferencia & "'," & pnImporte & "," & nMoneda & ") "
      
    oConect.Ejecutar sql
    InsertaMovOpeVarias = 0
    Exit Function
InsertaMovOpeVariasErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovOpeVarias Method")
End Function

Public Function InsertaNotaAbonoCargo(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pnEstado As NotaCargoAbonoEstado, _
                                    ByVal pnMotivo As Integer, ByVal pnMonto As Currency, Optional psObjetoCod As String = "", Optional psObjeto As String = "", Optional pnMoneda As Moneda = gMonedaNacional) As Integer
    On Error GoTo InsertaNotaAbonoCargoErr
    Dim sql As String
    InsertaNotaAbonoCargo = 1
    
    sql = " INSERT INTO NotaAbonoCargo(nDocTpo, cDocNro, nEstado, nMotivoCod, nMonto, cObjetoCodPadre, cObjetoCod, nMoneda)" _
        & " VALUES(" & pnTpoDoc & ",'" & psNroDoc & "'," & pnEstado & "," & pnMotivo & "," & pnMonto & ",'" & psObjetoCod & "','" & psObjeto & "'," & pnMoneda & ") "
    
    oConect.Ejecutar sql
    InsertaNotaAbonoCargo = 0
    Exit Function
InsertaNotaAbonoCargoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaNotaAbonoCargo Method")
End Function
Public Function InsertaNotaAbonoCargoEst(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pnEstado As NotaCargoAbonoEstado, _
                                        ByVal psMovNro As String) As Integer
    On Error GoTo InsertaNotaAbonoCargoEstErr
    Dim sql As String
    InsertaNotaAbonoCargoEst = 1
    
    sql = " INSERT INTO NotaAbonoCargoEst(nDocTpo, cDocNro, nEstado, cMovNro)" _
        & " VALUES(" & pnTpoDoc & ",'" & psNroDoc & "'," & pnEstado & ",'" & psMovNro & "') "
    
    oConect.Ejecutar sql
    InsertaNotaAbonoCargoEst = 0
    Exit Function
InsertaNotaAbonoCargoEstErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaNotaAbonoCargoEst Method")
End Function

Public Function InsertaMovArqueo(ByVal pnMovNro As Long, ByVal psAreaCod As String, ByVal psPersCod As String, ByVal psHoraIni As String, _
                                ByVal psHoraFin As String, ByVal pnEfectivo As Currency, ByVal pnOPImporte As Currency, _
                                ByVal pnTotDoc As Currency, ByVal pnTotDocRend As Currency, ByVal pnArendirPend As Currency, _
                                ByVal pnDiferencia As Currency) As Integer
                                
    On Error GoTo InsertaMovArqueoErr
    Dim sql As String
    InsertaMovArqueo = 1
    
    sql = " INSERT INTO MOVARQUEO(nMovNro, cAreaCod, cPersCod, cHoraIni, cHoraFin, " _
        & " nEfectivo, nOPImporte, nTotDoc, nTotDocRend, nArendirPend, nDiferencia)" _
        & " VALUES(" & pnMovNro & ",'" & psAreaCod & "','" & psPersCod & "','" & psHoraIni & "','" & psHoraFin & "'," _
        & pnEfectivo & "," & pnOPImporte & "," & pnTotDoc & "," & pnTotDocRend & "," & pnArendirPend & "," & pnDiferencia & ") "

    
    oConect.Ejecutar sql
    InsertaMovArqueo = 0
    Exit Function
InsertaMovArqueoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovArqueo Method")
End Function
Public Function InsertaMovArqueoOP(ByVal pnMovNro As Long, ByVal psNroOrden As String, ByVal pnImporte As String) As Integer
    On Error GoTo InsertaMovArqueoOPErr
    Dim sql As String
    InsertaMovArqueoOP = 1
    
    sql = " INSERT INTO MovArqueoOp(nMovNro, cNroOrden, nImporte) " _
        & " VALUES(" & pnMovNro & ",'" & psNroOrden & "'," & pnImporte & ")"

    
    oConect.Ejecutar sql
    InsertaMovArqueoOP = 0
    Exit Function
InsertaMovArqueoOPErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovArqueoOP Method")
End Function
Public Function InsertaMovArqueoEfectivo(ByVal pnMovNro As Long, ByVal psEfectivoCod As String, ByVal pnImporte As Currency) As Integer
    On Error GoTo InsertaMovArqueoEfectivoErr
    Dim sql As String
    InsertaMovArqueoEfectivo = 1
     
    sql = "INSERT INTO MovArqueoEfectivo(nMovNro, cEfectivoCod, nImporte ) " _
        & " VALUES(" & pnMovNro & ",'" & psEfectivoCod & "'," & pnImporte & ")"

    oConect.Ejecutar sql
    InsertaMovArqueoEfectivo = 0
    Exit Function
InsertaMovArqueoEfectivoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovArqueoEfectivo Method")
End Function

 
Public Function AgregaNuevaCaptacion(ByVal nProducto As Producto, ByVal nMoneda As Moneda, _
        ByVal sAgencia As String, ByVal nTasa As Double, ByVal nSaldo As Double, _
        ByVal dFecha As Date, ByVal nFirmas As Integer, ByVal nPersoneria As PersPersoneria, _
        ByVal nTipoCuenta As ProductoCuentaTipo, ByVal sMovNro As String, _
        ByVal nTipoTasa As CaptacTipoTasa, Optional bCheque As Boolean = False, _
        Optional ByVal bOrdPag As Boolean = False, Optional nPlazo As Long = 0, _
        Optional nFormaRetiro As CaptacPFFormaRetiro = gCapPFFormRetMensual, _
        Optional bCtaAboInt As Boolean = False, Optional sCtaCodAbono As String = "", _
        Optional nPorcRetCTS As Double = 0, Optional sInstitucion As String = "") As String
        
Dim clsGen As COMDConstSistema.DCOMGeneral
Dim sCuenta As String, sFecha As String
Dim sFecUltCierre As String, sOrdPag As String, sCtaAbonoIntPF As String
Dim nSaldoDisp As Double, nSaldRetCTS As Double
Set clsGen = New COMDConstSistema.DCOMGeneral
sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
Set clsGen = Nothing
sFecha = Format$(dFecha, oVarPublicas.gsFormatoFecha) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
sSql = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"
oConect.Ejecutar sSql


nSaldoDisp = IIf(bCheque, 0, nSaldo)
sFecUltCierre = Format$(DateAdd("d", -1, dFecha), oVarPublicas.gsFormatoFecha)
sSql = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion) " _
    & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
    & nTipoTasa & ",'" & sMovNro & "')"
oConect.Ejecutar sSql
Select Case nProducto
    Case gCapAhorros
        sOrdPag = IIf(bOrdPag, "1", "0")
        sSql = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto) " _
            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "')"
    Case gCapPlazoFijo
        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
        sSql = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho) " _
            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ")"
    Case gCapCTS
        nSaldRetCTS = Round(nSaldo * nPorcRetCTS / 100, 2)
        sSql = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst) " _
            & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "')"
End Select
oConect.Ejecutar sSql
If nProducto = gCapPlazoFijo And bCtaAboInt Then
    sSql = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
        & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
    oConect.Ejecutar sSql
End If
AgregaNuevaCaptacion = sCuenta
End Function
Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
sSql = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
oConect.Ejecutar sSql
End Sub

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String) As Recordset
Dim rsOP As Recordset
sSql = "SELECT E.cMovNro, D.nTpoDoc, D.cNroDoc, D.cCtaCod, D.nMonto, P.cPersNombre, E.cEstado " _
    & "D.cIFCodPers FROM DocRecOPEst E INNER JOIN DocRecOP D INNER JOIN " & vsServerPersonas & "Persona P " _
    & "ON D.cIFCodPers = P.cPersCod ON E.nTpoDoc = D.nTpoDoc AND E.cNroDoc = D.cNroDoc AND E.cCtaCod " _
    & "= D.cCtaCod WHERE E.cMovNro IN (SELECT MAX(E1.cMovNro) FROM DocRecOPEst E WHERE E1.nTpoDoc = " _
    & "E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cCtaCod = E.cCtaCod)"

Set rsOP = New Recordset
Set rsOP = oConect.CargaRecordSet(sSql)
Set GetDatosOrdenPago = rsOP
End Function
Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
sSql = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSql
sSql = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
oConect.Ejecutar sSql
End Sub

Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada)
Dim sFecha As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
Select Case nTpoDoc
    Case TpoDocCheque
        sSql = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cCtaCod,nMonto) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "','" & sCuenta & "'," & nMonto & ")"
    Case TpoDocOrdenPago
        sSql = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
    Case TpoDocNotaAbono, TpoDocNotaCargo
End Select
oConect.Ejecutar sSql
sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
oConect.Ejecutar sSql
End Sub
Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
sSql = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
oConect.Ejecutar sSql
End Sub
Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
        
sSql = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
oConect.Ejecutar sSql
End Sub
Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
sSql = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSql
End Sub
Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
oConect.Ejecutar sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, oVarPublicas.gsFormatoFecha) & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSql

End Sub

Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSql = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
oConect.Ejecutar sSql
sSql = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, oVarPublicas.gsFormatoFecha) & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSql
End Sub

Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
sSql = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSql
End Sub
Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
'Actualiza el ultimo estado a la tabla de producto
sSql = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, oVarPublicas.gsFormatoFecha) & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSql
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSql = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
oConect.Ejecutar sSql
End Sub
Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado)
'Actualiza el ultimo estado a la tabla de Documento Recibidos
sSql = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
oConect.Ejecutar sSql
'Agrega un registro mas de estado a la historia de estados del documento recibido
sSql = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado) " _
    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ")"
oConect.Ejecutar sSql
End Sub
Public Function GetMovExtorno(ByVal sDatoBus As String, Optional nTipoBus As Integer = 0, Optional pdFecSis As Date) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(pdFecSis, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
End If
sSql = "Select M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
    & "M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCap C " _
    & "INNER JOIN " & vsServerComunes & "OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro WHERE " & sWhere & " AND C.cOpeCod NOT IN ('" & gAhoEstInacAct & "') " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") " _
    & "ORDER BY M.nMovNro"
Set rsMov = New Recordset
Set rsMov = oConect.CargaRecordSet(sSql)
Set GetMovExtorno = rsMov
End Function

Public Sub ActualizaMovDatos(ByVal pnMovNro As Long, Optional psMovNro As String = "", Optional pnMovEstado As MovEstado, Optional pnMovFlag As MovFlag)
Dim lsFiltro As String
If pnMovEstado <> -1 Then
    lsFiltro = " nMovEstado  = " & pnMovEstado & ","
End If
If psMovNro <> "" Then
    lsFiltro = lsFiltro & " cMovNro = '" & psMovNro & "',"
End If
If pnMovFlag <> -1 Then
    lsFiltro = lsFiltro & " nMovFlag = " & pnMovFlag & ","
End If
If lsFiltro <> "" Then
    lsFiltro = Left(lsFiltro, Len(lsFiltro) - 1)
    sSql = "UPDATE Mov SET " & lsFiltro & " WHERE nMovNro = " & pnMovNro
    oConect.Ejecutar sSql
End If
If pnMovEstado = gMovEstContabMovContable Then
    ActualizaSaldoMovimiento psMovNro, "+"
End If
End Sub

Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
sSql = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
oConect.Ejecutar sSql
End Sub
Public Function EsOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Boolean
Dim rsOP As Recordset
Set rsOP = New Recordset
rsOP.CursorLocation = adUseClient
sSql = "Select OP.cCtaCod, OP.nInicio, OP.nFin FROM Mov M INNER JOIN MovDocEmitidoRango OP ON " _
    & "M.nMovNro = OP.nMovNro WHERE M.nMovFlag <> " & gMovFlagExtornado & " AND OP.cCtaCod = '" _
    & sCuenta & "' And " & nNumOP & " Between OP.nInicio And OP.nFin"

Set rsOP = oConect.CargaRecordSet(sSql)
If rsOP.EOF And rsOP.BOF Then
    EsOrdenPagoEmitida = False
Else
    EsOrdenPagoEmitida = True
End If
rsOP.Close
Set rsOP = Nothing
End Function
Public Function ActualizaNotaAbonoCargo(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, Optional ByVal pnEstado As NotaCargoAbonoEstado = -1, _
                                    Optional ByVal pnMotivo As Integer = -1, Optional ByVal pnMonto As Currency = -999999, _
                                    Optional psObjetoCod As String = "", Optional psObjeto As String = "") As Integer
    On Error GoTo ActualizaNotaAbonoCargoErr
    Dim sql As String
    Dim lsActualiza As String
    ActualizaNotaAbonoCargo = 1
    lsActualiza = ""
    
    If pnMotivo <> -1 Then
        lsActualiza = lsActualiza + " nMotivoCod=" & pnMotivo & ", "
    End If
    If pnEstado <> -1 Then
        lsActualiza = lsActualiza + " nEstado=" & pnEstado & ","
    End If
    If pnMonto <> -999999 Then
        lsActualiza = lsActualiza + " nMonto=" & pnMonto & ","
    End If
    If psObjetoCod <> "" Then
        lsActualiza = lsActualiza + " cObjetoCodPadre= '" & psObjetoCod & "',"
    End If
    If psObjeto <> "" Then
        lsActualiza = lsActualiza + " cObjetoCod='" & psObjeto & "',"
    End If
    If lsActualiza <> "" Then
        lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
        sql = " UPDATE NotaAbonoCargo SET " & lsActualiza & " " _
            & " WHERE nDocTpo=" & pnTpoDoc & " AND cDocNro='" & psNroDoc & "'"
    End If
        
    oConect.Ejecutar sql
    ActualizaNotaAbonoCargo = 0
    Exit Function
ActualizaNotaAbonoCargoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DCapMov:ActualizaNotaAbonoCargo Method")
End Function
Public Function InsertaCuentaIF(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                ByVal psCtaIfCod As String, ByVal psCtaIFDesc As String, _
                                ByVal pdCtaIFAper As Date, ByVal psCtaIFCap As String, ByVal psCtaIFInt As String, _
                                ByVal pnCtaIFPlazo As Integer, ByVal pnCtaIFEstado As CGEstadoCtaIF, _
                                ByVal psUltAct As String) As Integer
    On Error GoTo InsertaCuentaIFErr
    Dim sql As String
    InsertaCuentaIF = 1
    
    sql = "INSERT INTO CTAIF( cPersCod, cIFTpo, cCtaIFCod, cCtaIFDesc, " _
        & "                dCtaIFAper, dCtaIFCap, nCtaIFPlazo, cCtaIFEstado, cUltimaActualizacion , dCtaIfInt ) " _
        & " VALUES('" & psPersCod & "','" & psIFTpo & "','" & psCtaIfCod & "','" & psCtaIFDesc & "','" _
        & Format(pdCtaIFAper, oVarPublicas.gsFormatoFecha) & "'," & IIf(psCtaIFCap = "", "Null", "'" & psCtaIFCap & "'") & "," _
        & pnCtaIFPlazo & "," & pnCtaIFEstado & ",'" & psUltAct & "'," & IIf(psCtaIFInt = "", "Null", "'" & psCtaIFInt & "'") & ")"
    
    oConect.Ejecutar sql
    InsertaCuentaIF = 0
    Exit Function
InsertaCuentaIFErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaCuentaIF Method")
End Function
Public Function InsertaCuentaIFInteres(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                ByVal psCtaIfCod As String, ByVal pdFechaInt As Date, ByVal pnPeriodo As Integer, _
                                ByVal pnInteres As Currency, ByVal psMovNro As String) As Integer
    On Error GoTo InsertaCuentaIFInteresErr
    Dim sql As String
    InsertaCuentaIFInteres = 1
    
    sql = "INSERT INTO CtaIFInteres( cPersCod, cIFTpo, cCtaIFCod, dCtaIFIntRegistro , nCtaIFIntPeriodo, nCtaIFIntValor, cUltimaActualizacion  ) " _
        & " VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIfCod & "','" _
        & Format(pdFechaInt, oVarPublicas.gsFormatoFecha) & "'," & pnPeriodo & "," & pnInteres & ",'" & psMovNro & "')"
    
    oConect.Ejecutar sql
    InsertaCuentaIFInteres = 0
    Exit Function
InsertaCuentaIFInteresErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaCuentaIFInteres Method")
End Function

Public Function InsertaInteresCtaIF(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                    ByVal psCtaIfCod As String, _
                                    ByVal pdCtaIFIntRegistro As Date, ByVal pnCtaIFIntPeriodo As Long, _
                                    ByVal pnCtaIFIntValor As Currency, ByVal psUltAct As String) As Integer
    On Error GoTo InsertaInteresCtaIFErr
    Dim sql As String
    InsertaInteresCtaIF = 1
    
    sql = " INSERT INTO CtaIfInteres  " _
        & " (cPersCod, cIFTpo, cCtaIFCod, dCtaIFIntRegistro, nCtaIFIntPeriodo, nCtaIFIntValor, cUltimaActualizacion) " _
        & " VALUES('" & psPersCod & "','" & psIFTpo & "','" & psCtaIfCod & "','" _
        & Format(pdCtaIFIntRegistro, oVarPublicas.gsFormatoFecha) & "'," & pnCtaIFIntPeriodo & "," _
        & pnCtaIFIntValor & ",'" & psUltAct & "')"
    
    oConect.Ejecutar sql
    InsertaInteresCtaIF = 0
    Exit Function
InsertaInteresCtaIFErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaInteresCtaIF Method")
End Function

Public Function InsertaCuentaIFFiltro(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                    ByVal psCtaIfCod As String, ByVal psCtaContCod As String, _
                                    ByVal psCtaIfSubCta As String) As Integer
                                    
    On Error GoTo InsertaCuentaIFFiltroErr
    Dim sql As String
    InsertaCuentaIFFiltro = 1
    
    sql = "INSERT INTO CtaIFFiltro (cPersCod, cIFTpo, cCtaIfCod, cCtaContCod, cCtaIFSubCta)" _
        & " VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIfCod & "','" _
        & psCtaContCod & "','" & psCtaIfSubCta & "')"
    
    oConect.Ejecutar sql
    InsertaCuentaIFFiltro = 0
    Exit Function
InsertaCuentaIFFiltroErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaCuentaIFFiltro Method")
End Function

'Public Function InsertaMovCtaIF(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
'                                 ByVal psCtaIfCod As String, ByVal pnConCod As CGCtaIFConceptos, ByVal pnImporte As Currency, Optional pnDiasTrans As Integer = 0) As Integer
'
'    On Error GoTo InsertaMovCtaIFErr
'    Dim sql As String
'    InsertaMovCtaIF = 1
'
'    sql = "INSERT INTO MovCtaIF(nMovNro, cPersCod, cIFTpo, cCtaIFCod, nConceptoCod, nMovImporte,nDiasTrans) " _
'        & "     VALUES(" & pnMovNro & ",'" & psPersCod & "','" & Format(pnIFTpo, "00") & "','" _
'        & psCtaIfCod & "'," & pnConCod & "," & pnImporte & "," & pnDiasTrans & ")"
'
'    oConect.Ejecutar sql
'    InsertaMovCtaIF = 0
'    Exit Function
'InsertaMovCtaIFErr:
'    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovCtaIF Method")
'End Function

'Public Function ExtornaMovCtaIF(ByVal pnMovNro As Long) As Integer
'Dim sql As String
'Dim rs As adodb.Recordset
'Set rs = New adodb.Recordset
'
'sql = " SELECT  M.nMovNro,CONVERT(DATETIME , SUBSTRING(M.CMOVNRO,1,8)) AS FECHA, " _
'    & "         MC.cPersCod , MC.cIFTpo, MC.cCtaIFCod " _
'    & " FROM MOV M JOIN MOVCTAIF  MC ON MC.NMOVNRO =M.NMOVNRO  " _
'    & " WHERE M.NMOVNRO = " & pnMovNro & ""
'
'Set rs = oConect.CargaRecordSet(sql)
'If Not rs.EOF And Not rs.BOF Then
'    Do While Not rs.EOF
'        ActualizaSaldoCtaIF rs!Fecha, rs!cPersCod, rs!cIFTpo, rs!cCtaIFCod
'        rs.MoveNext
'    Loop
'End If
'rs.Close
'Set rs = Nothing
'End Function

Public Function ActualizaSaldoCtaIF(ByVal pdFecha As Date, _
                                    ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
                                    ByVal psCtaIfCod As String) As Integer
                                    
    On Error GoTo ActualizaSaldoCtaIFErr
    Dim sql As String
    ActualizaSaldoCtaIF = 1
    
    sql = "sp_ActualizaSaldoCtaIF '" & Format(pdFecha, oVarPublicas.gsFormatoFecha) & "','" _
                                    & psPersCod & "','" _
                                    & Format(pnIFTpo, "00") & "', '" & psCtaIfCod & "'"
    
    oConect.Ejecutar sql
    ActualizaSaldoCtaIF = 0
    Exit Function
ActualizaSaldoCtaIFErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaSaldoCtaIF Method")
End Function
Public Function ActualizaCtaIF(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                  ByVal psCtaIfCod As String, _
                                  Optional ByVal psCtaIFDesc As String = "", _
                                  Optional ByVal pdCtaIFAper As String = "", _
                                  Optional ByVal pdCtaIFCap As String = "", _
                                  Optional ByVal pdCtaIFInt As String = "", _
                                  Optional ByVal pnCtaIFPlazo As Integer = -1, _
                                  Optional ByVal pnCtaIFEstado As CGEstadoCtaIF = -1, _
                                  Optional ByVal pnInteres As Currency = -999999, _
                                  Optional ByVal psUltAct As String = "") As Integer
                                  
    On Error GoTo ActualizaCuentaIFErr
    Dim sql As String
    Dim lsFiltro As String
    ActualizaCtaIF = 1
    lsFiltro = ""
    If psCtaIFDesc <> "" Then
        lsFiltro = lsFiltro + " cCtaIFDesc ='" & psCtaIFDesc & "',"
    End If
    If pdCtaIFAper <> "" Then
        lsFiltro = lsFiltro + " dCtaIFAper= '" & pdCtaIFAper & "',"
    End If
    If pdCtaIFCap <> "" Then
        lsFiltro = lsFiltro + " dCtaIFCap= '" & pdCtaIFCap & "',"
    End If
    If pdCtaIFInt <> "" Then
        lsFiltro = lsFiltro + " dCtaIFInt= '" & pdCtaIFInt & "',"
    End If
    If pnCtaIFPlazo <> -1 Then
        lsFiltro = lsFiltro + " nCtaIFPlazo =" & pnCtaIFPlazo & ","
    End If
    If pnCtaIFEstado <> -1 Then
        lsFiltro = lsFiltro + " cCtaIFEstado =" & pnCtaIFEstado & ","
    End If
    If pnInteres <> -999999 Then
        lsFiltro = lsFiltro + " nInteres =" & pnInteres & ", "
    End If
    If psUltAct <> "" Then
        lsFiltro = lsFiltro + " cUltimaActualizacion ='" & psUltAct & "',"
    End If
    If lsFiltro <> "" Then
        lsFiltro = Mid(lsFiltro, 1, Len(lsFiltro) - 1)
        sql = "UPDATE CTAIF SET " & lsFiltro _
            & " WHERE cPersCod = '" & psPersCod & "' AND cIFTpo='" & Format(psIFTpo, "00") & "' AND cCtaIFCod='" & psCtaIfCod & "'"
    End If
    oConect.Ejecutar sql
    ActualizaCtaIF = 0
    Exit Function
ActualizaCuentaIFErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaCuentaIF Method")
End Function


Public Sub InsertaEfectivoAutomatico(ByVal pdFecha As Date, ByVal pcCodAge As String, ByVal pcUser As String)

On Error GoTo InsertaEfectivoAutomaticoErr
    Dim sql As String
    
    sql = "INSERT INTO EfectivoAutomatico (dFecha, cCodAge, cUser, nFlag )" _
        & " VALUES('" & Format(pdFecha, oVarPublicas.gsFormatoFecha) & "','" & Trim(pcCodAge) & "','" & Trim(pcUser) & "', 0)"
     
    oConect.Ejecutar sql
    
    Exit Sub
InsertaEfectivoAutomaticoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaEfectivoAutomatico Method")

End Sub
'EJVG20140809 ***
Public Sub ActualizaSaldoMovimiento(sMovNro As String, sSimbolo As String)
    Dim sql As String
    On Error GoTo ErrActualizaMovSaldo
    Dim ldFecha As Date
 
    If Left(sMovNro, 1) = "'" Then
        Dim lsMov As String
        Do While sMovNro <> ""
            If InStr(sMovNro, ",") > 0 Then
                lsMov = Left(sMovNro, InStr(sMovNro, ",") - 1)
                sMovNro = Mid(sMovNro, InStr(sMovNro, ",") + 1)
            Else
                lsMov = sMovNro
                sMovNro = ""
            End If
            sql = "INSERT tmpMovActualiza (cMovNro) VALUES (" & lsMov & ")"
            oConect.Ejecutar sql
        Loop
        ldFecha = GetFechaMov(Mid(lsMov, 2, 8), True)
        sql = "Exec spActualizaMovSaldo " & lsMov & ", '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "', '" & sSimbolo & "' "
        oConect.Ejecutar sql
        
        sql = "DELETE tmpMovActualiza"
        oConect.Ejecutar sql
    Else
        sql = "INSERT tmpMovActualiza (cMovNro) VALUES ('" & sMovNro & "')"
        oConect.Ejecutar sql
        
        ldFecha = GetFechaMov(Left(sMovNro, 8), True)
        sql = "spActualizaMovSaldo '" & sMovNro & "', '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "', '" & sSimbolo & "' "
        'oConect.CommandTimeOut 3500
        oConect.Ejecutar sql
        
        sql = "DELETE tmpMovActualiza"
        oConect.Ejecutar sql
    End If
    
    Exit Sub
ErrActualizaMovSaldo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Function GetFechaMov(cMovNro, lDia As Boolean) As String
    Dim lFec As Date
    lFec = Mid(cMovNro, 7, 2) & "/" & Mid(cMovNro, 5, 2) & "/" & Mid(cMovNro, 1, 4)
    If lDia Then
       GetFechaMov = Format(lFec, oVarPublicas.gsFormatoFechaView)
    Else
       GetFechaMov = Format(lFec, oVarPublicas.gsFormatoFecha)
    End If
End Function
'END EJVG *******
'Public Sub ActualizaSaldoMovimiento(sMovNro As String, sSimbolo As String)
' Dim sql As String
' Dim prs As adodb.Recordset
' Dim ldFecha As Date
' Dim ldFecSis As Date
' Dim oGen As New COMDConstSistema.DCOMGeneral
 
' ldFecha = oGen.GetFechaMov(Left(sMovNro, 8), True)
' sql = "spActualizaMovSaldo '" & sMovNro & "', '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "', '" & sSimbolo & "' "
' oConect.Ejecutar sql
 
 'CTASALDO
' sql = "SELECT left(m.cMovNro,8) as fecha, mc.cCtaContcod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dCtaSaldoFecha, ISNULL(sdo.nCtaSaldoImporte,0) nCtaSaldoImporte, ISNULL(sdo.nCtaSaldoImporteME,0) nCtaSaldoImporteME " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%' " _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte, nCtaSaldoImporteME FROM CtaSaldo cs WHERE cs.dCtaSaldoFecha = (SELECT Max(dCtaSaldoFecha) FROM CtaSaldo where cCtaContCod = cs.cCtaContCod and dCtaSaldoFecha <= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro like '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cmovnro,8), mc.cCtaContCod, cls.cCtaCaracter, sdo.dCtaSaldoFecha, sdo.nCtaSaldoImporte, sdo.nCtaSaldoImporteME " _
'     & "ORDER BY left(m.cmovnro,8) "
' Set prs = oConect.CargaRecordSet(sql)
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     If sSimbolo = "+" Then
'        If ldFecha <> prs!dCtaSaldofecha Then
'           sql = "INSERT CtaSaldo (cCtaContCod, dCtaSaldoFecha, nCtaSaldoImporte, nCtaSaldoImporteME) " _
'               & "VALUES ('" & prs!cCtaContCod & "', '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "', " & prs!nCtaSaldoImporte & ", " & prs!nCtaSaldoImporteME & ")"
'           oConect.Ejecutar sql
'         End If
'     End If
'    sql = "UPDATE CtaSaldo SET nCtaSaldoImporte = nCtaSaldoImporte " & sSimbolo & "(" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'        & "nCtaSaldoImporteME = nCtaSaldoImporteME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'        & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and dCtaSaldoFecha >= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "'"
'    oConect.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
'
' 'CTAOBJSALDO
' sql = "SELECT left(m.cmovnro,8) as fecha, mc.cCtaContcod, mo.nMovObjOrden, mo.cObjetoCod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dCtaObjSaldoFecha, ISNULL(sdo.nCtaObjSaldoImporte,0) nCtaObjSaldoImporte, ISNULL(sdo.nCtaObjSaldoImporteME,0) nCtaObjSaldoImporteMe " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%'" _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "     JOIN MovObj mo ON mo.nMovNro = mc.nMovNro and mo.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, nCtaObjOrden, cObjetoCod, dCtaObjSaldoFecha, nCtaObjSaldoImporte, nCtaObjSaldoImporteME " _
'     & "                FROM CtaObjSaldo cs WHERE cs.dCtaObjSaldoFecha = (SELECT Max(dCtaObjSaldoFecha) FROM CtaObjSaldo WHERE cCtaContCod = cs.cCtaContCod and nCtaObjOrden = cs.nCtaObjOrden and cObjetoCod = cs.cObjetoCod and dCtaObjSaldoFecha <= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod and sdo.nCtaObjOrden = mo.nMovObjOrden and sdo.cObjetoCod = mo.cObjetoCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro like '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cMovNro,8), mc.cCtaContCod, mo.nMovObjOrden, mo.cObjetoCod, cls.cCtaCaracter, sdo.dCtaObjSaldoFecha, ISNULL(sdo.nCtaObjSaldoImporte,0), ISNULL(sdo.nCtaObjSaldoImporteME,0) " _
'     & "ORDER BY left(m.cMovNro,8) "
' Set prs = oConect.CargaRecordSet(sql)
'
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     ldFecha = GetFechaMov(prs!Fecha, True)
'     If sSimbolo = "+" Then
'        If prs!dCtaObjSaldoFecha <> ldFecha Then
'            sql = "INSERT CtaObjSaldo (cCtaContCod, nCtaObjOrden, cObjetoCod, dCtaObjSaldoFecha, nCtaObjSaldoImporte, nCtaObjSaldoImporteME) " _
'                & " VALUES ('" & prs!cCtaContCod & "','" & prs!nMovObjOrden & "', '" & prs!cObjetoCod & "', '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "', " & prs!nCtaObjSaldoImporte & "," & prs!nCtaObjSaldoImporteME & " )"
'            oConect.Ejecutar sql
'        End If
'     End If
'     sql = "UPDATE CtaObjSaldo SET nCtaObjSaldoImporte = nCtaObjSaldoImporte " & sSimbolo & " (" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'         & "nCtaObjSaldoImporteME = nCtaObjSaldoImporteME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'         & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and nCtaObjOrden = " & prs!nMovObjOrden & " and cObjetoCod = '" & prs!cObjetoCod & "' and dCtaObjSaldoFecha >= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "' "
'     oConect.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
'
' 'CtaObjEfectivoSaldo
' sql = "SELECT LEFT(m.cMovNro,8) as fecha, mc.cCtaContcod, moef.nMovObjOrden, moef.cEfectivoCod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dEfectivoSaldoFecha, ISNULL(sdo.nEfectivoSaldoImporte,0) nEfectivoSaldoImporte, ISNULL(sdo.nEfectivoSaldoImporteME,0) nEfectivoSaldoImporteMe " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%'" _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "          JOIN MovObjEfectivo moef ON moef.nMovNro = mc.nMovNro and moef.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, nCtaObjOrden, cEfectivoCod, dEfectivoSaldoFecha, nEfectivoSaldoImporte, nEfectivoSaldoImporteME " _
'     & "                FROM CtaObjEfectivoSaldo cs WHERE cs.dEfectivoSaldoFecha = (SELECT Max(dEfectivoSaldoFecha) FROM CtaObjEfectivoSaldo WHERE cCtaContCod = cs.cCtaContCod and nCtaObjOrden = cs.nCtaObjOrden and cEfectivoCod = cs.cEfectivoCod and dEfectivoSaldoFecha <= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod and sdo.nCtaObjOrden = moef.nMovObjOrden and sdo.cEfectivoCod = moef.cEfectivoCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro like '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cMovNro,8), mc.cCtaContCod, moef.nMovObjOrden, moef.cEfectivoCod, cls.cCtaCaracter, sdo.dEfectivoSaldoFecha, ISNULL(sdo.nEfectivoSaldoImporte,0), ISNULL(sdo.nEfectivoSaldoImporteME,0) " _
'     & "ORDER BY left(m.cMovNro,8) "
' Set prs = oConect.CargaRecordSet(sql)
'
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     ldFecha = GetFechaMov(prs!Fecha, True)
'     If sSimbolo = "+" Then
'        If prs!dEfectivoSaldoFecha <> ldFecha Then
'            sql = "INSERT CtaObjEfectivoSaldo (cCtaContCod, nCtaObjOrden, cEfectivoCod, dEfectivoSaldoFecha, nEfectivoSaldoImporte, nEfectivoSaldoImporteME) " _
'                & " VALUES ('" & prs!cCtaContCod & "','" & prs!nMovObjOrden & "', '" & prs!cEfectivoCod & "', '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "', " & prs!nEfectivoSaldoImporte & "," & prs!nEfectivoSaldoImporteME & " )"
'            oConect.Ejecutar sql
'        End If
'     End If
'     sql = "UPDATE CtaObjEfectivoSaldo SET nEfectivoSaldoImporte = nEfectivoSaldoImporte " & sSimbolo & " (" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'         & "nEfectivoSaldoImporteME = nEfectivoSaldoImporteME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'         & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and nCtaObjOrden = " & prs!nMovObjOrden & " and cEfectivoCod = '" & prs!cEfectivoCod & "' and dEfectivoSaldoFecha >= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "' "
'     oConect.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
'
' 'CtaObjAreaAgenciaSaldo
' sql = "SELECT LEFT(m.cMovNro,8) as fecha, mc.cCtaContcod, moef.nMovObjOrden, moef.cAreaCod, moef.cAgeCod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dAreaSaldoFecha, ISNULL(sdo.nAreaSaldoImporte,0) nAreaSaldoImporte, ISNULL(sdo.nAreaSaldoImporteME,0) nAreaSaldoImporteMe " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%'" _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "          JOIN MovObjAreaAgencia moef ON moef.nMovNro = mc.nMovNro and moef.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, nCtaObjOrden, cAreaCod, cAgeCod, dAreaSaldoFecha, nAreaSaldoImporte, nAreaSaldoImporteME " _
'     & "                FROM CtaObjAreaAgenciaSaldo cs WHERE cs.dAreaSaldoFecha = (SELECT Max(dAreaSaldoFecha) FROM CtaObjAreaAgenciaSaldo WHERE cCtaContCod = cs.cCtaContCod and nCtaObjOrden = cs.nCtaObjOrden and cAreaCod = cs.cAreaCod and cAgeCod = cs.cAgeCod and dAreaSaldoFecha <= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod and sdo.nCtaObjOrden = moef.nMovObjOrden and sdo.cAreaCod = moef.cAreaCod and sdo.cAgeCod = moef.cAgeCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro like '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cMovNro,8), mc.cCtaContCod, moef.nMovObjOrden, moef.cAreaCod, moef.cAgeCod, cls.cCtaCaracter, sdo.dAreaSaldoFecha, ISNULL(sdo.nAreaSaldoImporte,0), ISNULL(sdo.nAreaSaldoImporteME,0) " _
'     & "ORDER BY left(m.cMovNro,8) "
' Set prs = oConect.CargaRecordSet(sql)
'
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     ldFecha = GetFechaMov(prs!Fecha, True)
'     If sSimbolo = "+" Then
'        If prs!dAreaSaldoFecha <> ldFecha Then
'            sql = "INSERT CtaObjAreaAgenciaSaldo (cCtaContCod, nCtaObjOrden, cAreaCod, cAgeCod, dAreaSaldoFecha, nAreaSaldoImporte, nAreaSaldoImporteME) " _
'                & " VALUES ('" & prs!cCtaContCod & "','" & prs!nMovObjOrden & "', '" & prs!cAreaCod & "','" & prs!CAGECOD & "', '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "', " & prs!nAreaSaldoImporte & "," & prs!nAreaSaldoImporteME & " )"
'            oConect.Ejecutar sql
'        End If
'     End If
'     sql = "UPDATE CtaObjAreaAgenciaSaldo SET nAreaSaldoImporte = nAreaSaldoImporte " & sSimbolo & " (" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'         & "nAreaSaldoImporteME = nAreaSaldoImporteME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'         & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and nCtaObjOrden = " & prs!nMovObjOrden & " and cAreaCod = '" & prs!cAreaCod & "' and cAgeCod = '" & prs!CAGECOD & "' and dAreaSaldoFecha >= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "' "
'     oConect.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
'
' 'CtaIFSaldo
' sql = "SELECT LEFT(m.cMovNro,8) as fecha, mc.cCtaContcod, moif.cPersCod, moif.cIFTpo, moif.cCtaIFCod, cls.cCtaCaracter, SUM(mc.nMovImporte) as nImporte, ISNULL(SUM(me.nMovMeImporte),0) as nImporteME, sdo.dCtaIFSaldo, ISNULL(sdo.nSaldo,0) nSaldo, ISNULL(sdo.nSaldoME,0) nSaldoME " _
'     & "FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro JOIN CtaContClase cls ON mc.cCtaContCod LIKE RTRIM(cls.cCtaContCod) + '%'" _
'     & "     LEFT JOIN MovMe me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
'     & "          JOIN MovObjIF moif ON moif.nMovNro = mc.nMovNro and moif.nMovItem = mc.nMovItem " _
'     & "     LEFT JOIN (SELECT cCtaContCod, cPersCod, cIFTpo, cCtaIFCod, dCtaIFSaldo, nSaldo, nSaldoME " _
'     & "                FROM CtaIFSaldo cs WHERE cs.dCtaIFSaldo = (SELECT Max(dCtaIFSaldo) FROM CtaIFSaldo WHERE cCtaContCod = cs.cCtaContCod and cPersCod = cs.cPersCod and cIFTpo = cs.cIFTpo and cCtaIFCod = cs.cCtaIFCod and dCtaIFSaldo <= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "') " _
'     & "               ) sdo ON sdo.cCtaContCod = mc.cCtaContCod and sdo.cPersCod = moif.cPersCod and sdo.cIFTpo = moif.cIFTpo and sdo.cCtaIFCod = moif.cCtaIFCod " _
'     & "WHERE m.nMovEstado = " & gMovEstContabMovContable & " and m.cMovNro LIKE '" & sMovNro & "%' and mc.cCtaContCod <> '' " _
'     & "GROUP BY left(m.cMovNro,8), mc.cCtaContCod, moif.cPersCod, moif.cIFTpo, moif.cCtaIFCod, cls.cCtaCaracter, sdo.dCtaIFSaldo, ISNULL(sdo.nSaldo,0), ISNULL(sdo.nSaldoME,0) " _
'     & "ORDER BY left(m.cMovNro,8) "
' Set prs = oConect.CargaRecordSet(sql)
' If Not prs.EOF Then
'   Do While Not prs.EOF
'     ldFecha = GetFechaMov(prs!Fecha, True)
'     If sSimbolo = "+" Then
'        If prs!dCtaIFSaldo <> ldFecha Then
'            sql = "INSERT CtaIFSaldo (cCtaContCod, cPersCod, cIFTpo, cCtaIFCod, dCtaIFSaldo, nSaldo, nSaldoME) " _
'                & " VALUES ('" & prs!cCtaContCod & "','" & prs!cPersCod & "','" & prs!cIFTpo & "', '" & prs!cCtaIFCod & "', '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "', " & prs!nSaldo & "," & prs!nSaldoME & " )"
'            oConect.Ejecutar sql
'        End If
'     End If
'     sql = "UPDATE CtaIFSaldo SET nSaldo = nSaldo " & sSimbolo & " (" & prs!nImporte * IIf(prs!cCtaCaracter = "A", -1, 1) & "), " _
'         & "nSaldoME = nSaldoME " & sSimbolo & "(" & prs!nImporteME * IIf(prs!cCtaCaracter = "A", -1, 1) & ") " _
'         & "WHERE  cCtaContCod = '" & prs!cCtaContCod & "' and cPersCod = '" & prs!cPersCod & "' and cIFTpo = '" & prs!cIFTpo & "' and cCtaIFCod = '" & prs!cCtaIFCod & "' and dCtaIFSaldo >= '" & Format(ldFecha, oVarPublicas.gsFormatoFecha) & "' "
'     oConect.Ejecutar sql
'     prs.MoveNext
'   Loop
' End If
' RSClose prs
'
'End Sub

'Public Sub InsertaDatosAdeudados(prsAdeud As ADODB.Recordset, psPersCod As String, psIFTpo As String, psCodAdeudado As String, pnCapital As Currency, pnNroCuotas As Integer, pnGracia As Currency, pnComisionIni As Currency, psInterno As String, pnCuotaPagoK As Currency, _
'                psFecha As String, pnMoneda As Integer)
'    Dim lbExiste As Boolean
'    Dim lsTipoCuota As String
'    Dim oDAdeudCal As New DAdeudCal
'    oDAdeudCal.Inicio oConect.ConexionActiva
'    lbExiste = oDAdeudCal.VerificaCtaIFAdeudado(psPersCod, psIFTpo, psCodAdeudado)
'    If lbExiste Then
'        oDAdeudCal.EliminarCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado
'    End If
'
'    oDAdeudCal.InsertaCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado, pnCapital, pnCapital, pnNroCuotas, _
'                pnGracia, pnComisionIni, psInterno, pnCuotaPagoK, _
'                psFecha, "1", pnMoneda, psFecha
'
'    'GRABAMOS EL CALENDARIO DE LOS ADEUDADOS
'    If Not prsAdeud Is Nothing Then
'        If Not prsAdeud.EOF Then
'            'GRABAMOS LOS POSIBLES PAGOS A REALIZARSE O PAGOS PROYECTADOS
'            Do While Not prsAdeud.EOF
'                'TipoCuota  : ---> A  Cuota Adelantada  ----> C Cuota Normal
'                lsTipoCuota = IIf(prsAdeud.Fields(1) = 0, "A", "C")
'                oDAdeudCal.InsertCalendario psPersCod, psIFTpo, psCodAdeudado, "A", prsAdeud.Fields(1), prsAdeud.Fields(0), "P", 0, prsAdeud.Fields(2), prsAdeud.Fields(3), 0
'                If prsAdeud.Fields(1) = 1 Then 'Si es primera Cuota
'                   oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado, pnCapital, pnCapital, _
'                                                  pnNroCuotas, pnGracia, pnComisionIni, psInterno, pnCuotaPagoK, _
'                                                    Format(prsAdeud.Fields(0), oVarPublicas.gsFormatoFecha), "1", pnMoneda
'                End If
'                prsAdeud.MoveNext
'            Loop
'        End If
'    End If
'    Set oDAdeudCal = Nothing
'End Sub

'Public Sub ActualizaAdeudadosProvision(psPersCod As String, psIFTpo As String, psCtaIfCod As String, psFecha As String, pnInteresPag As Currency, pnNroCuota As Integer, psMovNro As String)
'Dim oDAdeudCal As New DAdeudCal
'    oDAdeudCal.Inicio oConect.ConexionActiva
'    oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCtaIfCod, , , , , , , , , , , psFecha
'    oDAdeudCal.UpdateCalendario psPersCod, psIFTpo, psCtaIfCod, pnNroCuota, pnInteresPag, psMovNro
'    Set oDAdeudCal = Nothing
'End Sub
            
'Public Sub ActualizaAdeudadosPagoCuota(psPersCod As String, psIFTpo As String, psCtaIfCod As String, psFecha As String, pnCapital As Currency, pnInteresPag As Currency, pnNroCuota As Integer, psMovNro As String, Optional pbCancela As Boolean = False)
'Dim oDAdeudCal As New DAdeudCal
'    oDAdeudCal.Inicio oConect.ConexionActiva
'    oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCtaIfCod, , pnCapital, , , , , , , , , psFecha
'    oDAdeudCal.UpdateCalendario psPersCod, psIFTpo, psCtaIfCod, pnNroCuota, pnInteresPag, psMovNro, gTpoEstCuotaAdeudCanc
'Set oDAdeudCal = Nothing
'End Sub
      

Public Sub GrabaMovimientoEfectivo(ByVal pnMovNro As Long, ByRef pnMovItem As Long, rsEfectivo As ADODB.Recordset, psCuenta As String, psDH As String)
On Error GoTo GrabaMovimientoEfectivoErr
Dim lnMovOrden As Integer
'Guardamos el Billetaje
If Not rsEfectivo Is Nothing Then
    If rsEfectivo.State = adStateOpen Then
        If Not rsEfectivo.EOF And Not rsEfectivo.BOF Then
            Do While Not rsEfectivo.EOF
                If rsEfectivo!Monto <> 0 Then
                    pnMovItem = pnMovItem + 1: lnMovOrden = 0
                    InsertaMovCta pnMovNro, Format(pnMovItem, "#0"), psCuenta, rsEfectivo!Monto * IIf(psDH = "D", 1, -1)
                    lnMovOrden = lnMovOrden + 1
                    InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, ObjDescomEfectivo
                    InsertaMovObjEfectivo pnMovNro, pnMovItem, lnMovOrden, rsEfectivo!cEfectivoCod
                End If
                rsEfectivo.MoveNext
            Loop
        End If
        rsEfectivo.Close: Set rsEfectivo = Nothing
    End If
End If
Exit Sub
GrabaMovimientoEfectivoErr:
    Err.Raise Err.Number, "DMov: GrabaMovimientoEfectivo", Err.Description
End Sub

Public Sub GrabaMovOtrosCtaObj(pnMovNro As Long, pnMovItem As Long, prsOtro As ADODB.Recordset, prsObj As ADODB.Recordset, psDH As String)
Dim lnMovOrden As Long
Dim lsFiltro   As String

If Not prsOtro Is Nothing Then
    If prsOtro.State = adStateOpen Then
        Do While Not prsOtro.EOF
            pnMovItem = pnMovItem + 1: lnMovOrden = 0
            lsFiltro = ""
            If Not prsObj Is Nothing Then
                If prsObj.State = adStateOpen Then
                    Do While Not prsObj.EOF
                        If prsObj![#] = prsOtro![#] Then
                            lsFiltro = lsFiltro & prsObj!SubCta
                        End If
                        prsObj.MoveNext
                    Loop
                End If
            End If
            InsertaMovCta pnMovNro, pnMovItem, prsOtro!Cuenta & lsFiltro, prsOtro!Importe * IIf(psDH = "D", 1, -1)
            lsFiltro = ""
            If Not prsObj Is Nothing Then
                prsObj.MoveFirst
                lnMovOrden = 0
                Do While Not prsObj.EOF
                    If prsObj.Fields(0) = prsOtro.Fields(0) Then
                        Select Case prsObj!Objpadre
                            Case ObjCMACAgencias
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjAgenciaArea pnMovNro, pnMovItem, lnMovOrden, prsObj!Código, ""
                            Case ObjCMACArea
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjAgenciaArea pnMovNro, pnMovItem, lnMovOrden, "", prsObj!Código
                            Case ObjCMACAgenciaArea
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjAgenciaArea pnMovNro, pnMovItem, lnMovOrden, Mid(prsObj!Código, 4, 2), Mid(prsObj!Código, 1, 3)
                            Case ObjDescomEfectivo
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjEfectivo pnMovNro, pnMovItem, lnMovOrden, prsObj!Código
                            Case ObjEntidadesFinancieras
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Objpadre
                                InsertaMovObjIF pnMovNro, pnMovItem, lnMovOrden, Mid(prsObj!Código, 4, 13), Mid(prsObj!Código, 1, 2), Mid(prsObj!Código, 18, 10)
                            Case Else
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!Código
                        End Select
                    End If
                    prsObj.MoveNext
                Loop
            End If
            prsOtro.MoveNext
        Loop
    End If
End If

End Sub


Public Function ObtieneMovNroRef(ByVal pnMovNro As Long) As Long

    Dim lrst As ADODB.Recordset
    
    sSql = "Select nMovNro from MovRef where nMovNroRef =" & pnMovNro
    Set lrst = oConect.CargaRecordSet(sSql)
    
    If Not lrst.EOF Then
        ObtieneMovNroRef = lrst("nMovNro")
    End If
        
End Function

Public Function ObtenerMovimientosAculuados(ByVal fechaini As String, ByVal fechafin As String) As ADODB.Recordset
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = " Select TpoDoc,CodPers ,Cdoc=rtrim(isnull(Doc,'')), AFec , Cta , TpoOpe, MBase='', Monto=convert(varchar(20),Convert(decimal(18,2),Monto)), '' As CodPais  from "
    sql = sql & "    ( "
    sql = sql & " Select CodPers=ltrim(rtrim(CodPers)),Case tpo when 1 then '01' when 2 then '06' when 3 then '03' when 4 then '04' when 10 then '11' when 11 then '07' else '00' end TpoDoc, Doc, '01' AFec, '04' Cta, TpoOpe, Sum(monto) Monto   from"
    sql = sql & " (Select CodPers , Sum(nMontDet) Monto,TpoOpe, isnull(bb.cPersIDnro, isnull(   cc.cPersIDnro,isnull(  dd.cPersIDnro, isnull(  ee.cPersIDnro,isnull(  ff.cPersIDnro,isnull(gg.cPersIDnro,ii.cPersIDnro)))))) Doc ,  nperspersoneria, isnull(bb.cPersIDTpo, isnull(   cc.cPersIDTpo,isnull(  dd.cPersIDTpo, isnull(  ee.cPersIDTpo,isnull(ff.cPersIDTpo,isnull(gg.cPersIDTpo,ii.cPersIDTpo))))))  Tpo"
    sql = sql & " from ("
    sql = sql & "    SELECT  MC.CCTACOD, (SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 20) CodPers ,  Case Right(MC.COPECOD,2) When '01' Then '13' else '15' End TpoOpe, SUM(Case Substring(MC.CCTACOD,9,1) when '2' then  round(dbo.GetTCPonderado(LEFT(M.CMOVNRO,8)) * MD.nMonto,2) else round(MD.nMonto,2) end) as nMontDet"
    sql = sql & "    FROM    MOV M"
    sql = sql & "        JOIN MOVCOL MC ON MC.NMOVNRO = M.NMOVNRO"
    sql = sql & "        JOIN MOVCOLDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD"
    sql = sql & "        JOIN OPETPO O ON O.COPECOD = MC.COPECOD"
    sql = sql & "        JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nPrdConceptoCod"
    sql = sql & "    WHERE   MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND M.NMOVFLAG = 0"
    sql = sql & "    GROUP BY MC.CCTACOD,MC.COPECOD"
    sql = sql & "    Union"
    sql = sql & "    SELECT  MC.CCTACOD, (SELECT max(cperscod) FROM ProductoPersona pp where pp.cctacod = MC.cctacod And nPrdPersRelac = 10) CodPers,  Case Right(MC.COPECOD,2) When '01' Then '13' else '15' End TpoOpe, SUM(Case Substring(MC.CCTACOD,9,1) when '2' then  round(dbo.GetTCPonderado(LEFT(M.CMOVNRO,8)) * MD.nMonto,2) else round(MD.nMonto,2) end) as nMontDet"
    sql = sql & "    FROM    MOV M"
    sql = sql & "        JOIN MOVCAP MC ON MC.NMOVNRO = M.NMOVNRO"
    sql = sql & "        JOIN MOVCAPDET MD ON MD.NMOVNRO = MC.NMOVNRO AND MD.COPECOD = MC.COPECOD AND MD.CCTACOD = MC.CCTACOD"
    sql = sql & "        JOIN OPETPO O ON O.COPECOD = MC.COPECOD"
    sql = sql & "        JOIN PRODUCTOCONCEPTO PC ON PC.nPrdConceptoCod = MD.nConceptoCod"
    sql = sql & "    WHERE   MC.COPECOD LIKE '99%' AND LEFT(M.CMOVNRO,8) BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND M.NMOVFLAG = 0"
    'sql = sql & "        AND NOT M.COPECOD IN (100103, 100104)"
    sql = sql & "    GROUP BY MC.CCTACOD,MC.COPECOD"
    sql = sql & "    Union"
    'Sql = Sql & "    SELECT  '1089800000272' cCtaCod, '1089800000272' c, SUM(CASE WHEN SUBSTRING(MC.cCtaContCod,3,1) ='2' THEN ABS(ME.NMOVMEIMPORTE) ELSE ABS(MC.nMovImporte) END) as nMontDet"
    sql = sql & "   SELECT  '1090100012521' cCtaCod, '1090100012521' CodPers, '15' as TpoOpe, SUM(CASE WHEN SUBSTRING(MC.cCtaContCod,3,1) ='2' THEN round(ABS(ME.NMOVMEIMPORTE)*dbo.GetTCPonderado( LEFT(M.CMOVNRO,8)),2) ELSE round(ABS(MC.nMovImporte),2) END)  as nMontDet "
    sql = sql & "    FROM    MOV M"
    sql = sql & "        JOIN MOVCTA MC ON MC.NMOVNRO = M.NMOVNRO"
    sql = sql & "        LEFT JOIN MOVME ME ON ME.NMOVNRO = MC.NMOVNRO and ME.nMovItem = MC.nMovItem"
    sql = sql & "        JOIN OPETPO O ON O.COPECOD = M.COPECOD"
    sql = sql & "    WHERE   LEFT(M.CMOVNRO,8) BETWEEN '" & fechaini & "' AND '" & fechafin & "' AND M.NMOVFLAG = 0 AND M.COPECOD Not In ('701105')"
    sql = sql & "        AND MC.cCtaContCod LIKE '21_40207%' and MC.nMovImporte<0 ) As aaa"
    sql = sql & " left join persona pe on aaa.codpers =  pe.cperscod"
    sql = sql & " left join persid bb on aaa.codpers =  bb.cperscod and bb.cPersIDTpo = 1 " '--01
    sql = sql & " left join persid cc on aaa.codpers =  cc.cperscod and cc.cPersIDTpo = 2 " '--06
    sql = sql & " left join persid dd on aaa.codpers =  dd.cperscod and dd.cPersIDTpo = 4 " '--04
    sql = sql & " left join persid ee on aaa.codpers =  ee.cperscod and ee.cPersIDTpo = 11 " '--07
    sql = sql & " left join persid ff on aaa.codpers =  ff.cperscod and ff.cPersIDTpo = 10 " '--11
    sql = sql & " left join persid gg on aaa.codpers =  gg.cperscod and gg.cPersIDTpo = 3 "
    sql = sql & " left join persid ii on aaa.codpers =  ii.cperscod and ii.cPersIDTpo Like '1[35]' "
    sql = sql & "    Group by  CodPers,TpoOpe, isnull(bb.cPersIDnro, isnull(   cc.cPersIDnro,isnull(  dd.cPersIDnro, isnull(  ee.cPersIDnro,isnull(  ff.cPersIDnro,isnull( gg.cPersIDnro,ii.cPersIDnro)) )  )   )   ) ,  nperspersoneria , isnull(bb.cPersIDTpo, isnull(   cc.cPersIDTpo,isnull(  dd.cPersIDTpo, isnull(  ee.cPersIDTpo,isnull(  ff.cPersIDTpo,isnull( gg.cPersIDTpo,ii.cPersIDTpo))))))"
    sql = sql & " ) as aaa "
    'Sql = Sql & " where tpo in (1,2,4,10,11)"
    sql = sql & " group by Case tpo when 1 then '01' when 2 then '06' when 3 then '03' when 4 then '04' when 10 then '11' when 11 then '07' else '00' end ,CodPers , Doc,TpoOpe"
    sql = sql & " ) as bbb    "


    
    Set rs = oConect.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        Set ObtenerMovimientosAculuados = rs
    Else
        Set ObtenerMovimientosAculuados = Nothing
    End If
End Function

Public Sub InsertarArchMov(ByVal psMes As String, ByVal pscodpers As String, ByVal pstpodoc As String, ByVal pscod As String, ByVal psafec As String, ByVal pscta As String, ByVal pstpoope As String, ByVal psmonto As Double)
    Dim sql As String
    sql = "Insert ArchMovimientos values ('" & psMes & "','" & pscodpers & "','" & pstpodoc & "','" & pscod & "','" & psafec & "','" & pscta & "','" & pstpoope & "','','" & psmonto & "')"
    oConect.Ejecutar sql
    oConect.CierraConexion
End Sub


Public Function ObetenerMovRetCancAutoriza(ByVal psCtaCod As String, ByVal psOpeTpo As String, ByVal pdFecha As Date, ByVal pnMonto As Currency) As Long
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    sql = " SELECT nMovNro  FROM MovRetCancAutoriza " & _
          " WHERE cCtaCod ='" & psCtaCod & "' AND cOpeTpo ='" & psOpeTpo & "' AND nMonto=" & pnMonto & " AND" & _
          " DateDiff(Day,dFecha,'" & Format(pdFecha, "mm/dd/yyyy") & "') = 0"
    Set rs = oConect.CargaRecordSet(sql)
    If Not (rs.EOF And rs.BOF) Then
        ObetenerMovRetCancAutoriza = rs!nMovNro
    Else
        ObetenerMovRetCancAutoriza = 0
    End If
    oConect.CierraConexion
End Function

'**DAOR 20080125
'**Función para guardar datos maestros del registro de efectivo
Public Function InsertaMovRegEfectivo(ByVal pnMovNro As Long, ByVal psUser As String, ByVal pnNumRegEfec As Integer, ByVal pnIngresosMN As Currency, ByVal pnEgresosMN As Currency, ByVal pnEfectivoMN As Currency, ByVal pnSobrFaltMN As Currency, _
    ByVal pnIngresosME As Currency, ByVal pnEgresosME As Currency, ByVal pnEfectivoME As Currency, ByVal pnSobrFaltME As Currency) As Integer
Dim sql As String
On Error GoTo InsertaMovRegEfectivoErr
    InsertaMovRegEfectivo = 1
     
    sql = " exec stp_ins_MovRegEfectivo " & pnMovNro & ",'" & psUser & "'," & pnNumRegEfec & "," & pnIngresosMN & "," & pnEgresosMN & "," & pnEfectivoMN & "," & pnSobrFaltMN & "," & pnIngresosME & "," & pnEgresosME & "," & pnEfectivoME & "," & pnSobrFaltME

    oConect.Ejecutar sql
    InsertaMovRegEfectivo = 0
    Exit Function
InsertaMovRegEfectivoErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovUserEfectivo Method")
End Function
'MADM 20110125
Public Function InsertaMovRegEfectivoPreCuadre(ByVal pnMovNro As Long, ByVal psUser As String, ByVal pnNumRegEfec As Integer, ByVal pnIngresosMN As Currency, ByVal pnEgresosMN As Currency, ByVal pnEfectivoMN As Currency, ByVal pnSobrFaltMN As Currency, _
    ByVal pnIngresosME As Currency, ByVal pnEgresosME As Currency, ByVal pnEfectivoME As Currency, ByVal pnSobrFaltME As Currency) As Integer
Dim sql As String
On Error GoTo InsertaMovRegEfectivoPreCuadreErr
    InsertaMovRegEfectivoPreCuadre = 1
     
    sql = " exec stp_ins_MovRegEfectivoPreCuadre " & pnMovNro & ",'" & psUser & "'," & pnNumRegEfec & "," & pnIngresosMN & "," & pnEgresosMN & "," & pnEfectivoMN & "," & pnSobrFaltMN & "," & pnIngresosME & "," & pnEgresosME & "," & pnEfectivoME & "," & pnSobrFaltME

    oConect.Ejecutar sql
    InsertaMovRegEfectivoPreCuadre = 0
    Exit Function
InsertaMovRegEfectivoPreCuadreErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovUserEfectivo Method")
End Function
'END MADM

'ALPA 20091117********************************************
Public Function GetCtaAbono(ByVal pnMovNro As Long, ByVal pnCtaCodCargo As String) As ADODB.Recordset
Dim sql As String
On Error GoTo ErrInsertaCtaAbono
    sql = "exec stp_sel_ObtenerCtaAbonoTransferencia '" & pnCtaCodCargo & "'," & pnMovNro
    Set GetCtaAbono = oConect.CargaRecordSet(sql)
ErrInsertaCtaAbono:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovUserEfectivo Method")
End Function
'*********************************************************
'BRGO 20110914 **************************************************
Public Sub InsertaMovRedondeoITF(ByVal psMovNro As String, ByVal pnMovItem As Integer, ByVal pnITFCalculado As Currency, ByVal pnITFRedondeado As Currency, Optional ByVal pnMovNro As Long = 0)
Dim sql As String
Dim lnMovNro As Long
On Error GoTo InsertaMovRedondeoITFErr
    If pnMovNro > 0 Then
        lnMovNro = pnMovNro
    Else
        lnMovNro = GetnMovNro(psMovNro)
    End If
    
    sql = " exec stp_ins_MovRedondeoITF " & lnMovNro & "," & pnMovItem & "," & pnITFCalculado & "," & pnITFRedondeado
    oConect.Ejecutar sql
    Exit Sub
InsertaMovRedondeoITFErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovUserEfectivo Method")
End Sub
'END BRGO ********

'** Juez 20120723 ***********************************************************
Public Function ObtieneMovimientosOpeUsuario(ByVal psUser As String, ByVal pdFecha As Date) As ADODB.Recordset
Dim sql As String
On Error GoTo ErrObtieneMovimientosOpeUsuario
    sql = "exec stp_sel_ObtieneMovimientosOpeUsuarioxDia '" & psUser & "','" & Format(pdFecha, "yyyymmdd") & "'"
    Set ObtieneMovimientosOpeUsuario = oConect.CargaRecordSet(sql)
ErrObtieneMovimientosOpeUsuario:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ObtieneMovimientosOpeUsuario Method")
End Function
'** End Juez ****************************************************************

'MIOL 20130518, SEGUN RQ13521 **********************************
Public Function InsertaMovCambioGirador(ByVal pnMovNro As String, _
                            ByVal psMovNro As String, ByVal psNroDoc As String, _
                            ByVal psGirActual As String, _
                            ByVal psGirNuevo As String)
On Error GoTo InsertaMovCambioGiradorErr
    Dim sql As String
    
    sql = "INSERT INTO MovCambioGirador (nMovNro, cMovNro, cNroDoc,cPersCodGiradorAct,cPersCodGiradorNue)" _
        & " VALUES('" & pnMovNro & "','" & psMovNro & "','" & psNroDoc & "','" _
        & psGirActual & "','" & psGirNuevo & "')"
     
    oConect.Ejecutar sql
    
    Exit Function
InsertaMovCambioGiradorErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovCambioGirador Method")
End Function
'END MIOL ******************************************************
'EJVG20131109 ***
Public Sub registrarSeguimientoHabiltacionAgenciaCajaGeneral(ByVal psMovNroHab As String)
    Dim sql As String
    On Error GoTo ErrregistrarSeguimientoHabiltacionAgenciaCajaGeneral
    sql = "Exec stp_ins_RegistraSeguimientoHabiltacionAgenciaCajaGeneral '" & psMovNroHab & "'"
    oConect.Ejecutar sql
    Exit Sub
ErrregistrarSeguimientoHabiltacionAgenciaCajaGeneral:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:registrarSeguimientoHabiltacionAgenciaCajaGeneral Method")
End Sub
Public Sub actualizarSeguimientoHabiltacionAgenciaCajaGeneral(ByVal psMovNroHab As String, _
                                                              ByVal pnTipo As Integer, _
                                                              Optional ByVal psMovNroCofHab As String = "", _
                                                              Optional ByVal pbGenCofHab As Boolean = False, _
                                                              Optional ByVal psMovNroDepEF As String = "", _
                                                              Optional ByVal pbGenDepEF As Boolean = False)
    Dim sql As String
    On Error GoTo ErractualizarSeguimientoHabiltacionAgenciaCajaGeneral
    sql = "Exec stp_upd_ActualizaSeguimientoHabiltacionAgenciaCajaGeneral '" & psMovNroHab & "','" & psMovNroCofHab & "'," & IIf(pbGenCofHab = False, 0, 1) & ",'" & psMovNroDepEF & "'," & IIf(pbGenDepEF = False, 0, 1) & "," & pnTipo & ""
    oConect.Ejecutar sql
    Exit Sub
ErractualizarSeguimientoHabiltacionAgenciaCajaGeneral:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:actualizarSeguimientoHabiltacionAgenciaCajaGeneral Method")
End Sub
'END EJVG *******
'EJVG20140408 ***
Public Function InsertaChequeCabecera(ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCodIF As String, _
                                    ByVal psIFTpo As String, ByVal psIFCta As String, ByVal pnMonto As Currency, ByVal pdRegistro As Date, ByVal pdValorizacion As Date, _
                                    ByVal pnMoneda As Integer, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal psUser As String, _
                                    ByVal psPersCodGirador As String, ByVal psPersCodContacto As String, _
                                    ByVal pnConfCaja As Integer, ByVal pbGerencia As Boolean, ByVal psNroCheque1 As String, ByVal psNroCheque2 As String, ByVal psNroCheque3 As String, ByVal psNroCheque4 As String, ByVal pbDepositoEnRegistro As Boolean) As Long
    Dim lsSQL As String
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrInsertaChequeCabecera
    lsSQL = "Exec stp_ins_ERS1262013_InsertaChequeCabecera " & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCodIF & "','" & Format(psIFTpo, "00") & "','" & psIFCta & "'," & pnMonto & ",'" & Format(pdRegistro, "yyyymmdd") & "','" & Format(pdValorizacion, "yyyymmdd") & "'," & pnMoneda & ",'" & psAreaCod & "','" & psAgeCod & "','" & psUser & "','" & psPersCodGirador & "','" & psPersCodContacto & "'," & pnConfCaja & "," & IIf(pbGerencia, 1, 0) & ",'" & psNroCheque1 & "','" & psNroCheque2 & "','" & psNroCheque3 & "','" & psNroCheque4 & "'," & IIf(pbDepositoEnRegistro, 1, 0)
    Set rs = oConect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        InsertaChequeCabecera = rs!nId
    End If
    Set rs = Nothing
    Exit Function
ErrInsertaChequeCabecera:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaChequeCabecera Method")
End Function
Public Sub InsertaChequeRel(ByVal pnId As Long, ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCodIF As String, ByVal psIFTpo As String, ByVal psCtaIF As String)
    Dim lsSQL As String
    On Error GoTo ErrInsertaChequeRel
    lsSQL = "Exec stp_ins_ERS1262013_InsertaChequeRel " & pnId & "," & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCodIF & "','" & Format(psIFTpo, "00") & "','" & psCtaIF & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaChequeRel:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaChequeRel Method")
End Sub
Public Sub ActualizaChequexCab(ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCodIF As String, ByVal psIFTpo As String, ByVal psCtaIF As String, _
                                Optional ByVal pnTpoOperacion As Integer = -1, Optional ByVal psPersCodLista As String = "", Optional ByVal psCtaCodLista As String = "", Optional ByVal pnNroCliLote As Integer = -1, _
                                 Optional ByVal pbEsChequexAjuste As Boolean = False, Optional ByVal pnMonto As Currency = 0)
                                'PASI20161212 CCE agrego pbEsChequexAjuste,nMonto
    Dim lsSQL As String
    Dim lsActualiza As String
    On Error GoTo errActualizaChequexCab
    If pnTpoOperacion <> -1 Then
        lsActualiza = lsActualiza & " DR.nTipoOperacion = " & pnTpoOperacion & ","
    End If
    If psPersCodLista <> "" Then
        lsActualiza = lsActualiza & " DR.cPersCodDet = '" & psPersCodLista & "',"
    End If
    If psCtaCodLista <> "" Then
        lsActualiza = lsActualiza & " DR.cCtaCodDet = '" & psCtaCodLista & "',"
    End If
    If pnNroCliLote <> -1 Then
        lsActualiza = lsActualiza & " DR.nNroCliLote = " & pnNroCliLote & ","
    End If
    'PASI20161212 CCE****************************************
    If pbEsChequexAjuste Then
        lsActualiza = lsActualiza & " DR.nMonto = " & pnMonto & ","
    End If
    'PASI END************************************************
    If Len(lsActualiza) > 0 Then
        lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
        lsSQL = "UPDATE DR SET " & lsActualiza & " FROM DocRec DR WHERE nTpoDoc = " & pnTpoDoc & " AND cNroDoc = '" & psNroDoc & "' AND cPersCod = '" & psPersCodIF & "' AND cIFTpo = '" & Format(psIFTpo, "00") & "' AND cIFCta = '" & psCtaIF & "'"
        oConect.Ejecutar (lsSQL)
    End If
    Exit Sub
errActualizaChequexCab:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaChequexCab Method")
End Sub
'Public Function ActualizaChequeCabecera(ByVal pnId As Long, ByVal psPersCodGirador As String, ByVal psPersCodContacto As String) As Long /**Comentado PASI20161212 CCE**/
Public Function ActualizaChequeCabecera(ByVal pnId As Long, ByVal psPersCodGirador As String, ByVal psPersCodContacto As String, ByVal pnMonto As Currency) As Long 'PASI20161212 CCE
    Dim lsSQL As String
    On Error GoTo ErrActualizaChequeCabecera
    lsSQL = "Exec stp_upd_ERS1262013_ActualizaChequeCabecera " & pnId & ",'" & psPersCodGirador & "','" & psPersCodContacto & "'," & pnMonto
    oConect.Ejecutar lsSQL
    Exit Function
ErrActualizaChequeCabecera:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaChequeCabecera Method")
End Function
Public Function ActualizaDatosMovimientoCheque(ByVal pnTpoDoc As Integer, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String, ByVal psGlosa As String) As Long
    Dim lsSQL As String
    On Error GoTo ErrActualizaDatosMovimientoCheque
    lsSQL = "Exec stp_upd_ERS1262013_ActualizaDatosMovimientoCheque " & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psIFCta & "','" & psGlosa & "'"
    oConect.Ejecutar lsSQL
    Exit Function
ErrActualizaDatosMovimientoCheque:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaDatosMovimientoCheque Method")
End Function
Public Sub ExtornarCheque(ByVal pnId As Long, ByVal pnMovNroExtorno As Long)
    Dim lsSQL As String
    On Error GoTo ErrExtornarCheque
    lsSQL = "Exec stp_upd_ERS1262013_ExtornarCheque " & pnId & "," & pnMovNroExtorno
    oConect.Ejecutar lsSQL
    Exit Sub
ErrExtornarCheque:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ExtornarCheque Method")
End Sub
Public Sub RegistraDepositoCheque(ByVal pnId As Long, ByVal pdDepFechaReg As Date, ByVal psDepIFTpo As String, ByVal psDepPersCod As String, ByVal psDepIFCta As String)
    Dim lsSQL As String
    On Error GoTo ErrRegistraDepositoCheque
    lsSQL = "Exec stp_upd_ERS1262013_RegistraDepositoCheque " & pnId & ",'" & Format(pdDepFechaReg, "yyyymmdd") & "','" & Format(psDepIFTpo, "00") & "','" & psDepPersCod & "','" & psDepIFCta & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrRegistraDepositoCheque:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:RegistraDepositoCheque Method")
End Sub
Public Sub InsertaMovDocRec(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String)
    On Error GoTo InsertaMovDocErr
    Dim sql As String
    sql = "Exec stp_ins_ERS1262013_MovDocRec " & pnMovNro & "," & pnDocTpo & ",'" & psNroDoc & "','" & psPersCod & "','" & psIFTpo & "','" & psIFCta & "'"
    oConect.Ejecutar sql
    Exit Sub
InsertaMovDocErr:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovDocRec Method")
End Sub
Public Sub RevierteAsientoPendientexRecepcionCheque(ByVal pnMovNro As Long, ByVal pnMovNroPend As Long, ByVal pnMonto As Currency)
    On Error GoTo ErrRevierteAsientoPendientexRecepcionCheque
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_ERS1262013_RevierteAsientoPendientexRecepcionCheque " & pnMovNro & "," & pnMovNroPend & "," & pnMonto
    oConect.Ejecutar lsSQL
    Exit Sub
ErrRevierteAsientoPendientexRecepcionCheque:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:RevierteAsientoPendientexRecepcionCheque Method")
End Sub
Public Sub ActualizaNroMovDeposito(ByVal pnMovNro As Long, ByVal pnMovNroDeposito As Long)
    On Error GoTo ErrActualizaNroMovDeposito
    Dim lsSQL As String
    lsSQL = "Exec stp_upd_ERS1262013_ActualizaNroMovDeposito " & pnMovNro & "," & pnMovNroDeposito
    oConect.Ejecutar lsSQL
    Exit Sub
ErrActualizaNroMovDeposito:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ActualizaNroMovDeposito Method")
End Sub
'END EJVG *******
'EJVG20140709 ***
Public Sub InsertaMovHabilitacionTransp(ByVal pnMovNro As Long, ByVal pbPropioTransp As Boolean, ByVal psPersCodTransp As String)
    On Error GoTo ErrInsertaMovHabilitacionTransp
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0252014_MovHabilitacionTransp " & pnMovNro & "," & IIf(pbPropioTransp, 1, 0) & ",'" & psPersCodTransp & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovHabilitacionTransp:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovHabilitacionTransp Method")
End Sub
Public Function InsertaSolicitudAprobacionHabRemesa(ByVal psMovNroSolicitud As String, ByVal psAreaCodOrig As String, ByVal psAgeCodOrig As String, ByVal psAreaCodDest As String, ByVal psAgeCodDest As String, ByVal pnMoneda As Moneda, ByVal pnImporte As Currency) As Long
    On Error GoTo ErrInsertaSolicitudAprobacionHabRemesa
    Dim rsAprob As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_ins_ERS0252014_MovHabilitacionAprob '" & psMovNroSolicitud & "','" & psAreaCodOrig & "','" & psAgeCodOrig & "','" & psAreaCodDest & "','" & psAgeCodDest & "'," & pnMoneda & "," & pnImporte
    Set rsAprob = oConect.CargaRecordSet(lsSQL)
    If Not rsAprob.EOF Then
        InsertaSolicitudAprobacionHabRemesa = rsAprob!nId
    End If
    rsAprob.Close
    Set rsAprob = Nothing
    Exit Function
ErrInsertaSolicitudAprobacionHabRemesa:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaSolicitudAprobacionHabRemesa Method")
End Function
Public Function ConsultaSolicitudAprobacionRemesa(ByVal pnId As Long) As Integer
    On Error GoTo ErrConsultaSolicitudAprobacionRemesa
    Dim rsAprob As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "EXEC stp_sel_ERS0252014_ConsultaAprobHabRemesa " & pnId
    Set rsAprob = oConect.CargaRecordSet(lsSQL)
    ConsultaSolicitudAprobacionRemesa = -1
    If Not rsAprob.EOF Then
        ConsultaSolicitudAprobacionRemesa = rsAprob!nAprobacion
    End If
    rsAprob.Close
    Set rsAprob = Nothing
    Exit Function
ErrConsultaSolicitudAprobacionRemesa:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ConsultaSolicitudAprobacionRemesa Method")
End Function
Public Sub VincularSolicitudAprobHabRemesaOperacion(ByVal pnId As Long, ByVal pnMovNro As Long)
    On Error GoTo ErrVincularSolicitudAprobHabRemesaOperacion
    Dim lsSQL As String
    lsSQL = "EXEC stp_upd_ERS0252014_VincularSolicitudAprobHabRemesaOperacion " & pnId & "," & pnMovNro
    oConect.Ejecutar lsSQL
    Exit Sub
ErrVincularSolicitudAprobHabRemesaOperacion:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:VincularSolicitudAprobHabRemesaOperacion Method")
End Sub
Public Function ExisteMovimientosDespues(ByVal pnMovNro As Long, ByVal pnSoloVigentes As Boolean) As Boolean
    On Error GoTo ErrExisteMovimientosDespues
    Dim lsSQL As String
    Dim rs As New ADODB.Recordset
    lsSQL = "select dbo.fnc_ERS0252014_ExisteMovimientosDespues(" & pnMovNro & "," & IIf(pnSoloVigentes, 1, 0) & ") bExiste"
    Set rs = oConect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        ExisteMovimientosDespues = rs!bExiste
    End If
    Exit Function
ErrExisteMovimientosDespues:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:ExisteMovimientosDespues Method")
End Function
'END EJVG *******
'PASI20160610 CCE **************************************
Public Function CCE_InsertaCheque(ByVal pnId As Long, ByVal cNombrebenef As String, ByVal nTpoDocBenef As Integer, ByVal cNroDocBenef As String, _
                                    ByVal cNombreBcoDest As String, ByVal nMismoTit As Integer)
Dim lsSQL As String
On Error GoTo ErrInsertaChequeCCE
    lsSQL = "exec stp_ins_CCE_DocRec " & pnId & ",'" & cNombrebenef & "'," & nTpoDocBenef & ",'" & cNroDocBenef & "','" & cNombreBcoDest & "'," & nMismoTit
    oConect.Ejecutar lsSQL
    Exit Function
ErrInsertaChequeCCE:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaChequeCCE Method")
End Function
Public Sub CCE_ValorizaChequeNeg(ByVal pnId As Long, ByVal pdFechaValoriza As Date, ByVal psMovNro As String)
    Dim lsSQL As String
    lsSQL = "Exec stp_upd_ERS1262013_ValorizaCheque " & pnId & ",'" & Format(pdFechaValoriza, "yyyymmdd hh:mm:ss") & "','" & psMovNro & "'"
    oConect.Ejecutar lsSQL
End Sub
Public Sub CCE_RechazaChequeNeg(ByVal pnId As Long, ByVal pdFechaRechaza As Date)
    Dim lsSQL As String
    lsSQL = "Exec stp_upd_ERS1262013_RechazaCheque " & pnId & ",'" & Format(pdFechaRechaza, "yyyymmdd hh:mm:ss") & "'"
    oConect.Ejecutar lsSQL
End Sub
Public Sub CCE_ActualizaNroMovRechazo(ByVal pnMovNro As Long, ByVal pnMovNroDeposito As Long)
    Dim lsSQL As String
    lsSQL = "Exec stp_upd_ERS1262013_ActualizaNroMovRechazo " & pnMovNro & "," & pnMovNroDeposito
    oConect.Ejecutar lsSQL
    Exit Sub
End Sub
'end PASI**********************************************
'PASI20161212 CCE*****************************************
Public Function CCE_InsertaChequeAjuste(ByVal nId As Long, ByVal pnMontoAnterior As Currency, ByVal pnMonto As Currency, ByVal pnMovNro As Long)
Dim lsSQL As String
On Error GoTo ErrInsertaChequeAjusteCCE
    lsSQL = "exec stp_ins_CCE_DocRecAjuste " & nId & "," & pnMontoAnterior & "," & pnMonto & "," & pnMovNro
    oConect.Ejecutar lsSQL
    Exit Function
ErrInsertaChequeAjusteCCE:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:CCE_InsertaChequeAjuste Method")
End Function
'PASI END*************************************************
'***CTI3 (feirmoro) 19102018
Public Sub InsertaDetExtorno(ByVal pnMovNroAnt As Long, ByVal psOpeCod As String, Optional pscCtaCod As String = "", Optional ByVal psMotExtorno As Variant)
Dim sql As String

    On Error GoTo errInsertaDetExtorno
    
    sql = "INSERT MovDetExtornos (nMovNro,cOpedCod,cCtaCod,cMotExtorno,cDetMotExto) " _
        & "VALUES (" & pnMovNroAnt & ",'" & psOpeCod & "','" & pscCtaCod & "','" & psMotExtorno(0) & "','" & UCase(psMotExtorno(1)) & "')"
    oConect.Ejecutar sql
    
    
Exit Sub
errInsertaDetExtorno:
    Call oError.RaiseError(oError.MyUnhandledError, "DMov:InsertaMovModifica Method")
End Sub

