VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DMov"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A836C85007D"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Base 0
Option Explicit
Dim vsServerAdmin As String
Dim vsServerComunes As String
Dim vsServerImagenes As Variant
Dim vsServerNegocio As String
Dim vsServerPersonas As String
Dim oConect As DConecta
Dim sSQL As String
Dim rs   As ADODB.Recordset
Dim rsCta As ADODB.Recordset
Dim lbBitControlCuentasUltNivel As Boolean
Dim sDBComunes As String 'CTI2 20190502
Dim sDBPersona As String 'CTI2 20190502

Public Sub RollbackTrans()
    oConect.RollbackTrans
End Sub

Public Sub CommitTrans()
    oConect.CommitTrans
End Sub
Public Sub BeginTrans()
    oConect.BeginTrans
End Sub

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
    Call RaiseError(MyUnhandledError, "DMov:BuscarMov Method")
End Function

Public Function InsertaMovCont(ByVal pnMovNro As Long, ByVal pnMovMonto As Currency, ByVal pnMovCorrela As Integer, ByVal psMovParalelo As String) As Integer
    On Error GoTo InsertaMovContErr
    Dim sql As String
    InsertaMovCont = 1
    
    sql = "INSERT INTO MovCont (nMovNro, nMovMonto, nMovCorrela, cMovParalelo ) " _
         & "VALUES (" & pnMovNro & "," & pnMovMonto & "," & pnMovCorrela & ",'" & psMovParalelo & "')"
     
    oConect.Ejecutar sql
    InsertaMovCont = 0
    Exit Function
InsertaMovContErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCont Method")
End Function

Public Function CuentaEsPendiente(ByVal psCta As String, Optional ByRef psClaseCtaOut As String = "", Optional ByVal psClaseCtaIn As String) As Boolean
Dim lsSql As String
Dim lrs   As ADODB.Recordset
CuentaEsPendiente = False
lsSql = "SELECT cp.cCtaContCod, cCtaCaracter cCtaContClase " _
      & "FROM CtaContPend cp JOIN CtaContClase cc ON cp.cCtaContCod LIKE cc.cCtaContCod + '%' " _
      & "WHERE cp.cCtaContCod like '" & psCta & "%' " & IIf(psClaseCtaIn = "", "", " and cc.cCtaCaracter = '" & psClaseCtaIn & "' ")
Set lrs = oConect.CargaRecordSet(lsSql)
If Not lrs.EOF Then
    CuentaEsPendiente = True
    psClaseCtaOut = lrs!cCtaContClase
End If
RSClose lrs
End Function

Public Function InsertaMovPendientesRend(ByVal pnMovNro As Long, ByVal psCtaContCod As String, ByVal pnSaldo As Currency) As Integer
    On Error GoTo InsertaMovPendientesRendErr
    Dim sql As String
    InsertaMovPendientesRend = 1
    sql = "SELECT nMovNro FROM MovPendientesRend Where nMovNro = " & pnMovNro & " and cCtaContCod = '" & psCtaContCod & "' "
    Set rs = oConect.CargaRecordSet(sql)
    If rs.EOF Then
        sql = "INSERT INTO MovPendientesRend (nMovNro, cCtaContCod, nSaldo ) " _
             & "VALUES (" & pnMovNro & ", '" & psCtaContCod & "'," & pnSaldo & ") "
    Else
        sql = "UPDATE MovPendientesRend SET nSaldo = " & pnSaldo & " WHERE nMovNro = " & pnMovNro & " "
    End If
    oConect.Ejecutar sql
    InsertaMovPendientesRend = 0
    Exit Function
InsertaMovPendientesRendErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovPendientesRend Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovARendir Method")
End Function

Public Sub ActualizaMov(ByVal pnMovNro As Long, Optional ByVal psMovDesc As String = "", Optional ByVal pnMovEstado As MovEstado = -1, Optional ByVal psMovFlag As MovFlag = gMovFlagVigente, Optional ByVal psMovNro As String = "", Optional pbActualizaSaldo As Boolean = False, Optional psSimbolo As String = "+")
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
    If psMovNro <> "" Then
        lsActualiza = lsActualiza & " cMovNro = '" & psMovNro & "', "
    End If
    lsActualiza = lsActualiza & " nMovFlag=" & psMovFlag & ","
    lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    If lsActualiza = "" Then Exit Sub
    
    sql = " Update Mov set " & lsActualiza & " Where nMovNro =" & pnMovNro & ""
    oConect.Ejecutar sql
    
    If pbActualizaSaldo Then
        ActualizaSaldoMovimiento psMovNro, psSimbolo
    End If
    
    Exit Sub
ActualizaMovErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMov Method")
End Sub

Public Sub EliminaMov(ByVal psMovNro As String, Optional psMovNroActualiza As String)
Dim sql As String
   On Error GoTo EliminaMovErr
    sql = "UPDATE Mov SET nMovFlag = '" & gMovFlagEliminado & "' WHERE cMovNro = '" & psMovNro & "'"
    oConect.Ejecutar sql
    
    If Not psMovNroActualiza = "" Then
        InsertaMovModifica psMovNroActualiza, psMovNro
    End If
    ActualizaSaldoMovimiento psMovNro, "-"
Exit Sub

EliminaMovErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaMov Method")
End Sub

Public Sub EliminaMovUserEfectivo(ByVal psMovNro As String)
Dim sql As String
   On Error GoTo EliminaMovErr
    sql = "DELETE MovUserEfectivo WHERE nMovNro = " & psMovNro & " "
    oConect.Ejecutar sql
Exit Sub

EliminaMovErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaMovUserEfectivo Method")
End Sub


Public Sub InsertaMovModifica(ByVal psMovNro As String, ByVal psMovNroAnt As String, Optional psMovNroNew As String)
Dim sql As String
    On Error GoTo InsertaMovModificaErr
    sql = "INSERT MovModifica (cMovNro, cMovNroAnt,cMovNroNew) VALUES ('" & psMovNro & "','" & psMovNroAnt & "', '" & psMovNroNew & "')"
    oConect.Ejecutar sql
Exit Sub
InsertaMovModificaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovModifica Method")
End Sub
'##ModelId=3A848F08038A
'Public Function InsertaMovCta(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency) As Integer
'    On Error GoTo InsertaMovCtaErr
'    Dim sql As String
'    InsertaMovCta = 1
'
'    If lbBitControlCuentasUltNivel Then
'        sql = "Select cCtaContCod from ctacont where cCtaContCod Like '" & psCtaContCod & "%' And cCtaContCod <> '" & psCtaContCod & "'"
'        Set rsCta = oConect.CargaRecordSet(sql)
'
'        If rsCta.EOF And rsCta.BOF Then
'            sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte )" _
'                & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
'            oConect.Ejecutar sql
'            InsertaMovCta = 0
'            rsCta.Close
'            Exit Function
'        End If
'        rsCta.Close
'        Err.Number = 9999
'        Err.Description = "Se desea grabar una Cta Cont que no esta a Ult Nivel."
'    Else
'        sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte )" _
'            & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
'        oConect.Ejecutar sql
'        InsertaMovCta = 0
'        'rsCta.Close
'        Exit Function
'    End If
'InsertaMovCtaErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
'End Function

Public Function ObtieneValCuentaPrincipal(ByVal psCtaContCod As String) As String
Dim psSql As String
Dim rs As New ADODB.Recordset
Dim oConec As New DConecta
Dim psRegValCta As Currency
    oConec.AbreConexion
    psSql = "Exec stp_sel_ObtieneCtasPrincipales '" & psCtaContCod & "'"
    Set rs = oConec.CargaRecordSet(psSql)
    If Not rs.BOF And Not rs.EOF Then
        psRegValCta = rs!psTipoCons
    Else
        psRegValCta = "0"
    End If
oConec.CierraConexion
ObtieneValCuentaPrincipal = psRegValCta
Set oConec = Nothing
End Function 'NAGL 20190611 ACTA 018-2019

'Actualizado JHVP 06.12.2004 inclusión bit existen cuentas
'ALPA 20090317*******************************************
'Public Function InsertaMovCta(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency) As Integer
Public Function InsertaMovCta(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency, Optional pnTipoPago As Integer = 0, Optional ByRef psTextoError As String = "") As Integer
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    Dim psValCtaPrinc As String 'NAGL 20190611
    InsertaMovCta = 1
    
    psValCtaPrinc = ObtieneValCuentaPrincipal(psCtaContCod) 'NAGL 20190611
    
    If lbBitControlCuentasUltNivel Then
        If psValCtaPrinc = "1" Then 'Permite el Registro de Cuentas Principales Autorizadas
            sql = "Select cCtaContCod From CtaCont where cCtaContCod Like '" & psCtaContCod & "' And cCtaContCod <> '" & psCtaContCod & "' And nCtaEstado = 1"
            'NAGL Extrajó % en esta sección where cCtaContCod Like '" & psCtaContCod & "%' 20190201
        Else
            sql = "Select cCtaContCod From CtaCont where cCtaContCod Like '" & psCtaContCod & "%' And cCtaContCod <> '" & psCtaContCod & "' And nCtaEstado = 1"
        End If 'NAGL 20190611 Agregó Condicional
        
        Set rsCta = oConect.CargaRecordSet(sql)
        
        If rsCta.EOF And rsCta.BOF Then
            rsCta.Close

            sql = "Select cCtaContCod from CtaCont Where cCtaContCod = '" & psCtaContCod & "' And nCtaEstado = 1"
            Set rsCta = oConect.CargaRecordSet(sql)

            If rsCta.EOF And rsCta.BOF Then
                rsCta.Close
                Err.Number = 9999
                Err.Description = "Cuenta Contable no existe o no es valida " & psCtaContCod & "."
            Else
            'ALPA 20090317***************************************************
'                sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte)" _
'                    & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
                sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte,nTipoPago)" _
                   & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ", " & pnTipoPago & ") "
            '****************************************************************
                
                oConect.Ejecutar sql
                InsertaMovCta = 0
                rsCta.Close
                Exit Function
            End If
        Else
            rsCta.Close
            Err.Number = 9999
            Err.Description = "Se desea grabar una Cta Cont que no esta a Ult Nivel " & psCtaContCod & " ."
            
            '*** PEAC 20111115
            psTextoError = "A esta cuenta Contable " & psCtaContCod & " le falta asignar sub cuentas "
            'Exit Function 'ALPA20130719
            '*** FIN PEAC
            
        End If
    Else
'ALPA 20090330*******************************
'        sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte )" _
'            & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
        sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte,nTipoPago )"
        sql = sql & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ", " & pnTipoPago & ") "
'********************************************
        oConect.Ejecutar sql
        InsertaMovCta = 0
        Exit Function
    End If
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function

Public Function InsertaMovCtaPer(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency) As Integer
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    InsertaMovCtaPer = 1

    If lbBitControlCuentasUltNivel Then
        sql = "Select cCtaContCod From CtaCont Nolock where cCtaContCod Like '" & psCtaContCod & "%' And cCtaContCod <> '" & psCtaContCod & "' And nCtaEstado = 1"
        Set rsCta = oConect.CargaRecordSet(sql)

        If rsCta.EOF And rsCta.BOF Then
            rsCta.Close

            sql = "Select cCtaContCod from CtaCont Nolock Where cCtaContCod = '" & psCtaContCod & "' And nCtaEstado = 1"
            Set rsCta = oConect.CargaRecordSet(sql)

            If rsCta.EOF And rsCta.BOF Then
                rsCta.Close
                Err.Number = 9999
                Err.Description = "Cuenta Contable no existe o no es valida " & psCtaContCod & "."
            Else
                sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte)" _
                    & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
                oConect.Ejecutar sql
                InsertaMovCtaPer = 0
                rsCta.Close
                Exit Function
            End If
        Else
            rsCta.Close
            Err.Number = 9999
            Err.Description = "Se desea grabar una Cta Cont que no esta a Ult Nivel " & psCtaContCod & " ."
        End If
    Else
        sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte )" _
            & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
        oConect.Ejecutar sql
        InsertaMovCtaPer = 0
        Exit Function
    End If
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCartaFianza Method")
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
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovCta Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovMe Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
End Function


Public Function InsertaOpe(ByVal lnMovNro As Long, ByVal pnImporte As Currency, ByVal psNroDoc As String, ByVal psMovDesc As String, _
                            ByVal pnMoneda As Integer, ByVal psOpeCod As String) As Integer
    On Error GoTo InsertaMovObjErr
    Dim sql As String
    InsertaOpe = 1
'lnMovNro, pnImporte, psNroDoc, psMovdesc, pnMoneda, psopecod
    sSQL = "Insert Into MovOpeVarias Values("
    sSQL = sSQL & lnMovNro & ","
    sSQL = sSQL & pnImporte & ","
    sSQL = sSQL & "'" & psNroDoc & "',"
    sSQL = sSQL & "'" & psMovDesc & "',"
    sSQL = sSQL & pnMoneda & ","
    sSQL = sSQL & "'" & psOpeCod & "')"
    
    oConect.Ejecutar sSQL
    InsertaOpe = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjViaticos Method")
End Function

Public Function InsertaMovObjIF(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As String, ByVal psPersCod As String, ByVal psTipoIF As CGTipoIF, ByVal psCtaIFCod As String) As Integer
    On Error GoTo InsertaMovObjIFErr
    Dim sql As String
    InsertaMovObjIF = 1
     
    sql = "INSERT INTO MOVOBJIF(nMovNro, nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod ) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psPersCod & "','" & Format(psTipoIF, "00") & "','" & psCtaIFCod & "')"
    
    oConect.Ejecutar sql
    InsertaMovObjIF = 0
    Exit Function
InsertaMovObjIFErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjIF Method")
End Function

'##ModelId=3A848F4F002E
Public Function InsertaMovDoc(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psDocFecha As String, _
                              Optional ByVal pcLugar As String = "") As Integer
    On Error GoTo InsertaMovDocErr
    Dim sql As String
    InsertaMovDoc = 1
    '***Agregado por ELRO el 20130111, según OYP-RFC126-2012
    'sql = "INSERT INTO MOVDOC (nMovNro, nDocTpo, cDocNro, dDocFecha ) " _
    '    & "VALUES (" & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & psDocFecha & "')"
    sql = "exec stp_sel_RFC1262012_RegistrarMovDoc " & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & psDocFecha & "','" & pcLugar & "'"
    '***Fin Agregado por ELRO el 20130111*******************
    
    oConect.Ejecutar sql
    InsertaMovDoc = 0
    Exit Function
InsertaMovDocErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovDoc Method")
End Function

Public Function InsertaMovDocRef(ByVal pnMovNro As Long, ByVal psDocNro As String, ByVal psDocFecha As String, ByVal pnTipo As Integer) As Integer
    On Error GoTo InsertaMovDocRefErr
    Dim sql As String
    InsertaMovDocRef = 1
    
    sql = "INSERT INTO MOVDOCOtros (nMovNro, cNroConstancia, dFecha, nTipo ) " _
        & "VALUES (" & pnMovNro & ",'" & psDocNro & "','" & psDocFecha & "'," & pnTipo & ")"
        
    oConect.Ejecutar sql
    InsertaMovDocRef = 0
    Exit Function
InsertaMovDocRefErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovDoc Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCajaChica Method")
End Function

'MIOL 20130422, SEGUN RQ13152 **************************************************************************
Public Function InsertaMovCajaChicaCambResp(ByVal pnMovNro As Long, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnProcNro As Integer, ByVal psPersCodActual As String, ByVal psPersCodNuevo As String, ByVal pnSaldoActual As Currency, ByVal pcMovNro As String, ByVal pnEstado As Integer) As Integer
    On Error GoTo InsertaMovCajaChicaCambRespErr
    Dim sql As String
    InsertaMovCajaChicaCambResp = 1
    sql = " INSERT INTO MovCajaChicaCambioEncargado(nMovNro, cAreaCod, cAgeCod, nProNro,cPersCodActual,cPersCodNuevo,nSaldoActual,cMovNro,nEstado ) " _
        & " VALUES(" & pnMovNro & ",'" & psAreaCod & "','" & psAgeCod & "'," & pnProcNro & ",'" & psPersCodActual & "','" & psPersCodNuevo & "'," & pnSaldoActual & ",'" & pcMovNro & "'," & pnEstado & ")"
    
    oConect.Ejecutar sql
    
    InsertaMovCajaChicaCambResp = 0
    Exit Function
InsertaMovCajaChicaCambRespErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCajaChicaCambResp Method")
End Function
'END MIOL **********************************************************************************************
'MIOL 20130422, SEGUN RQ13153 **************************************************************************
Public Function ActualizaMovCajaChicaCambResp(ByVal psPersCodNuevo As String, ByVal psAreaCod As String, ByVal psAgeCod As String, ByVal pnProcNro As Integer) As Integer
    On Error GoTo ActualizaMovCajaChicaCambRespErr
    Dim sql As String
    ActualizaMovCajaChicaCambResp = 1
    
    sql = " UPDATE CAJACHICA SET cPersCod = '" & psPersCodNuevo & "' WHERE cAreaCod='" & psAreaCod & "' and cAgeCod='" & psAgeCod & "' and nProcNro=" & pnProcNro & ""
    
    oConect.Ejecutar sql
    
    ActualizaMovCajaChicaCambResp = 0
    Exit Function
ActualizaMovCajaChicaCambRespErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCajaChicaCambResp Method")
End Function
'END MIOL **********************************************************************************************
Public Function ActualizaMovCajaChicaNroProc(ByVal pnMovNro As Long, Optional ByVal pnProcNro As Integer = 0) As Integer
    On Error GoTo ActualizaMovCajaChicaNroProcErr
    Dim sql As String
    ActualizaMovCajaChicaNroProc = 1
    sql = " UPDATE MOVCAJACHICA SET nProcNro = " & pnProcNro & " WHERE nMovNro = " & pnMovNro & " "
    oConect.Ejecutar sql
    
    ActualizaMovCajaChicaNroProc = 0
    Exit Function
ActualizaMovCajaChicaNroProcErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovCajaChicaNroProc Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaCajaChica Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovBS Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)

    Set oImp = Nothing
    Set rsCta = New ADODB.Recordset
    
    Dim oIni As ClasIni
    
    Set oIni = New ClasIni
    
    lbBitControlCuentasUltNivel = IIf(oCon.LeeConstSistema(gConstSistBitControCtasContUltNiv) = "1", True, False)
    
    vsServerAdmin = oIni.BaseAdministracion
    vsServerComunes = oIni.BaseComunes
    vsServerImagenes = oIni.BaseImagenes
    vsServerNegocio = oIni.BaseNegocio
    vsServerPersonas = oIni.BasePersonas
    
    Set oConect = New DConecta
    If oConect.AbreConexion = False Then
        Call RaiseError(MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos")
    End If
    Set oIni = Nothing
End Sub

Private Sub Class_Terminate()
oConect.CierraConexion
Set oConect = Nothing
End Sub

Public Function RevierteGeneraMovME(ByVal pnMovNro As Long) As Boolean
Dim sSQL As String

sSQL = "DELETE MovCta WHERE nMovNro = " & pnMovNro & " and nMovItem NOT IN " _
     & "( SELECT nMovItem FROM MovMe WHERE nMovNro = " & pnMovNro & " ) "
oConect.Ejecutar sSQL

sSQL = "UPDATE mc SET nMovImporte = me.nMovMEImporte " _
     & "FROM MovME me JOIN MovCta mc ON mc.nMovNro = me.nMovNro and mc.nMovItem = me.nMovItem " _
     & "WHERE me.nMovNro = " & pnMovNro
oConect.Ejecutar sSQL

sSQL = "DELETE MovME WHERE nMovNro = " & pnMovNro
oConect.Ejecutar sSQL

sSQL = "DELETE MovTpoCambio WHERE nMovNro = " & pnMovNro
oConect.Ejecutar sSQL

End Function
'ALPA 20140318********************************************
'ALPA 20090430
'Public Function GeneraMovME(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional nTCComp1 As Currency, Optional ncOpeCod As String = "") As Boolean
'Public Function GeneraMovME(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional nTCComp1 As Currency, Optional ncOpeCod As String = "", Optional nTipoPago As Integer = 0) As Boolean
Public Function GeneraMovME(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional nTCComp1 As Currency, Optional ncOpeCod As String = "", Optional nTipoPago As Integer = 0, Optional ByVal bSpot As Boolean = False, _
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
Optional ByVal pnValVentTr As Currency = 0) As Boolean
    Dim nTCFijo As Currency
    Dim nTCVent As Currency
    Dim nTCComp As Currency
    Dim sSQL As String
    Dim rs   As New ADODB.Recordset
    Dim dFecha As Date
    Dim nImporteD As Currency
    Dim nImporteS As Currency
    On Error GoTo GeneraMovMEErr
    GeneraMovME = False
    dFecha = CDate(GetFechaMov(psMovNro, True))
    
    Dim oTC As New nTipoCambio
    nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
    
    If nTCMerc = 0 Then
        sSQL = "Select cOpeCod From Mov  Where nMovNro = " & pnMovNro        'X
        Set rs = oConect.CargaRecordSet(sSQL)                               'X
        If rs.EOF And rs.BOF Then                                           'X
            nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
            nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)
        Else                                                                'X
            If Left(rs!cOpeCod, 1) = "4" Or Left(rs!cOpeCod, 1) = "7" Then  'X
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCPondVenta)          'X
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCPonderado)          'X
            Else                                                            'X
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)              'X
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)             'X
            End If                                                          'X
        End If                                                              'X
    Else
        'ALPA 20140317****************************************************************
        InsertaMovTpoCambio pnMovNro, nTCMerc, pnValVent, pnValComp, pnValFijo, pnValFijoDia, pnValVentEsp, pnValCompEsp, pnValPond, pnValPondVenta, pnValPondREU, pnValSBSDia, pnValCompTr, pnValVentTr
        'InsertaMovTpoCambio pnMovNro, nTCMerc
            '*****************************************************************************
        
        If ncOpeCod = "702403" Then
            nTCVent = nTCMerc
            nTCComp = nTCComp1
            
        Else
            nTCVent = nTCMerc
            nTCComp = nTCMerc
        End If
    End If
    Set oTC = Nothing
    'ALPA 20090430 ***********************************
       'oConect.Ejecutar " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent
        oConect.Ejecutar " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent & "," & nTipoPago
    '**************************************************
    
   
    GeneraMovME = True

Exit Function
GeneraMovMEErr:
    Call RaiseError(MyUnhandledError, "DMov:GeneraMovME Method")
End Function

Public Function GeneraMovMEPerf(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional nTCFijo As Currency, Optional nTCVent As Currency, Optional nTCComp As Currency) As Boolean
'    Dim nTCFijo As Currency
'    Dim nTCVent As Currency
'    Dim nTCComp As Currency
    Dim sSQL As String
    Dim rs   As New ADODB.Recordset
    Dim dFecha As Date
    Dim nImporteD As Currency
    Dim nImporteS As Currency
    
    
    On Error GoTo GeneraMovMEErr
    GeneraMovMEPerf = False
   ' dFecha = CDate(GetFechaMov(psMovNro, True))
   
    If nTCMerc = 0 Then
        
    Else
        InsertaMovTpoCambio pnMovNro, nTCMerc
        nTCVent = nTCMerc
        nTCComp = nTCMerc
    End If
        oConect.Ejecutar " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent
   
    GeneraMovMEPerf = True

Exit Function
GeneraMovMEErr:
    Call RaiseError(MyUnhandledError, "DMov:GeneraMovME Method")
End Function

Public Function RecuperaTipoCambioF(ByVal pnFecha As Date, ByVal pnOpeCod As Long, Optional pnMovNro As Long = 0, Optional nTCMerc As Currency = 0, Optional ByRef nTCFijo As Currency, Optional ByRef nTCVent As Currency, Optional ByRef nTCComp) As Boolean
    Dim sSQL As String
    Dim rs   As New ADODB.Recordset
    Dim dFecha As Date
    Dim nImporteD As Currency
    Dim nImporteS As Currency
    Dim oTC As New nTipoCambio
    
    On Error GoTo GeneraMovMEErr
    RecuperaTipoCambioF = False
    
    'dFecha = CDate(GetFechaMov(psMovNro, True))
    dFecha = CDate(pnFecha)
        
    nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
        
        If pnMovNro <> 0 Then                                          'X
            nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
            nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)
        Else                                                                'X
            If Left(pnOpeCod, 1) = "4" Or Left(pnOpeCod, 1) = "7" Then  'X
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCPondVenta)          'X
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCPonderado)          'X
            Else                                                            'X
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)              'X
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)             'X
            End If                                                          'X
        End If                                                              'X

    Set oTC = Nothing
      
    RecuperaTipoCambioF = True

Exit Function
GeneraMovMEErr:
    Call RaiseError(MyUnhandledError, "DMov:GeneraMovME Method")
End Function



Private Sub GetClaseCtaCont(ByRef pnBalaMoneda As Boolean, ByRef psCtaClase, ByVal psCtaCod As String)
Dim oCta As New DCtaCont
Dim prs  As ADODB.Recordset
Set prs = oCta.CargaCtaContClase(psCtaCod)
If Not prs.EOF Then
   pnBalaMoneda = prs!nCtaMoneda
   psCtaClase = prs!cCtaCaracter
Else
   pnBalaMoneda = 0
   psCtaClase = ""
End If
RSClose prs
Set oCta = Nothing
End Sub

'ALPA 20090317**************************************************
'Public Function InsertaMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long, Optional psAgeCodRef As String = "") As Integer
Public Function InsertaMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long, Optional psAgeCodRef As String = "", Optional pnTipoPago As Integer = 0) As Integer
'***************************************************************
    On Error GoTo InsertaMovRefErr
    Dim sql As String
    InsertaMovRef = 1
'ALPA 20090317***********************************************************
'    sql = "INSERT INTO MovRef(nMovNro, nMovNroRef, cAgeCodRef) " _
'         & "VALUES (" & pnMovNro & "," & pnMovNroRef & ", '" & psAgeCodRef & "')"
    sql = "INSERT INTO MovRef(nMovNro, nMovNroRef, cAgeCodRef,nTipoPago ) " _
         & "VALUES (" & pnMovNro & "," & pnMovNroRef & ", '" & psAgeCodRef & "'," & pnTipoPago & ")"
'*************************************************************************
    oConect.Ejecutar sql
    InsertaMovRef = 0
    Exit Function
InsertaMovRefErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovContErr Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCajaChicaRend Method")
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
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovARendir Method")
End Function

Public Function ActualizaMovArendirSust(pnMovNroRend As Long, pnMovNroSust As Long)
    Dim sql As String

   sql = "UPDATE MOVARENDIRSUST SET nMovNroRend=" & pnMovNroRend & " where nMovNroSust =" & pnMovNroSust & " "
   oConect.Ejecutar sql

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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjPers Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjAgenciaArea Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjEfectivo Method")
End Function

Public Function InsertaMovObjEfectivoNew(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psEfectivoCod As String, ByRef psCadena As String) As Integer
    On Error GoTo InsertaMovObjEfectivoErr
    Dim sql As String
    InsertaMovObjEfectivoNew = 1
     
    sql = "INSERT INTO MovObjEfectivo(nMovNro, nMovItem, nMovObjOrden, cEfectivoCod ) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psEfectivoCod & "')"

    'oConect.Ejecutar sql
    psCadena = sql
    InsertaMovObjEfectivoNew = 0
    Exit Function
InsertaMovObjEfectivoErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjEfectivo Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertProv Method")
End Function

Public Function InsertaMovGasto(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal psPersCodDest As String, Optional psDestino As String = "") As Integer
    On Error GoTo InsertaMovGastoErr
    Dim sql As String
    InsertaMovGasto = 1
    
    sql = "INSERT INTO MovGasto (nMovNro, cPersCod, cPersCodDest, cDestino ) " _
        & "VALUES (" & pnMovNro & ", '" & psPersCod & "','" & psPersCodDest & "', '" & psDestino & "')"
    oConect.Ejecutar sql
    InsertaMovGasto = 0
    Exit Function
InsertaMovGastoErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovGasto Method")
End Function

Public Function InsertaMovGastoRef(ByVal pnMovNro As Long, ByVal psPersCod As String) As Integer
    On Error GoTo InsertaMovGastoRefErr
    Dim sql As String
    InsertaMovGastoRef = 1
    
    sql = "INSERT INTO MovGastoOtros(nMovNro, cPersCod) " _
        & "VALUES (" & pnMovNro & ", '" & psPersCod & "')"
    oConect.Ejecutar sql
    InsertaMovGastoRef = 0
    Exit Function
InsertaMovGastoRefErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovGasto Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovOtrosItem Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovArendirSust Method")
End Function
Public Function InsertaCheque(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psTipoIF As String, _
                              ByVal pnPlaza As ChequePlaza, ByVal psIFCta As String, ByVal pnMonto As Currency, _
                              ByVal pdValorizaRef As Date, ByVal pdValorizacion As Date, ByVal psDepIF As CGEstadosChq, _
                              ByVal pnConfCaja As CGEstadoConfCheque, ByVal pnMonedaCheque As Moneda, _
                              Optional ByVal pnEstado As ChequeEstado = gChqEstEnValorizacion, Optional psAreaCod As String, Optional psAgeCod As String)
    On Error GoTo InsertaChequeErr
    Dim sql As String
    InsertaCheque = 1
    
    sql = " INSERT INTO DOCREC(nTpoDoc, cNroDoc, cPersCod , cIFTpo, bPlaza, cIFCta, nMonto, dValorizaRef, dValorizacion, cDepIF, nConfCaja ,  nMoneda , nEstado, cAreaCod, cAgeCod ) " _
     & "    VALUES (" & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "'," & pnPlaza & ",'" & psIFCta & "'," & pnMonto & ",'" _
     & Format(pdValorizaRef, gsFormatoFecha) & "','" & Format(pdValorizacion, gsFormatoFecha) & "','" & psDepIF & "'," & pnConfCaja & "," & pnMonedaCheque & "," & pnEstado & ",'" & psAreaCod & "','" & psAgeCod & "')"
     
    oConect.Ejecutar sql
    
    InsertaCheque = 0
    Exit Function
InsertaChequeErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaCheque Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaNotaAbonoCargoEst Method")
End Function

Public Function InsertaDocRecCapta(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psTipoIF As String, _
                              ByVal psCodCta As String, ByVal pnMonto As Currency, ByVal psIFCta As String)
    On Error GoTo InsertaChequeErr
    Dim sql As String
    InsertaDocRecCapta = 1
    
    sql = "INSERT INTO DocRecCapta (nTpoDoc, cNroDoc, cPersCod, cIFTpo, cCtaCod, nMonto, cIFCta) VALUES " _
        & "(" & pnTpoDoc & ",'" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "', '" & psCodCta & "', " & pnMonto & ",'" & psIFCta & "') "
    oConect.Ejecutar sql
    
    InsertaDocRecCapta = 0
    Exit Function
InsertaChequeErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaDocRecCapta Method")
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
        lsFiltro = lsFiltro + " dValorizaRef  ='" & Format(pdValorizaRef, gsFormatoFecha) & "', "
    End If
    If pdValorizacion <> "" Then
        lsFiltro = lsFiltro + " dValorizacion ='" & Format(pdValorizacion, gsFormatoFecha) & "', "
    End If
    If pnPlaza <> -1 Then
        lsFiltro = lsFiltro + " bPlaza = " & pnPlaza & ","
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
     & "            cPersCod='" & psPersCod & "' AND cIFTpo ='" & psTipoIF & "' and cIFCta = '" & psIFCta & "' "

    oConect.Ejecutar sql
    
    ActualizaCheque = 0
    Exit Function
ActualizaChequeErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaCheque Method")
End Function

Public Function InsertaChequeEstado(ByVal psTpoDoc As TpoDoc, ByVal psNroDoc As String, psPersCod As String, ByVal psTipoIF As String, ByVal pdFechaEstado As Date, ByVal psEstado As ChequeEstado, ByVal psMovNro As String, psIFCta As String) As Integer
    On Error GoTo InsertaChequeEstadoErr
    Dim sql As String
    InsertaChequeEstado = 1
    
    sql = " INSERT INTO DocRecEst(nTpoDoc, cNroDoc, cPersCod, cIFTpo , nEstado , cMovNro, cIFCta ) " _
     & "    VALUES (" & psTpoDoc & ", '" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "'," _
     & psEstado & ",'" & psMovNro & "','" & psIFCta & "')"
     
    oConect.Ejecutar sql
    InsertaChequeEstado = 0
    Exit Function
InsertaChequeEstadoErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaChequeEstado Method")
End Function
Public Function ActualizaMovimiento(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long, Optional pnMovEstado As MovEstado = gMovEstContabMovContable, _
                                Optional plActualizaAsiento As Boolean = True, Optional pbInsertaMov As Boolean = True, Optional pbDeleteMovAnt As Boolean = True, _
                                Optional ByVal psOpeCod As String = "", Optional ByVal psMovDesc As String = "", Optional ByVal lbUpdaMovRef As Boolean = True, Optional pbActualizaSaldo As Boolean = True) As Integer
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
          If pbActualizaSaldo Then
            ActualizaSaldoMovimiento lsMovNroAnt, "-"
          End If
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
        sql = "UPDATE MOVREF       SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        sql = sql & "UPDATE MOVREF SET nMovNroRef =" & lnMovNronew & " Where nMovNroRef =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
   End If
   
   sql = "UPDATE MOVARENDIR     SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = sql & "UPDATE MOVARENDIRSUST SET nMovNroRend=" & lnMovNronew & " where nMovNroRend =" & pnMovNroAnt & "; "
   sql = sql & "UPDATE MOVARENDIRSUST SET nMovNroSust=" & lnMovNronew & " where nMovNroSust =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = sql & "UPDATE MOVOTROSITEM   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = sql & "UPDATE MOVCAJACHICA   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = sql & "UPDATE MOVCARTAFIANZA SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = sql & "UPDATE MOVTPOCAMBIO   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
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
        
        If pbActualizaSaldo Then
            ActualizaSaldoMovimiento psMovNroNew, "+"
        End If
   End If
   ActualizaMovimiento = 0
   Exit Function
ActualizaMovimientoErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovimiento Method")
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
   Call RaiseError(MyUnhandledError, "DMov:CargaMovDocAsiento Method")
End Function

Public Function CargaMovOpeAsiento(ByVal pnMovNro As Long, Optional ByVal psMovNro As String = "") As ADODB.Recordset
Dim psSql As String
   On Error GoTo CargaOpeMovAsientoErr
      psSql = "SELECT m.cMovNro, m.cOpeCod, o.cOpeDesc, m.cMovDesc, m.nMovEstado, m.nMovFlag, ISNULL(mg.cDestino,'') cDestino, ISNULL(mg.cPersCod,'') cPersCod, ISNULL(p.cPersNombre,'') cPersNombre, ISNULL(pid.cPersIDnro, '00000000') cRuc, ISNULL(mm.cMovNro,'') cMovNroModifica " _
            & "FROM Mov M JOIN " & vsServerComunes & "OpeTpo O ON O.cOpeCod = M.cOpeCod " _
            & "      LEFT JOIN MovGasto mg ON mg.nMovNro = m.nMovNro " _
            & "      LEFT JOIN " & vsServerPersonas & "Persona p ON p.cPersCod = mg.cPersCod LEFT JOIN  PersID pid ON pid.cPersCod = p.cPersCod and pid.cPersIDTpo = " & gPersIdRUC _
            & "      LEFT JOIN MovModifica mm ON mm.cMovNroAnt = m.cMovNro " _
            & "WHERE " & IIf(psMovNro <> "", " M.cMovNro = '" & psMovNro & "' ", " M.nMovNro = '" & pnMovNro & "'")
      Set CargaMovOpeAsiento = oConect.CargaRecordSet(psSql)
   Exit Function
CargaOpeMovAsientoErr:
   Call RaiseError(MyUnhandledError, "DMov:CargaMovOpeAsiento Method")
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
   Call RaiseError(MyUnhandledError, "DMov:CargaMovCtaAsiento Method")
End Function

Public Function CargaMovObjAsiento(ByVal pnMovNro As Long) As ADODB.Recordset
Dim psSql As String
Dim lnMovNro As Long
Dim prs      As ADODB.Recordset
Dim prsObj   As ADODB.Recordset
Dim prsRes   As ADODB.Recordset
   On Error GoTo CargaMovObjAsientoErr
    psSql = "SELECT cMovNro, m.nMovNro, mo.nMovItem, mo.nMovObjOrden, mo.cObjetoCod " _
          & "FROM Mov m JOIN MovObj mo ON m.nMovNro = mo.nMovNro " _
          & "WHERE M.nMovNro = '" & pnMovNro & "' "
    Set prs = oConect.CargaRecordSet(psSql)
    Do While Not prs.EOF
        Select Case prs!cObjetoCod
           Case ObjCMACAgencias, ObjCMACAgenciaArea, ObjCMACArea
              psSql = "SELECT mo.nMovNro, mo.nMovItem, nMovObjOrden, '" & prs!cObjetoCod & "' ObjPadre, mo.cAreacod+mo.cAgeCod cObjetoCod, ISNULL(ag.cAgeDescripcion,'')+' '+ISNULL(a.cAreaDescripcion,'') cObjetoDesc " _
                    & "FROM MovObjAreaAgencia mo JOIN AreaAgencia aa ON aa.cAreaCod = mo.cAreaCod and aa.cAgeCod = mo.cAgeCod " _
                    & "     LEFT JOIN Areas a ON a.cAreaCod = aa.cAreaCod " _
                    & "     LEFT JOIN Agencias ag ON ag.cAgeCod = aa.cAgeCod " _
                    & "WHERE mo.nMovNro = '" & prs!nMovNro & "' and mo.nMovItem = '" & prs!nMovItem & "' and mo.nMovObjOrden = " & prs!nMovObjOrden & ""
                    
           Case ObjEntidadesFinancieras
              psSql = "SELECT mif.nMovNro, mif.nMovItem, mif.nMovObjOrden, '" & prs!cObjetoCod & "' ObjPadre, cif.cIFTpo+'.'+cif.cPersCod +'.'+cif.cCtaIfCod cObjetoCod, RTRIM(p.cPersNombre) + ' ' + cCtaIFDesc cObjetoDesc " _
                    & "FROM MovObjIF mif " _
                    & "     JOIN Persona p ON p.cPersCod = mif.cPersCod " _
                    & "     JOIN CtaIF cif ON cif.cPersCod = mif.cPersCod and cif.cCtaIFCod = mif.cCtaIFCod " _
                    & "WHERE mif.nMovNro = '" & prs!nMovNro & "' and mif.nMovItem = " & prs!nMovItem & " and mif.nMovObjOrden = " & prs!nMovObjOrden
           
           Case ObjDescomEfectivo
              psSql = "SELECT mo.nMovNro, mo.nMovItem, nMovObjOrden, '" & prs!cObjetoCod & "' ObjPadre, aa.cEfectivoCod cObjetoCod, 'EFECTIVO: ' + convert(varchar(50),aa.nEfectivoValor) cObjetoDesc " _
                   & "FROM MovObjEfectivo mo JOIN Efectivo aa ON mo.cEfectivoCod = aa.cEfectivoCod " _
                   & "    " _
                   & "WHERE  mo.nMovNro = " & prs!nMovNro & " and mo.nMovItem = " & prs!nMovItem & " and mo.nMovObjOrden = " & prs!nMovObjOrden

           Case ObjBienesServicios
              psSql = "SELECT bs.nMovNro, bs.nMovItem, bs.nMovBsOrden nMovObjOrden, '" & prs!cObjetoCod & "' ObjPadre, bs.cBSCod cObjetoCod, b.cBSDescripcion cObjetoDesc " _
                   & "FROM MovBS bs JOIN BienesServicios b ON b.cBSCod = bs.cBsCod " _
                   & "WHERE  bs.nMovNro = " & prs!nMovNro & " and bs.nMovItem = " & prs!nMovItem
           Case Else
              psSql = "SELECT mo.nMovNro, mo.nMovItem, nMovObjOrden, '" & prs!cObjetoCod & "' ObjPadre, mo.cObjetoCod, o.cObjetoDesc " _
                   & "FROM MovObj mo JOIN  Objeto o ON o.cObjetoCod = mo.cObjetoCod " _
                   & "WHERE  mo.nMovNro = " & prs!nMovNro & " and mo.nMovItem = " & prs!nMovItem & " and not mo.cObjetoCod IN ('" & ObjPersona & "', '" & ObjEntidadesFinancieras & "', '" & ObjCMACAgenciaArea & "', '" & ObjCMACArea & "','" & ObjCMACAgencias & "','" & ObjBienesServicios & "') "
        End Select
         Set prsObj = oConect.CargaRecordSet(psSql)
        RecordSetAdiciona prsRes, prsObj
        prs.MoveNext
    Loop
    Set CargaMovObjAsiento = prsRes
Exit Function

CargaMovObjAsientoErr:
   Call RaiseError(MyUnhandledError, "DMov:CargaMovObjAsiento Method")
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
   Call RaiseError(MyUnhandledError, "DMov:CargaSumaMovAsiento Method")
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
    Call RaiseError(MyUnhandledError, "DMov:ActualizaArendirSust Method")
End Function
Public Function ExtornaMovimiento(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long, _
                                Optional ByVal psOpeCod As String = "", Optional ByVal psMovDesc As String = "", _
                                Optional ByVal pbEliminaMov As Boolean = False, Optional psMovActualiza As String = "") As Integer
On Error GoTo ExtornaMovimientoErr
    Dim sql As String
    Dim lnMovNronew As Long
    Dim lsMovNroAnt As String
    
    lsMovNroAnt = GetcMovNro(pnMovNroAnt)
    
    ExtornaMovimiento = 1
    If pbEliminaMov = True Then
        EliminaMov lsMovNroAnt, psMovActualiza
        'FRHU 20140531 ERS068-2014
        If gOpeCGOpeBancosDepositoXActivacionSegTarjetaMNExt = psOpeCod Then
            Call ActualizarSolicitudExtornarSegTarjeta(lsMovNroAnt, 1000, psMovNroNew)
        End If
        'FIN FRHU 20140531 ERS068-2014
        If psOpeCod = gOpeCGExtProvPagoEfectivo Or psOpeCod = gOpeCGExtProvPagoTransfer Or psOpeCod = gOpeCGExtProvPagoOPago Or psOpeCod = gOpeCGExtProvPagoCheque Or psOpeCod = gOpeCGExtProvPagoAbono Or psOpeCod = gOpeCGExtProvPagoRechazo Or psOpeCod = gOpeCGExtProvPagoSUNAT Then
            sql = " Update MP" _
                & " Set MP.nSaldo = Abs(MC.nMovImporte)" _
                & " From MovPendientesRend MP" _
                & " Inner Join MovCta MC On MP.nMovNro = MC.nMovNro And MP.cCtaContCod = MC.cCtaContCod" _
                & " Where MC.nMovNro in (Select nMovNroRef from movref where nMovNro = " & pnMovNroAnt & ") And MC.nMovImporte < 0"
            oConect.Ejecutar sql
            
            ActualizaEstadoCuotasCronogramaxExtPagoProveedor pnMovNroAnt 'EJVG20131209
            ActualizaEstadoCuotasCronogramaxExtPagoProveedorNew pnMovNroAnt 'PASIERS1242014
            ActualizaEstadoBienesxExtPagoProveedor pnMovNroAnt 'PASIERS1242014
            
            'Add by gitu 16-03-2009 para extorno de pago SUNAT
            If psOpeCod = gOpeCGExtProvPagoSUNAT Then
                sql = "Update MPS Set" _
                    & " MPS.nMovPagadoSunat = Null " _
                    & " From movControlPagoSunat MPS" _
                    & " Where MPS.nMovNro in (Select nMovNroRef from movref where nMovNro = " & pnMovNroAnt & ")"
                oConect.Ejecutar sql
            End If
            'End If
        End If
        Exit Function
    Else
        'Modi By GITU 2010-08-06 para que inserte el registro de extorno con el estado no contable
        sql = " INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag )" _
            & " SELECT '" & psMovNroNew & "'," & IIf(psOpeCod = "", "cOpeCod", "'" & psOpeCod & "'") & ", " _
            & IIf(psMovDesc = "", "cMovDesc", "'" & Replace(psMovDesc, "''", "'") & "'") & ",13, '" & gMovFlagDeExtorno & "'  " _
            & " FROM Mov WHERE nMovNro ='" & pnMovNroAnt & "' "
        oConect.Ejecutar sql
        lnMovNronew = GetnMovNro(psMovNroNew)
        
        InsertaMovRef lnMovNronew, pnMovNroAnt
        
        sql = "INSERT MOVCONT (nMovNro, nMovMonto, nMovCorrela, cMovParalelo) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nMovMonto, nMovCorrela, cMovParalelo FROM MovCont Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
       
        sql = "INSERT MOVDOC (nMovNro, nDocTpo, cDocNro, dDocFecha) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nDocTpo, cDocNro, dDocFecha FROM MovDoc Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVGASTO (nMovNro, cPersCod, cPersCodDest) " _
            & " SELECT nMovNro = " & lnMovNronew & ",cPersCod, cPersCodDest FROM MovGasto Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        If psOpeCod <> "" And Mid(psOpeCod, 3, 1) = "2" Then
            sql = "INSERT MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte ) " _
            & " SELECT " & lnMovNronew & ", mc.nMovItem, mc.cCtaContCod, me.nMovMEImporte * -1  " _
            & " FROM MOVCTA mc JOIN MovME me ON me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
            & " WHERE mc.nMovNro =" & pnMovNroAnt & " "
            oConect.Ejecutar sql
        Else
            sql = "INSERT MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, cCtaContCod, nMovImporte * -1  FROM MOVCTA Where nMovNro =" & pnMovNroAnt & " "
            oConect.Ejecutar sql
            
            sql = "INSERT MOVME (nMovNro, nMovItem, nMovMEImporte ) " _
                & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovMEImporte*-1 FROM MOVME Where nMovNro =" & pnMovNroAnt & " "
            oConect.Ejecutar sql
        End If
        sql = "INSERT MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cObjetoCod FROM MOVOBJ Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJIF (nMovNro, nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod FROM MOVOBJIF Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJEfectivo (nMovNro, nMovItem, nMovObjOrden, cEfectivoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cEfectivoCod FROM MOVOBJEfectivo Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJAreaAgencia (nMovNro, nMovItem, nMovObjOrden, cAreaCod, cAgeCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cAreaCod, cAgeCod FROM MOVOBJAreaAgencia Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT INTO MOVHABILITACION(nMovNro, cTranspCod, nMontoTrans) " _
            & " SELECT nMovNro=" & lnMovNronew & ", cTranspCod, nMontoTrans * -1 FROM MOVHABILITACION Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
'        sql = "INSERT INTO MOVHABDEVCAJERO(nMovNro, cUser, nMovImporte)  " _
'            & " SELECT nMovNro=" & lnMovNronew & ",cUser, nMovImporte *-1 FROM MOVHABDEVCAJERO Where nMovNro =" & pnMovNroAnt & " "
'        oConect.Ejecutar sql
        
        sql = "UPDATE MOV SET nMovFlag = '" & gMovFlagExtornado & "' where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
    
        If psOpeCod <> "" And Mid(psOpeCod, 3, 1) = "2" Then
             GeneraMovME lnMovNronew, psMovNroNew
        End If
        ActualizaSaldoMovimiento psMovNroNew, "+"
    End If
    ExtornaMovimiento = 0
    Exit Function
ExtornaMovimientoErr:
    Call RaiseError(MyUnhandledError, "DMov:ExtornaMovimiento Method")
End Function

Public Function ExtornaMovimientoRefNegocio(pnNumTran As Long, psMovNro As String, psAgeCod As String) As Boolean
Dim oLCon As New DConecta
Dim lsSql As String
On Error GoTo ExtornaMovimientoRefNegocioErr
If oLCon.AbreConexion() Then
    lsSql = "UPDATE TransRef SET cNroTran = NULL, cCodAge = NULL WHERE cNroTran = '" & psMovNro & "' and cNroTranRef = '" & pnNumTran & "' "
    oLCon.Ejecutar lsSql
    oLCon.CierraConexion
    ExtornaMovimientoRefNegocio = True
Else
    ExtornaMovimientoRefNegocio = False
    MsgBox "No se puede establecer conexión con Agencia " & psAgeCod & ". Imposible realizar Extorno. ", vbInformation, "¡Aviso!"
End If
Set oLCon = Nothing
Exit Function
ExtornaMovimientoRefNegocioErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ExtornaDepositoCheque(pnMovNro As Long) As Boolean
Dim lsSql As String
On Error GoTo ExtornaDepositoChequeErr
lsSql = "UPDATE dr set cDepIF = 0 FROM Mov m JOIN MovRef mr ON mr.nMovNroRef = m.nMovNro " _
      & "     JOIN MovDoc md ON md.nMovNro = m.nMovNro " _
      & "     JOIN DocRec dr ON dr.nTpoDoc = md.nDocTpo and dr.cNroDoc = md.cDocNro " _
      & "     JOIN DocRecEst dre ON dre.nTpoDoc = dr.nTpoDoc and dre.cNroDoc = dr.cNroDoc and dre.cPersCod = dr.cPersCod and dre.cMovNro = m.cMovNro " _
      & " WHERE mr.nMovNro = " & pnMovNro
Set rs = oConect.CargaRecordSet(lsSql)
Exit Function
ExtornaDepositoChequeErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


Public Function ExtornaConfirmacionChequeAgencia(psAgeCod As String, psNroDoc As String, pdRegChq As Date, psObjetoCod As String) As Boolean
Dim oLCon As New DConecta
Dim lsSql As String
Dim lsCodBco As String
On Error GoTo ExtornaConfirmacionChequeAgenciaErr
ExtornaConfirmacionChequeAgencia = False
lsSql = "SELECT cObjetoCod FROM AuxIF where cCodPers = '" & Mid(psObjetoCod, 4, 13) & "' "
Set rs = oConect.CargaRecordSet(lsSql)
If Not rs.EOF Then
    lsCodBco = rs!cObjetoCod
    If oLCon.AbreConexion Then 'Remota(psAgeCod, False, False, "01")
        lsSql = "SELECT nBcoCod FROM DBComunes..Bancos WHERE cValor2 = '" & Right(lsCodBco, 2) & "' "
        Set rs = oLCon.CargaRecordSet(lsSql)
        If Not rs.EOF Then
            lsSql = "UPDATE cheque SET cTransChq = NULL WHERE cNumChq = '" & psNroDoc & "' and datediff(d,dRegChq,'" & Format(pdRegChq, gsFormatoFecha) & "') = 0 " _
                  & "   and cCtaBco = '" & Mid(psObjetoCod, 18, 20) & "' and cCodBco = '" & Format(rs!nBcoCod, "00") & "' and cTransChq = 1 "
            oLCon.Ejecutar lsSql
            oLCon.CierraConexion
            ExtornaConfirmacionChequeAgencia = True
        Else
            Err.Raise 50001, "ExtornaConfirmacionCheque", "No se pudo determinar Codigo de Banco para Extornar Cheque. Consultar con Sistemas"
        End If
    Else
        Err.Raise 50001, "ExtornaConfirmacionCheque", "No se puede establecer conexión con Agencia " & psAgeCod & ". Imposible realizar Extorno. "
    End If
Else
    Err.Raise 50001, "ExtornaConfirmacionCheque", "No se pudo determinar Codigo de Banco para Extornar Cheque. Consultar con Sistemas"
End If
Set oLCon = Nothing
Exit Function
ExtornaConfirmacionChequeAgenciaErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ExtornaProvisionPlazoFijo(psMovNroProv As String, psOpeCod As String, pnImporte As Currency, psCtaPF As String) As Boolean
Dim lsSql As String
Dim lsMovUlt As String
Dim lnInteresPago As Currency
On Error GoTo ErrExtornaProvisionPF

lnInteresPago = 0
lsSql = "Select * From CtaIF where cIFTpo + '.' + cperscod + '.' + cCtaIFCod = '" & psCtaPF & "' "
Set rs = oConect.CargaRecordSet(lsSql)
If Not rs.EOF Then
    lnInteresPago = rs!nInteres
End If

lsSql = "Select Max(m.cMovNro) cMovUlt From mov m join MovObjif mif on mif.nMovNro = m.nMovNro " _
      & "Where cIFTpo + '.' + cPersCod + '.' + cCtaIFCod = '" & psCtaPF & "' " _
      & "  and m.cMovNro < '" & psMovNroProv & "' and m.nMovFlag = " & gMovFlagVigente & " and m.nMovEstado = " & gMovEstContabMovContable
Set rs = oConect.CargaRecordSet(lsSql)
If Not rs.EOF Then
    lsMovUlt = rs!cMovUlt
    lsSql = "Select cOpeCod FROM Mov WHERE cMovNro = '" & lsMovUlt & "' "
    Set rs = oConect.CargaRecordSet(lsSql)
    If rs!cOpeCod = gOpeCGOpeConfApertMN Or rs!cOpeCod = gOpeCGOpeConfApertME Then
        lnInteresPago = 0
    End If
    If rs!cOpeCod = gOpeCGOpeIntDevPFMN Or rs!cOpeCod = gOpeCGOpeIntDevPFME Or rs!cOpeCod = gOpeCGOpeCapIntPFMN Or rs!cOpeCod = gOpeCGOpeCapIntPFME Then
        lsSql = "Select " & IIf(Mid(rs!cOpeCod, 3, 1) = gMonedaExtranjera, " SUM(nMovMEImporte) ", " SUM(nMovImporte) ") & "nInteresPago FROM mov m join movcta mc on mc.nmovnro = m.nmovnro " _
              & IIf(Mid(rs!cOpeCod, 3, 1) = gMonedaExtranjera, " JOIN movme me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem ", "") _
              & " JOIN OpeCta oc ON oc.cOpeCod = '" & IIf(Mid(rs!cOpeCod, 3, 1) = gMonedaExtranjera, gOpeCGOpeCapIntPFME, gOpeCGOpeCapIntPFMN) & "' and mc.cCtaContCod LIKE oc.cCtaContCod + '%' " _
              & "WHERE m.cmovnro = '" & psMovNroProv & "' and oc.cOpeCtaOrden IN (1) and cOpeCtaDH = 'H'"
        Set rs = oConect.CargaRecordSet(lsSql)
        If Not rs.EOF Then
            lnInteresPago = lnInteresPago + rs!nInteresPago
        Else
            lnInteresPago = 0
        End If
    End If
End If

lsSql = "UPDATE ctaif set dCtaIFint = '" & GetFechaMov(lsMovUlt, False) & "', nInteres = " & lnInteresPago & ", cUltimaActualizacion = '" & lsMovUlt & "' " _
      & "Where cIFTpo + '.' + cPersCod + '.' + cCtaIFCod = '" & psCtaPF & "' "
oConect.Ejecutar lsSql

Exit Function
ErrExtornaProvisionPF:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ExtornaPagoCuotaAdeudados(psMovNroPago As String, psOpeCod As String, pnCapital As Currency, pdFechaVig As Date, pnMovNro As Long, psCtaIFPers As String) As Boolean
    Dim lnMontoLP   As Currency
    Dim lsSql       As String
    Dim lsMovNroUlt As String
    Dim lsOpeCod    As String
    Dim lsPersCod   As String
    Dim lsIFTpo     As String
    Dim lsCtaIFCod  As String
    Dim lnSaldo     As Currency
    Dim ldCuotaUltPago As Date
    Dim lnInteresPagado As Currency
    Dim lnInteresPagadoReal As Currency
    
    Dim rs1             As New ADODB.Recordset
    Dim rs2             As New ADODB.Recordset
    Dim rs3             As New ADODB.Recordset
    Dim rs4             As New ADODB.Recordset
    Dim rsCtaCancela    As New ADODB.Recordset
    
    Dim oAdeu As New DCaja_Adeudados
    Dim oCaja As New nCajaGeneral ' NCajaAdeudados
    Dim oAdeu1 As New NCajaAdeudados
    
    Dim lsMovNroAnt As String
    Dim lsMovNroAntRef As String
On Error GoTo ExtornaPagoCuotaAdeudadosErr
    
    Dim lsMovPagoGrabado As String
    Dim lsMovProvGrabado As String
    Dim lsIntPagCtaIF As Double
    
    'PROVISIONES DE ADEUDADOS
    If psOpeCod = gOpeCGAdeudaExtProvisiónMN Or psOpeCod = gOpeCGAdeudaExtProvisiónME Then
        lsSql = " Update a " _
              & " Set a.dCuotaUltModSaldos = b.dFecUltActAnt" _
              & " From CtaIFadeudados  a " _
              & " Inner Join CtaIFEstadistica b On a.cPersCod = b.cPersCod And a.cIFTpo = b.cIFTpo And a.cCtaIFCod = b.cCtaIFCod " _
              & " where b.nMovNro = " & pnMovNro
        oConect.Ejecutar lsSql

    ElseIf psOpeCod = gOpeCGAdeudaExtPagoCuotaMN Or psOpeCod = gOpeCGAdeudaExtPagoCuotaME Then
        
        lsSql = " Select a.dCancelacion " _
              & " From CtaIFadeudados  a " _
              & " Inner Join CtaIFEstadistica b On a.cPersCod = b.cPersCod And a.cIFTpo = b.cIFTpo And a.cCtaIFCod = b.cCtaIFCod " _
              & " where b.nMovNro = " & pnMovNro
        Set rsCtaCancela = oConect.CargaRecordSet(lsSql)
        
        If IsNull(rsCtaCancela!dCancelacion) Then
            lsSql = " Update a " _
                  & " Set a.dCuotaUltModSaldos = b.dFecUltActAnt, a.dCuotaUltPago = b.dFecUltPagoAnt, a.nSaldoCap = b.nSaldoAntReal, a.nSaldoCapLP = b.nSaldoLPAntReal" _
                  & " From CtaIFadeudados  a " _
                  & " Inner Join CtaIFEstadistica b On a.cPersCod = b.cPersCod And a.cIFTpo = b.cIFTpo And a.cCtaIFCod = b.cCtaIFCod " _
                  & " where b.nMovNro = " & pnMovNro
            oConect.Ejecutar lsSql
          
            lsSql = " Update a " _
                  & " Set a.cEstado = 2 " _
                  & " From CtaIFCalendario  a " _
                  & " Inner Join CtaIFEstadistica b On a.cPersCod = b.cPersCod And a.cIFTpo = b.cIFTpo And a.cCtaIFCod = b.cCtaIFCod And a.nNroCuota = b.nNroCuota " _
                  & " where b.nMovNro = " & pnMovNro
            oConect.Ejecutar lsSql
        Else
            lsSql = " Update a " _
                  & " Set a.dCuotaUltModSaldos = b.dFecUltActAnt, a.dCuotaUltPago = b.dFecUltPagoAnt, a.nSaldoCap = b.nSaldoAntReal, a.nSaldoCapLP = b.nSaldoLPAntReal, a.dCancelacion = Null " _
                  & " From CtaIFadeudados  a " _
                  & " Inner Join CtaIFEstadistica b On a.cPersCod = b.cPersCod And a.cIFTpo = b.cIFTpo And a.cCtaIFCod = b.cCtaIFCod " _
                  & " where b.nMovNro = " & pnMovNro
            oConect.Ejecutar lsSql
          
            lsSql = " Update a " _
                  & " Set a.cEstado = 2 " _
                  & " From CtaIFCalendario  a " _
                  & " Inner Join CtaIFEstadistica b On a.cPersCod = b.cPersCod And a.cIFTpo = b.cIFTpo And a.cCtaIFCod = b.cCtaIFCod And a.nNroCuota >= b.nNroCuota " _
                  & " where b.nMovNro = " & pnMovNro
            oConect.Ejecutar lsSql
        
            lsSql = " Update a " _
                  & " Set a.cCtaIFEstado = 1 " _
                  & " From CtaIF  a " _
                  & " Inner Join CtaIFEstadistica b On a.cPersCod = b.cPersCod And a.cIFTpo = b.cIFTpo And a.cCtaIFCod = b.cCtaIFCod " _
                  & " where b.nMovNro = " & pnMovNro
            oConect.Ejecutar lsSql
        End If
End If
    
Exit Function
ExtornaPagoCuotaAdeudadosErr:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function InsertaMovViaticos(ByVal pnViaticoMovNro As Long, ByVal pnMovNro As Long, _
                                   ByVal psDestinoCod As ViaticosDestino, _
                                   ByVal psTranspIda As ViaticosTransporte, _
                                   ByVal psTranspVuelta As ViaticosTransporte, ByVal pdPartida As Date, _
                                   ByVal pnMovViaticosDias As String, ByVal psDestinoDesc As String, _
                                   ByVal pdRetorno As Date) As Integer 'pdRetorno Agregado por PASI20140612
                                   
    On Error GoTo InsertaMovViaticosErr
    Dim sql As String
    InsertaMovViaticos = 1
    
    'Modificado PASI TI-ERS060 2014
'    Sql = "INSERT INTO MOVVIATICOS(nViaticoMovNro, nMovNro, " _
'        & "cDestinoCod, cTranspIda, cTranspVuelta,dPartida, nMovViaticosDias, cDestinoDesc  ) " _
'        & " VALUES (" & pnViaticoMovNro & "," & pnMovNro & ",'" _
'        & psDestinoCod & "','" & psTranspIda & "','" & psTranspVuelta & "','" _
'        & Format(pdPartida, gsFormatoFecha) & "'," & pnMovViaticosDias & ",'" & psDestinoDesc & "')"
    
    sql = "stp_ins_InsertaMovViaticos " & pnViaticoMovNro & "," & pnMovNro & ",'" & psDestinoCod & "','" & psTranspIda & "','" & psTranspVuelta & "','" & Format(pdPartida, gsFormatoFecha) & "'," & pnMovViaticosDias & ",'" & psDestinoDesc & "','" & Format(pdRetorno, gsFormatoFecha) & "'"
    'end PASI
     
    oConect.Ejecutar sql
    InsertaMovViaticos = 0
    Exit Function
InsertaMovViaticosErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovViaticos Method")
End Function
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
        lsFiltro = lsFiltro & " nTopeEgresos = " & pnMontoDesem & ","
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
    Call RaiseError(MyUnhandledError, "DMov:ActualizaCajaChica Method")
End Function

Public Function CargaBilletaje(lsFecha As String, lsMoneda As String, lsCodAge As String, lnMoneda As String, Optional lsEfectivoCod As String = "") As ADODB.Recordset
Dim psSql As String

psSql = "SELECT m.cMovNro, mue.* " _
    & "FROM   MovUserEfectivo MUE JOIN Mov M ON MUE.nMovNro = M.nMovNro " _
    & "WHERE  M.nMovEstado = " & gMovEstContabNoContable & " and M.nMovFlag = " & gMovFlagVigente & " and " _
    & "       m.cMovNro LIKE '" & lsFecha & "%' and " _
    & "       LEFT(mue.cEfectivoCod,1) = " & lsMoneda & " and m.cOpecod IN ('" & gOpeHabCajRegEfect & "','" & gOpeHabBoveRegEfect & "') and " _
    & "       SubString(m.cMovNro,18,2) = '" & lsCodAge & "' " & IIf(lsEfectivoCod = "", "", " and cEfectivoCod = '" & lsEfectivoCod & "'")

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
         & IIf(psFiltro = "", "", IIf(psMovNro = "", "WHERE ", "AND ")) & psFiltro & " and M.nMovFlag = " & gMovFlagVigente
         'NAGL 20191212 Agregó & " and M.nMovFlag = " & gMovFlagVigente
      Set CargaMovCta = oConect.CargaRecordSet(psSql)
   Exit Function
CargaMovCtaErr:
   Call RaiseError(MyUnhandledError, "DMov:CargaMovCta Method")
End Function
Public Sub ActualizaMovCta_Cuenta(psCtaCodant As String, psCtaCodNew As String, psFecha1 As String, psfecha2 As String)
    Dim sql As String
    On Error GoTo ActualizaMovCtaErr
    sql = "UPDATE mc SET cCtaContCod = '" & psCtaCodNew & "' FROM Mov m JOIN MovCta mc ON mc.nMovNro = m.nMovNro WHERE mc.cCtaContCod = '" & psCtaCodant & "' and LEFT(m.cMovNro,8) BETWEEN '" & psFecha1 & "' and '" & psfecha2 & "' "
    oConect.Ejecutar sql
    Exit Sub
ActualizaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovCta_Cuenta Method")

End Sub
'ALPA 20140225********************************************************************************************
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
'    Call RaiseError(MyUnhandledError, "DMov:InsertaMovTipCambio Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovTipCambio Method")
End Sub
'**********************************************************************************************************
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMov Method")
End Function
Public Function GetnMovNro(ByVal psMovNro As String) As Long
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

sql = "Select nMovNro From Mov where cMovNro ='" & psMovNro & "'"
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetnMovNro = rs!nMovNro
End If
rs.Close
Set rs = Nothing
End Function

Public Function GetSaldoVacAdeudado(pCtaIfCod As String, pcPersCod As String, pcIFTpo As String) As ADODB.Recordset
    Dim sql As String
    On Error GoTo BuscarMovErr
    
    sql = "Select nSaldoCap, dbo.GetVac(dCuotaUltPago) Vac, nSaldoCapLP, cMonedaPago  from ctaifadeudados where cctaifcod = '" & pCtaIfCod & "' And cPersCod = '" & pcPersCod & "' And cIFTpo = '" & pcIFTpo & "'"
    Set GetSaldoVacAdeudado = oConect.CargaRecordSet(sql)
    
    Exit Function
BuscarMovErr:
    Call RaiseError(MyUnhandledError, "DMov:BuscarMov Method")
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
Public Sub EliminaMovRef(ByVal pnMovNro As Long, Optional ByVal pnMovNroRef As Long = 0)
    On Error GoTo EliminaMovRefErr
    Dim sql As String
    sql = "DELETE MOVREF WHERE nMOVNRO =" & pnMovNro & " " _
        & IIf(pnMovNroRef = 0, "", " and nMovNroRef = " & pnMovNroRef)
    oConect.Ejecutar sql
    Exit Sub
EliminaMovRefErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaMovRef Method")
End Sub
Public Function InsertaMovHabilitaciones(ByVal pnMovNro As Long, ByVal psCodTransp As String, ByVal psAreaOrig As String, ByVal psAgeOrig As String, _
                                        ByVal psAreaDest As String, ByVal psAgeDest As String, _
                                        ByVal pnMontoTrans As Currency, ByVal pnMovImporte As Currency, ByVal psMoneda As String) As Integer
    On Error GoTo InsertaMovHabilitacionesErr
    Dim sql As String
    InsertaMovHabilitaciones = 1
    sql = "INSERT INTO MovHabilitacion (nMovNro, cTranspCod, cAreaOrig, cAgeOrig, cAreaDest, cAgeDest, nMontoTrans, nMovImporte, nMoneda)" _
        & " VALUES(" & pnMovNro & ",'" & psCodTransp & "','" & psAreaOrig & "','" & psAgeOrig & "','" _
        & psAreaDest & "','" & psAgeDest & "'," & pnMontoTrans & "," & pnMovImporte & ", " & psMoneda & ")"
     
    oConect.Ejecutar sql
    InsertaMovHabilitaciones = 0
    Exit Function
InsertaMovHabilitacionesErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovTipCambio Method")
End Function

Public Function InsertaMovHabDev(ByVal pnMovNro As Long, ByVal psUser As String, ByVal pnImporte As Currency) As Integer
    On Error GoTo InsertaMovHabDevErr
    Dim sql As String
    InsertaMovHabDev = 1
    sql = "Insert MovHabDevCajero(nMovNro, cUsuOrig, nMovImporte) Values(" & pnMovNro & ",'" & psUser & "'," & pnImporte & ")"
 
    oConect.Ejecutar sql
    InsertaMovHabDev = 0
    Exit Function
InsertaMovHabDevErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovHabDev Method")
End Function

Public Function InsertaMovHabilitaCaja(ByVal pnMovNro As Long, ByVal pnImporte As Currency, ByVal cCajaGenCod As String, ByVal psArea As String, ByVal psAgencia As String, ByVal pnMoneda As Integer) As Integer
    On Error GoTo InsertaMovHabDevErr
    Dim sql As String
    InsertaMovHabilitaCaja = 1
    sql = " Insert MovHabilitacion(nMovNro,cTranspCod,nMovImporte,cAreaOrig,cAgeOrig,cAreaDest,cAgeDest,nMontoTrans,nMoneda) " & _
          " Values(" & pnMovNro & ",''," & pnImporte & ",'" & cCajaGenCod & "','','" & psArea & "','" & psAgencia & "',0," & pnMoneda & ")"
 
    oConect.Ejecutar sql
    InsertaMovHabilitaCaja = 0
    Exit Function
InsertaMovHabDevErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovHabDev Method")
End Function

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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovUserEfectivo Method")
End Function

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
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovUserEfectivo Method")
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
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovCont Method")
End Sub

Public Sub ActualizaMovPendientesRend(ByVal pnMovNro As Long, ByVal psCtaCod As String, pnImporte As Currency, Optional pbAfectaSaldo As Boolean = True)
Dim sModif As String
Dim sql    As String
  On Error GoTo ActualizaMovPendientesRendErr
  sql = "UPDATE MovPendientesRend SET nSaldo = " & IIf(pbAfectaSaldo, " nSaldo - ", "") & pnImporte & " WHERE nMovNro = " & pnMovNro & " and cCtaContCod = '" & psCtaCod & "' "
  oConect.Ejecutar sql
Exit Sub
ActualizaMovPendientesRendErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovPendientesRend Method")
End Sub

Public Sub ActualizaMovPendientesRendCambio(ByVal pnMovNroAnt As Long, Optional psSimbolo As String = "+")
Dim sModif As String
Dim sql    As String
  On Error GoTo ActualizaMovPendientesRendCambioErr
  sql = "UPDATE mp set nSaldo = nSaldo " & psSimbolo & " abs(CASE WHEN SubString(mc.cCtaContCod,3,1) = '1' THEN mc.nMovImporte ELSE me.nMovMEImporte END) " _
      & "FROM MovPendientesRend mp join movref mr on mr.nmovnroref = mp.nmovnro " _
      & "      join Movcta mc on mc.nMovNro = mr.nmovnro and mp.cCtaContCod = mc.cCtaContCod " _
      & " left join MovMe  me on me.nMovNro = mc.nMovNro and me.nMovItem = mc.nMovItem " _
      & "WHERE mc.nMovNro = " & pnMovNroAnt
  oConect.Ejecutar sql
Exit Sub
ActualizaMovPendientesRendCambioErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovPendientesRendCambio Method")
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
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
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
    Call RaiseError(MyUnhandledError, "DMov:MovCompraVenta Method")
End Function

Public Function InsertaMovOpeVarias(ByVal pnMovNro As Long, ByVal psNroDoc As String, ByVal psReferencia As String, ByVal pnImporte As Currency) As Integer
    On Error GoTo InsertaMovOpeVariasErr
    Dim sql As String
    InsertaMovOpeVarias = 1
    
    sql = " INSERT INTO MovOpeVarias(nMovNro, cNroDoc, cReferencia, nMovImporte)" _
        & " VALUES(" & pnMovNro & ",'" & psNroDoc & "','" & psReferencia & "'," & pnImporte & ") "
      
    oConect.Ejecutar sql
    InsertaMovOpeVarias = 0
    Exit Function
InsertaMovOpeVariasErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovOpeVarias Method")
End Function

Public Function InsertaNotaAbonoCargo(ByVal pnTpoDoc As TpoDoc, ByVal psNroDoc As String, ByVal pnEstado As NotaCargoAbonoEstado, _
                                    ByVal pnMotivo As Integer, ByVal pnMonto As Currency, Optional psObjetoCod As String = "", Optional psObjeto As String = "", Optional pnMoneda As Moneda = gMonedaNacional) As Integer
    On Error GoTo InsertaNotaAbonoCargoErr
    Dim sql As String
    InsertaNotaAbonoCargo = 1

    sql = " INSERT INTO NotaAbonoCargo(nDocTpo, cDocNro, nEstado, nMotivoCod, nMonto, cObjetoCodPadre, cObjetoCod, nMoneda)" _
        & " VALUES(" & pnTpoDoc & ",'" & psNroDoc & "'," & pnEstado & "," & pnMotivo & "," & pnMonto & ",'" & Format(psObjeto, "00") & "','" & psObjetoCod & "'," & pnMoneda & ") "
    
    oConect.Ejecutar sql
    InsertaNotaAbonoCargo = 0
    Exit Function
InsertaNotaAbonoCargoErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaNotaAbonoCargo Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaNotaAbonoCargoEst Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovArqueo Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovArqueoOP Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovArqueoEfectivo Method")
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
Dim clsGen As DGeneral
Dim sCuenta As String, sFecha As String
Dim sFecUltCierre As String, sOrdPag As String, sCtaAbonoIntPF As String
Dim nSaldoDisp As Double, nSaldRetCTS As Double
Set clsGen = New DGeneral
sCuenta = clsGen.GeneraNuevaCuenta(sAgencia, nProducto, nMoneda)
Set clsGen = Nothing
sFecha = Format$(dFecha, gsFormatoFecha) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
sSQL = "Insert Producto (cCtaCod,nTasaInteres,nSaldo,nPrdEstado,dPrdEstado,nTransacc) " _
    & "Values ('" & sCuenta & "'," & nTasa & "," & nSaldo & "," & gCapEstActiva & ",'" & sFecha & "',0)"
oConect.Ejecutar sSQL


nSaldoDisp = IIf(bCheque, 0, nSaldo)
sFecUltCierre = Format$(DateAdd("d", -1, dFecha), gsFormatoFecha)
sSQL = "Insert Captaciones (cCtaCod,nSaldoDisp,nPersoneria,nFirmas,nIntAcum,dUltCierre,dApertura,nPrdCtaTpo,nPrdTasaInteres,cUltimaActualizacion) " _
    & "Values ('" & sCuenta & "'," & nSaldoDisp & "," & nPersoneria & "," & nFirmas & ",0,'" & sFecUltCierre & "','" & sFecha & "'," & nTipoCuenta & "," _
    & nTipoTasa & ",'" & sMovNro & "')"
oConect.Ejecutar sSQL
Select Case nProducto
    Case gCapAhorros
        sOrdPag = IIf(bOrdPag, "1", "0")
        sSQL = "Insert CaptacAhorros (cCtaCod,nSaldoAnterior,bOrdPag,nSobregiro,dUltContacto) " _
            & "Values ('" & sCuenta & "',0," & sOrdPag & ",0,'" & sFecha & "')"
    Case gCapPlazoFijo
        sCtaAbonoIntPF = IIf(bCtaAboInt, "1", "0")
        sSQL = "Insert CaptacPlazoFijo (cCtaCod,nPlazo,nIntPag,dRenovacion,nApertura,dAuxiliar,nFormaRetiro,nDuplicado,bAbonoIntCtaAho) " _
            & "Values ('" & sCuenta & "'," & nPlazo & ",0,'" & sFecha & "'," & nSaldo & ",'" & sFecha & "'," & nFormaRetiro & ",0," & sCtaAbonoIntPF & ")"
    Case gCapCTS
        nSaldRetCTS = Round(nSaldo * nPorcRetCTS / 100, 2)
        sSQL = "Insert CaptacCTS (cCtaCod,nSaldRetiro,nIntSaldo,cCodInst) " _
            & "Values ('" & sCuenta & "'," & nSaldRetCTS & ",0,'" & sInstitucion & "')"
End Select
oConect.Ejecutar sSQL
If nProducto = gCapPlazoFijo And bCtaAboInt Then
    sSQL = "Insert CaptacCtaAboIntPF (cCtaCod,cCtaCodAbono) " _
        & "Values ('" & sCuenta & "','" & sCtaCodAbono & "')"
    oConect.Ejecutar sSQL
End If
AgregaNuevaCaptacion = sCuenta
End Function
Public Sub AgregaNuevoProdPers(ByVal sCuenta As String, ByVal sPersona As String, ByVal nRelacion As CaptacRelacPersona)
sSQL = "INSERT ProductoPersona (cCtaCod,cPersCod,nPrdPersRelac) " _
    & "VALUES ('" & sCuenta & "','" & sPersona & "'," & nRelacion & ")"
oConect.Ejecutar sSQL
End Sub

Public Function GetDatosOrdenPago(ByVal sCuenta As String, ByVal sNroDoc As String) As Recordset
Dim rsOP As Recordset
sSQL = "SELECT E.cMovNro, D.nTpoDoc, D.cNroDoc, D.cCtaCod, D.nMonto, P.cPersNombre, E.cEstado " _
    & "D.cIFCodPers FROM DocRecOPEst E INNER JOIN DocRecOP D INNER JOIN " & vsServerPersonas & "Persona P " _
    & "ON D.cIFCodPers = P.cPersCod ON E.nTpoDoc = D.nTpoDoc AND E.cNroDoc = D.cNroDoc AND E.cCtaCod " _
    & "= D.cCtaCod WHERE E.cMovNro IN (SELECT MAX(E1.cMovNro) FROM DocRecOPEst E WHERE E1.nTpoDoc = " _
    & "E.nTpoDoc AND E1.cNroDoc = E.cNroDoc AND E1.cCtaCod = E.cCtaCod)"

Set rsOP = New Recordset
Set rsOP = oConect.CargaRecordSet(sSQL)
Set GetDatosOrdenPago = rsOP
End Function
Public Sub AgregaOrdenPagoEstado(ByVal sCuenta As String, ByVal sNroDoc As String, _
        ByVal sMovNro As String, ByVal nMonto As Double, ByVal nEstadoOP As CaptacOrdPagoEstado)
sSQL = "UPDATE DocRecOP Set nMonto = " & nMonto & " WHERE nTpoDoc = " & TpoDocOrdenPago & " " _
    & "AND cNroDoc = '" & sNroDoc & "' And cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSQL
sSQL = "INSERT DocRecOPEst (nTpoDoc,cNroDoc,cCtaCod,cMovNro,nMonto,nEstado) " _
    & "VALUES (" & TpoDocOrdenPago & ",'" & sNroDoc & "','" & sCuenta & "','" & sMovNro & "'," & nMonto & "," & nEstadoOP & ")"
oConect.Ejecutar sSQL
End Sub

Public Sub AgregaCuentaDocumento(ByVal sCuenta As String, ByVal nTpoDoc As TpoDoc, _
        ByVal sNroDoc As String, ByVal sCodIF As String, ByVal sMovNro As String, _
        ByVal nMovNro As Long, Optional nMonto As Double = 0, _
        Optional nEstadoOP As CaptacOrdPagoEstado = gCapOPEstCobrada)
Dim sFecha As String
sFecha = Mid(sMovNro, 1, 4) & "/" & Mid(sMovNro, 5, 2) & "/" & Mid(sMovNro, 7, 2) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
Select Case nTpoDoc
    Case TpoDocCheque
        sSQL = "INSERT DocRecCapta (nTpoDoc,cNroDoc,cPersCod,cIFTpo,cCtaCod,nMonto) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCodIF & "','" & Format$(gTpoIFBanco, "00") & "','" & sCuenta & "'," & nMonto & ")"
            
            'CTI2 20190531 ADD INI
            oConect.Ejecutar sSQL
            sSQL = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
                    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
            oConect.Ejecutar sSQL
            'CTI2 20190531 ADD FIN
            
    Case TpoDocOrdenPago
        sSQL = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"

            'CTI2 20190531 ADD INI
            oConect.Ejecutar sSQL
            sSQL = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
                    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
            oConect.Ejecutar sSQL
            'CTI2 20190531 ADD FIN

    Case TpoDocNotaAbono, TpoDocNotaCargo
    
End Select
'oConect.Ejecutar sSql CTI2 20190531 Comentado
'sSql = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
'    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
'oConect.Ejecutar sSql
End Sub
Public Sub AgregaMovCap(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nMonto As Double, _
        ByVal nSaldoDisp As Double, ByVal nSaldoCnt As Double)
sSQL = "INSERT MovCap (nMovNro,cOpeCod,cCtaCod,nMonto,nSaldoDisponible,nSaldoContable) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nMonto & "," & nSaldoDisp & "," & nSaldoCnt & ")"
oConect.Ejecutar sSQL
End Sub
Public Sub AgregaMovCapDet(ByVal nMovNro As Long, ByVal nOperacion As CaptacOperacion, _
        ByVal sCuenta As String, ByVal nConcepto As CaptacConcepto, ByVal nMonto As Double)
        
sSQL = "INSERT MovCapDet (nMovNro,cOpeCod,cCtaCod,nConceptoCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & nOperacion & "','" & sCuenta & "'," & nConcepto & "," & nMonto & ")"
oConect.Ejecutar sSQL
End Sub
Public Sub ActualizaSaldoAnteriorAho(ByVal sCuenta As String, ByVal nSaldoAnt As Double)
sSQL = "Update CaptacAhorros Set nSaldoAnterior = " & nSaldoAnt & "  WHERE cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSQL
End Sub
Public Sub ActualizaAbonoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSQL = "Update Producto Set nSaldo = nSaldo + " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
oConect.Ejecutar sSQL
sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp + " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, gsFormatoFecha) & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSQL

End Sub

Public Sub ActualizaCargoCaptacion(ByVal sCuenta As String, ByVal nMontoCnt As Double, _
            ByVal nMontoDisp As Double, ByVal nIntGanado As Double, ByVal dUltMov As Date, _
            ByVal sMovNro As String, Optional bActExtracto As Boolean = True)

If bActExtracto Then
    sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & ", nTransacc = nTransacc + 1 " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
Else
    sSQL = "Update Producto Set nSaldo = nSaldo - " & nMontoCnt & " " _
        & "WHERE cCtaCod = '" & sCuenta & "'"
End If
oConect.Ejecutar sSQL
sSQL = "Update Captaciones Set nSaldoDisp = nSaldoDisp - " & nMontoDisp & ", nIntAcum = nIntAcum + " & nIntGanado _
    & ", dUltCierre = '" & Format$(dUltMov, gsFormatoFecha) & "', cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSQL
End Sub

Public Sub UltimaActualizacionCuenta(ByVal sCuenta As String, ByVal sMovNro As String)
'Actualiza la fecha de ultima actualizacion de la cuenta de captaciones
sSQL = "Update Captaciones Set cUltimaActualizacion = '" & sMovNro & "' " _
    & "WHERE cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSQL
End Sub
Public Sub ActualizaEstadoCuenta(ByVal sCuenta As String, ByVal nEstado As CaptacEstado, _
        ByVal dFecha As Date, ByVal sMovNro As String)
'Actualiza el ultimo estado a la tabla de producto
sSQL = "Update Producto Set nPrdEstado = " & nEstado & ", " _
    & "dPrdEstado = '" & Format$(dFecha, gsFormatoFecha) & "' WHERE " _
    & "cCtaCod = '" & sCuenta & "'"
oConect.Ejecutar sSQL
'Agrega un registro mas de estado a la historia de estados de la cuenta
sSQL = "INSERT CaptacEstado (cCtaCod,nPrdEstado,cMovNro) " _
    & "VALUES ('" & sCuenta & "'," & nEstado & ",'" & sMovNro & "')"
oConect.Ejecutar sSQL
End Sub
Public Sub ActualizaEstadoDocRecEst(ByVal nTipoDoc As TpoDoc, ByVal sNroDoc As String, _
        ByVal sCodIF As String, ByVal sMovNro As String, ByVal nEstado As ChequeEstado)
'Actualiza el ultimo estado a la tabla de Documento Recibidos
sSQL = "Update DocRec Set nEstado = " & nEstado & " WHERE " _
    & "nTpoDoc = '" & nTipoDoc & "' AND cNroDoc = '" & sNroDoc & "' AND cPersCod = '" & sCodIF & "'"
oConect.Ejecutar sSQL
'Agrega un registro mas de estado a la historia de estados del documento recibido
sSQL = "INSERT DocRecEst (nTpoDoc,cNroDoc,cPersCod,cMovNro,nEstado) " _
    & "VALUES ('" & nTipoDoc & "','" & sNroDoc & "','" & sCodIF & "','" & sMovNro & "'," & nEstado & ")"
oConect.Ejecutar sSQL
End Sub
Public Function GetMovExtorno(ByVal sDatoBus As String, Optional nTipoBus As Integer = 0, Optional pdFecsis As Date) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(pdFecsis, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
End If
sSQL = "Select M.cMovNro, O.cOpeDesc, C.cCtaCod, ISNULL(MD.nDocTpo,'') nDocTpo, ISNULL(MD.cDocNro,'') cDocNro, " _
    & "M.cMovDesc, C.cOpeCod, C.nMonto, M.nMovNro FROM MovDoc MD RIGHT JOIN Mov M INNER JOIN MovCap C " _
    & "INNER JOIN " & vsServerComunes & "OpeTpo O ON C.cOpeCod = O.cOpeCod ON M.nMovNro = C.nMovNro " _
    & "ON MD.nMovNro = M.nMovNro WHERE " & sWhere & " AND C.cOpeCod NOT IN ('" & gAhoEstInacAct & "') " _
    & "AND M.nMovFlag NOT IN (" & gMovFlagExtornado & "," & gMovFlagEliminado & "," & gMovFlagDeExtorno & ")" _
    & "AND M.nMovEstado IN (" & gMovEstContabMovContable & ") " _
    & "ORDER BY M.nMovNro"
Set rsMov = New Recordset
Set rsMov = oConect.CargaRecordSet(sSQL)
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
    sSQL = "UPDATE Mov SET " & lsFiltro & " WHERE nMovNro = " & pnMovNro
    oConect.Ejecutar sSQL
End If
If pnMovEstado = gMovEstContabMovContable Then
    ActualizaSaldoMovimiento psMovNro, "+"
End If
End Sub

Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
sSQL = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
oConect.Ejecutar sSQL
End Sub

Public Sub ActualizaMovOtrosItemItem(ByVal nMovNro As Long, ByVal cVariable As String, nMovItem As Long)
    sSQL = " Update MovOtrosItem" _
         & " Set nMovItem = " & nMovItem & "" _
         & " Where nMovNro = " & nMovNro & "  And cMovOtroVariable = '" & cVariable & "'"
    oConect.Ejecutar sSQL
End Sub

Public Function EsOrdenPagoEmitida(ByVal sCuenta As String, ByVal nNumOP As Long) As Boolean
Dim rsOP As Recordset
Set rsOP = New Recordset
rsOP.CursorLocation = adUseClient
sSQL = "Select OP.cCtaCod, OP.nInicio, OP.nFin FROM Mov M INNER JOIN MovDocEmitidoRango OP ON " _
    & "M.nMovNro = OP.nMovNro WHERE M.nMovFlag <> " & gMovFlagExtornado & " AND OP.cCtaCod = '" _
    & sCuenta & "' And " & nNumOP & " Between OP.nInicio And OP.nFin"

Set rsOP = oConect.CargaRecordSet(sSQL)
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
    Call RaiseError(MyUnhandledError, "DCapMov:ActualizaNotaAbonoCargo Method")
End Function
Public Function InsertaCuentaIF(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, _
                                ByVal psCtaIFCod As String, ByVal psCtaIFDesc As String, _
                                ByVal pdCtaIFAper As Date, ByVal psCtaIFCap As String, ByVal psCtaIFInt As String, _
                                ByVal pnCtaIFPlazo As Integer, ByVal pdCtaIFVenc As Date, ByVal pnCtaIFEstado As CGEstadoCtaIF, _
                                ByVal psUltAct As String, Optional psMontoEuros As Currency = 0, Optional pnTpoPF As Integer = 0, Optional pbContable As Integer = 0, Optional pnMovNroOriginal As Long = 0) As Integer
    On Error GoTo InsertaCuentaIFErr
    Dim sql As String
    InsertaCuentaIF = 1
    
    sql = "INSERT INTO CTAIF( cPersCod, cIFTpo, cCtaIFCod, cCtaIFDesc, " _
        & "                dCtaIFAper, dCtaIFCap, nCtaIFPlazo, dCtaIFVenc, cCtaIFEstado, cUltimaActualizacion , dCtaIfInt,nMontoEuros,nTpoPF,bContable,nMovNroOriginal) " _
        & " VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIFCod & "','" & psCtaIFDesc & "','" _
        & Format(pdCtaIFAper, gsFormatoFecha) & "'," & IIf(psCtaIFCap = "", "Null", "'" & psCtaIFCap & "'") & "," _
        & pnCtaIFPlazo & ",'" & Format(pdCtaIFVenc, gsFormatoFecha) & "', " & pnCtaIFEstado & ",'" & psUltAct & "'," & IIf(psCtaIFInt = "", "Null", "'" & psCtaIFInt & "'") & ", " & psMontoEuros & "," & pnTpoPF & "," & pbContable & "," & pnMovNroOriginal & ") "
    'ALPA 20150206 Se agregó el campo bContable
    oConect.Ejecutar sql
    InsertaCuentaIF = 0
    Exit Function
InsertaCuentaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaCuentaIF Method")
End Function

'RIRO20140530 ERS017 Se modificó parámetro pnInteres, se cambio de tipo currency a double
Public Function InsertaCuentaIFInteres(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, _
                                ByVal psCtaIFCod As String, ByVal pdFechaInt As Date, ByVal pnPeriodo As Integer, _
                                ByVal pnInteres As Double, ByVal psMovNro As String) As Integer
    On Error GoTo InsertaCuentaIFInteresErr
    Dim sql As String
    InsertaCuentaIFInteres = 1
    
    sql = "INSERT INTO CtaIFInteres( cPersCod, cIFTpo, cCtaIFCod, dCtaIFIntRegistro , nCtaIFIntPeriodo, nCtaIFIntValor, cUltimaActualizacion  ) " _
        & " VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIFCod & "','" _
        & Format(pdFechaInt, gsFormatoFecha) & "'," & pnPeriodo & "," & pnInteres & ",'" & psMovNro & "')"
    
    oConect.Ejecutar sql
    InsertaCuentaIFInteres = 0
    Exit Function
InsertaCuentaIFInteresErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaCuentaIFInteres Method")
End Function

Public Function EliminaCuentaIFPresup(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal psCtaIFCod As String) As Integer
    On Error GoTo EliminaCuentaIFPresupErr
    Dim sql As String
    EliminaCuentaIFPresup = 1
    sql = "DELETE CtaIFCalendario " _
        & " WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and cCtaIFCod = '" & psCtaIFCod & "' "
    oConect.Ejecutar sql
    EliminaCuentaIFPresup = 0
    Exit Function
EliminaCuentaIFPresupErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaCuentaIFPresup Method")
End Function

Public Function EliminaCuentaIFAdeudado(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal psCtaIFCod As String) As Integer
    On Error GoTo EliminaCuentaIFAdeudadoErr
    Dim sql As String
    EliminaCuentaIFAdeudado = 1
    sql = "DELETE CtaIFAdeudados " _
        & " WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and cCtaIFCod = '" & psCtaIFCod & "' "
    oConect.Ejecutar sql
    EliminaCuentaIFAdeudado = 0
    Exit Function
EliminaCuentaIFAdeudadoErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaCuentaIFAdeudado Method")
End Function

Public Function EliminaCuentaIF(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal psCtaIFCod As String) As Integer
    On Error GoTo EliminaCuentaIFErr
    Dim sql As String
    EliminaCuentaIF = 1
    sql = "DELETE CtaIF " _
        & " WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and cCtaIFCod = '" & psCtaIFCod & "' "
    oConect.Ejecutar sql
    EliminaCuentaIF = 0
    Exit Function
EliminaCuentaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaCuentaIF Method")
End Function

Public Function EliminaCuentaIFInteres(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, _
                                ByVal psCtaIFCod As String, Optional ByVal pdFechaInt As Date = 0) As Integer
    On Error GoTo EliminaCuentaIFInteresErr
    Dim sql As String
    EliminaCuentaIFInteres = 1
    
    sql = "DELETE CtaIFInteres " _
        & " WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and cCtaIFCod = '" & psCtaIFCod & "' " _
        & IIf(pdFechaInt = 0, "", "   and dCtaIFregistro = '" & Format(pdFechaInt, gsFormatoFecha) & "' ")
    oConect.Ejecutar sql
    EliminaCuentaIFInteres = 0
    Exit Function
EliminaCuentaIFInteresErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaCuentaIFInteres Method")
End Function

Public Function ActualizaCuentaIFInteres(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, _
                                ByVal psCtaIFCod As String, ByVal pdFechaInt As Date, ByVal pnPeriodo As Integer, _
                                ByVal pnInteres As Currency, ByVal psMovNro As String) As Integer
    On Error GoTo ActualizaCuentaIFInteresErr
    Dim sql As String
    ActualizaCuentaIFInteres = 1
    
    sql = "UPDATE CtaIFInteres SET nCtaIFIntPeriodo = " & pnPeriodo & ", nCtaIFIntValor = " & pnInteres & ", cUltimaActualizacion = '" & psMovNro & "' " _
        & " WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and cCtaIFCod = '" & psCtaIFCod & "' and dCtaIFIntRegistro = '" _
        & Format(pdFechaInt, gsFormatoFecha) & "' "
    oConect.Ejecutar sql
    ActualizaCuentaIFInteres = 0
    Exit Function
ActualizaCuentaIFInteresErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaCuentaIFInteres Method")
End Function

Public Function InsertaInteresCtaIF(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                    ByVal psCtaIFCod As String, _
                                    ByVal pdCtaIFIntRegistro As Date, ByVal pnCtaIFIntPeriodo As Long, _
                                    ByVal pnCtaIFIntValor As Currency, ByVal psUltAct As String) As Integer
    On Error GoTo InsertaInteresCtaIFErr
    Dim sql As String
    InsertaInteresCtaIF = 1
    
    sql = " INSERT INTO CtaIfInteres  " _
        & " (cPersCod, cIFTpo, cCtaIFCod, dCtaIFIntRegistro, nCtaIFIntPeriodo, nCtaIFIntValor, cUltimaActualizacion) " _
        & " VALUES('" & psPersCod & "','" & psIFTpo & "','" & psCtaIFCod & "','" _
        & Format(pdCtaIFIntRegistro, gsFormatoFecha) & "'," & pnCtaIFIntPeriodo & "," _
        & pnCtaIFIntValor & ",'" & psUltAct & "')"
    
    oConect.Ejecutar sql
    InsertaInteresCtaIF = 0
    Exit Function
InsertaInteresCtaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaInteresCtaIF Method")
End Function

Public Function InsertaCuentaIFFiltro(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                    ByVal psCtaIFCod As String, ByVal psCtaContCod As String, _
                                    ByVal psCtaIfSubCta As String, Optional ByVal pnUsoAgencia As Integer = 1) As Integer
                                    
    On Error GoTo InsertaCuentaIFFiltroErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    InsertaCuentaIFFiltro = 1
    'GITU
    
    '*** PEAC 20111122
    'If psCtaContCod = "51110301" Or psCtaContCod = "51210301" Or psCtaContCod = "51210303" Or psCtaContCod = "51110303" Then
    'EJVG20120113
    'If (psCtaContCod = "51110301" Or psCtaContCod = "51210301" Or psCtaContCod = "51210303" Or psCtaContCod = "51110303") And Right(psCtaIfSubCta, 2) <> "01" Then
    If (psCtaContCod = "51110301" Or psCtaContCod = "51210301" Or psCtaContCod = "51210303" Or psCtaContCod = "51110303" Or psCtaContCod = "51210302" Or psCtaContCod = "51110302") And Right(psCtaIfSubCta, 2) <> "01" Then

        'sql = "Select cCtaIfSubCta from CtaIfFiltro " _
            & "Where cperscod = '" & psPersCod & "' And Left(cCtaIfCod,2) = '03' and cIFTpo = '" & psIFTpo & "' and cCtaContCod = '" & psCtaContCod & "' " _
            & "and cCtaIFCOd = (Select Max(cCtaIFCOd) From CtaIfFiltro " _
            & "Where cperscod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaContCod = '" & psCtaContCod & "')"
        
        'sql = "exec stp_sel_BuscaFiltroCtaIF '" & psPersCod & "','" & psIFTpo & "','" & psCtaContCod & "' "
        sql = "exec stp_sel_BuscaFiltroCtaIF '" & psPersCod & "','" & psIFTpo & "','" & psCtaContCod & "','" & psCtaIfSubCta & "'" 'EJVG20120524
        
        Set rs = oConect.CargaRecordSet(sql)
        
        psCtaIfSubCta = rs!cCtaIFSubCta 'EJVG20120524
        'psCtaIfSubCta = IIf(Right(psCtaIfSubCta, 2) = "02", "01", IIf(Right(psCtaIfSubCta, 2) = "03", "02", IIf(Right(psCtaIfSubCta, 2) = "04", "05", ""))) + rs!cCtaIFSubCta
    End If
    '*** FIN PEAC

    sql = " select cCtaIFCod from CtaIFFiltro where cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIfCod='" & psCtaIFCod & "' And cCtaContCod = '" & psCtaContCod & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If RSVacio(rs) Then
        sql = "INSERT INTO CtaIFFiltro (cPersCod, cIFTpo, cCtaIfCod, cCtaContCod, cCtaIFSubCta, bUsoAgencia)" _
            & " VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIFCod & "','" _
            & psCtaContCod & "','" & psCtaIfSubCta & "'," & pnUsoAgencia & ")"
        oConect.Ejecutar sql
    End If
    InsertaCuentaIFFiltro = 0
    Exit Function
InsertaCuentaIFFiltroErr:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
'EJVG20120801 ***
Public Function InsertaCuentaIFFiltro2(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                    ByVal psCtaIFCod As String, ByVal psCtaContCod As String, _
                                    ByVal psCtaIfSubCta As String, Optional ByVal pnUsoAgencia As Integer = 1) As Integer
                                    
    On Error GoTo InsertaCuentaIFFiltroErr
    Dim sql As String
    Dim rs As ADODB.Recordset

    InsertaCuentaIFFiltro2 = 1
    sql = " select cCtaIFCod from CtaIFFiltro where cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIfCod='" & psCtaIFCod & "' And cCtaContCod = '" & psCtaContCod & "'"
    Set rs = oConect.CargaRecordSet(sql)
    If RSVacio(rs) Then
        sql = "INSERT INTO CtaIFFiltro (cPersCod, cIFTpo, cCtaIfCod, cCtaContCod, cCtaIFSubCta, bUsoAgencia)" _
            & " VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIFCod & "','" _
            & psCtaContCod & "','" & psCtaIfSubCta & "'," & pnUsoAgencia & ")"
        oConect.Ejecutar sql
    End If
    InsertaCuentaIFFiltro2 = 0
    Exit Function
InsertaCuentaIFFiltroErr:
   Err.Raise Err.Number, Err.Source, Err.Description
End Function
'END EJVG *******
Public Function EliminaCuentaIFFiltroPersona(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, Optional ByVal psCtaIFCod As String) As Integer
    On Error GoTo EliminaCuentaIFErr
    Dim sql As String
    EliminaCuentaIFFiltroPersona = 1
    sql = "DELETE CtaIFFiltro " _
        & " WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and cCtaIFCod = '" & psCtaIFCod & "' "
    oConect.Ejecutar sql
    EliminaCuentaIFFiltroPersona = 0
    Exit Function
EliminaCuentaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaCuentaIFFiltro Method")
End Function


Public Function ActualizaSaldoCtaIF(ByVal pdFecha As Date, _
                                    ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
                                    ByVal psCtaIFCod As String) As Integer
                                    
    On Error GoTo ActualizaSaldoCtaIFErr
    Dim sql As String
    ActualizaSaldoCtaIF = 1
    
    sql = "sp_ActualizaSaldoCtaIF '" & Format(pdFecha, gsFormatoFecha) & "','" _
                                    & psPersCod & "','" _
                                    & Format(pnIFTpo, "00") & "', '" & psCtaIFCod & "'"
    
    oConect.Ejecutar sql
    ActualizaSaldoCtaIF = 0
    Exit Function
ActualizaSaldoCtaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaSaldoCtaIF Method")
End Function
Public Function ActualizaCtaIF(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                  ByVal psCtaIFCod As String, _
                                  Optional ByVal psCtaIFDesc As String = "", _
                                  Optional ByVal pdCtaIFAper As String = "", _
                                  Optional ByVal pdCtaIFCap As String = "", _
                                  Optional ByVal pdCtaIFInt As String = "", _
                                  Optional ByVal pnCtaIFPlazo As Integer = -1, _
                                  Optional ByVal pdCtaIFVenc As String = "", _
                                  Optional ByVal pnCtaIFEstado As CGEstadoCtaIF = -1, _
                                  Optional ByVal pnInteres As Currency = -999999, _
                                  Optional ByVal psUltAct As String = "", Optional ByVal pbAcumulaInteres As Boolean = False, _
                                  Optional ByVal pnMontoEuros As Currency, Optional ByVal pbContable As Integer = -1) As Integer
                                  
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
    If pdCtaIFVenc <> "" Then
        lsFiltro = lsFiltro + " dCtaIFVenc = '" & pdCtaIFVenc & "', "
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
        If pbAcumulaInteres Then
            lsFiltro = lsFiltro + " nInteres = nInteres + " & pnInteres & ", "
        Else
            lsFiltro = lsFiltro + " nInteres =" & pnInteres & ", "
        End If
    End If
       
    If psUltAct <> "" Then
        lsFiltro = lsFiltro + " cUltimaActualizacion ='" & psUltAct & "',"
    End If
    'ALPA 20150209******************************
    If pbContable <> -1 Then
        lsFiltro = lsFiltro + " bContable ='" & pbContable & "',"
    End If
    '*******************************************
    If pnMontoEuros <> 0 Then
        lsFiltro = lsFiltro + " nMontoEuros = " & pnMontoEuros & " "
    End If
    If lsFiltro <> "" Then
        lsFiltro = Mid(Trim(lsFiltro), 1, Len(Trim(lsFiltro)) - 1)
        sql = "UPDATE CTAIF SET " & lsFiltro _
            & " WHERE cPersCod = '" & psPersCod & "' AND cIFTpo='" & Format(psIFTpo, "00") & "' AND cCtaIFCod='" & psCtaIFCod & "'"
    End If
    oConect.Ejecutar sql
    ActualizaCtaIF = 0
    Exit Function
ActualizaCuentaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaCuentaIF Method")
End Function

Public Sub ActualizaCtaIFAdeudado(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, ByVal psCtaIFCod As String, _
                                Optional ByVal pnMontoPrestado As Currency = -1, _
                                Optional ByVal pnSaldoCap As Currency = -1, Optional ByVal pnNroCuotas As Integer = -1, Optional ByVal pnPeriodoGracia As Currency = -1, Optional ByVal pnComisionInicial As Currency = -1, Optional ByVal pnComisionMonto As Currency = -1, _
                                Optional ByVal psPlaza As String = "", Optional ByVal pnCuotaPagoCap As Integer = -1, Optional ByVal psVencimiento As String = "", _
                                Optional ByVal pnEjecutado As Integer = -1, Optional ByVal psMonedaPago As String = "", Optional ByVal psCuotaUltPago As String = "", Optional ByVal pnTpoCuota As CGAdeudCalTpoCuota = -1, Optional pnTramo As Currency = -1, _
                                Optional pnComisionCuota As Currency = -1, Optional psLinCredCod As String = "", Optional pnFechaFija As Integer = -1, Optional pbAfectaSaldoCap As Boolean = False, _
                                Optional pnSaldoCapLP As Currency = -1)
Dim lsDatos As String
Dim sql As String
On Error GoTo ErrorUpdateCtaIFAdeud
lsDatos = ""
If pnMontoPrestado <> -1 Then
    lsDatos = lsDatos & " nMontoPrestado=" & pnMontoPrestado & ","
End If

If pnSaldoCap <> -1 Then
    If pbAfectaSaldoCap Then
        lsDatos = lsDatos & " nSaldoCap = nSaldoCap - " & pnSaldoCap & ","
    Else
        lsDatos = lsDatos & " nSaldoCap = " & pnSaldoCap & ","
    End If
End If
If pnPeriodoGracia <> -1 Then
    lsDatos = lsDatos & " nPeriodoGracia=" & pnPeriodoGracia & ","
End If
If pnComisionInicial <> -1 Then
    lsDatos = lsDatos & " nComisionInicial= " & pnComisionInicial & ","
End If
If pnComisionMonto <> -1 Then
    lsDatos = lsDatos & " nComisionMonto= " & pnComisionMonto & ","
End If
If psPlaza <> "" Then
    lsDatos = lsDatos & " cPlaza= '" & psPlaza & "',"
End If
If pnCuotaPagoCap <> -1 Then
    lsDatos = lsDatos & " nCuotaPagoCap=" & pnCuotaPagoCap & ","
End If
If psVencimiento <> "" Then
    lsDatos = lsDatos & " dVencimiento='" & psVencimiento & "',"
End If
If pnEjecutado <> -1 Then
    lsDatos = lsDatos & "bEjecutado=" & pnEjecutado & ","
End If
If psMonedaPago <> "" Then
    lsDatos = lsDatos & "cMonedaPago= '" & psMonedaPago & "',"
End If
If psCuotaUltPago <> "" Then
    lsDatos = lsDatos & "dCuotaUltPago = '" & psCuotaUltPago & "',"
End If
If pnTpoCuota <> -1 Then
    lsDatos = lsDatos & " nTpoCuota = " & pnTpoCuota & ","
End If
If pnTramo <> -1 Then
    lsDatos = lsDatos & " nTramoConcesion = " & pnTramo & ","
End If
If pnComisionCuota <> -1 Then
    lsDatos = lsDatos & " nComisionCuota = " & pnComisionCuota & ","
End If
If psLinCredCod <> "" Then
    lsDatos = lsDatos & " cCodLinCred = '" & psLinCredCod & "',"
End If
If pnFechaFija <> -1 Then
    lsDatos = lsDatos & " nFechaFija = " & pnFechaFija & ","
End If
If pnNroCuotas <> -1 Then
    lsDatos = lsDatos & " nCtaIFCuotas = " & pnNroCuotas & ","
End If
If pnSaldoCapLP <> -1 Then
    lsDatos = lsDatos & " nSaldoCapLP = " & pnSaldoCapLP & ","
End If
If lsDatos <> "" Then
    lsDatos = Left(lsDatos, Len(lsDatos) - 1)
    sql = "UPDATE CTAIFADEUDADOS SET " & lsDatos _
        & " WHERE  cPersCod = '" & psPersCod & "' and cIFTpo = '" & Format(psIFTpo, "00") & "' and cCtaIfCod = '" & psCtaIFCod & "'"
    oConect.Ejecutar sql
End If
Exit Sub
ErrorUpdateCtaIFAdeud:
   Err.Raise Err.Number, "DMov", Err.Description
End Sub

Public Sub ActualizaSaldoMovimiento(sMovNro As String, sSimbolo As String)
 Dim sql As String
 On Error GoTo ErrActualizaMovSaldo
' Dim prs As adodb.Recordset
 Dim ldFecha As Date
' Dim ldFecSis As Date
 
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
    sql = "Exec spActualizaMovSaldo " & lsMov & ", '" & Format(ldFecha, gsFormatoFecha) & "', '" & sSimbolo & "' "
    oConect.Ejecutar sql
    
    sql = "DELETE tmpMovActualiza"
    oConect.Ejecutar sql

 Else
    sql = "INSERT tmpMovActualiza (cMovNro) VALUES ('" & sMovNro & "')"
    oConect.Ejecutar sql
    
    ldFecha = GetFechaMov(Left(sMovNro, 8), True)
    sql = "spActualizaMovSaldo '" & sMovNro & "', '" & Format(ldFecha, gsFormatoFecha) & "', '" & sSimbolo & "' "
    'oConect.CommandTimeOut 3500
    oConect.Ejecutar sql
    
    sql = "DELETE tmpMovActualiza"
    oConect.Ejecutar sql
 End If

Exit Sub
ErrActualizaMovSaldo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'' Nuevo Pepe Trujillo
Public Sub ActualizaColocAdeudadoSaldo(psLinCredCod As String, pdFecha As Date, pnMonto As Currency)
Dim sql As String
sql = "UPDATE ColAdeudadoSaldo SET nSaldo = nSaldo + " & pnMonto & " WHERE cLineaCred = '" & psLinCredCod & "' and dFecha >= '" & Format(pdFecha, gsFormatoFecha) & "' "
oConect.Ejecutar sql
End Sub

'' Nuevo Pepe Trujillo
Public Sub InsertaDatosAdeudados(prsAdeud As ADODB.Recordset, psPersCod As String, psIFTpo As CGTipoIF, psCodAdeudado As String, pnCapital As Currency, pnNroCuotas As Integer, pnGracia As Currency, pnComisionIni As Currency, pnComisionMonto As Currency, psInterno As String, pnCuotaPagok As Integer, _
                psFecVencimiento As String, pnMoneda As Integer, psFecUltPago As String, ByVal pnTpoCuota As CGAdeudCalTpoCuota, pnTramo As Currency, _
                ByVal pnComisionCuota As Currency, psLinCredCod As String, pnFechaFija As Integer, ByVal psMovUltAct As String, Optional pnMalPg As Integer = 0, Optional pnCapitalVac As Currency = 0, Optional ByVal pnPorcentAfect As Currency = 0, Optional ByVal pnMontoEuros As Currency, Optional pnMontoPresConce As Currency = 0)

    Dim lbExiste As Boolean
    Dim lsTipoCuota As String
    Dim oDAdeudCal As New DAdeudCal
    Dim lnCapPagado As Currency
    'ALPA20130617*************************************
    Dim lnCapitalConcesionado As Currency
    Dim lbGrabarCapitalConcesionado As Boolean
    lnCapitalConcesionado = 0
    lbGrabarCapitalConcesionado = False
    
    If Not prsAdeud Is Nothing Then
    Do While Not prsAdeud.EOF
        If lnCapitalConcesionado = 0# Then
            lbGrabarCapitalConcesionado = True
        End If
        If prsAdeud!cEstado = "0" Then
            lnCapitalConcesionado = lnCapitalConcesionado + prsAdeud!CapitalCon
        End If
        
     prsAdeud.MoveNext
    Loop
    prsAdeud.MoveFirst
    End If
    
    '*************************************************
    oDAdeudCal.Inicio oConect.ConexionActiva
    lbExiste = oDAdeudCal.VerificaCtaIFAdeudado(psPersCod, psIFTpo, psCodAdeudado)
    If pnCapitalVac = 0 Then
        pnCapitalVac = pnCapital
    End If
    If lbExiste Then
        'Insertamos Historico Mal Pagador
        oDAdeudCal.InsertaMalPagadorAdeudado psPersCod, psIFTpo, psCodAdeudado, pnMalPg, psMovUltAct

        'oDAdeudCal.EliminarCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado
        oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado, pnCapital, IIf(Not prsAdeud Is Nothing, pnCapitalVac, -1), pnNroCuotas, pnGracia, pnComisionIni, _
                pnComisionMonto, psInterno, pnCuotaPagok, psFecVencimiento, "1", pnMoneda, psFecUltPago, pnTpoCuota, pnTramo, pnComisionCuota, psLinCredCod, pnFechaFija, , , pnPorcentAfect, pnMalPg, pnMontoEuros, pnMontoPresConce, lnCapitalConcesionado, lbGrabarCapitalConcesionado
    Else
        'Insertamos Historico Mal Pagador
        oDAdeudCal.InsertaMalPagadorAdeudado psPersCod, psIFTpo, psCodAdeudado, pnMalPg, psMovUltAct

        oDAdeudCal.InsertaCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado, pnCapital, pnCapitalVac, pnNroCuotas, _
                pnGracia, pnComisionIni, pnComisionMonto, psInterno, pnCuotaPagok, _
                psFecVencimiento, "1", pnMoneda, psFecUltPago, pnTpoCuota, pnTramo, pnComisionCuota, psLinCredCod, pnFechaFija, pnPorcentAfect, pnMalPg, pnMontoEuros, pnMontoPresConce, lnCapitalConcesionado
    End If



    'GRABAMOS EL CALENDARIO DE LOS ADEUDADOS
    Dim lbCuotaPend As Boolean
    Dim sql As String
    lbCuotaPend = False
    lnCapPagado = 0
    If Not prsAdeud Is Nothing Then
        If Not prsAdeud.EOF Then
            lbExiste = oDAdeudCal.VerificaCtaIFAdeudadoCalendario(psPersCod, psIFTpo, psCodAdeudado)
            If lbExiste Then
'                oDAdeudCal.EliminarCalendario psPersCod, psCodAdeudado
                    sql = "Update CtaIFCalendario Set bVigente=0"
                    sql = sql & " Where cPersCod='" & psPersCod & "' And cIFTpo=" & psIFTpo & " And cCtaIFCod = '" & psCodAdeudado & "' And bVigente = 1"
                    oConect.Ejecutar sql
            End If
            'GRABAMOS LOS POSIBLES PAGOS A REALIZARSE O PAGOS PROYECTADOS
            If Not RSVacio(prsAdeud) Then
                Do While Not prsAdeud.EOF
                    'TipoCuota  : ---> A  Cuota Adelantada  ----> C Cuota Normal
                    'gCGTipoCuotCalIFCuota = 2
                    'gCGTipoCuotCalIFDesem = 1

                    'lsTipoCuota = IIf(prsAdeud.Fields(1) = 0, "A", "C")
                    lsTipoCuota = prsAdeud!cTpoCuota
                    
                    oDAdeudCal.InsertCalendario psPersCod, psIFTpo, psCodAdeudado, lsTipoCuota, prsAdeud.Fields(1), prsAdeud.Fields(0), prsAdeud!cEstado, 0, prsAdeud!Capital, prsAdeud!Interes, prsAdeud!Comisión, nVal(prsAdeud!nInteresPagado), nVal(prsAdeud!Total), prsAdeud!cMovNro, 1, psMovUltAct, prsAdeud!CapitalCon, prsAdeud!InteresCon, prsAdeud!ComisionC
                    If prsAdeud!cEstado = gTpoEstCuotaAdeudPend And lbCuotaPend = False And lsTipoCuota = gCGTipoCuotCalIFCuota Then 'Si es primera Cuota
                       lbCuotaPend = True
                       'Actualizamos Fecha de Vencimiento del Proxima Cuota
                       oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado, , , _
                                                       , , , , , , Format(prsAdeud.Fields(0), gsFormatoFecha), , , , , , , , , , , , , , pnMontoPresConce, lnCapitalConcesionado
                    ElseIf prsAdeud!cEstado = gTpoEstCuotaAdeudCanc Then
                        lnCapPagado = lnCapPagado + prsAdeud!Capital
                    End If
                    prsAdeud.MoveNext
                Loop
            End If
            If lnCapPagado > 0 Then
                oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCodAdeudado, , pnCapitalVac - lnCapPagado, , , , , , , , , , , , , , , , , , , , , pnMontoPresConce
            End If
        End If
    End If
    Set oDAdeudCal = Nothing
End Sub

'' Nuevo Pepe Trujillo
Public Sub ActualizaAdeudadosProvision(psPersCod As String, psIFTpo As String, psCtaIFCod As String, psFecha As String, pnInteresPag As Currency, pnInteresPagReal As Currency, pnNroCuota As Integer, psMovNro As String, pnNroCuota6 As Integer)
Dim oDAdeudCal As New DAdeudCal
    oDAdeudCal.Inicio oConect.ConexionActiva
    oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCtaIFCod, , , , , , , , , , , , Format(psFecha, "yyyy/mm/dd")
    oDAdeudCal.UpdateCalendario psPersCod, psIFTpo, psCtaIFCod, pnNroCuota, gCGTipoCuotCalIFCuota, pnInteresPag, pnInteresPagReal, psMovNro
    Set oDAdeudCal = Nothing
End Sub

'' Nuevo Pepe Trujillo
Public Sub ActualizaAdeudadosPagoCuota(psPersCod As String, psIFTpo As String, psCtaIFCod As String, psFecha As String, pnCapital As Currency, pnInteresPag As Currency, pnNroCuota As Integer, psMovNro As String, Optional pbCancela As Boolean = False, Optional pnNroCuota6 As String = "")
Dim oDAdeudCal As New DAdeudCal
    oDAdeudCal.Inicio oConect.ConexionActiva
    oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCtaIFCod, , pnCapital, , , , , , , , , , Format(psFecha, "yyyy/mm/dd"), , , , , , True
    oDAdeudCal.UpdateCalendario psPersCod, psIFTpo, psCtaIFCod, pnNroCuota, gCGTipoCuotCalIFCuota, pnInteresPag, 0, psMovNro, gTpoEstCuotaAdeudCanc, pbCancela, True, pnNroCuota6 ', IIf(pbCancela = True, False, True)
    If pbCancela Then
        ActualizaCtaIF psPersCod, psIFTpo, psCtaIFCod, , , , , , , gEstadoCtaIFCancelada
    End If
Set oDAdeudCal = Nothing
End Sub
    
'' Nuevo Pepe Trujillo
Public Sub ActualizaAdeudadosSaldoCap(psPersCod As String, psIFTpo As String, psCtaIFCod As String, Optional pnSaldoCap As Currency = -1, Optional pnSaldoCapLP As Currency = -1)
Dim oDAdeudCal As New DAdeudCal
    oDAdeudCal.Inicio oConect.ConexionActiva
    oDAdeudCal.UpdateCtaIFAdeudado psPersCod, psIFTpo, psCtaIFCod, , pnSaldoCap, , , , , , , , , , , , , , , , True, pnSaldoCapLP
Set oDAdeudCal = Nothing
End Sub

'' Nuevo Pepe Trujillo
Public Function CancelaCuotaConcesional(psPersCod As String, psIFTpo As String, psCtaIFCod As String, psFecha As String, psTpoCuota As CGTipoCuotCalIF, psMovNro As String) As Currency
Dim oDAdeudCal As New DAdeudCal
    CancelaCuotaConcesional = 0
    oDAdeudCal.Inicio oConect.ConexionActiva
    sSQL = "SELECT nNroCuota, nCapital FROM CtaIFCalendario WHERE cPersCod = '" & psPersCod & "' and cIFTpo = '" & psIFTpo & "' and cCtaIFCod = '" & psCtaIFCod & "' and dVencimiento <= '" & psFecha & "' and cEstado = '" & gTpoEstCuotaAdeudPend & "' and cTpoCuota = '" & CGTipoCuotCalIF.gCGTipoCuotCalIFNoConcesional & "'"
    Set rs = oConect.CargaRecordSet(sSQL)
    If Not rs.EOF Then
        oDAdeudCal.UpdateCalendario psPersCod, psIFTpo, psCtaIFCod, rs!nNroCuota, psTpoCuota, 0, psMovNro, gTpoEstCuotaAdeudCanc
        CancelaCuotaConcesional = CancelaCuotaConcesional + rs!nCapital
    End If
    RSClose rs
Set oDAdeudCal = Nothing
End Function

'' Fin Nuevos Pepe Trujillo

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

Public Sub GrabaMovimientoEfectivoNew(ByVal pnMovNro As Long, ByRef pnMovItem As Long, rsEfectivo As ADODB.Recordset, psCuenta As String, psDH As String)
On Error GoTo GrabaMovimientoEfectivoErr
Dim lnMovOrden As Integer
'Guardamos el Billetaje
If Not rsEfectivo Is Nothing Then
    If rsEfectivo.State = adStateOpen Then
        If Not rsEfectivo.EOF And Not rsEfectivo.BOF Then
            Do While Not rsEfectivo.EOF
                If rsEfectivo!Monto <> 0 Then
                    pnMovItem = pnMovItem + 1: lnMovOrden = 0
                    InsertaMovCtaPer pnMovNro, Format(pnMovItem, "#0"), psCuenta, rsEfectivo!Monto * IIf(psDH = "D", 1, -1)
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
                        Select Case prsObj!ObjPadre
                            Case ObjCMACAgencias
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
                                InsertaMovObjAgenciaArea pnMovNro, pnMovItem, lnMovOrden, prsObj!Código, ""
                            Case ObjCMACArea
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
                                InsertaMovObjAgenciaArea pnMovNro, pnMovItem, lnMovOrden, "", prsObj!Código
                            Case ObjCMACAgenciaArea
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
                                InsertaMovObjAgenciaArea pnMovNro, pnMovItem, lnMovOrden, Mid(prsObj!Código, 4, 2), Mid(prsObj!Código, 1, 3)
                            Case ObjDescomEfectivo
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
                                InsertaMovObjEfectivo pnMovNro, pnMovItem, lnMovOrden, prsObj!Código
                            Case ObjEntidadesFinancieras
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObj pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
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

Public Function GetFactorCreditoFiscal(pdIni As Date, pdFin As Date) As Double
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sql = "Select ISNULL( dbo.GetPorcentajeGrabado('" & Format(pdIni, gsFormatoMovFecha) & "','" & Format(pdFin, gsFormatoMovFecha) & "'),0)"
    oCon.AbreConexion
    
    GetFactorCreditoFiscal = oCon.CargaRecordSet(sql).Fields(0)
End Function

Public Function GetFactorCreditoFiscalCuadro(pdIni As Date, pdFin As Date) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    sql = "Select Periodo, IngGrabado, IngNoGrabado, IngTotal From dbo.GetIngresos('" & Format(pdIni, gsFormatoMovFecha) & "','" & Format(pdFin, gsFormatoMovFecha) & "')"
    oCon.AbreConexion
    
    Set GetFactorCreditoFiscalCuadro = oCon.CargaRecordSet(sql)
End Function


Public Function GetMovCredFiscal(pdFecha As Date, pnFactor As Double, psOpeCod As String) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Dim lsCtaIGV As String
    Dim oConst As NConstSistemas
    Set oConst = New NConstSistemas
    Set oCon = New DConecta
    lsCtaIGV = oConst.LeeConstSistema(gConstCtaIGV)
    sql = " Select MC.cCtaContCod, Sum(Round(nMovOtroImporte * " & pnFactor & ",2)) Monto From Mov M" _
        & " Inner Join MovCta MC On M.nMovNro = MC.nMovNro" _
        & " Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
        & " Inner Join MovOtrosItem MO On M.nMovNro = MO.nMovNro And MC.nMovItem = MO.nMovItem" _
        & " Where MO.cMovOtroVariable = '" & lsCtaIGV & "' And MC.cCtaContCod Not Like '18%' And M.nMovEstado = 10 And M.nMovFlag  = 0" _
        & " And ndoctpo In (Select nDocTpo From OpeDoc Where cOpeCod = '" & psOpeCod & "') And MC.cCtaContCod <> '" & lsCtaIGV & "' " _
        & " And M.cMovNro Like '" & Format(pdFecha, "yyyymm") & "%' And m.nmovNro In" _
        & "     (Select M.nMovNro From Mov M" _
        & "         Inner Join MovCta MC On M.nMovNro = MC.nMovNro" _
        & "         Inner Join MovDoc MD On M.nMovNro = MD.nMovNro" _
        & "         Where MC.cCtaContCod = '" & lsCtaIGV & "' And M.nMovEstado = " & MovEstado.gMovEstContabMovContable & " And M.nMovFlag  = " & MovFlag.gMovFlagVigente & "" _
        & "         And nDocTpo In (Select nDocTpo From OpeDoc Where cOpeCod = '" & psOpeCod & "')) Group By MC.cCtaContCod"
    oCon.AbreConexion
    
    Set GetMovCredFiscal = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set oConst = Nothing
End Function

Public Function GetMovCredFiscalAF(pdIni As Date, pdFin As Date, pnFactor As Double) As ADODB.Recordset
    Dim sql As String
    Dim oCon As DConecta
    Dim lsCtaIGV As String
    Dim oConst As New NConstSistemas
    Set oCon = New DConecta
    
    sql = " Select Sum(Round(nIGV * " & pnFactor & ",2)), Cta" _
        & " From" _
        & " ( Select cBSCod, nBSValor, nIGV," _
        & "    (Select cCtaContCodD From CtaBS CBS" _
        & "        Inner Join OpeCtaCta OCC On CBS.cOpeCod = OCC.cOpeCod And OCC.cCtaContCodH = CBS.cCtaContCod" _
        & "        Where OCC.cOpeCod = '" & gnAlmaSalXAtencion & "' And BSA.cBSCod Like cObjetoCod + '%') Cta From BSActivoFijo BSA" _
        & "    Where dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & "'" _
        & "    And nAnio = " & Year(pdIni) & " And ltrim(rtrim(str(nMovNro)))+cBSCod+cSerie In" _
        & "        (Select Ltrim(Rtrim(Str(nMovNro)))+cBSCod+cSerie" _
        & "            From BSControlSal BSCS" _
        & "            Where BSCS.cBsCod Like '112%' And dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & "')) As BBSS" _
        & "    Group By Cta"
    
    oCon.AbreConexion
    
    Set GetMovCredFiscalAF = oCon.CargaRecordSet(sql)
    oCon.CierraConexion
    Set oCon = Nothing
    Set oConst = Nothing
End Function

Public Function SetCredFiscalAF(pdIni As Date, pdFin As Date, pnFactor As Double) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Dim lsCtaIGV As String
    Dim oConst As New NConstSistemas
    Set oCon = New DConecta
    
    sql = " Update BSActivoFijo" _
        & " Set nBSValor = nBSValor + Round(nIGV * " & pnFactor & ", 2)" _
        & " Where dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & "'" _
        & "     And nAnio = " & Year(pdIni) & " And ltrim(rtrim(str(nMovNro)))+cBSCod+cSerie In" _
        & "         (Select Ltrim(Rtrim(Str(nMovNro)))+cBSCod+cSerie" _
        & "             From BSControlSal BSCS" _
        & "             Where BSCS.cBsCod Like '112%' And" _
        & "             dActivacion Between '" & Format(pdIni, gsFormatoFecha) & "' And '" & Format(pdFin, gsFormatoFecha) & "')"
    
    oCon.AbreConexion
    oCon.Ejecutar sql
    oCon.CierraConexion
    Set oCon = Nothing
    Set oConst = Nothing
End Function


Public Function VerfRevCredFiscal(psOpeCod As String, pdFecha As Date) As Boolean
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    oCon.AbreConexion
    
    sql = "Select nMovNro From Mov Where cOpeCod = '" & psOpeCod & "' And cMovNro Like '" & Format(pdFecha, gsFormatoMovFecha) & "%' ANd nMovFlag = " & MovFlag.gMovFlagVigente & " And nMovEstado = " & MovEstado.gMovEstContabMovContable & ""
    Set rs = oCon.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        VerfRevCredFiscal = False
    Else
        VerfRevCredFiscal = True
    End If
    
    rs.Close
End Function

Public Function GetConexion() As ADODB.Connection
    Set GetConexion = oConect.ConexionActiva
End Function

Public Function InsertaAsientoRuta(ByVal psRuta As String)
                           
                           
On Error GoTo InsertaMovErr
    Dim sql As String
    sql = "delete from FINANC_ASIENTOS "
    oConect.Ejecutar sql
    
    sql = "INSERT INTO FINANC_ASIENTOS (cdirectorio)" _
        & " VALUES('" & psRuta & "')"
    oConect.Ejecutar sql
    Exit Function
InsertaMovErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaAsientoRuta Method")
End Function


Public Function GetCuentaPagoAdeudado(ByVal psCtaContCodRef As String) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    On Error GoTo BuscarMovErr
    
    sql = "SELECT cCtaContCodGastosFinan From CtaIFAdeudadosRef Where cCtaContCod =  '" & psCtaContCodRef & "'"
    Set rs = oConect.CargaRecordSet(sql)
    
    If Not rs.EOF And Not rs.BOF Then
        GetCuentaPagoAdeudado = rs.Fields(0)
    Else
        GetCuentaPagoAdeudado = ""
    End If
    
    rs.Close
    Set rs = Nothing
    Exit Function
BuscarMovErr:
    Call RaiseError(MyUnhandledError, "DMov:BuscarMov Method")
End Function

Public Function GetDatosCheque(ByVal psCodCheque As String) As Recordset
    Dim sSQL As String
    Dim oConec As DConecta
    
    sSQL = "Select MD.nMovNro,MD.cDocNro as cCheque,Pers.cPersNombre,AG.cAgeDescripcion as cAgencia,           "
    sSQL = sSQL & "            MOV.nMovImporte as nMonto,MOV.nMoneda,Ag.cAgeCod,MOI.cCtaIfCod"
    sSQL = sSQL & " From MovDoc MD"
    sSQL = sSQL & "           Inner Join MovObjIF MOI on MOI.nMovNro=MD.nMovNro"
    sSQL = sSQL & "           Inner Join Persona Pers on Pers.cPersCod=MOI.cPersCod"
    sSQL = sSQL & "           Inner Join MovObjAreaAgencia MOA on MOA.nMovNro=MD.nMovNro"
    sSQL = sSQL & "           Inner Join Agencias  AG on AG.cAgeCod=MOA.cAgeCod"
    sSQL = sSQL & "           Inner Join MovOpeVarias MOV on MOV.nMovNro=MD.nMovNro and MOV.cNroDoc= MD.cDocNro"
    sSQL = sSQL & "           Inner Join Mov M on M.nMovNro=MD.nMovNro"
    sSQL = sSQL & "  Where MD.cDocNro='" & psCodCheque & "' and MD.nDocTpo=47 and M.nMovFlag=0 and Not Exists(Select * From MovRef Where nMovNroRef=M.nMovNro)"
    
    Set oConec = New DConecta
    oConec.AbreConexion
    Set GetDatosCheque = oConec.CargaRecordSet(sSQL)
    oConec.CierraConexion
    Set oConec = Nothing
End Function

Public Function ActualizaOpeCheq(ByVal lnMovNro As Long, ByVal pnMovItem As Integer, ByVal nMovObjOrden As Integer, ByVal cAreaCod As String, _
                            ByVal cAgeCod As String) As Integer
    On Error GoTo InsertaMovObjErr
    Dim psSql As String
   Dim nMovNro As Long
   
    ActualizaOpeCheq = 1

    psSql = "UPDATE MovObjAreaAgencia Set cAgeCod = '" & cAgeCod & "' Where nmovnro=" & lnMovNro & "  "

    oConect.Ejecutar psSql
    ActualizaOpeCheq = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
End Function


Public Function ActualizaMovCheq(ByVal lnMovNro As Long, ByVal txtglosa As String) As Integer
    On Error GoTo InsertaMovObjErr
    Dim psSql As String
   Dim nMovNro As Long
   
    ActualizaMovCheq = 1

    psSql = "UPDATE Mov Set cMovdesc = '" & txtglosa & "' Where nmovnro=" & lnMovNro & "  "

    oConect.Ejecutar psSql
    ActualizaMovCheq = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
End Function



Public Function AnulaMovCheq(ByVal lnMovNro As Long) As Integer
    On Error GoTo InsertaMovObjErr
    Dim psSql As String
   Dim nMovNro As Long
   
    AnulaMovCheq = 1

    psSql = "UPDATE Mov Set nMovflag = 1 Where nmovnro=" & lnMovNro & "  "

    oConect.Ejecutar psSql
    AnulaMovCheq = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
End Function

Public Function InsertaMovModificaNoCont(ByVal pnMovNro As Long, ByVal psMovNroMod As String, ByVal psPersCod As String, ByVal psMovDes As String) As Integer
    On Error GoTo InsertaMovPendientesRendErr
    Dim sql As String
    InsertaMovModificaNoCont = 1
    
    sql = " Insert MovModificaNoCont(nMovNro, cMovNroMod, cPerscod, cMovDesc)" _
        & " Values (" & pnMovNro & ",'" & psMovNroMod & "','" & psPersCod & "','" & psMovDes & "')"
    oConect.Ejecutar sql
    
    InsertaMovModificaNoCont = 0
    Exit Function
InsertaMovPendientesRendErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovPendientesRend Method")
End Function
Public Sub ActualizaMovGasto(ByVal pnMovNro As Long, ByVal psPersCod As String)
    On Error GoTo ActualizaMovErr
    Dim sql As String
    Dim rsPers As ADODB.Recordset
    Set rsPers = New ADODB.Recordset
    
    sql = " Select nMovNro From MovGasto  Where nMovNro = " & pnMovNro
    Set rsPers = oConect.CargaRecordSet(sql)
    
    If rsPers.EOF And rsPers.BOF Then
        sql = " insert movgasto (nMovNro, cPersCod, cPersCodDest, cDestino)" _
            & " values(" & pnMovNro & ",'" & psPersCod & "','','')"
    Else
        sql = " Update MovGasto " _
            & " Set cPersCod = '" & psPersCod & "' " _
            & " Where nMovNro = " & pnMovNro
    End If
    
    oConect.Ejecutar sql
    
    rsPers.Close
    Set rsPers = Nothing
    
    Exit Sub
ActualizaMovErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMov Method")
End Sub

Public Sub ActualizaMovDescripcion(ByVal pnMovNro As Long, ByVal psMovDescripcion As String)
    On Error GoTo ActualizaMovErr
    
    Dim sql As String
    
    sql = " Update Mov " _
        & " Set cMovDesc = '" & psMovDescripcion & "' " _
        & " Where nMovNro = " & pnMovNro
    
    oConect.Ejecutar sql
    
    Exit Sub
ActualizaMovErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMov Method")
End Sub

'##ModelId=3A848F4F002E
Public Function UpdateDocNoCont(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pcMovOtroVariable As String, ByVal pnMovOtroImp As Double, ByVal pcMovOtroTexto As String) As Integer
    On Error GoTo InsertaMovDocErr
    Dim sql As String
    UpdateDocNoCont = 1
    
    sql = "UPDATE MovOtrosItem set nMovOtroImporte = " & pnMovOtroImp & ", cMovOtroTexto = '" & pcMovOtroTexto & "'" & _
       " where NMovNro = " & pnMovNro & " and nMovItem = '" & pnMovItem & "' and cMovOtroVariable = '" & pcMovOtroVariable & "'"
        
    oConect.Ejecutar sql
    UpdateDocNoCont = 0
    Exit Function
InsertaMovDocErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovDoc Method")
End Function


Public Function InsertaMovModificaNoContDoc(ByVal pnMovNro As Long, ByVal psMovNroMod As String) As Integer
    On Error GoTo InsertaMovPendientesRendErr
    Dim sql As String
    
    InsertaMovModificaNoContDoc = 1
    
    sql = " Insert MovModificaNoContDoc (nMovNro, cMovNroMod, nDocTpo, cDocNro, dDocFecha) " _
        & " Select nMovNro, '" & psMovNroMod & "' MovDodMod , nDocTpo, cDocNro, dDocFecha from MovDoc Where nMovNro  = " & pnMovNro & " "
    oConect.Ejecutar sql
    
    sql = " Delete MovDoc Where nMovNro = " & pnMovNro
    oConect.Ejecutar sql
    
    InsertaMovModificaNoContDoc = 0
    Exit Function
InsertaMovPendientesRendErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovPendientesRend Method")
End Function

Public Function CargaMovDocNoContable(ByVal pnMovNro As Long) As ADODB.Recordset
Dim psSql As String
   On Error GoTo CargaMovDocAsientoErr
   psSql = " Select cCtaContCod, cCtaContDesc, nMovOtroImporte, nMovItem, cMovOtroTexto " _
           & " from MovOtrosItem I " _
           & " inner join CtaCont C on I.cMovOtroVariable = C.cCtaContCod " _
           & "WHERE nMovNro = " & pnMovNro
           
   Set CargaMovDocNoContable = oConect.CargaRecordSet(psSql)
   Exit Function
CargaMovDocAsientoErr:
   Call RaiseError(MyUnhandledError, "DMov:CargaMovDocAsiento Method")
End Function
Public Function fbACGRegistrarGarantia(ByVal pnMoNro As Long, ByVal psCodPers As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, ByVal psCodPersPF As String, ByVal psIFTpoPF As String, ByVal psCtaIfCodPF As String, ByVal psMovNro As String) As Boolean
    Dim sSQL As String
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    
    On Error GoTo fbACGRegistrarGarantiaErr
    sSQL = " insert into cgCtaIFAdeudadosGarantia(nMovNro,cPersCod,cIFTpo,cCtaIFCod,cPersCodPF,cIFTpoPF,cCtaIFCodPF,cmovnroIni)"
    sSQL = sSQL & " values(" & pnMoNro & ",'" & psCodPers & "','" & psIFTpo & "','" & psCtaIFCod & "','" & psCodPersPF & "','" & psIFTpoPF & "','" & psCtaIfCodPF & "','" & psMovNro & "')"
    oConect.Ejecutar (sSQL)
    fbACGRegistrarGarantia = True
    Exit Function
fbACGRegistrarGarantiaErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Registrar Garantias ")
    fbACGRegistrarGarantia = False
End Function
Public Function fbACGRegistrarCtaIFFiltro(ByVal psCodPers As String, ByVal psIFTpo As String, ByVal psCtaIFCod As String, ByVal psCtaContCod As String, ByVal psSubCtaCont As String) As Boolean
    Dim sSQL As String
    Dim rs As New ADODB.Recordset

    On Error GoTo fbACGRegistrarGarantiaErr
    sSQL = " select * from CtaIFfiltro where cPersCod='" & psCodPers & "' and cIfTpo='" & psIFTpo & "' and cCtaIFCod='" & psCtaIFCod & "' and cctacontcod like '11" & Mid(psCtaIFCod, 3, 1) & "7%'"
    Set rs = oConect.CargaRecordSet(sSQL)
    If RSVacio(rs) Then
        sSQL = " Insert Into CtaIFfiltro(cPersCod,cIFTpo,cCtaIfCod,cCtaContCod,cCtaIFSubCta,bUsoAgencia)"
        sSQL = sSQL & " Values('" & psCodPers & "','" & psIFTpo & "','" & psCtaIFCod & "','" & psCtaContCod & "','" & psSubCtaCont & "',0)"
        oConect.Ejecutar sSQL
    End If
    fbACGRegistrarCtaIFFiltro = True
    Exit Function
fbACGRegistrarGarantiaErr:
    Call RaiseError(MyUnhandledError, "DACGAdeuGarantias:Registrar CtaIFFiltro ")
    'fbACGRegis
End Function
'##ModelId=3A84A3FE038A
Public Function ExtonrnaRendCH(ByVal psAgencia As String, ByVal psAreaCod As String, ByVal pnProcesoNro As Integer) As Integer
    On Error GoTo ExtonrnaRendCHErr
    Dim sql As String
    ExtonrnaRendCH = 1
    
    sql = " Exec spCCH_RendicionExtorno '" & psAgencia & "','" & psAreaCod & "'," & pnProcesoNro

    oConect.Ejecutar sql
    
    ExtonrnaRendCH = 0
    Exit Function
ExtonrnaRendCHErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCajaChica Method")
End Function

Public Function GetDevuelvePorcentajeDetraccion(ByVal psCtaContCod As String, ByVal nCodDoc As Integer, ByVal nMonto As Double) As Double
Dim sql As String
Dim cCtaDetrac As String
Dim nPorc As Double
Dim oCon As New DConecta
Dim rs As New ADODB.Recordset
Dim nBan As Integer

nBan = 0
nPorc = 0

sql = "Select CDD.cCtaContCod as cCtaDetrac, CC.cCtaContCod, CC.cCtaContDesc "
sql = sql & " From CtaCont CC "
sql = sql & " Inner Join CtaDetraccionesDet CDD On CC.cCtaContCod like CDD.cCtaContCod + '%' "
sql = sql & " Where substring(CC.cCtaContCod,1,2) + '" & Mid(psCtaContCod, 3, 1) & "' + substring(CC.cCtaContCod,4, Case When Len(CC.cCtaContCod)-2 < 0 Then 0 Else Len(CC.cCtaContCod)-2 End )='" & psCtaContCod & "'"

oCon.AbreConexion

Set rs = oCon.CargaRecordSet(sql)
If rs.BOF Then
    nBan = 1
Else
    nBan = 2
    cCtaDetrac = rs!cCtaDetrac
End If
rs.Close
If nBan = 2 Then
    sql = "select CD.nDetraPorc " 'D.nDocTpo, D.cDocDesc, CDD.cCtaContCod, CD.nDetraPorc, CD.nDetraPorc "
    sql = sql & " From CtaDetracciones CD "
    sql = sql & " Inner Join CtaDetraccionesDet CDD On CD.nDetraCod = CDD.nDetraCod "
    sql = sql & " Inner Join Documento D On CD.nDocDetra = D.nDocTpo "
    sql = sql & " Where D.nDocTpo=" & nCodDoc & " and CDD.cCtaContCod='" & cCtaDetrac & "'"
    sql = sql & " And Rango1<=" & nMonto & " and Rango2 >=" & nMonto
    Set rs = oCon.CargaRecordSet(sql)
    If rs.BOF Then
        nPorc = 0
    Else
        nPorc = rs!nDetraPorc
    End If
    rs.Close
End If
Set rs = Nothing
GetDevuelvePorcentajeDetraccion = nPorc

End Function

Public Function InsertaMovDetra(ByVal pnMovNro As Long, Optional ByVal pnEstado As Integer = 1, Optional gsCodUserDetra As String = "") As Integer
    On Error GoTo InsertaMovContErr
    Dim sql As String
    InsertaMovDetra = 1
    
    sql = "INSERT INTO MovDetra (nMovNro, nEstado, cUserDetra ) " _
         & "VALUES (" & pnMovNro & "," & pnEstado & ", '" & gsCodUserDetra & "')"
     
    oConect.Ejecutar sql
    InsertaMovDetra = 0
    Exit Function
InsertaMovContErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCont Method")
End Function


Public Function InsertaMovAjusteCV_Soles_Doc_MonedaExt(ByVal pnMovNro As Long, pnItem As Long, psCtaCVIng As String, psCtaCVEgr As String) As Integer
    On Error GoTo InsertaMovContErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    InsertaMovAjusteCV_Soles_Doc_MonedaExt = 1
    
    sql = "Select Sum(nMovImporte) Monto From MovCta Where nMovNro = " & pnMovNro & ""
    Set rs = oConect.CargaRecordSet(sql)
    
    If IsNull(rs!Monto) Then
        
    Else
        If rs!Monto < 0 Then
            InsertaMovCta pnMovNro, pnItem, psCtaCVEgr, rs!Monto * -1
        Else
            InsertaMovCta pnMovNro, pnItem, psCtaCVIng, rs!Monto * -1
        End If
    End If
    
    InsertaMovAjusteCV_Soles_Doc_MonedaExt = 0
    Exit Function
InsertaMovContErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCont Method")
End Function


Public Function ActualizaDatosFAdeudado(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal pscCtaIFCod As String, _
                                     ByVal bCuotaUltPago As Boolean, ByVal dCuotaUltPago As Date, _
                                     ByVal bcuotaultmodsaldos As Boolean, ByVal dcuotaultmodsaldos As Date, _
                                     ByVal bSaldoCap As Boolean, ByVal nSaldoCap As Double, _
                                     ByVal bSaldoCapLP As Boolean, ByVal nSaldoCapLP As Double, _
                                     Optional ByVal bCancelacion As Boolean = False, Optional dCancelacion As Date = "01/01/1900", Optional bExtornoCancelacion As Boolean = False, _
                                     Optional bPrePago As Boolean = False) As Integer
                                    
On Error GoTo InsertaCtaIFADeudadoErr
    Dim sql As String
    Dim sqltemp As String
    
    sql = ""
    sqltemp = ""
    
    sql = "UPDATE CtaIFAdeudados "
    sql = sql & " SET "
    
    If bCancelacion = True Or bPrePago = True Then
        sqltemp = " dCancelacion='" & Format(dCancelacion, "MM/dd/YYYY") & "' "
    End If
    If bExtornoCancelacion = True Then
        sqltemp = " dCancelacion=null "
    End If
    
    If bCuotaUltPago = True Then
        If Len(Trim(sqltemp)) > 0 Then
            sqltemp = sqltemp & ", dCuotaUltPago='" & Format(dCuotaUltPago, "MM/dd/YYYY") & "'"
        Else
            sqltemp = " dCuotaUltPago='" & Format(dCuotaUltPago, "MM/dd/YYYY") & "'"
        End If
    End If
    
    If bcuotaultmodsaldos = True Then
        If Len(Trim(sqltemp)) > 0 Then
            sqltemp = sqltemp & ", dcuotaultmodsaldos='" & Format(dcuotaultmodsaldos, "YYYY/MM/DD") & "'"
            'sqltemp = sqltemp & ", dcuotaultmodsaldos='" & Format(dcuotaultmodsaldos, "MM/dd/YYYY") & "'"
        Else
            sqltemp = " dcuotaultmodsaldos='" & Format(dcuotaultmodsaldos, "YYYY/MM/DD") & "'"
            'sqltemp = " dcuotaultmodsaldos='" & Format(dcuotaultmodsaldos, "MM/dd/YYYY") & "'"
        End If
    End If
    
    If bSaldoCap = True Then
        If Len(Trim(sqltemp)) > 0 Then
            sqltemp = sqltemp & ", nSaldoCap=" & Format(nSaldoCap, "0.00")
        Else
            sqltemp = " nSaldoCap=" & Format(nSaldoCap, "0.00")
        End If
    End If
    
    If bSaldoCapLP = True Then
        If Len(Trim(sqltemp)) > 0 Then
            sqltemp = sqltemp & ", nSaldoCapLP=" & Format(nSaldoCapLP, "0.00")
        Else
            sqltemp = " nSaldoCapLP=" & Format(nSaldoCapLP, "0.00")
        End If
    End If
    
    sql = sql & sqltemp & " Where cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
    
    oConect.Ejecutar sql
    
    Exit Function
InsertaCtaIFADeudadoErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaCtaIFADeudado Method")
End Function


Public Function ActualizaDatosCalendario(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal pscCtaIFCod As String, _
                                    ByVal bCuota2 As Boolean, ByVal cTpoCuota2 As String, ByVal nNroCuota2 As Long, _
                                    ByVal bCuota6 As Boolean, ByVal cTpoCuota6 As String, ByVal nNroCuota6 As Long, _
                                    ByVal cEstado As String, Optional ByVal bCancelacion As Boolean = False, Optional bExtornoCancelacion As Boolean = False, _
                                    Optional ByVal bPrePago As Boolean = False, Optional lstLista As Variant, _
                                    Optional cMovNroAnt As String = "", Optional pnSaldoMesExtorno As Currency = -1, Optional pnSaldoMesProvision As Currency = -1, Optional pnSaldoMesExtornoPago As Currency = -1, _
                                    Optional nSaldoMesConce As Currency = -1, Optional pnPagadoConcesion As Integer = 0) As Integer
                                    
On Error GoTo ActualizaDatosCalendarioErr
    Dim sql As String
    Dim sqltemp As String
    Dim rs As New ADODB.Recordset
    Dim nSaldoAnt As Double
    
    sql = ""
    sqltemp = ""
    
    If bCancelacion = True Then
        
        sql = "UPDATE CtaIFCalendario "
        sql = sql & " SET "
        sql = sql & " cEstado='" & cEstado & "'" 'Se quito la coma "," NAGL 20180612 Según INC1806120013
        
        'ALPA 20110816
        If pnSaldoMesProvision <> -1 Then
            sql = sql & " ,nSaldoMes=isnull(nSaldoMes,0)+" & pnSaldoMesProvision & ""
        End If
        If pnSaldoMesExtorno <> -1 Then
            sql = sql & " ,nSaldoMesExtor=" & pnSaldoMesExtorno & ""
        End If
        If pnSaldoMesExtornoPago <> -1 Then
            sql = sql & " ,nSaldoMesExtorPago=" & pnSaldoMesExtornoPago & ""
        End If
        '**********************
        'ALPA20130618*********************************
        If nSaldoMesConce <> -1 Then
            sql = sql & " ,nProvisionConce=" & nSaldoMesConce & ""
        End If
        '**********************
        'ALPA20130708*********************************
        sql = sql & " ,nPagadoConcesio =" & pnPagadoConcesion & ""
        '*********************************************
        sql = sql & " Where ("
        If bCuota2 = True Then
            sqltemp = " (nNroCuota>=" & nNroCuota2 & " And cTpoCuota='" & cTpoCuota2 & "') "
        End If
        If bCuota6 = True Then
            If Len(Trim(sqltemp)) > 0 Then
                sqltemp = sqltemp & " Or (nNroCuota>=" & nNroCuota6 & " And cTpoCuota='" & cTpoCuota6 & "')"
            Else
                sqltemp = " (nNroCuota>=" & nNroCuota6 & " And cTpoCuota='" & cTpoCuota6 & "') "
            End If
        End If
        sql = sql & sqltemp & ")"
        sql = sql & " And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
        sql = sql & " And bVigente=1 "
        oConect.Ejecutar sql
    
    
    ElseIf bExtornoCancelacion = True Then
        
        'La cuota que correspondia la volvemos como antes
        sql = "UPDATE CtaIFCalendario "
        sql = sql & " SET "
        sql = sql & " cEstado='" & cEstado & "'"
         'ALPA20130708*********************************
        sql = sql & " ,nPagadoConcesio =0"
        '*********************************************
        sql = sql & " Where ("
        If bCuota2 = True Then
            sqltemp = " (nNroCuota=" & nNroCuota2 & " And cTpoCuota='" & cTpoCuota2 & "') "
        End If
        If bCuota6 = True Then
            If Len(Trim(sqltemp)) > 0 Then
                sqltemp = sqltemp & " Or (nNroCuota=" & nNroCuota6 & " And cTpoCuota='" & cTpoCuota6 & "')"
            Else
                sqltemp = " (nNroCuota=" & nNroCuota6 & " And cTpoCuota='" & cTpoCuota6 & "') "
            End If
        End If
        sql = sql & sqltemp & ")"
        sql = sql & " And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
        sql = sql & " And bVigente=1 "
        oConect.Ejecutar sql
    
        ' Las demas cuotas las pasamos a cero
        sql = "UPDATE CtaIFCalendario "
        sql = sql & " SET "
        sql = sql & " cEstado='0'"
         'ALPA20130708*********************************
        If nSaldoMesConce <> -1 Then
            sql = sql & " ,nPagadoConcesio =0"
        End If
        '*********************************************
        sql = sql & " Where ("
        If bCuota2 = True Then
            sqltemp = " (nNroCuota>" & nNroCuota2 & " And cTpoCuota='" & cTpoCuota2 & "') "
        End If
        If bCuota6 = True Then
            If Len(Trim(sqltemp)) > 0 Then
                sqltemp = sqltemp & " Or (nNroCuota>" & nNroCuota6 & " And cTpoCuota='" & cTpoCuota6 & "')"
            Else
                sqltemp = " (nNroCuota>" & nNroCuota6 & " And cTpoCuota='" & cTpoCuota6 & "') "
            End If
        End If
        sql = sql & sqltemp & ")"
        sql = sql & " And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
        sql = sql & " And bVigente=1 "
        oConect.Ejecutar sql
    
    ElseIf bPrePago = True Then
    
        'Salvamos el antiguo Calendario
    
        sql = "Insert Into CtaIFCalendario (cPersCod, cIFTpo, cCtaIFCod, bVigente, cTpoCuota, dVencimiento, nDiasPago, nNroCuota, nCapital, nInteres, nInteresGracia, nComision, nTotalCuota, nSaldoCap, nOrigen, cEstado, cUltimaRegeneracion) "
        sql = sql & " Select cPersCod, cIFTpo, cCtaIFCod, bVigente, cTpoCuota, dVencimiento, nDiasPago, nNroCuota, nCapital, nInteres, nInteresGracia, nComision, nTotalCuota, nSaldoCap, nOrigen, cEstado, '" & cMovNroAnt & "' "
        sql = sql & " From CGCtaIFCalendario "
        sql = sql & " Where cPersCod='" & psPersCod & "' And cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
        sql = sql & " And bVigente=1 "
        
        oConect.Ejecutar sql
    
        sql = "Update CtaIFCalendario set bVigente=0"
        sql = sql & " Where cPersCod='" & psPersCod & "' And cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
        sql = sql & " And cUltimaRegeneracion<>'" & cMovNroAnt & "' "
    
        oConect.Ejecutar sql
    
        'Sacamos el monto pendiente de la cuota que no se paga en su totalidad
        
        sql = "Select nCapital "
        sql = sql & " From CtaIFCalendario "
        sql = sql & " Where "
        
        sql = sql & " (cTpoCuota='2' And nNroCuota=" & lstLista.ListItems(1).SubItems(6)
        sql = sql & " ) "
        
        sql = sql & " And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
        sql = sql & " And bVigente=1 "
        Set rs = oConect.CargaRecordSet(sql)
        If rs.BOF Then
            nSaldoAnt = 0
        Else
            nSaldoAnt = rs!nCapital
        End If
        rs.Close
        Set rs = Nothing
        
        'Las que estan en el rango las pasamos a pagadas
        sql = "Update CtaIFCalendario "
        sql = sql & " SET "
        sql = sql & " cEstado ='" & cEstado & "' "
        sql = sql & " Where ("
        
        sql = sql & " (cTpoCuota='2' And nNroCuota>=" & lstLista.ListItems(1).SubItems(3)
        sql = sql & "                And nNroCuota<=" & lstLista.ListItems(1).SubItems(4)
        sql = sql & " ) "
        sql = sql & " Or "
        sql = sql & " (cTpoCuota='6' And nNroCuota>=" & lstLista.ListItems(1).SubItems(9)
        sql = sql & "                And nNroCuota<=" & lstLista.ListItems(1).SubItems(10)
        sql = sql & " ) "
        
        sql = sql & ") And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
        sql = sql & " And bVigente=1 "
    
        oConect.Ejecutar sql
    
        If Val(lstLista.ListItems(1).SubItems(7)) > 0 Then
        
            'Las que se paga la mitad nomas (SOn las tipo 2 unicamente)
            'Actualizamos el Capital
        
            sql = "Update CtaIFCalendario "
            sql = sql & " SET "
            sql = sql & " nCapital=" & lstLista.ListItems(1).SubItems(7)
            sql = sql & " Where "
            
            sql = sql & " (cTpoCuota='2' And nNroCuota=" & lstLista.ListItems(1).SubItems(6)
            sql = sql & " ) "
            
            sql = sql & " And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
            sql = sql & " And bVigente=1 "
            
            oConect.Ejecutar sql
            
            'Actualizamos la cuota siguiente
            
            sql = "Update CtaIFCalendario "
            sql = sql & " SET "
            sql = sql & " nCapital= ncapital + " & nSaldoAnt - Val(lstLista.ListItems(1).SubItems(7))
            sql = sql & " Where "
            
            sql = sql & " (cTpoCuota='2' And nNroCuota=" & Val(lstLista.ListItems(1).SubItems(6)) + 1
            sql = sql & " ) "
            
            sql = sql & " And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
            sql = sql & " And bVigente=1 "
            
            oConect.Ejecutar sql
        End If
        
        'Cambiamos estados de calendario a 3
        sql = "Update CtaIFCalendario "
        sql = sql & " SET "
        sql = sql & " cEstado ='3' "
        sql = sql & " Where ("
        
        sql = sql & " (cTpoCuota='2' And nNroCuota>" & lstLista.ListItems(1).SubItems(4)
        sql = sql & " ) "
        sql = sql & " Or "
        sql = sql & " (cTpoCuota='6' And nNroCuota>" & lstLista.ListItems(1).SubItems(10)
        sql = sql & " ) "
        
        sql = sql & ") And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
        sql = sql & " And bVigente=1 "
    
        oConect.Ejecutar sql
    
        If Val(lstLista.ListItems(1).SubItems(7)) > 0 Then
        
            'Las que se paga la mitad nomas (SOn las tipo 2 unicamente)
            'Actualizamos el Estado Despues de haber hecho todas las actualizaciones
            
            sql = "Update CtaIFCalendario "
            sql = sql & " SET "
            sql = sql & " cEstado ='" & cEstado & "' "
            sql = sql & " Where "
            
            sql = sql & " (cTpoCuota='2' And nNroCuota=" & lstLista.ListItems(1).SubItems(6)
            sql = sql & " ) "
            
            sql = sql & " And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
            sql = sql & " And bVigente=1 "
            
            oConect.Ejecutar sql
            
        End If
    
    Else
    
        sql = "UPDATE CtaIFCalendario "
        sql = sql & " SET "
        sql = sql & " cEstado='" & cEstado & "'"
        sql = sql & " ,nPagadoConcesio =" & pnPagadoConcesion & ""
        If pnSaldoMesProvision <> -1 Then
            If pnSaldoMesExtorno = -1 Then
                sql = sql & " ,nSaldoMes=" & pnSaldoMesProvision & ""
            Else
                sql = sql & " ,nSaldoMes=isnull(nSaldoMes,0)+" & pnSaldoMesProvision & ""
            End If
        End If
        If pnSaldoMesExtorno <> -1 Then
            sql = sql & " ,nSaldoMesExtor=" & pnSaldoMesExtorno & ""
        End If
        If pnSaldoMesExtornoPago <> -1 Then
            sql = sql & " ,nSaldoMesExtorPago=" & pnSaldoMesExtornoPago & ""
        End If
        'ALPA20130618***************************
        If nSaldoMesConce <> -1 Then
            If pnSaldoMesExtorno = -1 Then
                sql = sql & " ,nProvisionConce=" & nSaldoMesConce & ""
            Else
                sql = sql & " ,nProvisionConce=isnull(nProvisionConce,0)+" & nSaldoMesConce & ""
            End If
        End If
        '***************************************
        sql = sql & " Where ("
        If bCuota2 = True Then
            sqltemp = " (nNroCuota=" & nNroCuota2 & " And cTpoCuota='" & cTpoCuota2 & "') "
        End If
        If bCuota6 = True Then
            If Len(Trim(sqltemp)) > 0 Then
                sqltemp = sqltemp & " Or (nNroCuota=" & nNroCuota6 & " And cTpoCuota='" & cTpoCuota6 & "')"
            Else
                sqltemp = " (nNroCuota=" & nNroCuota6 & " And cTpoCuota='" & cTpoCuota6 & "') "
            End If
        End If
        sql = sql & sqltemp & ")"
    
        sql = sql & " And cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
        sql = sql & " And bVigente=1 "
        
        oConect.Ejecutar sql
    
    End If
    
    Exit Function
ActualizaDatosCalendarioErr:
    Call RaiseError(MyUnhandledError, "DMov:CGActualizaDatosCalendario Method")
End Function

'desde aqui adeudados
Public Function GrabaCtaIFEstadistica(ByVal cPerscod As String, cIFTpo As String, cCtaIFCod As String, _
                                        ByVal cTpoCuota As String, ByVal nNroCuota As Long, ByVal nMovNro As Long, _
                                        ByVal nTipoEstadistica As Integer, _
                                        ByVal nCapitalMovido As Double, _
                                        ByVal nInteres1Movido As Double, ByVal nInteres2Movido As Double, _
                                        ByVal nComisionMovida As Double, _
                                        ByVal nTrasladoCPLP As Double, ByVal nTrasladoLPCP As Double, _
                                        ByVal nSaldoAnt As Double, ByVal nSaldoNew As Double, _
                                        ByVal nSaldoLPAnt As Double, ByVal nSaldoLPNew As Double, _
                                        ByVal Vac As Integer, _
                                        ByVal nCapitalMovidoReal As Double, _
                                        ByVal nInteres1MovidoReal As Double, ByVal nInteres2MovidoReal As Double, _
                                        ByVal nComisionMovidaReal As Double, _
                                        ByVal nTrasladoCPLPReal As Double, ByVal nTrasladoLPCPReal As Double, _
                                        ByVal nSaldoAntReal As Double, ByVal nSaldoNewReal As Double, _
                                        ByVal nSaldoLPAntReal As Double, ByVal nSaldoLPNewReal As Double, _
                                        ByVal dFecUltActAnt As Date, ByVal dFecUltActNew As Date, _
                                        ByVal dFecUltPagoAnt As Date, ByVal dFecUltPagoNew As Date, _
                                        ByVal cEstAnt As String) As Integer
                                        
' ntipoestadistica --1=Registro Pagare 2=Confirmacion 3=Pago 4=Provision
Dim sSQL As String
                                

On Error GoTo GrabaCtaIFEstadisticaErr

    If Vac = 0 Then
        If nSaldoAnt <> nSaldoAntReal Or nSaldoNew <> nSaldoNewReal Or nSaldoLPAnt <> nSaldoLPAntReal And nSaldoLPNew <> nSaldoLPNewReal Then
            MsgBox "Error en grabacion de saldos ... ", vbInformation, "Aviso"
        End If
    End If
    
    
    sSQL = "Insert Into CtaIFEstadistica "
    sSQL = sSQL & "        ( cPersCod, cIFTpo, cCtaIFCod, cTpoCuota, nNroCuota, nMovNro, "
    sSQL = sSQL & "          nTipoEstadistica, nCapitalMovido, nInteres1Movido, nInteres2Movido, "
    sSQL = sSQL & "          nComisionMovida, nTrasladoCPLP, nTrasladoLPCP, nSaldoAnt, nSaldoNew, "
    sSQL = sSQL & "          nSaldoLPAnt, nSaldoLPNew, bVac, dFecUltActAnt, dFecUltActNew, dFecUltPagoAnt, "
    sSQL = sSQL & "         dFecUltPagoNew, cEstAnt, nCapitalMovidoReal, nInteres1MovidoReal, "
    sSQL = sSQL & "          nInteres2MovidoReal, nComisionMovidaReal, nTrasladoCPLPReal, "
    sSQL = sSQL & "          nTrasladoLPCPReal, nSaldoAntReal, nSaldoNewReal, nSaldoLPAntReal, "
    sSQL = sSQL & "          nSaldoLPNewReal )"
    
    sSQL = sSQL & " Values ( '" & cPerscod & "', '" & cIFTpo & "', '" & cCtaIFCod & "', '" & cTpoCuota & "', " & nNroCuota & ", " & nMovNro & ", "
    sSQL = sSQL & nTipoEstadistica & ", " & nCapitalMovido & ", " & nInteres1Movido & ", " & nInteres2Movido & ", "
    sSQL = sSQL & nComisionMovida & ", " & nTrasladoCPLP & ", " & nTrasladoLPCP & ", " & nSaldoAnt & ", " & nSaldoNew & ", "
    sSQL = sSQL & nSaldoLPAnt & ", " & nSaldoLPNew & ", " & Vac & ", '" & Format(dFecUltActAnt, "MM/dd/YYYY") & "', '" & Format(dFecUltActNew, "MM/dd/YYYY") & "', '"
    sSQL = sSQL & Format(dFecUltPagoAnt, "MM/dd/YYYY") & "', '" & Format(dFecUltPagoNew, "MM/dd/YYYY") & "', '"
    sSQL = sSQL & cEstAnt & "', " & nCapitalMovidoReal & ", " & nInteres1MovidoReal & ", "
    sSQL = sSQL & nInteres2MovidoReal & ", " & nComisionMovidaReal & ", " & nTrasladoCPLPReal & ", "
    sSQL = sSQL & nTrasladoLPCPReal & ", " & nSaldoAntReal & ", " & nSaldoNewReal & ", " & nSaldoLPAntReal & ", "
    sSQL = sSQL & nSaldoLPNewReal & " )"
    oConect.Ejecutar sSQL
    
    '** Cambio 1

    sSQL = "Insert Into CtaIFEstadisticaDet (nMovNro, nCapitalMovidoReal, nInteres1MovidoReal, "
    sSQL = sSQL & "          nInteres2MovidoReal, nComisionMovidaReal, nTrasladoCPLPReal, "
    sSQL = sSQL & "          nTrasladoLPCPReal, nSaldoAntReal, nSaldoNewReal, nSaldoLPAntReal, "
    sSQL = sSQL & "          nSaldoLPNewReal )"

    sSQL = sSQL & " Values (" & nMovNro & ", " & nCapitalMovidoReal & ", " & nInteres1MovidoReal & ", "
    sSQL = sSQL & nInteres2MovidoReal & ", " & nComisionMovidaReal & ", " & nTrasladoCPLPReal & ", "
    sSQL = sSQL & nTrasladoLPCPReal & ", " & nSaldoAntReal & ", " & nSaldoNewReal & ", " & nSaldoLPAntReal & ", "
    sSQL = sSQL & nSaldoLPNewReal & " )"

    
     
    oConect.Ejecutar sSQL
                                        
Exit Function
GrabaCtaIFEstadisticaErr:
    Call RaiseError(MyUnhandledError, "DMov:GrabaCtaIFEstadistica Method")
                                        
End Function


Public Function CGActualizaDatosFAdeudado(ByVal psPersCod As String, ByVal psIFTpo As String, ByVal pscCtaIFCod As String, _
                                     ByVal bCuotaUltPago As Boolean, ByVal dCuotaUltPago As Date, _
                                     ByVal bcuotaultmodsaldos As Boolean, ByVal dcuotaultmodsaldos As Date, _
                                     ByVal bSaldoCap As Boolean, ByVal nSaldoCap As Double, _
                                     ByVal bSaldoCapLP As Boolean, ByVal nSaldoCapLP As Double, _
                                     Optional ByVal bCancelacion As Boolean = False, Optional dCancelacion As Date = "01/01/1900", Optional bExtornoCancelacion As Boolean = False, _
                                     Optional bPrePago As Boolean = False) As Integer
                                    
On Error GoTo CGInsertaCtaIFADeudadoErr
    Dim sql As String
    Dim sqltemp As String
    
    sql = ""
    sqltemp = ""
    
    sql = "UPDATE CtaIFAdeudados "
    sql = sql & " SET "
    
    If bCancelacion = True Or bPrePago = True Then
        sqltemp = " dCancelacion='" & Format(dCancelacion, "MM/dd/YYYY") & "' "
    End If
    If bExtornoCancelacion = True Then
        sqltemp = " dCancelacion=null "
    End If
    
    If bCuotaUltPago = True Then
        If Len(Trim(sqltemp)) > 0 Then
            sqltemp = sqltemp & ", dCuotaUltPago='" & Format(dCuotaUltPago, "YYYY/DD/MM") & "'"
            'sqltemp = sqltemp & ", dCuotaUltPago='" & Format(dCuotaUltPago, "MM/dd/YYYY") & "'"
        Else
            sqltemp = " dCuotaUltPago='" & Format(dCuotaUltPago, "YYYY/MM/DD") & "'"
            'sqltemp = " dCuotaUltPago='" & Format(dCuotaUltPago, "MM/dd/YYYY") & "'"
        End If
    End If
    
    If bcuotaultmodsaldos = True Then
        If Len(Trim(sqltemp)) > 0 Then
            sqltemp = sqltemp & ", dcuotaultmodsaldos='" & Format(dcuotaultmodsaldos, "YYYY/MM/DD") & "'"
            'sqltemp = sqltemp & ", dcuotaultmodsaldos='" & Format(dcuotaultmodsaldos, "MM/dd/YYYY") & "'"
        Else
            sqltemp = " dcuotaultmodsaldos='" & Format(dcuotaultmodsaldos, "YYYY/MM/DD") & "'"
            'sqltemp = " dcuotaultmodsaldos='" & Format(dcuotaultmodsaldos, "MM/dd/YYYY") & "'"
        End If
    End If
    
    If bSaldoCap = True Then
        If Len(Trim(sqltemp)) > 0 Then
            sqltemp = sqltemp & ", nSaldoCap=" & Format(nSaldoCap, "0.00")
        Else
            sqltemp = " nSaldoCap=" & Format(nSaldoCap, "0.00")
        End If
    End If
    
    If bSaldoCapLP = True Then
        If Len(Trim(sqltemp)) > 0 Then
            sqltemp = sqltemp & ", nSaldoCapLP=" & Format(nSaldoCapLP, "0.00")
        Else
            sqltemp = " nSaldoCapLP=" & Format(nSaldoCapLP, "0.00")
        End If
    End If
    
    sql = sql & sqltemp & " Where cPersCod='" & psPersCod & "' and cIFTpo='" & psIFTpo & "' and cCtaIFCod='" & pscCtaIFCod & "' "
    
    oConect.Ejecutar sql
    
    Exit Function
CGInsertaCtaIFADeudadoErr:
    Call RaiseError(MyUnhandledError, "DMov:CGInsertaCtaIFADeudado Method")
End Function



Public Function Ejecutar(pSql As String) As ADODB.Recordset
On Error GoTo ErrEjecutar
Set Ejecutar = oConect.Ejecutar(pSql)
Exit Function
ErrEjecutar:
   Call RaiseError(MyUnhandledError, "DConecta:Ejecutar Method")
End Function

Public Function EjecutarNew(pSql As String)
On Error GoTo ErrEjecutar
     Call oConect.ConexionActiva.Execute(pSql)
Exit Function
ErrEjecutar:
   Call RaiseError(MyUnhandledError, "DConecta:Ejecutar Method")
End Function



Public Function CGActualizaCuentaIFInteres(ByVal psPersCod As String, ByVal psIFTpo As CGTipoIF, _
                                ByVal psCtaIFCod As String, ByVal pdFechaInt As Date, ByVal pnPeriodo As Integer, _
                                ByVal pnInteres As Currency, ByVal psMovNro As String) As Integer
    On Error GoTo CGActualizaCuentaIFInteresErr
    Dim sql As String
    CGActualizaCuentaIFInteres = 1

    sql = "Update CtaifInteres set nVigente=0 "
    sql = sql & " Where cPersCod='" & psPersCod & "' and cIFTpo='" & Format(psIFTpo, "00") & "' and cCtaIFCod='" & psCtaIFCod & "' "
    oConect.Ejecutar sql

    sql = "Delete CtaifInteres "
    sql = sql & " Where cPersCod='" & psPersCod & "' and cIFTpo='" & Format(psIFTpo, "00") & "' and cCtaIFCod='" & psCtaIFCod & "' "
    sql = sql & " And Convert(varchar(8), dCtaIFIntRegistro, 112)='" & Format(pdFechaInt, "YYYYMMdd") & "' "
    oConect.Ejecutar sql


    sql = "INSERT INTO CtaIFInteres( cPersCod, cIFTpo, cCtaIFCod, dCtaIFIntRegistro , nCtaIFIntPeriodo, nCtaIFIntValor, cUltimaActualizacion, nVigente) "
    sql = sql & " VALUES('" & psPersCod & "','" & Format(psIFTpo, "00") & "','" & psCtaIFCod & "','"
    sql = sql & Format(pdFechaInt, gsFormatoFecha) & "'," & pnPeriodo & "," & pnInteres & ",'" & psMovNro & "', 1)"

    oConect.Ejecutar sql
    CGActualizaCuentaIFInteres = 0
    Exit Function
CGActualizaCuentaIFInteresErr:
    Call RaiseError(MyUnhandledError, "DMov:CGActualizaCuentaIFInteres Method")
End Function

'--------  JEOM ------------
Public Function InsertaMovCtaNew(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency, ByRef psCadena As String) As Integer
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    InsertaMovCtaNew = 1

    If lbBitControlCuentasUltNivel Then
        sql = "Select cCtaContCod From CtaCont Nolock where cCtaContCod Like '" & psCtaContCod & "%' And cCtaContCod <> '" & psCtaContCod & "' And nCtaEstado = 1"
        Set rsCta = oConect.CargaRecordSet(sql)

        If rsCta.EOF And rsCta.BOF Then
            rsCta.Close

            sql = "Select cCtaContCod from CtaCont Nolock Where cCtaContCod = '" & psCtaContCod & "' And nCtaEstado = 1"
            Set rsCta = oConect.CargaRecordSet(sql)

            If rsCta.EOF And rsCta.BOF Then
                rsCta.Close
                Err.Number = 9999
                Err.Description = "Cuenta Contable no existe o no es valida " & psCtaContCod & "."
            Else
                sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte)" _
                    & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
                'oConect.Ejecutar sql
                psCadena = sql
                InsertaMovCtaNew = 0
                rsCta.Close
                Exit Function
            End If
        Else
            rsCta.Close
            Err.Number = 9999
            Err.Description = "Se desea grabar una Cta Cont que no esta a Ult Nivel " & psCtaContCod & " ."
        End If
    Else
        sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte )" _
            & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
        'oConect.Ejecutar sql
        psCadena = sql
        InsertaMovCtaNew = 0
        Exit Function
    End If
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function


Public Function InsertaMovObjNew(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psObjetoCod As String, ByRef psMovObj As String) As Integer
    On Error GoTo InsertaMovObjErr
    Dim sql As String
    InsertaMovObjNew = 1
     
    sql = "INSERT INTO MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & Format(psObjetoCod, "00") & "')"

    'oConect.Ejecutar sql
    psMovObj = sql
    InsertaMovObjNew = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjNew Method")
End Function


Public Function GeneraMovMENew(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional ByRef psCadena As String) As Boolean
    Dim nTCFijo As Currency
    Dim nTCVent As Currency
    Dim nTCComp As Currency
    Dim sSQL As String
    Dim rs   As New ADODB.Recordset
    Dim dFecha As Date
    Dim nImporteD As Currency
    Dim nImporteS As Currency
    On Error GoTo GeneraMovMEErr
    GeneraMovMENew = False
    dFecha = CDate(GetFechaMov(psMovNro, True))
    
    Dim oTC As New nTipoCambio
    nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
    If nTCMerc = 0 Then
        sSQL = "Select cOpeCod From Mov Where nMovNro = " & pnMovNro        'X
        Set rs = oConect.CargaRecordSet(sSQL)                               'X
        If rs.EOF And rs.BOF Then                                           'X
            nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
            nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)
        Else                                                                'X
            If Left(rs!cOpeCod, 1) = "4" Or Left(rs!cOpeCod, 1) = "7" Then  'X
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCPondVenta)          'X
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCPonderado)          'X
            Else                                                            'X
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)              'X
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)             'X
            End If                                                          'X
        End If                                                              'X
    Else
        InsertaMovTpoCambio pnMovNro, nTCMerc
        nTCVent = nTCMerc
        nTCComp = nTCMerc
    End If
    Set oTC = Nothing
    'oConect.Ejecutar " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent
    psCadena = " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent
   
    GeneraMovMENew = True

Exit Function
GeneraMovMEErr:
    Call RaiseError(MyUnhandledError, "DMov:GeneraMovME Method")
End Function

Public Sub ActualizaSaldoMovimientoNew(ByVal sMovNro As String, ByVal sSimbolo As String, ByRef psCadena1 As String, ByRef psCadena2 As String, ByRef psCadena3 As String)
 Dim sql As String
 On Error GoTo ErrActualizaMovSaldo
' Dim prs As adodb.Recordset
 Dim ldFecha As Date
' Dim ldFecSis As Date
 
 'psCadena = " "
 
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
         psCadena1 = "INSERT tmpMovActualiza (cMovNro) VALUES (" & lsMov & ")"
        'oConect.Ejecutar sql
    Loop
    ldFecha = GetFechaMov(Mid(lsMov, 2, 8), True)
     psCadena2 = "Exec spActualizaMovSaldo " & lsMov & ", '" & Format(ldFecha, gsFormatoFecha) & "', '" & sSimbolo & "' "
    'oConect.Ejecutar sql
    
     psCadena3 = "DELETE tmpMovActualiza"
    'oConect.Ejecutar sql

 Else
    psCadena1 = "INSERT tmpMovActualiza (cMovNro) VALUES ('" & sMovNro & "')"
    'oConect.Ejecutar sql
    
    ldFecha = GetFechaMov(Left(sMovNro, 8), True)
    
    psCadena2 = "spActualizaMovSaldo '" & sMovNro & "', '" & Format(ldFecha, gsFormatoFecha) & "', '" & sSimbolo & "' "
    'oConect.Ejecutar sql
    
    psCadena3 = "DELETE tmpMovActualiza"
    'oConect.Ejecutar sql
 End If

Exit Sub
ErrActualizaMovSaldo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function GetTodoMov(ByVal psMovNro As String) As ADODB.Recordset
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta

sql = "Select * From Mov where cMovNro ='" & psMovNro & "'"
oCon.AbreConexion
Set GetTodoMov = oConect.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function

Public Sub ActualizaExtMovSUNAT(ByVal pnMovNro As Long)
    On Error GoTo ActualizaMovSUNAT
    Dim sql As String
         
    sSQL = "Update MovControlPagoSunat set nMovPagadoSunat= NULL Where nMovPagadoSunat = " & pnMovNro & " and bvigente=1 and bvalido=1"
    oConect.Ejecutar sSQL
Exit Sub
ActualizaMovSUNAT:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovSUNAT Method")
End Sub
 
Public Sub ActualizaMovSUNAT(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long)
    On Error GoTo ActualizaMovSUNAT
    Dim sql As String
 
    sSQL = "Update MovControlPagoSunat set nMovPagadoSunat=" & pnMovNroRef & " Where nMovNro=" & pnMovNro & " and bvigente=1 and bvalido=1"
    oConect.Ejecutar sSQL
     
    Exit Sub
ActualizaMovSUNAT:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovSUNAT Method")
End Sub

'*** PEAC 20101011
Public Sub ActualizaMovMe(ByVal pnMovNro As Long, ByVal pnMovItem As Long, Optional pnMovImporte As Currency = 0)
    On Error GoTo ActualizaMovMeErr
    Dim sql As String
    Dim lsActualiza As String
    lsActualiza = ""
    
    If pnMovImporte <> 0 Then
       lsActualiza = lsActualiza & " nMovMEImporte = " & pnMovImporte & ","
    End If
    If lsActualiza <> "" Then
      lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    End If
    If lsActualiza = "" Then Exit Sub
    
    sql = " UPDATE MovMe SET " & lsActualiza & " WHERE nMovNro =" & pnMovNro & " and nMovItem = " & pnMovItem & ""
    oConect.Ejecutar sql
    Exit Sub
ActualizaMovMeErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovMe Method")
End Sub

'*** PEAC 20110314
Public Function GetCuentaOrdenBienNoDepre(ByVal psMoneda As String, ByVal psAgencia As String) As ADODB.Recordset
Dim sql As String
Dim oCon As DConecta
Set oCon = New DConecta

sql = "exec stp_sel_CuentaOrdenBienNoDepre '" & psMoneda & "','" & psAgencia & "'"

oCon.AbreConexion
Set GetCuentaOrdenBienNoDepre = oConect.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function

'JACA 20110912*******************************************************************
Public Sub InsertaInversion(pnMovNro As Long, pnOpeCod As String, psPersCodAdm As String, psIFTipoAdm As String, psCtaIfCodAdm As String, psCtaIFDescAdm As String, _
                                     pnTpoInversion As Integer, pnMoneda As Integer, pnCondicion As Integer, pnImporte As Currency, pnPlazo As Integer, _
                                     pdFecApe As Date, pdFecVenc As Date, pdFecRenov As Date, pnTasa As Double, pnIntPactado As Currency, _
                                     psPersCodTran As String, psIFTipoTran As String, psCtaIfCodTran As String, psGlosa As String, psUltimaActualizacion, Optional pnIntereses As Currency = 0#)
    On Error GoTo MsgError
    Dim sql As String
       
       
    sql = " exec stp_ins_Inversion " & pnMovNro & ",'" & pnOpeCod & "','" & psPersCodAdm & "','" & psIFTipoAdm & "','" & psCtaIfCodAdm & "','" & psCtaIFDescAdm & "',"
    sql = sql + " " & pnTpoInversion & "," & pnMoneda & ",'" & pnCondicion & "'," & pnImporte & "," & pnPlazo & ","
    sql = sql + " '" & Format(pdFecApe, "yyyymmdd") & "','" & Format(pdFecVenc, "yyyymmdd") & "','" & Format(pdFecRenov, "yyyymmdd") & "'," & pnTasa & "," & pnIntPactado & ","
    sql = sql + " '" & psPersCodTran & "','" & psIFTipoTran & "','" & psCtaIfCodTran & "','" & psGlosa & "','" & psUltimaActualizacion & "'," & pnIntereses & ""
    oConect.Ejecutar sql
    Exit Sub
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:Guardar Inversion")
End Sub
Public Sub actualizaInteresInversion(pnMovNro As Long, pnOpeCod As String, pnIntereses As Currency)
    On Error GoTo MsgError
    Dim sql As String
       
       
    sql = " exec stp_upd_InversionInteres " & pnMovNro & ",'" & pnOpeCod & "'," & pnIntereses & ""
    
    oConect.Ejecutar sql
    Exit Sub
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:Actualiza Intereses Inversion")
End Sub
Public Sub actualizaEstadoInversion(pnMovNro As Long, pnOpeCod As String, pnEstado As Integer)
    On Error GoTo MsgError
    Dim sql As String
       
    sql = " exec stp_upd_InversionEstado " & pnMovNro & ",'" & pnOpeCod & "'," & pnEstado & ""
    
    oConect.Ejecutar sql
    Exit Sub
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:Actualiza Intereses Inversion")
End Sub

Public Function getUltFecProvInversion(pnMovNro As Long, pnOpeCod As String, psIFTipo As String, psPersCod As String, psCtaIFCod As String) As Recordset
    On Error GoTo MsgError
    Dim sql As String
       
    sql = " exec stp_sel_UltFecProvInversion " & pnMovNro & ",'" & pnOpeCod & "','" & psIFTipo & "','" & psPersCod & "','" & psCtaIFCod & "'"
    
    Set getUltFecProvInversion = oConect.Ejecutar(sql)
    Exit Function
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:Obtener Ultima Fecha de Provision Inversion")
End Function
Public Sub actualizaInversionDatos(pnMovNro As Long, pnOpeCod As String, psCtaIFDescAdm As String, pnTpoInversion As Integer, _
                                    pnCondicion As Integer, pnImporte As Currency, pnPlazo As Integer, pdFecApe As Date, _
                                    pdFecVenc As Date, pnTasa As Double, pnIntPactado As Currency, psGlosa As String)
    On Error GoTo MsgError
    Dim sql As String
       
    sql = " exec stp_upd_InversionDatos " & pnMovNro & ",'" & pnOpeCod & "','" & psCtaIFDescAdm & "'," & pnTpoInversion & ","
    sql = sql & pnCondicion & "," & pnImporte & "," & pnPlazo & ",'" & Format(pdFecApe, "yyyymmdd") & "','" & Format(pdFecVenc, "yyyymmdd") & "',"
    sql = sql & pnTasa & "," & pnIntPactado & ",'" & psGlosa & "'"
    
    oConect.Ejecutar sql
    Exit Sub
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:Actualiza datos Inversion")
End Sub

'JACA END************************************************************************

'JACA 20111201*******************************************************************
Public Sub InsertaInversionFondosMutuos(nMovNro As Long, dFecha As Date, psPersCod As String, psIFTipo As String, psCtaIFCod As String, _
                                        pnImporte As Currency, pnValorCuota As Double, pnNroCuotas As Currency, _
                                        pnImpRenta As Currency)
    On Error GoTo MsgError
    Dim sql As String
       
    sql = " exec stp_ins_InversionFondosMutuos " & nMovNro & ",'" & Format(dFecha & " " & oConect.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & psPersCod & "','" & psIFTipo & "','" & psCtaIFCod & "'," & pnImporte & ","
    sql = sql + " " & pnValorCuota & "," & pnNroCuotas & "," & pnImpRenta & ""
    
    oConect.Ejecutar sql
    Exit Sub
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:Guardar Inversion")
End Sub
Public Sub actualizaEstadoFondoMutuo(pnMovNro As Long, psIFTipo As String, psPersCod As String, psCtaIFCod As String)
    On Error GoTo MsgError
    Dim sql As String
       
    sql = " exec stp_upd_InversionEstadoFondoMutuo " & pnMovNro & ",'" & psIFTipo & "','" & psPersCod & "','" & psCtaIFCod & "'"
    
    oConect.Ejecutar sql
    Exit Sub
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:Actualiza Intereses Inversion")
End Sub
Public Sub actualizaFondoMutuoMov(pnMovNro As Long, pnMovNroAnt As Long, psIFTipo As String, psPersCod As String, psCtaIFCod As String)
    On Error GoTo MsgError
    Dim sql As String
       
    sql = " exec stp_upd_InversionFondoMutuoMov " & pnMovNro & "," & pnMovNroAnt & ",'" & psIFTipo & "','" & psPersCod & "','" & psCtaIFCod & "'"
    
    oConect.Ejecutar sql
    Exit Sub
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:Actualiza Intereses Inversion")
End Sub
Public Sub actualizaInversionFondosMutuos(nMovNro As Long, dFecha As Date, psPersCod As String, psIFTipo As String, psCtaIFCod As String, _
                                        pnImporte As Currency, pnValorCuota As Double, pnNroCuotas As Currency, _
                                        pnImpRenta As Currency)
    On Error GoTo MsgError
    Dim sql As String
       
    sql = " exec stp_upd_InversionFondosMutuos " & nMovNro & ",'" & Format(dFecha & " " & oConect.GetHoraServer, "yyyymmdd hh:mm:ss") & "','" & psPersCod & "','" & psIFTipo & "','" & psCtaIFCod & "'," & pnImporte & ","
    sql = sql + " " & pnValorCuota & "," & pnNroCuotas & "," & pnImpRenta & ""
    
    oConect.Ejecutar sql
    Exit Sub
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:Guardar Inversion")
End Sub
'JACA END**********************************************************************
'***Agregado por ELRO el 201203, según OYP-RFC005-2012 y OYP-RFC016-2012
Public Sub actualizarConceptoSolicitudARendirCta(ByVal pnMovNro As Long, ByVal psMovDesc As String)
Dim sql As String

On Error GoTo ErractualizarConceptoSolicitudARendirCta
    
    sql = " exec stp_upd_ActualizarConceptoSolicitudARendirCta " & pnMovNro & ", '" & psMovDesc & "'"
    oConect.Ejecutar sql
    Exit Sub

ErractualizarConceptoSolicitudARendirCta:

    Err.Raise Err.Number, "DMov:actualizarConceptoSolicitudARendirCta", Err.Description
End Sub

Public Sub actualizarMovCont(ByVal pnMovNro As Long, ByVal nMovMonto As Currency)
Dim sql As String

On Error GoTo MsgError

    sql = "stp_upd_ActualizarMovCont " & pnMovNro & ", " & nMovMonto & ""
    oConect.Ejecutar sql
    Exit Sub
    
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:actualizarMovCont")
End Sub

Public Sub actualizarMovDoc(ByVal pnMovNro As Long, ByVal pcDocNro As String)
Dim sql As String

On Error GoTo MsgError

    sql = "exec stp_upd_ActualizaMovDoc " & pnMovNro & ", '" & pcDocNro & "'"
    oConect.Ejecutar sql
    Exit Sub

MsgError:
    Call RaiseError(MyUnhandledError, "DMov:actualizarMovDoc")
End Sub

Public Sub actualizarMovARendir(ByVal pnMovNro As Long, ByVal pcCategCod As String)
Dim sql As String

On Error GoTo MsgError

    sql = "exec stp_upd_ActualizaMovARendir " & pnMovNro & ", '" & pcCategCod & "'"
    oConect.Ejecutar sql
    Exit Sub

MsgError:
    Call RaiseError(MyUnhandledError, "DMov:actualizarcCategCod")
End Sub

Public Sub actualizarMovObjViaticos(ByVal pnMovNro As Long, ByVal pcObjetoCod As String, ByVal pnMovImporte As Currency)
Dim sql As String

On Error GoTo MsgError

    sql = "exec stp_upd_ActulizaMovObjViaticos " & pnMovNro & ", '" & pcObjetoCod & "'," & pnMovImporte & ""
    oConect.Ejecutar sql
    Exit Sub
    
MsgError:
    Call RaiseError(MyUnhandledError, "DMov:actualizarcCategCod")
End Sub
Public Sub actualizarMovViaticos(ByVal pnMovNro As Long, ByVal psDestinoCod As String, ByVal psTranspIda As String, ByVal psTranspVuelta As String, ByVal psPartida As String, ByVal pnMovViaticosDias As Currency, ByVal psDestinoDesc As String)
Dim sql As String

On Error GoTo MsgError

    sql = "exec stp_upd_ActualizaMovViaticos " & pnMovNro & ", '" & psDestinoCod & "', '" & psTranspIda & "', '" & psTranspVuelta & "', '" & psPartida & "', " & pnMovViaticosDias & ", '" & psDestinoDesc & "'"
    oConect.Ejecutar sql
    Exit Sub

MsgError:
    Call RaiseError(MyUnhandledError, "DMov:actualizarMovViaticos")
End Sub

Public Sub actualizarGlosaViaticos(ByVal pnMovNro As Long, ByVal psMovDesc As String)
Dim sql As String

On Error GoTo MsgError

sql = "exec stp_upd_ActualizaGlosaViaticos " & pnMovNro & ", '" & psMovDesc & "'"
oConect.Ejecutar sql
Exit Sub

MsgError:
    Call RaiseError(MyUnhandledError, "DMov:actualizarGlosaViaticos")
End Sub

'***Fin Agregado por ELRO*********************************
'***ALPA 20120418*****************************************
Public Function GeneraMovMECierreAnual(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional nTCComp1 As Currency, Optional ncOpeCod As String = "", Optional nTipoPago As Integer = 0) As Boolean
    Dim nTCFijo As Currency
    Dim nTCVent As Currency
    Dim nTCComp As Currency
    Dim sSQL As String
    Dim rs   As New ADODB.Recordset
    Dim dFecha As Date
    Dim nImporteD As Currency
    Dim nImporteS As Currency
    On Error GoTo GeneraMovMECierreAnualErr
    GeneraMovMECierreAnual = False
    dFecha = CDate(GetFechaMov(psMovNro, True))
    
    Dim oTC As New nTipoCambio
    nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
    
    If nTCMerc = 0 Then
        sSQL = "Select cOpeCod From Mov  Where nMovNro = " & pnMovNro        'X
        Set rs = oConect.CargaRecordSet(sSQL)                               'X
        If rs.EOF And rs.BOF Then                                           'X
            nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
            nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)
        Else                                                                'X
            If Left(rs!cOpeCod, 1) = "4" Or Left(rs!cOpeCod, 1) = "7" Then  'X
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCPondVenta)          'X
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCPonderado)          'X
            Else                                                            'X
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)              'X
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)             'X
            End If                                                          'X
        End If                                                              'X
    Else
        InsertaMovTpoCambio pnMovNro, nTCMerc
        
        If ncOpeCod = "702403" Then
            nTCVent = nTCMerc
            nTCComp = nTCComp1
            
        Else
            nTCVent = nTCMerc
            nTCComp = nTCMerc
        End If
    End If
    Set oTC = Nothing
    'ALPA 20090430 ***********************************
       'oConect.Ejecutar " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent
        oConect.Ejecutar " sp_GeneraMovMECierreAnual " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent & "," & nTipoPago
    '**************************************************
    
   
    GeneraMovMECierreAnual = True

Exit Function
GeneraMovMECierreAnualErr:
    Call RaiseError(MyUnhandledError, "DMov:GeneraMovMECierreAnual Method")
End Function
'*********************************************************

'*** PEAC 20120503
Public Sub ModificaActivosFijos(ByVal psMovNro As String, ByVal psOpeCod As String)
Dim sql As String
   On Error GoTo ModificaActivosFijosErr

    sql = "exec stp_upd_ModificaActivosFijos '" & psMovNro & "','" & psOpeCod & "' "
    oConect.Ejecutar sql

Exit Sub

ModificaActivosFijosErr:
    Call RaiseError(MyUnhandledError, "DMov:ModificaActivosFijos Method")
End Sub
'***Agregado por ELRO el 20120531, según OYP-RFC047-2012
Public Sub actualizarCajaChicaSinAprobar(ByVal psPersCod As String, _
                                         ByVal pnMontoAsig As Currency, _
                                         ByVal pnTopeEgresos As Currency, _
                                         ByVal psAreaCod As String, _
                                         ByVal psAgeCod As String, _
                                         ByVal pnProcNro As Integer)
Dim sql As String

On Error GoTo MsgError

sql = "exec stp_upd_ActualizarCajaChicaSinAprobar '" & psPersCod & "', " & pnMontoAsig & ", " & pnTopeEgresos & ", '" & psAreaCod & "', '" & psAgeCod & "', " & pnProcNro & ""
oConect.Ejecutar sql
Exit Sub

MsgError:
    Call RaiseError(MyUnhandledError, "DMov:actualizarGlosaViaticos")
End Sub
Public Sub registrarEntregaEfectivoCH(ByVal pnMovNro As Long, _
                                      ByVal psPersCod As String, _
                                      ByVal psAreCod As String, _
                                      ByVal psAgeCod As String, _
                                      ByVal pnProcNro As Integer, _
                                      ByVal pnImporte As Currency)
Dim sql As String

On Error GoTo MsgError

sql = "exec stp_ins_RegistrarCajaChicaEntregaEfectivo " & pnMovNro & ",'" & psAreCod & "', '" & psAgeCod & "', " & pnProcNro & ",'" & psPersCod & "'," & pnImporte & ""
oConect.Ejecutar sql
Exit Sub

MsgError:
    Call RaiseError(MyUnhandledError, "DMov:actualizarGlosaViaticos")
End Sub

Public Function obtenerMovRef(ByVal pnMovNro As Long) As Long
Dim sql As String
Dim rsMovNro As New ADODB.Recordset

On Error GoTo MsgError

sql = "exec stp_sel_ObtenerMovRef " & pnMovNro & ""
Set rsMovNro = oConect.Ejecutar(sql)

If Not rsMovNro.BOF And Not rsMovNro.EOF Then
    obtenerMovRef = rsMovNro!nMovNroRef
End If
Exit Function

MsgError:
    Call RaiseError(MyUnhandledError, "DMov:obtenerMovRef")
End Function
'***Fin Agregado por ELRO*******************************
'***Agregado por ELRO el 20120901, según OYP-RFC104-2012
Public Function devolverTotalEgresos_2(ByVal psAreaCod As String, _
                                       ByVal psAgeCod As String, _
                                       ByVal pnProcNro As Integer) As Currency
Dim sql As String
Dim rsTE As ADODB.Recordset
Set rsTE = New Recordset

sql = "exec stp_sel_DevolverTotalEgresos '" & psAreaCod & "', '" & psAgeCod & "', " & pnProcNro & ""
Set rsTE = oConect.Ejecutar(sql)

If Not rsTE.BOF And Not rsTE.EOF Then
    devolverTotalEgresos_2 = rsTE!nTEEE
Else
    devolverTotalEgresos_2 = 0
End If
Set rsTE = Nothing
End Function
'***Fin Agregado por ELRO el 20120901*******************
'***Agregado por ELRO el 20120926, según TIC1209270003
Public Function DevolverCorrelativoMovDoc(ByVal pnDocTpo As Long) As Long
Dim sql As String
Dim rsMD As ADODB.Recordset
Set rsMD = New Recordset

sql = "exec stp_sel_DevolverCorrelativoMovDoc " & pnDocTpo & ""
Set rsMD = oConect.Ejecutar(sql)

If Not rsMD.BOF And Not rsMD.EOF Then
    DevolverCorrelativoMovDoc = rsMD!cDocCorrela + 1
Else
    DevolverCorrelativoMovDoc = 1
End If
Set rsMD = Nothing
End Function
'***Fin Agregado por ELRO el 20120926*****************
'***Agregado por ELRO el 20121001, según OYP-RFC106-2012
Public Function devolverMovimientoTransferencia(ByVal pnMovNro As Long) As ADODB.Recordset
Dim sql As String

sql = "exec stp_sel_DevolverMovimientoTransferencia " & pnMovNro & ""
Set devolverMovimientoTransferencia = oConect.Ejecutar(sql)

End Function

Public Function DevolverPlantillaNegocioTransferencia(ByVal psOpeCod As String, _
                                                      ByVal pnConcepto As Integer, _
                                                      ByVal pnPersoneria As Integer) As ADODB.Recordset
Dim sql As String

sql = "exec stp_sel_devolverPlantillaNegocioTransferencia '" & psOpeCod & "', " & pnConcepto & ", " & pnPersoneria & " "
Set DevolverPlantillaNegocioTransferencia = oConect.Ejecutar(sql)

End Function

Public Function devolverVariableInstitucionFinanciera(ByVal psCodCta As String, ByVal pnTipoIF As CGTipoIF) As String
    Dim sSQL As String
    Dim RegCta As ADODB.Recordset

    sSQL = "exec stp_sel_VarInstitucionFinanciera '" & psCodCta & "','" & Format$(pnTipoIF, "00") & "'"

    Set RegCta = New ADODB.Recordset

    Set RegCta = oConect.Ejecutar(sSQL)

    If (RegCta.BOF Or RegCta.EOF) Then
        devolverVariableInstitucionFinanciera = ""
    Else
        devolverVariableInstitucionFinanciera = Trim(RegCta!Campo)
    End If

    RegCta.Close
    Set RegCta = Nothing
End Function
'***Fin Agregado por ELRO el 20121001*******************
'***Agregado por ELRO el 20120928, según INC1209250017 y INC1209260002
Public Function verificarAprobacionARendirCuenta(ByVal pnMovNroRef As Long) As Integer
    Dim sSQL As String
    Dim rsAprobacion As ADODB.Recordset
    Set rsAprobacion = New ADODB.Recordset

    sSQL = "exec stp_sel_VerificarAprobacionARendirCuenta " & pnMovNroRef & ""
    Set rsAprobacion = oConect.Ejecutar(sSQL)

    If Not (rsAprobacion.BOF Or rsAprobacion.EOF) Then
        verificarAprobacionARendirCuenta = rsAprobacion!nAprobado
    Else
        verificarAprobacionARendirCuenta = 0
    End If

    rsAprobacion.Close
    Set rsAprobacion = Nothing
End Function
Public Function verificarAprobacionARendirViatico(ByVal pnMovNroRef As Long) As Integer
    Dim sSQL As String
    Dim rsAprobacion As ADODB.Recordset
    Set rsAprobacion = New ADODB.Recordset

    sSQL = "exec stp_sel_VerificarAprobacionARendirViaticos " & pnMovNroRef & ""
    Set rsAprobacion = oConect.Ejecutar(sSQL)

    If Not (rsAprobacion.BOF Or rsAprobacion.EOF) Then
        verificarAprobacionARendirViatico = rsAprobacion!nAprobado
    Else
        verificarAprobacionARendirViatico = 0
    End If

    rsAprobacion.Close
    Set rsAprobacion = Nothing
End Function
'***Fin Agregado por ELRO el 20121001*********************************
'EJVG20131116 ***
Public Sub ActualizaComprobanteLog(ByVal pnMovNro As Long, ByVal pnDocTpo As Integer, ByVal psDocNro As String, ByVal pdDocFecha As Date)
    Dim lsSql As String
    lsSql = "Exec stp_upd_ERS0622013_Comprobante " & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & Format(pdDocFecha, "yyyymmdd") & "'"
    oConect.Ejecutar lsSql
End Sub
Public Sub RevierteAsientoActaConformidad(ByVal pnMovNro As Long, ByVal pnMovNroComprobante As Long)
    Dim lsSql As String
    lsSql = "Exec stp_ins_ERS0622013_RevierteAsientoActaConformidad " & pnMovNro & "," & pnMovNroComprobante
    oConect.Ejecutar lsSql
End Sub
Public Sub InsertaVenta(ByVal psOpeTpo As String, ByVal pnDocTpo As Long, ByVal psDocNro As String, ByVal pdDocFecha As Date, ByVal psPersCod As String, ByVal psCtaCod As String, ByVal psDescrip As String, ByVal pnVVenta As Currency, ByVal pnIGV As Currency, ByVal pnPVentas As Currency, ByVal psDocNroRefe As String, ByVal pdDocFecRef As Date)
    Dim lsSql As String
    lsSql = "INSERT RegVenta (cOpeTpo,nDocTpo,cDocNro, dDocFecha, cPersCod, cCtaCod, cDescrip, nVVenta, nIGV, nPVenta,cDocNroRefe, dDocRefeFec) "
    lsSql = lsSql & "VALUES   ('" & psOpeTpo & "','" & pnDocTpo & "','" & psDocNro & "','" & Format(pdDocFecha, "mm/dd/yyyy hh:mm:ss") & "', '"
    lsSql = lsSql & psPersCod & "','" & psCtaCod & "','" & psDescrip & "'," & pnVVenta & "," & pnIGV & "," & pnPVentas & ",'" & psDocNroRefe & "', '" & Format(pdDocFecRef, "mm/dd/yyyy hh:mm:ss") & "')"
    oConect.Ejecutar lsSql
End Sub
Public Sub ActualizaEstadoCuotasCronogramaxProvision(ByVal pnMovNroComprobante As Long, Optional ByVal pbExtorno As Boolean = False)
    Dim lsSql As String
    On Error GoTo ErrActualizaEstadoCuotasCronogramaxProvision
    lsSql = "Exec stp_upd_ERS0622013_ActualizaEstadoCuotasCronogramaxProvision " & pnMovNroComprobante & "," & IIf(pbExtorno, 1, 0)
    oConect.Ejecutar lsSql
    Exit Sub
ErrActualizaEstadoCuotasCronogramaxProvision:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotasCronogramaxProvision Method")
End Sub
Public Sub ActualizaEstadoCuotasCronogramaxPagoProveedor(ByVal pnMovNroProvision As Long)
    Dim lsSql As String
    On Error GoTo ErrActualizaEstadoCuotasCronogramaxPagoProveedor
    lsSql = "Exec stp_upd_ERS0622013_ActualizaEstadoCuotasCronogramaxPagoProveedor " & pnMovNroProvision
    oConect.Ejecutar lsSql
    Exit Sub
ErrActualizaEstadoCuotasCronogramaxPagoProveedor:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotasCronogramaxPagoProveedor Method")
End Sub
Public Sub ActualizaEstadoCuotasCronogramaxExtPagoProveedor(ByVal pnMovNroPagoProv As Long)
    Dim lsSql As String
    On Error GoTo ErrActualizaEstadoCuotasCronogramaxExtPagoProveedor
    lsSql = "Exec stp_upd_ERS0622013_ActualizaEstadoCuotasCronogramaxExtPagoProveedor " & pnMovNroPagoProv
    oConect.Ejecutar lsSql
    Exit Sub
ErrActualizaEstadoCuotasCronogramaxExtPagoProveedor:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotasCronogramaxExtPagoProveedor Method")
End Sub
Public Function ListaItemsPenalidad(ByVal pnMovNro As Long) As ADODB.Recordset
    Dim lsSql As String
    On Error GoTo ErrListaItemsPenalidad
    lsSql = "Exec stp_sel_ERS0622013_ListaItemsPenalidad " & pnMovNro
    Set ListaItemsPenalidad = oConect.CargaRecordSet(lsSql)
    Exit Function
ErrListaItemsPenalidad:
    Call RaiseError(MyUnhandledError, "DMov:ListaItemsPenalidad Method")
End Function
'END EJVG *******
'EJVG20140128 ***
Public Function ChequeDetxID(ByVal pnId As Long) As ADODB.Recordset
    Dim lsSql As String
    lsSql = "Exec stp_sel_ERS1262013_ChequeDetxId " & pnId
    Set ChequeDetxID = oConect.CargaRecordSet(lsSql)
End Function
Public Sub ValorizaChequeNeg(ByVal pnId As Long, ByVal pdFechaValoriza As Date, ByVal psMovNro As String) 'psMovNro agregado PASI20140530
    Dim lsSql As String
    lsSql = "Exec stp_upd_ERS1262013_ValorizaCheque " & pnId & ",'" & Format(pdFechaValoriza, "yyyymmdd hh:mm:ss") & "','" & psMovNro & "'"
    oConect.Ejecutar lsSql
End Sub
Public Sub RechazaChequeNeg(ByVal pnId As Long, ByVal pdFechaRechaza As Date)
    Dim lsSql As String
    lsSql = "Exec stp_upd_ERS1262013_RechazaCheque " & pnId & ",'" & Format(pdFechaRechaza, "yyyymmdd hh:mm:ss") & "'"
    oConect.Ejecutar lsSql
End Sub
Public Sub InsertaMovDocRec(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psNroDoc As String, ByVal psPersCod As String, ByVal psIFTpo As String, ByVal psIFCta As String)
    On Error GoTo InsertaMovDocErr
    Dim sql As String
    sql = "Exec stp_ins_ERS1262013_MovDocRec " & pnMovNro & "," & pnDocTpo & ",'" & psNroDoc & "','" & psPersCod & "','" & psIFTpo & "','" & psIFCta & "'"
    oConect.Ejecutar sql
    Exit Sub
InsertaMovDocErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovDocRec Method")
End Sub
Public Sub ExtornaValorizaChequeNeg(ByVal pnId As Long)
    Dim lsSql As String
    lsSql = "Exec stp_upd_ERS1262013_ExtornaValorizaCheque " & pnId
    oConect.Ejecutar lsSql
End Sub
Public Sub ActualizaNroMovRechazo(ByVal pnMovNro As Long, ByVal pnMovNroDeposito As Long)
    Dim lsSql As String
    lsSql = "Exec stp_upd_ERS1262013_ActualizaNroMovRechazo " & pnMovNro & "," & pnMovNroDeposito
    oConect.Ejecutar lsSql
    Exit Sub
End Sub
Public Sub ExtornaRechazoChequeNeg(ByVal pnId As Long)
    Dim lsSql As String
    lsSql = "Exec stp_upd_ERS1262013_ExtornaRechazoCheque " & pnId
    oConect.Ejecutar lsSql
End Sub
'END EJVG *******
'EJVG20140407 ***
Public Function ExisteMovRef(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long) As Boolean
    Dim rs As New ADODB.Recordset
    Dim lsSql As String
    lsSql = "SELECT nCuenta = count(*) FROM MovRef MR WHERE MR.nMovNro=" & pnMovNro & " AND nMovNroRef=" & pnMovNroRef
    Set rs = oConect.CargaRecordSet(lsSql)
    If Not rs.EOF Then
        ExisteMovRef = IIf(rs!nCuenta > 0, True, False)
    End If
    rs.Close
    Set rs = Nothing
End Function
'END EJVG *******
'*** PASI TI-ERS001-2014
Public Function InsertaMovIntgAmort(ByVal pnAnio As Long, ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psIntCod As String, ByVal pnNroAmort As Long, ByVal pnMovNroRef As Long, ByVal pbEstadistico As Integer, ByVal pdFechaRef As Date) As Integer
    On Error GoTo ErrorInsertaMovIntgAmort
    Dim sql As String
    sql = "stp_ins_ERS0012014_InsertarMovIntgAmort " & pnAnio & "," & pnMovNro & "," & pnMovItem & ",'" & psIntCod & "'," & pnNroAmort & "," & pnMovNroRef & "," & pbEstadistico & ",'" & Format(pdFechaRef, "yyyyMMdd") & "'"
    oConect.Ejecutar sql
    Exit Function
ErrorInsertaMovIntgAmort:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovIntgAmort Method")
End Function
Public Function GeneraAsientoRes(pnMovNro As Long, pnMovNroBase As Long)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = "stp_sel_ERS0012014_GeneraAsientoRes " & pnMovNroBase
    Set rs = oConect.CargaRecordSet(sql)
     
    While Not rs.EOF
        InsertaMovCta pnMovNro, rs.Bookmark, rs!cCtaContCod, rs!Monto
        rs.MoveNext
    Wend
End Function
'*** end PASI
'EJVG20140513 ***
Public Function InsertaMovHabilitacionTransp(ByVal pnMovNro As Long, ByVal pbPropioTransp As Boolean, ByVal psPersCodTransp As String) As Integer
    On Error GoTo ErrInsertaMovHabilitacionTransp
    Dim lsSql As String
    lsSql = "EXEC stp_ins_ERS0252014_MovHabilitacionTransp " & pnMovNro & "," & IIf(pbPropioTransp, 1, 0) & ",'" & psPersCodTransp & "'"
    oConect.Ejecutar lsSql
    Exit Function
ErrInsertaMovHabilitacionTransp:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovHabilitacionTransp Method")
End Function
'END EJVG *******
'FRHU 20140530 ERS068-2014
Public Sub InsertaSegTarjetaEstado(ByVal psNumSolicitud As String, ByVal pdPrdEstado As Date, ByVal pnPrdEstado As Integer, ByVal pcDescripcion As String, ByVal pnMonto As Currency, ByVal psUltimaActualizacion As String)
On Error GoTo InsertaErr
    Dim lsSql As String

    lsSql = "exec stp_ins_SegTarjetaSolicitudEstado '" & psNumSolicitud & "','" & Format(pdPrdEstado, "yyyymmdd") & "'," & pnPrdEstado & ",'" & pcDescripcion & "'," & pnMonto & ",'" & psUltimaActualizacion & "'"
    oConect.Ejecutar lsSql
    
        Exit Sub
InsertaErr:
    Call RaiseError(MyUnhandledError, "DSeguros:InsertaSegTarjeta Method")
End Sub
Public Sub ActualizarSegTarjetaSolicitudAproAcep(ByVal psNumSolicitud As String, ByVal psMovNro As String, ByVal pnEstado As Integer, ByVal psCtaDeposito As String, ByVal pnMonto As Currency)
On Error GoTo ActualizarErr
    Dim lsSql As String

    lsSql = "exec stp_upd_SegTarjetaSolicitudAproAcep '" & psNumSolicitud & "','" & psMovNro & "'," & pnEstado & ",'" & psCtaDeposito & "'," & pnMonto
    oConect.Ejecutar lsSql
    
    Exit Sub
ActualizarErr:
    Call RaiseError(MyUnhandledError, "DSeguros:ActualizarSegTarjeta Method")
End Sub
Public Function ObtenerSolicitudesParaExtornarSegTarjeta(ByVal psMovNro As String, ByVal pnOperacion As Integer) As ADODB.Recordset
    Dim lsSql As String
    lsSql = "Exec stp_sel_SegTarjetaSolicitudParaExtornar '" & psMovNro & "'," & pnOperacion
    Set ObtenerSolicitudesParaExtornarSegTarjeta = oConect.CargaRecordSet(lsSql)
End Function
Public Sub ActualizarSolicitudExtornarSegTarjeta(ByVal psMovNro As String, ByVal pnEstado As Integer, ByVal psMovNroExt As String)
On Error GoTo ActualizarErr
    Dim lsSql As String
    Dim rs As ADODB.Recordset
    
    Set rs = ObtenerSolicitudesParaExtornarSegTarjeta(psMovNro, 1)
    
    If Not rs.BOF And Not rs.EOF Then
        Do While Not rs.EOF
            lsSql = "exec stp_upd_SegTarjetaSolicitudParaExtornar '" & rs!cNumSolicitud & "','" & psMovNroExt & "'," & pnEstado & "," & CCur(rs!nMonto) & ",'" & psMovNro & "'"
            oConect.Ejecutar lsSql
            rs.MoveNext
        Loop
    End If
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ActualizarErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizarSegTarjetaExtornar Method")
End Sub
'FIN FRHU 20140530
'JUEZ 20140711 **************************************************************
Public Sub InsertaSegTarjetaOpeBancos(ByVal pnMovNro As Long, ByVal psAnio As String, ByVal psMes As String, ByVal psFechaPago As String, ByVal psTime As String, ByVal pnMonto As Double, ByVal pnTipoOpe As Integer)
On Error GoTo InsertaErr
    Dim lsSql As String

    lsSql = "exec stp_ins_SegTarjetaOpeBancos " & pnMovNro & ",'" & psAnio & "','" & psMes & "','" & Format(psFechaPago & " " & psTime, "mm/dd/yyyy hh:mm:ss") & "'," & pnMonto & "," & pnTipoOpe
    oConect.Ejecutar lsSql
    
        Exit Sub
InsertaErr:
    Call RaiseError(MyUnhandledError, "DSeguros:InsertaSegTarjeta Method")
End Sub
'END JUEZ *******************************************************************
'APRI20180130 ERS028-2017
Public Sub InsertaSegurosOpeBancos(ByVal pnMovNro As Long, ByVal psAnio As String, ByVal psMes As String, ByVal psFechaPago As String, ByVal psTime As String, ByVal pnMonto As Double, ByVal pnMoneda As Integer, ByVal pnTipoSeg As Integer)
On Error GoTo InsertaErr
    Dim lsSql As String

    lsSql = "exec stp_ins_SegurosOpeBancos " & pnMovNro & ",'" & psAnio & "','" & psMes & "','" & Format(psFechaPago & " " & psTime, "mm/dd/yyyy hh:mm:ss") & "'," & pnMonto & "," & pnMoneda & "," & pnTipoSeg
    oConect.Ejecutar lsSql
    
        Exit Sub
InsertaErr:
    Call RaiseError(MyUnhandledError, "DSeguros:InsertaSegTarjeta Method")
End Sub
'END APRI
'EJVG20140727 ***
Public Sub InsertaMovProvRetencSistPens(ByVal pnMovNro As Long, ByVal pnTpoSistPens As Integer, ByVal pnMontoBase As Currency, ByVal pnMontoRetencion As Currency)
    On Error GoTo ErrInsertaMovProvRetencSistPens
    Dim lsSql As String
    lsSql = "EXEC stp_ins_ERS1022014_MovProvRetencSistPens " & pnMovNro & "," & pnTpoSistPens & "," & pnMontoBase & "," & pnMontoRetencion
    oConect.Ejecutar lsSql
    Exit Sub
ErrInsertaMovProvRetencSistPens:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovProvRetencSistPens Method")
End Sub
Public Sub InsertaMovProvRetencSistPensDet(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal pnConceptoRetencion As Integer, ByVal pnMontoCalculado As Currency, ByVal pnMontoCobrado As Currency)
    On Error GoTo ErrInsertaMovProvRetencSistPens
    Dim lsSql As String
    lsSql = "EXEC stp_ins_ERS1022014_MovProvRetencSistPensDet " & pnMovNro & "," & pnMovItem & "," & pnConceptoRetencion & "," & pnMontoCalculado & "," & pnMontoCobrado
    oConect.Ejecutar lsSql
    Exit Sub
ErrInsertaMovProvRetencSistPens:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovProvRetencSistPens Method")
End Sub
'END EJVG *******
'EJVG20140811 ***
Public Function ExisteMovimientosDespues(ByVal pnMovNro As Long, ByVal pnSoloVigentes As Boolean) As Boolean
    On Error GoTo ErrExisteMovimientosDespues
    Dim lsSql As String
    Dim rs As New ADODB.Recordset
    lsSql = "select dbo.fnc_ERS0252014_ExisteMovimientosDespues(" & pnMovNro & "," & IIf(pnSoloVigentes, 1, 0) & ") bExiste"
    Set rs = oConect.CargaRecordSet(lsSql)
    If Not rs.EOF Then
        ExisteMovimientosDespues = rs!bExiste
    End If
    Exit Function
ErrExisteMovimientosDespues:
    Call RaiseError(MyUnhandledError, "DMov:ExisteMovimientosDespues Method")
End Function
'END EJVG *******
'FRHU 20140814 ANEXO-03
Public Sub ExtornarMovimientoDepositoSegTarjetaReferencia(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long)

On Error GoTo ExtornarMovimientoDepositoSegTarjetaReferenciaErr
    Dim sql As String
    Dim lnMovNronew As Long
    Dim lsMovNroAnt As String
    
    lsMovNroAnt = GetcMovNro(pnMovNroAnt)

    sql = "UPDATE Mov SET nMovFlag = '" & gMovFlagEliminado & "' WHERE cMovNro = '" & lsMovNroAnt & "'"
    oConect.Ejecutar sql

    Call ActualizarSolicitudExtornarSegTarjeta(lsMovNroAnt, 1000, psMovNroNew)

Exit Sub
ExtornarMovimientoDepositoSegTarjetaReferenciaErr:
    Call RaiseError(MyUnhandledError, "DMov:ExtornaMovimiento Method")
    
End Sub
'FIN FRHU
'PASIERS0772014
Public Sub ActualizaComprobanteLogNew(ByVal pnMovNro As Long, ByVal pnDocTpo As Integer, ByVal psDocNro As String, ByVal pdDocFecha As Date)
    Dim lsSql As String
    lsSql = "Exec stp_upd_ERS0772014_ComprobanteReg " & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & Format(pdDocFecha, "yyyymmdd") & "'"
    oConect.Ejecutar lsSql
End Sub
Public Sub ActualizaEstadoCuotasCronogramaxProvisionNew(ByVal pnMovNroComprobante As Long, Optional ByVal pbExtorno As Boolean = False)
    Dim lsSql As String
    On Error GoTo ErrActualizaEstadoCuotasCronogramaxProvision
    lsSql = "Exec stp_upd_ERS0772014_ActualizaEstadoCuotasCronogramaxProvision " & pnMovNroComprobante & "," & IIf(pbExtorno, 1, 0)
    oConect.Ejecutar lsSql
    Exit Sub
ErrActualizaEstadoCuotasCronogramaxProvision:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotasCronogramaxProvision Method")
End Sub
Public Sub ActualizaEstadoBienesxProvision(ByVal pnMovNroComprobante As Long, Optional ByVal pbExtorno As Boolean = False)
    Dim lsSql As String
    On Error GoTo ErrActualizaEstadoBienesxProvision
    lsSql = "Exec stp_upd_ERS0772014_ActualizaEstadoBienesxProvision " & pnMovNroComprobante & "," & IIf(pbExtorno, 1, 0)
    oConect.Ejecutar lsSql
    Exit Sub
ErrActualizaEstadoBienesxProvision:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotasCronogramaxProvision Method")
End Sub
Public Sub ActualizaEstadoCuotasCronogramaxPagoProveedorNew(ByVal pnMovNroProvision As Long)
    Dim lsSql As String
    On Error GoTo ErrActualizaEstadoCuotasCronogramaxPagoProveedorNew
    lsSql = "stp_upd_ERS0772014_ActualizaEstadoCuotasCronogramaxPagoProveedor " & pnMovNroProvision
    oConect.Ejecutar lsSql
    Exit Sub
ErrActualizaEstadoCuotasCronogramaxPagoProveedorNew:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotasCronogramaxPagoProveedorNew Method")
End Sub
Public Sub ActualizaEstadoBienesxPagoProveedor(ByVal pnMovNroProvision As Long)
    Dim lsSql As String
    On Error GoTo ErrActualizaEstadoBienesxPagoProveedor
    lsSql = "stp_upd_ERS0772014_ActualizaEstadoBienesxPagoProveedor " & pnMovNroProvision
    oConect.Ejecutar lsSql
    Exit Sub
ErrActualizaEstadoBienesxPagoProveedor:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoBienescPagoProveedor Method")
End Sub
Public Sub ActualizaEstadoCuotasCronogramaxExtPagoProveedorNew(ByVal pnMovNroPagoProv As Long)
    Dim lsSql As String
    On Error GoTo ErrActualizaEstadoCuotasCronogramaxExtPagoProveedorNew
    lsSql = "stp_upd_ERS0772014_ActualizaEstadoCuotasCronogramaxExtPagoProveedor " & pnMovNroPagoProv
    oConect.Ejecutar lsSql
    Exit Sub
ErrActualizaEstadoCuotasCronogramaxExtPagoProveedorNew:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotasCronogramaxExtPagoProveedorNew Method")
End Sub
Public Sub ActualizaEstadoBienesxExtPagoProveedor(ByVal pnMovNroPagoProv As Long)
    Dim lsSql As String
    On Error GoTo ErrActualizaEstadoBienesxExtPagoProveedor
    lsSql = "stp_upd_ERS0772014_ActualizaEstadoBienesxExtPagoProveedor " & pnMovNroPagoProv
    oConect.Ejecutar lsSql
    Exit Sub
ErrActualizaEstadoBienesxExtPagoProveedor:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoBienesxExtPagoProveedor Method")
End Sub
'END PASI
'ALPA 20150106*******************************************************************
Public Function InsertaMovContMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovMonto As Currency, ByVal pnMovCorrela As Integer, ByVal psMovParalelo As String) As Integer
    On Error GoTo InsertaMovContMIVIVIENDAErr
    Dim sql As String
    InsertaMovContMIVIVIENDA = 1
    
    sql = "INSERT INTO MovContMIVIVIENDA (nMovNro, nMovMonto, nMovCorrela, cMovParalelo ) " _
         & "VALUES (" & pnMovNro & "," & pnMovMonto & "," & pnMovCorrela & ",'" & psMovParalelo & "')"
     
    oConect.Ejecutar sql
    InsertaMovContMIVIVIENDA = 0
    Exit Function
InsertaMovContMIVIVIENDAErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovContMIVIVIENDA Method")
End Function
Public Function InsertaMovCtaMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency, Optional pnTipoPago As Integer = 0, Optional ByRef psTextoError As String = "") As Integer
    On Error GoTo InsertaMovCtaMIVIVIENDAErr
    Dim sql As String
    InsertaMovCtaMIVIVIENDA = 1

    If lbBitControlCuentasUltNivel Then
        sql = "Select cCtaContCod From CtaCont where cCtaContCod Like '" & psCtaContCod & "%' And cCtaContCod <> '" & psCtaContCod & "' And nCtaEstado = 1"
        Set rsCta = oConect.CargaRecordSet(sql)

        If rsCta.EOF And rsCta.BOF Then
            rsCta.Close

            sql = "Select cCtaContCod from CtaCont Where cCtaContCod = '" & psCtaContCod & "' And nCtaEstado = 1"
            Set rsCta = oConect.CargaRecordSet(sql)

            If rsCta.EOF And rsCta.BOF Then
                rsCta.Close
                Err.Number = 9999
                Err.Description = "Cuenta Contable no existe o no es valida " & psCtaContCod & "."
            Else
                sql = " INSERT INTO MOVCTAMIVIVIENDA(nMovNro, nMovItem, cCtaContCod, nMovImporte,nTipoPago)" _
                   & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ", " & pnTipoPago & ") "
                oConect.Ejecutar sql
                InsertaMovCtaMIVIVIENDA = 0
                rsCta.Close
                Exit Function
            End If
        Else
            rsCta.Close
            Err.Number = 9999
            Err.Description = "Se desea grabar una Cta Cont que no esta a Ult Nivel " & psCtaContCod & " ."
            psTextoError = "A esta cuenta Contable " & psCtaContCod & " le falta asignar sub cuentas "
        End If
    Else
        sql = " INSERT INTO MOVCTAMIVIVIENDA(nMovNro, nMovItem, cCtaContCod, nMovImporte,nTipoPago )"
        sql = sql & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ", " & pnTipoPago & ") "
        oConect.Ejecutar sql
        InsertaMovCtaMIVIVIENDA = 0
        Exit Function
    End If
InsertaMovCtaMIVIVIENDAErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCtaMIVIVIENDA Method")
End Function
Public Function InsertaMovObjMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psObjetoCod As String) As Integer
    On Error GoTo InsertaMovObjMIVIVIENDAErr
    Dim sql As String
    InsertaMovObjMIVIVIENDA = 1
     
    sql = "INSERT INTO MOVOBJMIVIVIENDA (nMovNro, nMovItem, nMovObjOrden, cObjetoCod) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & Format(psObjetoCod, "00") & "')"

    oConect.Ejecutar sql
    InsertaMovObjMIVIVIENDA = 0
    Exit Function
InsertaMovObjMIVIVIENDAErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjMIVIVIENDA Method")
End Function
Public Function InsertaMovObjIFMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As String, ByVal psPersCod As String, ByVal psTipoIF As CGTipoIF, ByVal psCtaIFCod As String) As Integer
    On Error GoTo InsertaMovObjIFMIVIVIENDAErr
    Dim sql As String
    InsertaMovObjIFMIVIVIENDA = 1
     
    sql = "INSERT INTO MOVOBJIFMIVIVIENDA(nMovNro, nMovItem, nMovObjOrden, cPersCod, cIFTpo, cCtaIfCod ) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psPersCod & "','" & Format(psTipoIF, "00") & "','" & psCtaIFCod & "')"
    
    oConect.Ejecutar sql
    InsertaMovObjIFMIVIVIENDA = 0
    Exit Function
InsertaMovObjIFMIVIVIENDAErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjIFMIVIVIENDA Method")
End Function
Public Sub GrabaMovimientoEfectivoMIVIVIENDA(ByVal pnMovNro As Long, ByRef pnMovItem As Long, rsEfectivo As ADODB.Recordset, psCuenta As String, psDH As String)
On Error GoTo GrabaMovimientoEfectivoMIVIVIENDAErr
Dim lnMovOrden As Integer
'Guardamos el Billetaje
If Not rsEfectivo Is Nothing Then
    If rsEfectivo.State = adStateOpen Then
        If Not rsEfectivo.EOF And Not rsEfectivo.BOF Then
            Do While Not rsEfectivo.EOF
                If rsEfectivo!Monto <> 0 Then
                    pnMovItem = pnMovItem + 1: lnMovOrden = 0
                    InsertaMovCtaMIVIVIENDA pnMovNro, Format(pnMovItem, "#0"), psCuenta, rsEfectivo!Monto * IIf(psDH = "D", 1, -1)
                    lnMovOrden = lnMovOrden + 1
                    InsertaMovObjMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, ObjDescomEfectivo
                    InsertaMovObjEfectivo pnMovNro, pnMovItem, lnMovOrden, rsEfectivo!cEfectivoCod
                End If
                rsEfectivo.MoveNext
            Loop
        End If
        rsEfectivo.Close: Set rsEfectivo = Nothing
    End If
End If
Exit Sub
GrabaMovimientoEfectivoMIVIVIENDAErr:
    Err.Raise Err.Number, "DMov: GrabaMovimientoEfectivoMIVIVIENDA", Err.Description
End Sub
Public Function InsertaMovObjEfectivoMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psEfectivoCod As String) As Integer
    On Error GoTo InsertaMovObjEfectivoMIVIVIENDAErr
    Dim sql As String
    InsertaMovObjEfectivoMIVIVIENDA = 1
     
    sql = "INSERT INTO MovObjEfectivoMIVIVIENDA(nMovNro, nMovItem, nMovObjOrden, cEfectivoCod ) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psEfectivoCod & "')"

    oConect.Ejecutar sql
    InsertaMovObjEfectivoMIVIVIENDA = 0
    Exit Function
InsertaMovObjEfectivoMIVIVIENDAErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjEfectivoMIVIVIENDA Method")
End Function
Public Function InsertaMovObjAgenciaAreaMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovObjOrden As Long, ByVal psAgeCod As String, ByVal psAreaCod As String) As Integer
    On Error GoTo InsertaMovObjAgenciaAreaMIVIVIENDAErr
    Dim sql As String
    InsertaMovObjAgenciaAreaMIVIVIENDA = 1
    
    sql = "INSERT INTO MovObjAreaAgenciaMIVIVIENDA(nMovNro, nMovItem, nMovObjOrden, cAreaCod, cAgeCod) " _
         & "VALUES (" & pnMovNro & "," & pnMovItem & "," & pnMovObjOrden & ",'" & psAreaCod & "','" & psAgeCod & "')"
     
    oConect.Ejecutar sql
    InsertaMovObjAgenciaAreaMIVIVIENDA = 0
    Exit Function
InsertaMovObjAgenciaAreaMIVIVIENDAErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjAgenciaAreaMIVIVIENDA Method")
End Function
Public Function InsertaMovRefMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long, Optional psAgeCodRef As String = "", Optional pnTipoPago As Integer = 0) As Integer
    On Error GoTo InsertaMovRefMIVIVIENDAErr
    Dim sql As String
    InsertaMovRefMIVIVIENDA = 1
    sql = "INSERT INTO MovRefMIVIVIENDA(nMovNro, nMovNroRef, cAgeCodRef,nTipoPago ) " _
         & "VALUES (" & pnMovNro & "," & pnMovNroRef & ", '" & psAgeCodRef & "'," & pnTipoPago & ")"

    oConect.Ejecutar sql
    InsertaMovRefMIVIVIENDA = 0
    Exit Function
InsertaMovRefMIVIVIENDAErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovContMIVIVIENDAErr Method")
End Function
Public Sub GrabaMovOtrosCtaObjMIVIVIENDA(pnMovNro As Long, pnMovItem As Long, prsOtro As ADODB.Recordset, prsObj As ADODB.Recordset, psDH As String)
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
            InsertaMovCtaMIVIVIENDA pnMovNro, pnMovItem, prsOtro!Cuenta & lsFiltro, prsOtro!Importe * IIf(psDH = "D", 1, -1)
            lsFiltro = ""
            If Not prsObj Is Nothing Then
                prsObj.MoveFirst
                lnMovOrden = 0
                Do While Not prsObj.EOF
                    If prsObj.Fields(0) = prsOtro.Fields(0) Then
                        Select Case prsObj!ObjPadre
                            Case ObjCMACAgencias
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObjMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
                                InsertaMovObjAgenciaAreaMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, prsObj!Código, ""
                            Case ObjCMACArea
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObjMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
                                InsertaMovObjAgenciaAreaMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, "", prsObj!Código
                            Case ObjCMACAgenciaArea
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObjMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
                                InsertaMovObjAgenciaAreaMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, Mid(prsObj!Código, 4, 2), Mid(prsObj!Código, 1, 3)
                            Case ObjDescomEfectivo
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObjMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
                                InsertaMovObjEfectivoMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, prsObj!Código
                            Case ObjEntidadesFinancieras
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObjMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, prsObj!ObjPadre
                                InsertaMovObjIFMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, Mid(prsObj!Código, 4, 13), Mid(prsObj!Código, 1, 2), Mid(prsObj!Código, 18, 10)
                            Case Else
                                lnMovOrden = lnMovOrden + 1
                                InsertaMovObjMIVIVIENDA pnMovNro, pnMovItem, lnMovOrden, prsObj!Código
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
Public Function InsertaMovDocMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psDocFecha As String, _
                              Optional ByVal pcLugar As String = "") As Integer
    On Error GoTo InsertaMovDocMIVIVIENDAErr
    Dim sql As String
    InsertaMovDocMIVIVIENDA = 1

    sql = "exec stp_sel_RFC1262012_RegistrarMovDocMIVIVIENDA " & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & psDocFecha & "','" & pcLugar & "'"
    
    oConect.Ejecutar sql
    InsertaMovDocMIVIVIENDA = 0
    Exit Function
InsertaMovDocMIVIVIENDAErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovDocMIVIVIENDA Method")
End Function
Public Function ObtenerMov(ByVal pnMovNro As Long) As ADODB.Recordset
    Dim lsSql As String
    lsSql = "select * from Mov where nMovNro= " & pnMovNro
    Set ObtenerMov = oConect.CargaRecordSet(lsSql)
End Function
Public Sub ObtenerMovContMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovNroOriginal As Long)
    Dim lsSql As String
    lsSql = "insert into MovCont "
    lsSql = lsSql & "select " & pnMovNro & ",nMovMonto,nMovCorrela,cMovParalelo from MovContMIVIVIENDA where nMovNro= " & pnMovNroOriginal
    oConect.Ejecutar lsSql
End Sub
Public Sub ObtenerMovCtaMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovNroOriginal As Long)
    Dim lsSql As String
    lsSql = "insert into MovCta "
    lsSql = lsSql & "select " & pnMovNro & ",nMovItem,cCtaContCod,nMovImporte,nTipoPago from MovCtaMIVIVIENDA where nMovNro= " & pnMovNroOriginal
    oConect.Ejecutar lsSql
End Sub
Public Sub ObtenerMovObjMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovNroOriginal As Long)
    Dim lsSql As String
    lsSql = "insert into MovObj "
    lsSql = lsSql & "select " & pnMovNro & ",nMovItem,nMovObjOrden,cObjetoCod from MovObjMIVIVIENDA where nMovNro= " & pnMovNroOriginal
    oConect.Ejecutar lsSql
End Sub
Public Sub ObtenerMovObjIFMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovNroOriginal As Long)
    Dim lsSql As String
    lsSql = "insert into MovObjIF "
    lsSql = lsSql & "select " & pnMovNro & ",nMovItem,nMovObjOrden,cPersCod,cIFTpo,cCtaIfCod from MovObjIFMIVIVIENDA where nMovNro= " & pnMovNroOriginal
    oConect.Ejecutar lsSql
End Sub
Public Sub ObtenerMovObjEfectivoMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovNroOriginal As Long)
    Dim lsSql As String
    lsSql = "insert into MovObjEfectivo "
    lsSql = lsSql & "select " & pnMovNro & ",nMovItem,nMovObjOrden,cEfectivoCod from MovObjEfectivoMIVIVIENDA where nMovNro= " & pnMovNroOriginal
    oConect.Ejecutar lsSql
End Sub
Public Sub ObtenerMovObjAreaAgenciaMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovNroOriginal As Long)
    Dim lsSql As String
    lsSql = "insert into MovObjAreaAgencia "
    lsSql = lsSql & "select " & pnMovNro & ",nMovItem,nMovObjOrden,cAreaCod,cAgeCod from MovObjAreaAgenciaMIVIVIENDA where nMovNro= " & pnMovNroOriginal
    oConect.Ejecutar lsSql
End Sub
Public Sub ObtenerMovRefMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovNroOriginal As Long)
    Dim lsSql As String
    lsSql = "insert into MovRef "
    lsSql = lsSql & "select " & pnMovNro & ",nMovNroRef,cAgeCodRef,cAgeCodBen,nTipoPago from MovRefMIVIVIENDA where nMovNro= " & pnMovNroOriginal
    oConect.Ejecutar lsSql
End Sub
Public Sub ObtenerMovDocMIVIVIENDA(ByVal pnMovNro As Long, ByVal pnMovNroOriginal As Long)
    Dim lsSql As String
    lsSql = "insert into MovDoc "
    lsSql = lsSql & "select " & pnMovNro & ",nDocTpo,cDocNro,dDocFecha,cMovDocDesc from MovDocMIVIVIENDA where nMovNro= " & pnMovNroOriginal
    oConect.Ejecutar lsSql
End Sub
'********************************************************************************
'PASIERS0872014********************************************************************
Public Function InsertaMovRetPagoSeguro(ByVal pnMovNro As Long, ByVal psAnio As String, ByVal psMes As String, ByVal pdFecPago As Date, ByVal pnMonto As Currency, ByVal pnMoneda As Integer) As Integer
    On Error GoTo ErrInsertaMovRetPagoSeg
    Dim sql As String
    InsertaMovRetPagoSeguro = 1
        sql = "stp_ins_ERS1362014_MovRetPagoSeguro " & pnMovNro & ",'" & psAnio & "','" & psMes & "','" & Format(pdFecPago, "yyyyMMdd") & "'," & pnMonto & "," & pnMoneda
        oConect.Ejecutar sql
    InsertaMovRetPagoSeguro = 0
    Exit Function
ErrInsertaMovRetPagoSeg:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovRetPagoSeguro Method")
End Function
Public Function InsertaMovRetPagoSeguroDet(ByVal pnMovNro As Long, ByVal psAnio As String, ByVal psMes As String, ByVal psCtaCod As String, ByVal pnCuota As Integer, ByVal pnNroCalen As Integer, ByVal pnMonto As Currency, ByVal psAgeCod As String) As Integer
    On Error GoTo ErrInsertaMovRetPagoSegDet
    Dim sql As String
    InsertaMovRetPagoSeguroDet = 1
        sql = "stp_ins_ERS1362014_MovRetPagoSeguroDet " & pnMovNro & ",'" & psAnio & "','" & psMes & "','" & psCtaCod & "'," & pnCuota & "," & pnNroCalen & "," & pnMonto & ",'" & psAgeCod & "'"
        oConect.Ejecutar sql
    InsertaMovRetPagoSeguroDet = 0
 Exit Function
ErrInsertaMovRetPagoSegDet:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovRetPagoSeguroDet Method")
End Function
'end PASI**************************************************************************
'PASI20150417 INC1504160009
Public Function GeneraMovMExProvision(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional nTCComp1 As Currency, Optional ncOpeCod As String = "", Optional nTipoPago As Integer = 0, Optional ByVal bSpot As Boolean = False, _
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
Optional ByVal pnValVentTr As Currency = 0, _
Optional ByVal pdFechaEmisionComprobante As String = "") As Boolean
    Dim nTCFijo As Currency
    Dim nTCVent As Currency
    Dim nTCComp As Currency
    Dim sSQL As String
    Dim rs   As New ADODB.Recordset
    Dim dFecha As Date
    Dim nImporteD As Currency
    Dim nImporteS As Currency
    On Error GoTo GeneraMovMEErr
    GeneraMovMExProvision = False
    
    dFecha = CDate(GetFechaMov(psMovNro, True))
   
    
    
    
    Dim oTC As New nTipoCambio
    nTCFijo = oTC.EmiteTipoCambio(dFecha, TCFijoDia)
    
    If nTCMerc = 0 Then
        sSQL = "Select cOpeCod From Mov  Where nMovNro = " & pnMovNro
        Set rs = oConect.CargaRecordSet(sSQL)
        If rs.EOF And rs.BOF Then
            nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
            nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)
        Else
            If Left(rs!cOpeCod, 1) = "4" Or Left(rs!cOpeCod, 1) = "7" Then
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCPondVenta)
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCPonderado)
            Else
                nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
                nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)
            End If
        End If
    Else
        InsertaMovTpoCambio pnMovNro, nTCMerc, pnValVent, pnValComp, pnValFijo, pnValFijoDia, pnValVentEsp, pnValCompEsp, pnValPond, pnValPondVenta, pnValPondREU, pnValSBSDia, pnValCompTr, pnValVentTr
        If ncOpeCod = "702403" Then
            nTCVent = nTCMerc
            nTCComp = nTCComp1
        ElseIf ncOpeCod = "702406" Then 'TORE RFC17122260002
            dFecha = CDate(pdFechaEmisionComprobante)
            nTCVent = oTC.EmiteTipoCambio(dFecha, TCPondVenta)
            nTCComp = oTC.EmiteTipoCambio(dFecha, TCPondVenta)
        Else 'END TORE RFC17122260002
            nTCVent = nTCMerc
            nTCComp = nTCMerc
        End If
    End If
    Set oTC = Nothing
    oConect.Ejecutar " sp_GeneraMovMExProvision " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent & "," & nTipoPago
    GeneraMovMExProvision = True

Exit Function
GeneraMovMEErr:
    Call RaiseError(MyUnhandledError, "DMov:GeneraMovME Method")
End Function
'end PASI
'EJVG20151007 ***
Public Function RequiereReLiberaGarantiasCreditoDiferenteMoneda(ByVal pdFecha As Date) As Boolean
    Dim lsSql As String
    Dim rs As New ADODB.Recordset
    On Error GoTo ErrRequiere
    lsSql = "SELECT bRealizarReLiberacion = DBO.fnc_ERS0632014_RequiereReLiberaGarantiaCreditoDiferenteMoneda('" & Format(pdFecha, "yyyymmdd") & "')"
    Set rs = oConect.CargaRecordSet(lsSql)
    If Not rs.EOF Then
        RequiereReLiberaGarantiasCreditoDiferenteMoneda = rs!bRealizarReLiberacion
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
ErrRequiere:
    Err.Raise Err.Number, "Error al consultar si se va a reliberar garantías", Err.Description
End Function
Public Sub ReLiberaGarantiasCreditoDiferenteMoneda(ByVal pdFecha As Date)
    Dim lsSql As String
    On Error GoTo ErrReLiberaGarantiasCreditoDiferenteMoneda
    lsSql = "EXEC stp_upd_ERS0632014_ReLiberaGarantiaCreditoDiferenteMoneda '" & Format(pdFecha, "yyyymmdd") & "'"
    oConect.Ejecutar lsSql
    Exit Sub
ErrReLiberaGarantiasCreditoDiferenteMoneda:
    Err.Raise Err.Number, "Error al ReLiberar las garantías de los créditos de difernte monedas", Err.Description
End Sub
'END EJVG *******

Public Sub InsertaMovAprobacionER_Viaticos(ByVal pdFecha As Date, ByVal psMovNro As String, ByVal psOpeCod As String, ByVal psOpeCodITF As String, ByVal pnMoneda As String, ByVal psCtaCod As String, ByVal pnImporte As Double)
    Dim lsSql As String
    On Error GoTo ErrInsertaMovAprobacionER_Viaticos
        lsSql = "EXEC stp_ins_ERS0752017_MovAprobacionER_Viaticos '" & Format(pdFecha & " " & oConect.GetHoraServer, "mm/dd/yyyy hh:mm:ss") & "', '" & psMovNro & "', '" & psOpeCod & "','" & psOpeCodITF & "', '" & pnMoneda & "','" & psCtaCod & "'," & pnImporte & ""
        oConect.Ejecutar lsSql
    Exit Sub
ErrInsertaMovAprobacionER_Viaticos:
    Err.Raise Err.Number, "Error al Insertar los Movimientos correspondientes a la Aprobación", Err.Description
End Sub
'END NAGL ERS075-2017 20171125 *******

Public Sub InsertaMovCtaNew_Update(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal psCtaCont As String, ByVal pnImporte As Double)
    Dim lsSql As String
    On Error GoTo ErrInsertaMovCtaNew
        lsSql = "EXEC stp_ins_MovCtaNew " & pnMovNro & ", " & pnMovItem & ",'" & psCtaCont & "', " & pnImporte & ""
        oConect.Ejecutar lsSql
    Exit Sub
ErrInsertaMovCtaNew:
    Err.Raise Err.Number, "Error al Insertar los Movimientos correspondientes a la Cuenta Contable", Err.Description
End Sub
'END NAGL ERS075-2017 20171128 *******

Public Sub InsertaMovMeNew(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal psCtaCont As String, ByVal pnImporte As Double)
    Dim lsSql As String
    On Error GoTo ErrInsertaMovMeNew
        lsSql = "EXEC stp_ins_MovMENew " & pnMovNro & ", " & pnMovItem & ",'" & psCtaCont & "', " & pnImporte & ""
        oConect.Ejecutar lsSql
    Exit Sub
ErrInsertaMovMeNew:
    Err.Raise Err.Number, "Error al Insertar los Movimientos correspondientes a la Cuenta Contable en ME", Err.Description
End Sub
'END NAGL ERS075-2017 20171128 *******

'CTI2 ADD 20190502 Mejorar proceso de pago de proveedores

Public Function GetDatosCuentaAho(ByVal sCuenta As String) As Recordset
Dim sSQL As String
Dim rsCta As Recordset

sSQL = "Select P.cCtaCod, P.nTasaInteres, P.nSaldo, P.nPrdEstado, P.nTransacc, " _
        & "C.nSaldoDisp, C.nPersoneria, C.nFirmas, C.nIntAcum, C.dUltCierre, C.dApertura, " _
        & "C.nPrdCtaTpo, C.nPrdTasaInteres, A.bOrdPag, A.dUltContacto, T.cConsDescripcion cEstado, " _
        & "T1.cConsDescripcion cTipoCuenta, T2.cConsDescripcion cTipoTasa, A.bInactiva, A.bInmovilizada FROM Producto P " _
        & "INNER JOIN Captaciones C INNER JOIN CaptacAhorros A ON C.cCtaCod = A.cCtaCod ON " _
        & "P.cCtaCod = C.cCtaCod INNER JOIN Constante T ON P.nPrdEstado  = T.nConsValor " _
        & "INNER JOIN " & sDBComunes & "Constante T1 ON C.nPrdCtaTpo = T1.nConsValor INNER JOIN " & sDBComunes & "Constante T2 " _
        & "ON C.nPrdTasaInteres = T2.nConsValor WHERE P.cCtaCod = '" & sCuenta & "' AND " _
        & "T.nConsCod = " & gCaptacEstado & " AND T1.nConsCod = " & gProductoCuentaTipo & " " _
        & "AND T2.nConsCod = " & gCaptacTipoTasa

Set rsCta = oConect.CargaRecordSet(sSQL)
Set GetDatosCuentaAho = rsCta
Set rsCta = Nothing
End Function
Public Function GetInteres(ByVal nCapital As Double, ByVal nTasa As Double, _
            ByVal nPlazo As Long, Optional nTipoInteres As TipoCalculoInteres = TpoCalcIntSimple) As Double
If nTipoInteres = TpoCalcIntSimple Then
    GetInteres = Round((nTasa / 36000) * nPlazo * nCapital, 2)
ElseIf nTipoInteres = TpoCalcIntCompuesto Then
    GetInteres = Round((((nTasa / 36000) + 1) ^ nPlazo - 1) * nCapital, 2)
End If
End Function
Public Sub AgregaMov(ByVal sMovNro As String, ByVal nOperacion As CaptacOperacion, _
        ByVal sGlosa As String, Optional nMovEstado As MovEstado = gMovEstContabMovContable, _
        Optional nMovFlag As MovFlag = gMovFlagVigente)
sSQL = "INSERT Mov (cMovNro,cOpeCod,cMovDesc,nMovEstado,nMovFlag) " _
    & "VALUES ('" & sMovNro & "','" & nOperacion & "','" & sGlosa & "'," & nMovEstado & "," & nMovFlag & ")"
oConect.Ejecutar sSQL
End Sub
Public Sub ActualizadUltContactoAho(ByVal sCuenta As String, ByVal dFecha As Date)
    sSQL = "Update CaptacAhorros Set dUltContacto = '" & Format(dFecha, gsFormatoFecha) & "'  WHERE cCtaCod = '" & sCuenta & "'"
    oConect.Ejecutar sSQL
End Sub
Public Function LeeConstSistema(ByVal psConstSistCod As ConstSistemas) As String
Dim rsVar As Recordset
Dim sSQL As String

sSQL = "Select nConsSisDesc, nConsSisValor From ConstSistema where nConsSisCod =" & psConstSistCod & ""
Set rsVar = New Recordset
Set rsVar = oConect.CargaRecordSet(sSQL)
LeeConstSistema = ""
If Not rsVar.EOF And Not rsVar.BOF Then
    LeeConstSistema = rsVar("nConsSisValor")
End If
rsVar.Close
Set rsVar = Nothing
End Function
Public Sub ActualizaInmovilizadaAct(ByVal sCuenta As String)
    sSQL = "Update CaptacAhorros Set bInactiva = 0,  bInmovilizada = 0 WHERE cCtaCod = '" & sCuenta & "'"
    oConect.Ejecutar sSQL
End Sub
Public Sub AgregaMovTipoCambio(ByVal nMovNro As Long, ByVal pnTipoCambio As Currency)
sSQL = " INSERT MovTpoCambio (nMovNro, nMovTpoCambio) " _
     & " VALUES (" & nMovNro & "," & pnTipoCambio & ")"
oConect.Ejecutar sSQL
End Sub
Public Sub AgregaMovCMAC(ByVal nMovNro As Long, ByVal sPersCod As String, _
        ByVal nMoneda As Moneda, ByVal nOperacion As CaptacOperacion, _
        ByVal nMonto As Double, Optional sDocumento As String = "NULL", _
        Optional sCuenta As String = "NULL")
        
Dim sIFTipo As String
Dim sCta As String, sDoc As String
sCta = "NULL"
sDoc = "NULL"
sIFTipo = Format$(gTpoIFCmac, "00")
If sCuenta <> "NULL" Then sCta = "'" & sCuenta & "'"
If sDocumento <> "NULL" Then sDoc = "'" & sDocumento & "'"
sSQL = "INSERT MovCMAC (nMovNro,cPersCod,cIFTpo,nMoneda,cCtaIF,cDocumento,cOpeCod,nMonto) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "','" & sIFTipo & "'," & nMoneda & "," & sCta & "," & sDoc & ",'" & Trim(nOperacion) & "'," & Trim(nMonto) & ")"
oConect.Ejecutar sSQL

End Sub
Public Sub ActualizaInactivaAct(ByVal sCuenta As String)
    sSQL = "Update CaptacAhorros Set bInactiva = 0 WHERE cCtaCod = '" & sCuenta & "'"
    oConect.Ejecutar sSQL
End Sub
Public Sub AgregaMovRef(ByVal nMovNro As Long, ByVal nMovRef As Long)
sSQL = "INSERT MovRef (nMovNro,nMovNroRef) " _
    & "VALUES (" & nMovNro & "," & nMovRef & ")"
oConect.Ejecutar sSQL
End Sub
Public Sub AgregaMovLavDinero(ByVal nMovNro As Long, ByVal sPersCod As String)
sSQL = "INSERT MovLavDinero (nMovNro,cPersCod) " _
    & "VALUES (" & nMovNro & ",'" & sPersCod & "')"
oConect.Ejecutar sSQL
End Sub
Public Function GetNombreTitulares(ByVal sCuenta As String) As String
Dim sSQL As String
Dim rsTit As ADODB.Recordset
Dim sCadena As String
Dim sCadAux As String
Dim sCadAux1 As String
Dim sCadAux2 As String

Dim nPosApe As Long
Dim nPosNom  As Long
Dim sCadAux3 As String
Dim nContador As Long

Dim lstOpe As String * 30

sSQL = "SELECT P.cPersNombre FROM " & sDBPersona & "Persona P INNER JOIN ProductoPersona PP " _
    & "INNER JOIN Producto D ON PP.cCtaCod = D.cCtaCod ON P.cPersCod = PP.cPersCod " _
    & "WHERE D.cCtaCod = '" & sCuenta & "' AND PP.nPrdPersRelac = " & gCapRelPersTitular
Set rsTit = oConect.CargaRecordSet(sSQL)
If rsTit.EOF And rsTit.BOF Then
    GetNombreTitulares = ""
Else
    nContador = 1
    lstOpe = PstaNombre(rsTit("cPersNombre"), False)
    sCadAux = Trim(lstOpe)
    sCadAux3 = rsTit("cPersNombre")
    rsTit.MoveNext
    If Not rsTit.EOF Then sCadAux = ""
    Do While Not rsTit.EOF And nContador <> 4
        If sCadAux <> "" Then
            sCadAux3 = rsTit("cPersNombre")
        End If
        nPosApe = InStr(1, sCadAux3, "/", vbTextCompare)
        nPosNom = InStr(1, sCadAux3, ",", vbTextCompare)
        sCadAux1 = Left(Left(sCadAux3, nPosApe), 18)
        sCadAux2 = Left(Mid(sCadAux3, nPosNom + 1), 18)
        nPosNom = InStr(1, Trim(sCadAux2), " ", vbTextCompare)
        If nPosNom <> 0 Then
            sCadAux2 = Mid(sCadAux2, 1, nPosNom)
        End If
        sCadAux2 = Trim(sCadAux2)
        If sCadAux <> "" Then
            If nContador Mod 2 = 0 Then
                sCadAux = sCadAux & "*" & Left(sCadAux1 & sCadAux2, 20) & Space(20 - Len(Left(sCadAux1 & sCadAux2, 20)))
            Else
                sCadAux = sCadAux & "," & Left(sCadAux1 & sCadAux2, 20) & Space(20 - Len(Left(sCadAux1 & sCadAux2, 20)))
            End If
            nContador = nContador + 1
            rsTit.MoveNext
        Else
            sCadAux = Left(sCadAux1 & sCadAux2, 20) & Space(23 - Len(Left(sCadAux1 & sCadAux2, 20)))
        End If
    Loop
    GetNombreTitulares = sCadAux
End If
rsTit.Close
Set rsTit = Nothing
End Function
Public Function ImprimeBoleta(ByVal psTit As String, ByVal psText As String, ByVal psCodOpe As String, ByVal pnMonto As String, _
        ByVal psCliente As String, ByVal psCodCta As String, ByVal psNumDoc As String, ByVal pnSaldo As Double, _
        ByVal pnInteresA As String, NomDoc As String, ByVal pnNumExt As Long, ByVal pnSaldoC As Double, _
        Optional pbOpSaldoC As Boolean = True, Optional pbSaldoInt As Boolean = True, Optional psNumDias As String = "----", _
        Optional psNomAgeRem As String = "", Optional psCodUsuRem As String = "", Optional pbCuenta As Boolean = False, _
        Optional psComCmac As String = "XXX", Optional psLin3 As String = "XXX", Optional psTexto As String = "XXX", _
        Optional pdFecsis As Date, Optional sNomAge As String = "", Optional sCodUser As String = "", _
        Optional sLPT As String = "") As String

Dim nFicSal As Integer
Dim sFecha As String
Dim sHora As String
Dim sSep As Integer
Dim sIni As Integer
Dim sMonto As String
Dim sSDisp As String
Dim sIntAcum As String
Dim sMax As Integer
Dim sAux As Integer
Dim lsNegritaOn As String
Dim lsNegritaOff As String
Dim lsNroExt As String
Dim lnTope As Integer
Dim lsSaldoC As String
Dim lsCadArg As String
Dim lsInteres As String
Dim sCad As String
Dim lnCliAux As Long
Dim lsCliAux1 As String
Dim lsCliAux2 As String

Dim lnChq As Long
Dim lsChqAux1 As String
Dim lsChqAux2 As String
Dim lsNomAge As String

Dim lnNumLinCmac As Integer

Dim lsMensaje As String * 39

ETIQ:

On Error GoTo ERROR

lnTope = 0 '6 'Tope de lineas en Boleta

lsNegritaOn = Chr$(27) + Chr$(71)
lsNegritaOff = Chr$(27) + Chr$(72)

lsNroExt = Str(pnNumExt)

nFicSal = FreeFile
Open sLPT For Output As nFicSal

Print #nFicSal, Chr$(27) & Chr$(64);

Print #nFicSal, Chr$(27) & Chr$(50);   'espaciamiento lineas 1/6 pulg.
Print #nFicSal, Chr$(27) & Chr$(67) & Chr$(22);  'Longitud de página a 22 líneas'
Print #nFicSal, Chr$(27) & Chr$(77);   'Tamaño 10 cpi
Print #nFicSal, Chr$(27) + Chr$(107) + Chr$(0);     'Tipo de Letra Sans Serif
Print #nFicSal, Chr$(27) + Chr$(18) ' cancela condensada
Print #nFicSal, Chr$(27) + Chr$(72) ' desactiva negrita

sSep = 15
sIni = 1
sMax = 33
sAux = 5

sFecha = Format$(pdFecsis, "dd/mm/yyyy")
sHora = Format$(Time, "hh:mm:ss")
sMonto = Format$(pnMonto, "#,##0.00")
sSDisp = Format$(pnSaldo, "#,##0.00")
lsSaldoC = Format$(pnSaldoC, "#,##0.00")

lsNomAge = sNomAge

'Print #nFicSal, Chr$(10);
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni); "CMACT - AHORRO"; Space(19 + sSep + sAux); "CMACT - AHORRO"

If Mid(psCodCta, 9, 1) = 1 Then
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt; Space(sSep); Trim(sNomAge) & "-SOLES"; Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-SOLES")) + lsNroExt;
Else
    Print #nFicSal, Tab(sIni); Trim(sNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) & lsNroExt; Space(sSep); Trim(sNomAge) & "-DOLARES"; Space(sAux + sMax - Len(Trim(sNomAge)) - Len(lsNroExt) - Len("-DOLARES")) + lsNroExt;
End If

If psNomAgeRem = "" Then
    Print #nFicSal, ""
Else
    Print #nFicSal, Tab(sIni); "Ag.Rem: " & Trim(psNomAgeRem); Space(sAux + sMax + sSep - Len("Ag.Rem:") - Len(Trim(psNomAgeRem))); "Ag.Rem: " & Trim(psNomAgeRem)
End If

If psComCmac = "XXX" Then
    If psLin3 = "XXX" Then
        Print #nFicSal, lsNegritaOff; 'Desactiva Negrita
    Else
        Print #nFicSal, Tab(sIni); psLin3 & Space(sAux + sSep + sMax - Len(psLin3)) & psLin3 & lsNegritaOff;  'Desactiva Negrita
        lnNumLinCmac = 1
    End If
    lnNumLinCmac = 0
Else
    Print #nFicSal, Tab(sIni); "NroDocCmac:" & psComCmac & Space(sAux + sSep + sMax - Len("NroDocCmac:" & psComCmac)) & "NroDocCmac:" & psComCmac & lsNegritaOff;  'Desactiva Negrita
    lnNumLinCmac = 1
End If

Print #nFicSal, Tab(sIni); "Fecha:" & sFecha; Space(10); "Hora:" & sHora; Space(sAux + sSep - 6); "Fecha:" & sFecha; Space(10); "Hora:" & sHora

lnCliAux = InStr(1, psCliente, "*", vbTextCompare)

If lnCliAux = 0 Then
    If sAux + sMax - Len(psCliente) < 0 Then psCliente = Mid(psCliente, 1, sMax + sAux)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(psCliente); Space(sAux + sMax + sSep - Len(psCliente)); ImpreCarEsp(psCliente)
Else
    lsCliAux1 = (Mid(psCliente, 1, lnCliAux - 1))
    lsCliAux2 = (Mid(psCliente, lnCliAux + 1))
    
    If sMax - Len(lsCliAux1) < 2 Then lsCliAux1 = Mid(lsCliAux1, 1, sMax + sAux)
    If sMax - Len(lsCliAux2) < 2 Then lsCliAux2 = Mid(lsCliAux2, 1, sMax + sAux)
    
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux1); Space(sAux + sMax + sSep - Len(lsCliAux1)); ImpreCarEsp(lsCliAux1)
    Print #nFicSal, Tab(sIni); ImpreCarEsp(lsCliAux2); Space(sAux + sMax + sSep - Len(lsCliAux2)); ImpreCarEsp(lsCliAux2)
    
    lnCliAux = 1
End If

If pbSaldoInt Or pbCuenta Then
    Print #nFicSal, Tab(sIni); "Cuenta:" & psCodCta; Space(8 + sSep + sAux); "Cuenta:" & psCodCta
Else
    Print #nFicSal, ""
End If

psTit = Trim(psTit)
psTit = CentrarCadena(psTit, 28)
Print #nFicSal, lsNegritaOn; 'Activa Negrita
Print #nFicSal, Tab(sIni + 1); "-----" & psTit & "-----"; Space(-1 + sSep); "-----" & psTit & "-----"


lnChq = InStr(1, psText, "*", vbTextCompare)

If psTexto = "XXX" Then
    If lnChq = 0 Then
        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(Mid(psText, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(Mid(psText, 1, 28))); Space(sMax + 6 - Len(Trim(Mid(psText, 1, 28))) - Len(sMonto)); sMonto
        Print #nFicSal, ""
    Else
        lsChqAux1 = (Mid(psText, 1, lnChq - 1))
        lsChqAux2 = (Mid(psText, lnChq + 1))
        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(lsChqAux1)); Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)); sMonto; Space(-1 + sSep); ImpreCarEsp(Trim(lsChqAux1)); Space(sMax + 6 - Len(Trim(lsChqAux1)) - Len(sMonto)); sMonto
        Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(lsChqAux2)); Space(sMax + 6 - Len(Trim(lsChqAux2))); Space(-1 + sSep); ImpreCarEsp(Trim(lsChqAux2)); Space(sMax + 6 - Len(Trim(lsChqAux2)))
    End If
Else
    Print #nFicSal, Tab(sIni); ImpreCarEsp(Trim(psTexto)); Space(sAux + sSep + sMax - Len(Trim(psTexto))); ImpreCarEsp(Trim(psTexto))
    Print #nFicSal, ""
End If


Print #nFicSal, lsNegritaOff; 'Desactiva Negrita

If pbSaldoInt Then
    If MsgBox("Desea Imprimir los Saldos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Print #nFicSal, Tab(sIni); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp; Space(-1 + sSep); "Saldo Disponible"; Space(sMax + 6 - Len("Saldo Disponible") - Len(sSDisp)); sSDisp
        If pbOpSaldoC Then
            Print #nFicSal, Tab(sIni); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC; Space(-1 + sSep); "Saldo Contable"; Space(sMax + 6 - Len("Saldo Contable") - Len(lsSaldoC)); lsSaldoC
        Else
            Print #nFicSal, ""
        End If
    Else
        Print #nFicSal, ""
        Print #nFicSal, ""
        pbSaldoInt = False
    End If
Else
    Print #nFicSal, ""
End If

lsInteres = pnInteresA

If pbSaldoInt Then
    If lsInteres <> "No Valido" Then
        lsInteres = Format(lsInteres, "#,##0.00")
        Print #nFicSal, Tab(sIni); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres; Space(-1 + sSep); "Interes del Mes"; Space(sMax + 6 - Len("Interes del Mes") - Len(lsInteres)); lsInteres
    End If
Else
    Print #nFicSal, ""
End If

If Not psNumDoc = "" Then
    Print #nFicSal, Tab(sIni); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc; Space(-1 + sSep); NomDoc; Space(sMax + 6 - Len(NomDoc) - Len(psNumDoc)); psNumDoc
Else
    Print #nFicSal, ""
End If

If Not psNumDias = "----" Then
    Print #nFicSal, Tab(sIni); "Nro Dias Interes"; Space(sMax + 6 - Len("Nro Dias Interes") - Len(psNumDias)); psNumDias; Space(-1 + sSep); "Nro Dias Interes"; Space(sMax + 6 - Len(psNumDias) - Len("Nro Dias Interes")); psNumDias
    lnTope = 4 - lnCliAux
Else
    lnTope = 3 - lnCliAux
End If

Print #nFicSal, Tab(sIni); "---------------------------------------"; Space(-1 + sSep); "---------------------------------------"
If psCodUsuRem = "" Then
    Print #nFicSal, Tab(sIni); ImpreCarEsp(sCodUser); Space(29 + sSep + sAux); ImpreCarEsp(sCodUser)
Else
    Print #nFicSal, Tab(sIni); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(sCodUser) & "/"; ImpreCarEsp(psCodUsuRem); Space(sSep); "Loc/Rem"; Space(sMax + sAux - Len("Loc/Rem") - 1 - 8); ImpreCarEsp(sCodUser) & "/"; ImpreCarEsp(psCodUsuRem)
End If
Dim clsGen As DGeneral
Set clsGen = New DGeneral
lsMensaje = clsGen.GetMensajeBoletas(psCodCta)
Set clsGen = Nothing
Print #nFicSal, Tab(sIni); lsNegritaOn & ImpreCarEsp(lsMensaje); Space(-1 + sSep); ImpreCarEsp(lsMensaje); lsNegritaOff

lnNumLinCmac = lnNumLinCmac + 1
 
'Print #nFicSal, Chr$(12)
For sAux = 1 To (lnTope - lnNumLinCmac)
    Print #nFicSal, ""
Next sAux
Close nFicSal
Exit Function
ERROR:
    Close nFicSal
    If MsgBox("Comprueba la conexion de su impresora, " + Err.Description & " Desea Reintentar?", vbCritical + vbYesNo, "Aviso") = vbYes Then
        GoTo ETIQ
    End If
End Function
Public Sub ActualizaMovPendientesRend2(ByVal nMovNro As Long, ByVal pnMonto As Currency)
    sSQL = " UPDATE MovPendientesRend " _
         & " Set nSaldo = nSaldo - " & pnMonto & "" _
         & " Where nMovNro = " & nMovNro & ""
    oConect.Ejecutar sSQL
End Sub
Public Function CapAbonoCuentaAho(ByVal sCuenta As String, ByVal nMonto As Double, _
                                  ByVal nOperacion As CaptacOperacion, ByVal sMovNro As String, _
                                  ByVal sGlosa As String, Optional nTipoDoc As TpoDoc = TpoDocCheque, _
                                  Optional sNroDoc As String = "", Optional sCodIF As String = "", _
                                  Optional dFechaValor As Date, Optional bActivaCta As Boolean = True, _
                                  Optional bNumExtracto As Boolean = True, Optional sCodCMAC As String = "", _
                                  Optional sNomCMAC As String = "", Optional sNomAge As String = "", _
                                  Optional sLPT As String = "", Optional sPersLavDinero As String = "", _
                                  Optional btransaccion As Boolean = False, Optional bImprimeBoleta As Boolean = True, _
                                  Optional pnMovNroTransfer As Long = -1, Optional pnMonedaTrans As Moneda = gMonedaNacional, _
                                  Optional pCon As ADODB.Connection = Nothing, Optional pbImprimeBoleta As Boolean = True, _
                                  Optional pbITFAplica As Boolean = True, Optional pnITFValor As Double = 0) As Double
    
    Dim rsCta As Recordset
    Dim nEstado As CaptacEstado
    Dim nSaldoDisp As Double, nSaldoCnt As Double, nIntAcum As Double
    Dim nIntSaldo As Double, nIntGanado As Double, nTasa As Double
    Dim nSaldoInac As Double
    Dim dUltMov As Date
    Dim bInactiva As Boolean, bTrans As Boolean, bInamovilizada As Boolean
    Dim nNumExtracto As Long, nDiasTranscurridos As Long, nExtracto As Long
    Dim sMsgOpe As String, sCodUser As String
    Dim nMovNro As Long
    Dim dFecSis As Date
    Dim lnMontoCV As Currency
    
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
    bInamovilizada = IIf(rsCta("bInmovilizada") = 1, True, False)
    nExtracto = rsCta("nTransacc") + 1
    rsCta.Close
    Set rsCta = Nothing

    'Inicia la transaccion
        
    On Error GoTo ErrGraba

    'If Not btransaccion Then clsCap.dbCmact.BeginTrans
    'If Not pCon Is Nothing Then clsCap.SetConexion pCon
    'bTrans = True
    
    If sGlosa = "" Then sGlosa = "Abono Cuenta = " & sCuenta
    If bActivaCta Then
        If bInactiva Then
            nSaldoInac = nSaldoCnt
        End If
    End If
    'Calcula intereses
    nDiasTranscurridos = DateDiff("d", dUltMov, dFecSis) - 1
    nIntGanado = GetInteres(nSaldoDisp, nTasa, nDiasTranscurridos, TpoCalcIntSimple)
    dUltMov = DateAdd("d", -1, dFecSis)
    nSaldoCnt = nSaldoCnt + nMonto
    AgregaMov sMovNro, nOperacion, sGlosa
    nMovNro = GetnMovNro(sMovNro)
                 
    If sNroDoc = "" Then ' Si la operacion es en efectivo
        ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
        nSaldoDisp = nSaldoDisp + nMonto
        ActualizadUltContactoAho sCuenta, dFecSis
    Else
        If nTipoDoc = TpoDocCheque Then 'Si el abono es con cheque
            nSaldoDisp = nSaldoDisp + 0
            nIntGanado = 0
            ActualizadUltContactoAho sCuenta, dFecSis
        ElseIf nTipoDoc = TpoDocNotaAbono Then 'Si el abono es con nota de abono
            ActualizaSaldoAnteriorAho sCuenta, nSaldoDisp
            nSaldoDisp = nSaldoDisp + nMonto
        End If
    End If
    AgregaMovCap nMovNro, nOperacion, sCuenta, nMonto, nSaldoDisp, nSaldoCnt
    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcCapital, nMonto

    If sNroDoc = "" Then ' Si la operacion es en efectivo
        ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
        If nOperacion = gAhoDepEfec Then
            sMsgOpe = "Depósito Efectivo"
        ElseIf nOperacion = gAhoDepTransf Then
            sMsgOpe = "Depósito Transferencia"
        ElseIf nOperacion = gAhoDepDevFondoGar Then
            sMsgOpe = "Depósito Dev Fondo Gar"
        End If
    Else
        If nTipoDoc = TpoDocCheque Then 'Si el abono es con cheque
            ActualizaAbonoCaptacion sCuenta, nMonto, 0, nIntGanado, dUltMov, sMovNro, True
            AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, sCodIF, sMovNro, nMovNro, nMonto
            sMsgOpe = "Depósito Cheque"
        ElseIf nTipoDoc = TpoDocNotaAbono Then 'Si el abono es con Nota Abono
            ActualizaAbonoCaptacion sCuenta, nMonto, nMonto, nIntGanado, dUltMov, sMovNro, True
            AgregaCuentaDocumento sCuenta, nTipoDoc, sNroDoc, "", sMovNro, nMovNro, nMonto
            sMsgOpe = "Depósito N.Abono"
        End If
    End If

    If pbITFAplica And pnITFValor > 0 Then
        ActualizaAbonoCaptacion sCuenta, pnITFValor * -1, pnITFValor * -1, nIntGanado, dUltMov, sMovNro, True
        If pbITFAplica Then
            AgregaMovCap nMovNro, gITFCobroCargo, sCuenta, pnITFValor, nSaldoDisp - pnITFValor, nSaldoCnt - pnITFValor
            AgregaMovCapDet nMovNro, gITFCobroCargo, sCuenta, gConcITFCliente, pnITFValor
        End If
    End If

    If bInamovilizada Then
        If LeeConstSistema(gConstSistBitCtrCtasInmovil) = 1 Then
            AgregaMovCapDet nMovNro, gAhoEstInmovAct, sCuenta, gConcCapital, nSaldoInac
            ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
            ActualizaInmovilizadaAct sCuenta
        End If
    End If

    If bInactiva And Not bInamovilizada Then
        AgregaMovCapDet nMovNro, gAhoEstInacAct, sCuenta, gConcCapital, nSaldoInac
        ActualizaEstadoCuenta sCuenta, gCapEstActiva, dFecSis, sMovNro
        ActualizaInactivaAct sCuenta
    End If

    If pnMovNroTransfer <> -1 Then
        AgregaMovRef nMovNro, pnMovNroTransfer
        If pnMonedaTrans <> Mid(sCuenta, 9, 1) Then
            GetTipCambio dFecSis
            
            If Mid(sCuenta, 9, 1) = gMonedaNacional Then
                ActualizaMovPendientesRend2 pnMovNroTransfer, Round(nMonto / gnTipCambioC, 2)
                lnMontoCV = Round((nMonto / gnTipCambioC) * gnTipCambio, 2)
                AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeCompra, lnMontoCV
                AgregaMovTipoCambio nMovNro, gnTipCambioC
                If gnTipCambioC < gnTipCambio Then
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - nMonto)
                Else
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - nMonto)
                End If
            Else
                ActualizaMovPendientesRend2 pnMovNroTransfer, Round(nMonto * gnTipCambioV, 2)
                lnMontoCV = Round(nMonto * gnTipCambio, 2)
                AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeVenta, lnMontoCV
                AgregaMovTipoCambio nMovNro, gnTipCambioV
                If gnTipCambioV > gnTipCambio Then
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedeUtilidad, Abs(lnMontoCV - Round(nMonto * gnTipCambioV, 2))
                Else
                    AgregaMovCapDet nMovNro, nOperacion, sCuenta, gConcOpedePerdida, Abs(lnMontoCV - Round(nMonto * gnTipCambioV, 2))
                End If
            End If
        Else
            ActualizaMovPendientesRend2 pnMovNroTransfer, nMonto
        End If
    End If

    If sCodCMAC <> "" Then
        AgregaMovCMAC nMovNro, sCodCMAC, CLng(Mid(sCuenta, 9, 1)), nOperacion, nMonto
    End If
    UltimaActualizacionCuenta sCuenta, sMovNro
    If sPersLavDinero <> "" Then
        AgregaMovLavDinero nMovNro, sPersLavDinero
    End If
    
    'If Not btransaccion Then clsCap.dbCmact.CommitTrans
    'bTrans = False
    
    CapAbonoCuentaAho = nSaldoCnt

    If Not bImprimeBoleta Then Exit Function
    Dim sTipDep As String, sCodOpe As String
    Dim sModDep As String, sTipApe As String
    Dim sNomTit As String
    sTipDep = IIf(Mid(sCuenta, 9, 1) = Moneda.gMonedaNacional, "SOLES", "DOLARES")
    sCodOpe = Trim(nOperacion)
    sModDep = sMsgOpe
    sTipApe = "DEPOSITO AHORROS"
    Dim clsTit As NCapMantenimiento
    Set clsTit = New NCapMantenimiento
    sNomTit = ImpreCarEsp(GetNombreTitulares(sCuenta))
    
    If pbImprimeBoleta Then
        Do
            If sNroDoc <> "" Then
                Select Case nTipoDoc
                    Case TpoDocCheque
                        If sCodCMAC <> "" Then
                            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLPT
                        Else
                            ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, Format$(dFechaValor, "dd/mm/yyyy"), nSaldoDisp, nIntGanado, "Fecha Valor", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLPT
                        End If
                    Case TpoDocNotaAbono
                        ImprimeBoleta sTipApe, ImpreCarEsp(sMsgOpe) & " No. " & sNroDoc, Trim(nOperacion), Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", nExtracto, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLPT
                End Select
            Else
                If sCodCMAC <> "" Then
                    sTipApe = "DEPOSITO CMAC AHORROS"
                    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , sNomCMAC, , dFecSis, sNomAge, sCodUser, sLPT
                Else
                    ImprimeBoleta sTipApe, ImpreCarEsp(sModDep), sCodOpe, Trim(nMonto), sNomTit, sCuenta, "", nSaldoDisp, nIntGanado, "", 1, nSaldoCnt, True, , , , , , , , , dFecSis, sNomAge, sCodUser, sLPT
                End If
            End If
        Loop Until MsgBox("Desea reimprimir ?? ", vbQuestion + vbYesNo, "Aviso") = vbNo
    End If
    Exit Function
ErrGraba:
        Err.Raise Err.Number, "", Err.Description
        CapAbonoCuentaAho = 0
    End Function
'END CTI2

Public Sub EliminaCuentasProAsiento(ByVal pnMovNro As Long, ByVal psMovNro As String, ByVal psOpeCod As String)
Dim sql As String
On Error GoTo EliminaCuentasProAsientoErr
    sql = "exec stp_del_CuentasProAsientos " & pnMovNro & ", '" & psMovNro & "','" & psOpeCod & "'"
    oConect.Ejecutar sql
Exit Sub
EliminaCuentasProAsientoErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaCuentasProAsiento Method")
End Sub 'NAGL 202008 Según ACTA N°063-2020

Public Function PermiteModAsiContxCodOperación(ByVal psOpeCod As String) As Boolean
Dim sql As String
Dim oRs As New ADODB.Recordset
On Error GoTo PermiteModAsiContxCodOperaciónErr
    sql = "exec stp_sel_PermiteModAsiContxCodOpe '" & psOpeCod & "'"
    Set oRs = oConect.CargaRecordSet(sql)
    If Not oRs.BOF And Not oRs.EOF Then
        PermiteModAsiContxCodOperación = oRs!pbValida
    Else
        PermiteModAsiContxCodOperación = True
    End If
Exit Function
PermiteModAsiContxCodOperaciónErr:
    Call RaiseError(MyUnhandledError, "DMov:PermiteModAsiContxCodOperación Method")
End Function 'NAGL 202008 Según ACTA N°063-2020

