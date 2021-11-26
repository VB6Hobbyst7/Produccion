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
Dim vsConexion As String
Dim vsServerAdmin As String
Dim vsServerComunes As String
Attribute vsServerComunes.VB_VarDescription = "Se guarda la caden aen la cual se guardan el nombre del Servidor y la base de datos que posee la tabla de comunes."
Dim vsServerImagenes As Variant
Dim vsServerNegocio As String
Dim vsServerPersonas As String
Dim vsFormatoFecha As String
Dim oConect As DConecta
Dim sSQL As String
Dim lbBitControlCuentasUltNivel As Boolean
Dim lbBitControlAsientoCuadrado As Boolean
Dim lnMovNroAsientoCuadrado As Long

Public Sub RollBackTrans()
    oConect.RollBackTrans
End Sub

Public Sub CommitTrans()
    oConect.CommitTrans
End Sub

Public Sub BeginTrans()
    oConect.BeginTrans
End Sub

Public Sub Inicio(ByVal psFormatoFecha As String)
Attribute Inicio.VB_Description = "Utilizado para inicializar las variables utilizadas en la clase de datos"
    vsFormatoFecha = psFormatoFecha
End Sub

Public Function GetConexion() As ADODB.Connection
    Set GetConexion = oConect.ConexionActiva
End Function

Public Function BuscarMov(ByVal psMovNro As String, Optional psFiltro As String = "") As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    Dim oConect As New DConecta
    On Error GoTo BuscarMovErr
    BuscarMov = False
    
    If oConect.AbreConexion = False Then
        Exit Function
    End If
    sql = "SELECT cMovnro from Mov WHere cMovNro LIKE '" & psMovNro & "%'" & IIf(psFiltro <> "", " AND ", "") & psFiltro
    Set rs = oConect.CargaRecordSet(sql)
    BuscarMov = Not rs.EOF
    rs.Close
    Set rs = Nothing
    oConect.CierraConexion
    Set oConect = Nothing
    Exit Function
BuscarMovErr:
    Call RaiseError(MyUnhandledError, "DMov:BuscarMov Method")
End Function


'##ModelId=3A836CB1037A
Public Sub BuscarMovCta()
    On Error GoTo BuscarMovCtaErr

    'your code goes here...

    Exit Sub
BuscarMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:BuscarMovCta Method")
End Sub

'##ModelId=3A837105004E
Public Sub BuscarMovObj()
    On Error GoTo BuscarMovObjErr

    'your code goes here...

    Exit Sub
BuscarMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:BuscarMovObj Method")
End Sub

'##ModelId=3A837109002E
Public Sub BuscarMovDoc()
    On Error GoTo BuscarMovDocErr

    'your code goes here...

    Exit Sub
BuscarMovDocErr:
    Call RaiseError(MyUnhandledError, "DMov:BuscarMovDoc Method")
End Sub

'##ModelId=3A83710C0251
Public Sub BuscarMovCabObj()
    On Error GoTo BuscarMovCabObjErr

    'your code goes here...

    Exit Sub
BuscarMovCabObjErr:
    Call RaiseError(MyUnhandledError, "DMov:BuscarMovCabObj Method")
End Sub
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovContErr Method")
End Function

'##ModelId=3A848EF203A9
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
Public Sub ActualizaMov(ByVal pnMovNro As Long, Optional ByVal psMovDesc As String = "", Optional ByVal psMovEstado As MovEstado = gMovEstContabMovContable, Optional ByVal psMovFlag As MovFlag = gMovFlagVigente)
    On Error GoTo ActualizaMovErr
    Dim sql As String
    Dim lsActualiza As String
    lsActualiza = ""
    If psMovDesc <> "" Then
        lsActualiza = lsActualiza & " cMovDesc='" & Replace(psMovDesc, "'", "''") & "',"
    End If
    lsActualiza = lsActualiza & " nMovEstado=" & psMovEstado & ","
    lsActualiza = lsActualiza & " nMovFlag=" & psMovFlag & ","
    lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    If lsActualiza = "" Then Exit Sub
    
    sql = " Update Mov set " & lsActualiza & " Where nMovNro =" & pnMovNro & ""
    oConect.Ejecutar sql
    Exit Sub
ActualizaMovErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMov Method")
End Sub
Public Sub EliminaMov(ByVal pnMovNro As String)
Dim sql As String
   On Error GoTo EliminaMovErr
    sql = "UPDATE Mov SET nMovFlag = '" & gMovFlagEliminado & "' WHERE nMovNro = " & pnMovNro & ""
    oConect.Ejecutar sql
Exit Sub
EliminaMovErr:
    Call RaiseError(MyUnhandledError, "DMov:EliminaMov Method")
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
Public Function InsertaMovCta(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psCtaContCod As String, ByVal pnMovImporte As Currency) As Integer
    On Error GoTo InsertaMovCtaErr
    Dim rsCta As ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Dim sql As String
    InsertaMovCta = 1
    
    If lbBitControlCuentasUltNivel Then
        
        sql = "Select cCtaContCod from ctacont where cCtaContCod Like '" & psCtaContCod & "%' And cCtaContCod <> '" & psCtaContCod & "' And nCtaEstado = 1"
        Set rsCta = oConect.CargaRecordSet(sql)
        
        If rsCta.EOF And rsCta.BOF Then
            'EJVG20130701 ***
            'rsCta.Close
            sql = "Select cCtaContCod from CtaCont Where cCtaContCod = '" & psCtaContCod & "' And nCtaEstado = 1"
            Set rsCta = oConect.CargaRecordSet(sql)
            If rsCta.EOF And rsCta.BOF Then
            Else
            'END EJVG *******
            sql = " INSERT INTO MOVCTA(nMovNro, nMovItem, cCtaContCod, nMovImporte )" _
                & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
            oConect.Ejecutar sql
            InsertaMovCta = 0
            rsCta.Close
            Exit Function
            End If
        End If
        rsCta.Close
        Err.Number = 9999
        Err.Description = "Se desea grabar una Cta Cont que no esta a Ult Nivel."
    Else
        sql = " INSERT INTO MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte )" _
            & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psCtaContCod & "'," & pnMovImporte & ") "
        
        oConect.Ejecutar sql
        InsertaMovCta = 0
        Exit Function
    End If
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function

'##ModelId=3A848F08038A
Public Function GetOpeCtaCta(ByVal psOpeCod As String, ByVal psCtaDebe As String, ByVal psCtaHaber As String, Optional psAlmacen As String = "") As String
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If psAlmacen = "" Then
        If psCtaDebe = "" Then
            sql = " Select cCtaContCodD From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                & " And cCtaContCodH like '" & psCtaHaber & "%'"
        Else
            sql = " Select cCtaContCodH From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                & " And cCtaContCodD like '" & psCtaDebe & "%'"
        End If
    Else
        If psCtaDebe = "" Then
            sql = " Select cCtaContCodD From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                & " And Replace(cCtaContCodH,'AG','" & psAlmacen & "') like '" & psCtaHaber & "%'"
        Else
            sql = " Select cCtaContCodH From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                & " And Replace(cCtaContCodD,'AG','" & psAlmacen & "') like '" & psCtaDebe & "%'"
        End If
    End If
    
    Set rs = oConect.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetOpeCtaCta = ""
    Else
        GetOpeCtaCta = rs.Fields(0)
    End If
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function

'##ModelId=3A848F08038A
Public Function GetCtaOpeBS(ByVal psOpeCod As String, ByVal psBSCod As String) As String
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select Top 1 cCtaContCod From CtaOpeBS" _
        & " Where '" & psBSCod & "' Like cBSCod + '%' And cOpeCod = '" & psOpeCod & "' Order By Len(cBSCod) Desc "
    Set rs = oConect.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetCtaOpeBS = ""
    Else
        GetCtaOpeBS = rs.Fields(0)
    End If
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function


'##ModelId=3A848F08038A
Public Function GetOpeCtaCtaOtro(ByVal psOpeCod As String, ByVal psCtaDebe As String, ByVal psCtaHaber As String, Optional pbCtaContDebe As Boolean, Optional psAgencia As String = "") As String
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If psAgencia = "" Then
        If psCtaDebe = "" Then
            If Not pbCtaContDebe Then
                sql = " Select cCtaContCodOtroD From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                    & " And cCtaContCodH like '" & psCtaHaber & "%'"
            Else
                sql = " Select cCtaContCodOtroH From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                    & " And cCtaContCodH like '" & psCtaHaber & "%'"
            End If
        Else
            If Not pbCtaContDebe Then
                sql = " Select cCtaContCodOtroD From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                    & " And cCtaContCodD like '" & psCtaDebe & "%'"
            Else
                sql = " Select cCtaContCodOtroH From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                    & " And cCtaContCodD like '" & psCtaDebe & "%'"
            End If
        End If
    Else
        If psCtaDebe = "" Then
            If Not pbCtaContDebe Then
                sql = " Select cCtaContCodOtroD From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                    & " And Replace(cCtaContCodH,'AG','" & psAgencia & "') like '" & psCtaHaber & "%'"
            Else
                sql = " Select cCtaContCodOtroH From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                    & " And Replace(cCtaContCodH,'AG','" & psAgencia & "') like '" & psCtaHaber & "%'"
            End If
        Else
            If Not pbCtaContDebe Then
                sql = " Select cCtaContCodOtroD From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                    & " And Replace(cCtaContCodD,'AG','" & psAgencia & "') like '" & psCtaDebe & "%'"
            Else
                sql = " Select cCtaContCodOtroH From OpeCtaCta Where cOpeCod = '" & psOpeCod & "'" _
                    & " And Replace(cCtaContCodD,'AG','" & psAgencia & "') like '" & psCtaDebe & "%'"
            End If
        End If
    End If
    
    Set rs = oConect.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetOpeCtaCtaOtro = ""
    Else
        GetOpeCtaCtaOtro = rs.Fields(0)
    End If
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function


'##ModelId=3A848F08038A
Public Function GetCtaAreaAge(ByVal psAreaCod As String, ByVal psAgenciaCod As String) As String
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    
    sql = " Select cSubCtaCod From AreaAgencia Where cAreaCod = '" & psAreaCod & "' And cAgeCod = '" & psAgenciaCod & "'"
    
    GetCtaAreaAge = oConect.CargaRecordSet(sql).Fields(0)
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function

'##ModelId=3A848F08038A
Public Function InsertaMovCotizacDet(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psDescrip As String, ByVal psItemPlazo As String) As Integer
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    InsertaMovCotizacDet = 1
    
    sql = " INSERT INTO MOVCOTIZACDET (nMovNro,nMovItem,cDescrip,dItemPlazo)" _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psDescrip & "','" & psItemPlazo & "')"
      
    oConect.Ejecutar sql
    InsertaMovCotizacDet = 0
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function

'##ModelId=3A848F08038A
Public Function InsertaMovCant(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovCant As Long) As Integer
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    InsertaMovCant = 1
    
    sql = " INSERT INTO MOVCANT (nMovNro,nMovItem,nMovCant)" _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovCant & ") "
      
    oConect.Ejecutar sql
    InsertaMovCant = 0
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function


'##ModelId=3A848F08038A
'Public Function InsertaMovCotizac(ByVal pnMovNro As Long, ByVal psCotizacNro As String, ByVal psAreaCod As String, ByVal psMovFormaPago As String, ByVal psMovMoneda As String, ByVal psMovPlazo As String, ByVal psMovLugarEntrega As String, ByVal psPresuCod As String, ByVal pnPresuPuestado As Integer, ByVal psAgeCodRef As String, psAreaCodRef As String, psApruebaADM As String) As Integer
'    On Error GoTo InsertaMovCtaErr
'    Dim sql As String
'    InsertaMovCotizac = 1
'
'    sql = " INSERT INTO MOVCOTIZAC (nMovNro,cCotizacNro,cAreaCod,cMovFormaPago,cMovMoneda,dMovPlazo,cMovLugarEntrega,cPresuCod,nPresupuestado,cAgeCodRef,cAreaCodRef, cApruebaAdm)" _
'        & " VALUES(" & pnMovNro & ",'" & psCotizacNro & "','" & psAreaCod & "','" & psMovFormaPago & "','" & psMovMoneda & "','" & psMovPlazo & "','" & psMovLugarEntrega & "','" & psPresuCod & "'," & pnPresuPuestado & ",'" & psAgeCodRef & "','" & psAreaCodRef & "','" & psApruebaADM & "')"
'
'    oconect.Ejecutar sql
'    InsertaMovCotizac = 0
'    Exit Function
'InsertaMovCtaErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCotizac Method")
'End Function

'Public Function InsertaMovCotizac(ByVal pnMovNro As Long, ByVal psCotizacNro As String, ByVal psAreaCod As String, ByVal psMovFormaPago As String, ByVal psMovMoneda As String, ByVal psMovPlazo As String, ByVal psMovLugarEntrega As String, ByVal psPresuCod As String, ByVal psTipoOC As String, ByVal pnNroDiasAtraso As Currency) As Integer
'    On Error GoTo InsertaMovCtaErr
'    Dim Sql As String
'    InsertaMovCotizac = 1
'
'    Sql = " INSERT INTO MOVCOTIZAC (nMovNro,cCotizacNro,cAreaCod,cMovFormaPago,cMovMoneda,dMovPlazo,cMovLugarEntrega,cPresuCod,ctipoOC,nNroDiasAtraso)" _
'        & " VALUES(" & pnMovNro & ",'" & psCotizacNro & "','" & psAreaCod & "','" & psMovFormaPago & "','" & psMovMoneda & "','" & psMovPlazo & "','" & psMovLugarEntrega & "','" & psPresuCod & "','" & psTipoOC & "','" & pnNroDiasAtraso & "')"
'
'    oconect.Ejecutar Sql
'    InsertaMovCotizac = 0
'    Exit Function
'InsertaMovCtaErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
'End Function

Public Function InsertaMovCotizac(ByVal pnMovNro As Long, ByVal psCotizacNro As String, ByVal psAreaCod As String, ByVal psMovFormaPago As String, ByVal psMovMoneda As String, ByVal psMovPlazo As String, ByVal psMovLugarEntrega As String, ByVal psPresuCod As String, ByVal psTipoOC As String, ByVal pnNroDiasAtraso As Currency) As Integer
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    InsertaMovCotizac = 1
    
    sql = " INSERT INTO MOVCOTIZAC (nMovNro,cCotizacNro,cAreaCod,cMovFormaPago,cMovMoneda,dMovPlazo,cMovLugarEntrega,cPresuCod,ctipoOC,nNroDiasAtraso)" _
        & " VALUES(" & pnMovNro & ",'" & psCotizacNro & "','" & psAreaCod & "','" & psMovFormaPago & "','" & psMovMoneda & "','" & psMovPlazo & "','" & psMovLugarEntrega & "','" & psPresuCod & "','" & psTipoOC & "','" & pnNroDiasAtraso & "')"
      
    oConect.Ejecutar sql
    InsertaMovCotizac = 0
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCta Method")
End Function


'##ModelId=3A848F08038B
Public Function ActualizaMovCotizac(ByVal pnMovNro As Long, pnLogPoderes As Integer, psMovActualiza As String) As Integer
    On Error GoTo InsertaMovCtaErr
    Dim sql As String
    ActualizaMovCotizac = 1
    
    If pnLogPoderes = LogPoderesAprobacionADM Then
        sql = "cApruebaAdm = '" & psMovActualiza & "'"
    ElseIf pnLogPoderes = LogPoderesAprobacionJTer Then
        sql = "cApruebaJTer = '" & psMovActualiza & "'"
    ElseIf pnLogPoderes = LogPoderesAprobacionJLog Then
        sql = "cApruebaJLog = '" & psMovActualiza & "'"
    ElseIf pnLogPoderes = LogPoderesConfirmaADM Then
        sql = "cConformidadAdm = '" & psMovActualiza & "'"
    ElseIf pnLogPoderes = LogPoderesConfirmaJLog Then
        sql = "cConformidadJlog = '" & psMovActualiza & "'"
    End If
    
    sql = " UPDATE MOVCOTIZAC " _
        & " SET " & sql & " " _
        & " WHERE nMovNro = " & pnMovNro
      
    oConect.Ejecutar sql
    ActualizaMovCotizac = 0
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCotizac Method")
End Function

Public Function InsertaMovCotizacPenalidad(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal psMonto As String, ByVal psMovNro As String, ByVal psConcepto As String)
    On Error GoTo InsertaMovCtaErr
    Dim Sql1 As String
    Dim Sql2 As String
    Dim Sql3 As String
    Dim Sql4 As String
    Dim rs As New ADODB.Recordset
    
    InsertaMovCotizacPenalidad = 1
    
    Sql3 = " Update MOVCOTIZACPENALIDAD set bvigente=0 where nmovnro=" & pnMovNro & " and nmovitem=" & pnMovItem & ""
    
    Sql2 = " Select max(nmovaut) nItem From MOVCOTIZACPENALIDAD Where nmovnro=" & pnMovNro & " and nmovitem=" & pnMovItem & ""
    Set rs = oConect.CargaRecordSet(Sql2)
      
    If IsNull(rs!nItem) Then
        Sql1 = " INSERT INTO MOVCOTIZACPENALIDAD (nMovNro,nMovItem,nmovaut,nImporte,bvigente,cMovNro,cConcepto)" _
            & " VALUES(" & pnMovNro & "," & pnMovItem & ",1," & psMonto & ",1,'" & psMovNro & "','" & psConcepto & "')"
    
    Else
        Sql1 = " INSERT INTO MOVCOTIZACPENALIDAD (nMovNro,nMovItem,nmovaut,nImporte,bvigente,cMovNro,cConcepto)" _
            & " VALUES(" & pnMovNro & "," & pnMovItem & "," & rs!nItem + 1 & "," & psMonto & ",1,'" & psMovNro & "','" & psConcepto & "')"
    End If
    oConect.Ejecutar Sql3
    oConect.Ejecutar Sql1
   
    InsertaMovCotizacPenalidad = 0
    Exit Function
InsertaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCotizacPenalidad Method")

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
    
    If lsActualiza = "" Then Exit Sub
    
    lsActualiza = Mid(lsActualiza, 1, Len(lsActualiza) - 1)
    
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

'##ModelId=3A848F3301A5
Public Function InsertaMovGastoPersonas(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psPersCod As String, ByVal pcAgeCod As String, ByVal pcProducto As String, ByVal pnImporte As Currency) As Integer
    On Error GoTo InsertaMovGastoPersonasErr
    Dim sql As String
    InsertaMovGastoPersonas = 1
     
    sql = " INSERT INTO MovGastoPersonas (nMovNro, nMovItem, cPersCod, cAgeCod, cProductoCod, nImporte) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psPersCod & "','" & pcAgeCod & "','" & pcProducto & "'," & pnImporte & " )"
    
    oConect.Ejecutar sql
    InsertaMovGastoPersonas = 0
    Exit Function
InsertaMovGastoPersonasErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovGastoPersonasObj Method")
End Function

'##ModelId=3A848F3301A5
Public Function InsertaMovGastoPersonasPorcent(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psPersCod As String, ByVal pcAgeCod As String, ByVal pcProducto As String, ByVal pnImporte As Currency) As Integer
    On Error GoTo InsertaMovGastoPersonasPorcentErr
    Dim sql As String
    InsertaMovGastoPersonasPorcent = 1
     
    sql = "DELETE FROM MovGastoPersonasPorcent WHERE nMovNro = " & pnMovNro & " and nMovItem = " & pnMovItem & " and cPersCod = '" & psPersCod & "'"
    oConect.Ejecutar sql
    
    sql = " INSERT INTO MovGastoPersonasPorcent (nMovNro, nMovItem, cPersCod, cAgeCod, cProductoCod, nImporte) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psPersCod & "','" & pcAgeCod & "','" & pcProducto & "'," & pnImporte & " )"
    oConect.Ejecutar sql
    
    InsertaMovGastoPersonasPorcent = 0
    Exit Function
InsertaMovGastoPersonasPorcentErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovGastoPersonasObj Method")
End Function

'##ModelId=3A848F3301A5
Public Function InsertaMovGastoPersonasPorcentAnt(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovNroAnt As Long, ByVal pnMovItemAnt As Long) As Integer
    On Error GoTo InsertaMovGastoPersonasPorcentAntErr
    Dim sql As String
    InsertaMovGastoPersonasPorcentAnt = 1
     
    sql = " Insert MovGastoPersonasPorcent (nMovNro, nMovItem, cPersCod, cAgeCod, cProductoCod, nImporte)" _
        & " Select " & pnMovNro & " nMovNro, " & pnMovItem & " nMovItem, cPersCod, cAgeCod, cProductoCod, nImporte from MovGastoPersonasPorcent" _
        & " Where nMovNro = " & pnMovNroAnt & " And nMovItem = " & pnMovItemAnt & ""
    
    oConect.Ejecutar sql
    InsertaMovGastoPersonasPorcentAnt = 0
    Exit Function
InsertaMovGastoPersonasPorcentAntErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovGastoPersonasObj Method")
End Function

'EJVG20140321 Se agregó pnMovAlmacenDest
Public Function InsertaMovBS(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovAlmacen As Long, ByVal psBSCod As String, Optional ByVal pnMovAlmacenDest As Integer = 0) As Integer
    On Error GoTo InsertaMovBSErr
    Dim sql As String
    InsertaMovBS = 1
     
    sql = "INSERT INTO MOVBS (nMovNro, nMovItem, nMovBSOrden, cBSCod, nMovBSOrdenDest) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & "," & pnMovAlmacen & ",'" & psBSCod & "'," & pnMovAlmacenDest & ")"
    oConect.Ejecutar sql
    InsertaMovBS = 0
    Exit Function
InsertaMovBSErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovBS Method")
End Function

Public Function InsertaMovBSActivoFijo(ByVal pnAnio As Long, ByVal pnMovNro As Long, ByVal psBSCod As String, ByVal psSerie As String, pnValorIni As Currency, pnMontoDeprecia As Currency, pdActivacion As Date, psArea As String, psAgencia As String, psPeriodos As String, psPerioIni As String, ByVal psDescripcion As String, ByVal pdCompra As Date, ByVal pdBase As Date, ByVal pBan As Integer) As Integer
    On Error GoTo InsertaMovBSErr
    Dim sql As String
    InsertaMovBSActivoFijo = 1

    sql = " Insert BSActivoFijo (nAnio,nMovNro,cBSCod,cSerie,nBSValor,nBSDeprecia,dActivacion, cAreCod, cAgeCod, nBSPerDeprecia, nBSPerDepAct,cDescripcion, dCompra, dBase, Ban)" _
        & " Values(" & pnAnio & "," & pnMovNro & ",'" & psBSCod & "','" & psSerie & "'," & pnValorIni & "," & pnMontoDeprecia & ",'" & Format(pdActivacion, gsFormatoFecha) & "','" & psArea & "','" & psAgencia & "','" & psPeriodos & "','" & psPerioIni & "','" & psDescripcion & "','" & Format(pdCompra, gsFormatoFecha) & "', '" & Format(pdBase, gsFormatoFecha) & "', " & pBan & ")"
    oConect.Ejecutar sql
    
    InsertaMovBSActivoFijo = 0
    Exit Function
InsertaMovBSErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovBS Method")
End Function

Public Function InsertaMovBSActivoFijoUnico(ByVal pnAnio As Long, ByVal pnMovNro As Long, ByVal pcBSCod As String, ByVal pcSerie As String, ByVal pnBSValor As Currency, ByVal pnBSDeprecia As Currency, ByVal pbDepreciado As String, ByVal pdActivacion As Date, ByVal pcAreCod As String, ByVal pcAgeCod As String, ByVal pnBSPerDeprecia As Long, ByVal pnBSPerDepAct As Long, ByVal pcDescripcion As String, ByVal pbDeprecia As String, ByVal pbIntangible As String, ByVal pbForsado As String, ByVal pcNroID As String, ByVal pcNroPatrimonio As String, ByVal pdCompra As Date, ByVal pdBase As Date, ByVal pBan As String, ByVal pbBaja As String, ByVal pcPersCod As String, ByVal pcCategoBien As String, ByVal pnTasaDepre As Currency, _
                            Optional ByVal psMarca As String = "", Optional ByVal psModelo As String = "", Optional psSeriePlaca As String = "", Optional ByVal pnBSPerDepreciaTrib As Integer = 0, Optional ByVal pnMoneda As Moneda = gMonedaNacional) As Integer
    '*** PEAC 20100506 - SE AGREGO EL PARAMETRO (pcCategoBien,pnTasaDepre)
    '*** PEAC 20120326 - SE AGREGO EL PARAMETRO MORCRA , MODELO, SERIE
    'EJVG20130531 - SE AGREGO EL PARAMETRO nBSPerDepreciaTrib y nMoneda
    On Error GoTo InsertaMovBSErr
    Dim sql As String
    InsertaMovBSActivoFijoUnico = 1

    sql = " Insert BSActivoFijo (nAnio,nMovNro,cBSCod,cSerie,nBSValor,nBSDeprecia,bDepreciado,dActivacion,cAreCod,cAgeCod,nBSPerDeprecia,nBSPerDepAct,cDescripcion,bDeprecia,bIntangible,bForsado,cNroID,cNroPatrimonio,dCompra,dBase,Ban,bBaja,cPersCod,cCategoBien,nTasaDepre,vMarca,vModelo,vSerie,nBSPerDepreciaTrib,nMoneda)" _
        & " Values(" & pnAnio & "," & pnMovNro & ",'" & pcBSCod & "','" & pcSerie & "'," & pnBSValor & "," & pnBSDeprecia & "," & pbDepreciado & ",'" & Format(pdActivacion, gsFormatoFecha) & "','" & pcAreCod & "','" & pcAgeCod & "'," & pnBSPerDeprecia & "," & pnBSPerDepAct & ",'" & pcDescripcion & "'," & pbDeprecia & "," & pbIntangible & "," & pbForsado & ",'" & pcNroID & "','" & pcNroPatrimonio & "','" & Format(pdCompra, gsFormatoFecha) & "','" & Format(pdBase, gsFormatoFecha) & "'," & pBan & "," & pbBaja & ",'" & pcPersCod & "','" & pcCategoBien & "'," & pnTasaDepre & ",'" & psMarca & "','" & psModelo & "','" & psSeriePlaca & "'," & pnBSPerDepreciaTrib & "," & pnMoneda & ")"
    oConect.Ejecutar sql
    'EJVG20130621 ***
    sql = " INSERT INTO BSActivoFijoVidaUtil (nMovNroAF,nMovNro,nBSPerDeprecia)" _
        & " VALUES(" & pnMovNro & "," & pnMovNro & "," & pnBSPerDeprecia & ")"
    oConect.Ejecutar sql
    'END EJVG *******
    InsertaMovBSActivoFijoUnico = 0
    Exit Function
InsertaMovBSErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovBS Method")
End Function

Public Function InsertaMovBSAF(ByVal pnAnio As Long, ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psBSCod As String, ByVal psSerie As String, pnMovNroRef As Long, Optional psTipo As String) As Integer
    On Error GoTo InsertaMovBSErr
    Dim sql As String
    InsertaMovBSAF = 1

    sql = " Insert MovBSAF(nAnio, nMovNro, nMovItem, cBSCod, cSerie, nMovNroRef, cTipoActivo)" _
        & " Values(" & pnAnio & "," & pnMovNro & "," & pnMovItem & ",'" & psBSCod & "','" & psSerie & "'," & pnMovNroRef & ",'" & psTipo & "')"
    oConect.Ejecutar sql
    
    InsertaMovBSAF = 0
    Exit Function
InsertaMovBSErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovBS Method")
End Function

Public Function InsertaMovBSBND(ByVal psBSCod As String, ByVal psSerie As String, pnMovNroRef As Long) As Integer
    On Error GoTo InsertaMovBSErr
    Dim sql As String
    InsertaMovBSBND = 1

    sql = " Insert MovBSBND(cBSCod, cSerie, nMovNro)" _
        & " Values('" & psBSCod & "','" & psSerie & "'," & pnMovNroRef & ")"
    oConect.Ejecutar sql
    
    InsertaMovBSBND = 0
    Exit Function
InsertaMovBSErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovBS Method")
End Function

'##ModelId=3A848F3301A5
'Public Function InsertaMovBSSerie(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psBSCod As String, ByVal psSerie As String, Optional pnIGVActivo As Currency = 0, Optional pnValorActivo As Currency = 0, Optional psMovNro As String = "2004") As Integer
'    On Error GoTo InsertaMovObjErr
'    Dim Sql As String
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'    InsertaMovBSSerie = 1
'
'    Sql = " Select cBSCod From BSSerie Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
'    Set rs = oconect.CargaRecordSet(Sql)
'
'    If rs.EOF And rs.BOF Then
'        Sql = " Insert BSSerie (cBSCod,cSerie,nBienEstado, nIGV, nValor)" _
'            & " Values('" & psBSCod & "','" & psSerie & "','0'," & pnIGVActivo & "," & pnValorActivo & ")"
'    Else
'        Sql = " Update BSSerie" _
'            & " Set nBienEstado = '0',  nIGV = " & pnIGVActivo & ", nValor = " & pnValorActivo & "" _
'            & " Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
'    End If
'    oconect.Ejecutar Sql
'
'    Sql = " Update MBS" _
'        & " Set nIGV = " & pnIGVActivo & " , nValor = " & pnValorActivo & "" _
'        & " From Mov M" _
'        & " Inner Join MovBSSerie MBS On M.nMovNro = MBS.nMovNro" _
'        & " Where copecod like '5912%' and M.nMovFlag = " & MovFlag.gMovFlagVigente & " And M.nMovEstado = " & MovEstado.gMovEstContabMovContable & "" _
'        & " and left(M.cMovNro,6) = '" & Left(psMovNro, 6) & "' and MBS.cBSCod = '" & psBSCod & "' And MBS.cSerie = '" & psSerie & "'"
'    oconect.Ejecutar Sql
'
'    Sql = " INSERT INTO MOVBSSerie (nMovNro, nMovItem, cBSCod, cSerie, nIGV, nValor) " _
'        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psBSCod & "','" & psSerie & "'," & pnIGVActivo & "," & pnValorActivo & ")"
'    oconect.Ejecutar Sql
'
'    InsertaMovBSSerie = 0
'    Exit Function
'InsertaMovObjErr:
'    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
'End Function

Public Function InsertaMovBSSerie(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psBSCod As String, ByVal psSerie As String, Optional pnIGVActivo As Currency = 0, Optional pnValorActivo As Currency = 0, Optional ByVal psSerieFisica As String) As Integer
    On Error GoTo InsertaMovObjErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    InsertaMovBSSerie = 1
    
    sql = " Select cBSCod From BSSerie Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    Set rs = oConect.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        sql = " Insert BSSerie (cBSCod,cSerie,nBienEstado, nIGV, nValor,cSerieFisica)" _
            & " Values('" & psBSCod & "','" & psSerie & "','0'," & pnIGVActivo & "," & pnValorActivo & ",'" & psSerieFisica & "')"
    Else
        sql = " Update BSSerie" _
            & " Set nBienEstado = '0',  nIGV = " & pnIGVActivo & ", nValor = " & pnValorActivo & ",cSerieFisica ='" & psSerieFisica & "' " _
            & " Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    End If
    oConect.Ejecutar sql
    
    sql = " INSERT INTO MOVBSSerie (nMovNro, nMovItem, cBSCod, cSerie, nIGV, nValor,cSerieFisica) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psBSCod & "','" & psSerie & "'," & pnIGVActivo & "," & pnValorActivo & ",'" & psSerieFisica & "')"
    oConect.Ejecutar sql
    
    InsertaMovBSSerie = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
End Function


'##ModelId=3A848F3301A5
Public Function InsertaMovSalidaBSSerie(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psBSCod As String, ByVal psSerie As String, Optional pnIGVActivo As Currency = 0, Optional pnValorActivo As Currency = 0) As Integer
    On Error GoTo InsertaMovObjErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    InsertaMovSalidaBSSerie = 1
    
    sql = " Update BSSerie" _
        & " Set nBienEstado = '1'" _
        & " Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    oConect.Ejecutar sql
    
    sql = " INSERT INTO MOVBSSerie (nMovNro, nMovItem, cBSCod, cSerie, nIGV, nValor) " _
        & " VALUES(" & pnMovNro & "," & pnMovItem & ",'" & psBSCod & "','" & psSerie & "'," & pnIGVActivo & "," & pnValorActivo & ")"
    oConect.Ejecutar sql
    
    InsertaMovSalidaBSSerie = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
End Function

'##ModelId=3A848F3301A5
Public Function EliminaMovBSSerieparaActualizar(ByVal pnMovNroActualizar As Long) As Integer
    On Error GoTo InsertaMovObjErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    EliminaMovBSSerieparaActualizar = 1
    
    sql = " Update BSSerie Set nBienEstado = '2' Where  cBSCod + cSerie" _
        & " In (Select  cBSCod + cSerie From MovBSSerie where nMovNro  = '" & pnMovNroActualizar & "')"
    Set rs = oConect.CargaRecordSet(sql)
    
    oConect.Ejecutar sql
    
    EliminaMovBSSerieparaActualizar = 0
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
'##ModelId=3A848F3301A5
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObjIF Method")
End Function

'##ModelId=3A848F4F002E
Public Function InsertaMovDoc(ByVal pnMovNro As Long, ByVal pnDocTpo As TpoDoc, ByVal psDocNro As String, ByVal psDocFecha As String) As Integer
    On Error GoTo InsertaMovDocErr
    Dim sql As String
    InsertaMovDoc = 1
    
    sql = "INSERT INTO MOVDOC(nMovNro, nDocTpo, cDocNro, dDocFecha ) " _
        & "VALUES (" & pnMovNro & "," & pnDocTpo & ",'" & psDocNro & "','" & psDocFecha & "')"
        
    oConect.Ejecutar sql
    InsertaMovDoc = 0
    Exit Function
InsertaMovDocErr:
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
Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    Dim oCon As NConstSistemas
    Set oCon = New NConstSistemas
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    lbBitControlCuentasUltNivel = IIf(oCon.LeeConstSistema(gConstSistBitControCtasContUltNiv) = "1", True, False)
    
    Set oImp = Nothing

    Dim oIni As ClasIni
    
    Set oIni = New ClasIni
    
    vsConexion = oIni.CadenaConexion
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
    
     CargaVarSistema False
End Sub
Private Sub Class_Terminate()
'oConect.CierraConexion
'Set oConect = Nothing
End Sub

'##ModelId=3A837398005D
'Public Function GeneraMovME(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional pbEsCompra As Boolean = False) As Boolean
Public Function GeneraMovME(ByVal pnMovNro As Long, ByVal psMovNro As String, Optional nTCMerc As Currency = 0, Optional pbEsCompra As Boolean = False, Optional pnTipoPago As Integer = 0) As Boolean
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
        nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
        nTCComp = oTC.EmiteTipoCambio(dFecha, TCCompra)
   Else
        If pbEsCompra Then
            nTCVent = oTC.EmiteTipoCambio(dFecha, TCVenta)
            nTCComp = nTCMerc
        Else
            nTCVent = nTCMerc
            nTCComp = nTCMerc
        End If
   End If
   Set oTC = Nothing
   'ALPA 20090405******************************
   'oConect.Ejecutar " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent
   oConect.Ejecutar " sp_GeneraMovME " & pnMovNro & ", " & nTCFijo & ", " & nTCComp & " , " & nTCVent & "," & pnTipoPago
   '*******************************************
   GeneraMovME = True
Exit Function
GeneraMovMEErr:
    Call RaiseError(MyUnhandledError, "DMov:GeneraMovME Method")
End Function

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
Public Function InsertaMovGasto(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal psDestino As String, Optional psPersCodDest As String = "") As Integer
    On Error GoTo InsertaMovGastoErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select nMovNro from movGasto where nMovNro = " & pnMovNro
    Set rs = oConect.CargaRecordSet(sql)
    
    InsertaMovGasto = 1
    
    If rs.BOF And rs.EOF Then
        sql = "INSERT INTO MovGasto (nMovNro, cPersCod, cDestino, cPersCodDest) " _
         & "VALUES (" & pnMovNro & ", '" & psPersCod & "','" & psDestino & "','" & psPersCodDest & "')"
        
        oConect.Ejecutar sql
    End If
    
    rs.Close
    Set rs = Nothing
    
    InsertaMovGasto = 0
    Exit Function
InsertaMovGastoErr:
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
     & Format(pdValorizaRef, vsFormatoFecha) & "','" & Format(pdValorizacion, vsFormatoFecha) & "','" & psDepIF & "'," & pnConfCaja & "," & pnMonedaCheque & "," & pnEstado & ",'" & psAreaCod & "','" & psAgeCod & "')"
     
    oConect.Ejecutar sql
    InsertaCheque = 0
    Exit Function
InsertaChequeErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaCheque Method")
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
        lsFiltro = lsFiltro + " dValorizaRef  ='" & Format(pdValorizaRef, vsFormatoFecha) & "', "
    End If
    If pdValorizacion <> "" Then
        lsFiltro = lsFiltro + " dValorizacion ='" & Format(pdValorizacion, vsFormatoFecha) & "', "
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
    Call RaiseError(MyUnhandledError, "DMov:ActualizaCheque Method")
End Function

Public Function InsertaChequeEstado(ByVal psTpoDoc As TpoDoc, ByVal psNroDoc As String, psPersCod As String, ByVal psTipoIF As String, ByVal pdFechaEstado As Date, ByVal psEstado As ChequeEstado, ByVal psMovNro As String) As Integer
    On Error GoTo InsertaChequeEstadoErr
    Dim sql As String
    InsertaChequeEstado = 1
    
    sql = " INSERT INTO DocRecEst(nTpoDoc, cNroDoc, cPersCod, cIFTpo , nEstado , cMovNro ) " _
     & "    VALUES (" & psTpoDoc & ", '" & psNroDoc & "','" & psPersCod & "','" & Format(psTipoIF, "00") & "'," _
     & psEstado & ",'" & psMovNro & "')"
     
    oConect.Ejecutar sql
    InsertaChequeEstado = 0
    Exit Function
InsertaChequeEstadoErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaChequeEstado Method")
End Function
Public Function ActualizaMovimiento(ByVal psMovNroNew As String, ByVal pnMovNroAnt As Long, Optional psMovEstado As MovEstado = gMovEstContabMovContable, _
                                Optional plActualizaAsiento As Boolean = True, Optional pbInsertaMov As Boolean = True, Optional pbDeleteMovAnt As Boolean = True, _
                                Optional ByVal psOpeCod As String = "", Optional ByVal psMovDesc As String = "", Optional ByVal lbUpdaMovRef As Boolean = True) As Integer
On Error GoTo ActualizaMovimientoErr
   Dim sql As String
   Dim lnNovNronew As Long
   Dim lnMovNronew As Long
   ActualizaMovimiento = 1
   sql = ""
   
   If plActualizaAsiento Then
      If pbInsertaMov Then
      
        sql = " INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag )" _
              & " SELECT '" & psMovNroNew & "'," & IIf(psOpeCod = "", "cOpeCod", "'" & psOpeCod & "'") & "," _
            & IIf(psMovDesc = "", "cMovDesc", "'" & psMovDesc & "'") & ",'" & psMovEstado & "', nMovFlag  " _
            & " FROM Mov WHERE nMovNro =" & pnMovNroAnt & " "
        
        oConect.Ejecutar sql
        lnMovNronew = GetnMovNro(psMovNroNew)
      Else
        lnMovNronew = GetnMovNro(psMovNroNew)
        'lnMovNroNew = psMovNroNew
      End If
      sql = ""
      sql = " UPDATE MOVCONT        SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
      oConect.Ejecutar sql
   End If
   sql = " UPDATE MOVGASTO       SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   'sql = "UPDATE MOVARQUEO      SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
   
   sql = "INSERT INTO MOVARQUEO(nMovNro, cAreaCod, cPersCod, cHoraIni, cHoraFin, nEfectivo, nOPImporte, nTotDoc, nTotDocRend, nArendirPend, nDiferencia )      " _
    & " SELECT " & lnMovNronew & ",cAreaCod, cPersCod, cHoraIni, cHoraFin, nEfectivo, nOPImporte, nTotDoc, nTotDocRend, nArendirPend, nDiferencia FROM MOVARQUEO Where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   
   sql = "UPDATE MOVARQUEOOP SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   
   sql = "UPDATE MOVARQUEOEFECTIVO      SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   
   sql = "DELETE MOVARQUEO Where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   
   sql = "UPDATE MOVDOC         SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   If lbUpdaMovRef Then
        sql = "UPDATE MOVREF         SET nMovNro    =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        sql = "UPDATE MOVREF         SET nMovNroRef =" & lnMovNronew & " Where nMovNroRef =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
   End If
   sql = "UPDATE MOVARENDIR     SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = "UPDATE MOVARENDIRSUST SET nMovNroRend=" & lnMovNronew & " where nMovNroSust =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = "UPDATE MOVOTROSITEM   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = "UPDATE MOVCAJACHICA   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = "UPDATE MOVCARTAFIANZA SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = "UPDATE MOVTPOCAMBIO   SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
   sql = "UPDATE MOVVIATICOS    SET nMovNro    =" & lnMovNronew & " where nMovNro =" & pnMovNroAnt & " "
   oConect.Ejecutar sql
    
   If plActualizaAsiento Then
        sql = " INSERT MOVCTA (nMovNro, nMovItem, cCtaContCod, nMovImporte ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, cCtaContCod, nMovImporte FROM MOVCTA Where nMovNro =" & pnMovNroAnt & " "
        
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJ (nMovNro, nMovItem, nMovObjOrden, cObjetoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cObjetoCod FROM MOVOBJ Where nMovNro =" & pnMovNroAnt & " "
        
        oConect.Ejecutar sql
        
        sql = "UPDATE MOVOBJPers        SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        sql = "UPDATE MOVOBJIF          SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        sql = "UPDATE MOVOBJEfectivo    SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        sql = "UPDATE MOVOBJBS          SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        sql = "UPDATE MOVOBJAreaAgencia SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        sql = "UPDATE MOVME             SET nMovNro =" & lnMovNronew & " Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        'sql = "UPDATE MOVOBJ            SET nMovNro =" & lnMovNroNew & " Where nMovNro =" & pnMovNroAnt & " "
        sql = "DELETE MOVOBJ            Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        'sql = "UPDATE MOVCTA            SET nMovNro =" & lnMovNroNew & " Where nMovNro =" & pnMovNroAnt & " "
        sql = "DELETE MOVCTA            Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        If pbDeleteMovAnt Then
            sql = "Update MOV set nMovFlag=" & gMovFlagEliminado & " where nMovNro =" & pnMovNroAnt & " "
            oConect.Ejecutar sql
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
Public Function CargaMovOpeAsiento(ByVal psMovNro As String) As ADODB.Recordset
Dim psSql As String
   On Error GoTo CargaOpeMovAsientoErr
      psSql = "SELECT m.cMovNro, m.cOpeCod, o.cOpeDesc, m.cMovDesc, m.nMovEstado, m.nMovFlag, ISNULL(mg.cPersCod,'') cPersCod, ISNULL(p.cPersNombre,'') cPersNombre " _
            & "FROM Mov M JOIN " & vsServerComunes & "OpeTpo O ON O.cOpeCod = M.cOpeCod " _
            & "      LEFT JOIN MovGasto mg ON mg.nMovNro = m.nMovNro " _
            & "      LEFT JOIN " & vsServerPersonas & "Persona p ON p.cPersCod = mg.cPersCod " _
            & "WHERE M.cMovNro = '" & psMovNro & "'"
      Set CargaMovOpeAsiento = oConect.CargaRecordSet(psSql)
   Exit Function
CargaOpeMovAsientoErr:
   Call RaiseError(MyUnhandledError, "DMov:CargaMovOpeAsiento Method")
End Function

Public Function CargaMovCtaAsiento(ByVal psMovNro As String) As ADODB.Recordset
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
         & "WHERE  cMovNro = '" & psMovNro & "'" _
         & "GROUP BY cMovNro, a.nMovItem, a.cCtaContCod, b.cCtaContDesc ORDER BY a.nMovItem "
      Set CargaMovCtaAsiento = oConect.CargaRecordSet(psSql)
   Exit Function
CargaMovCtaAsientoErr:
   Call RaiseError(MyUnhandledError, "DMov:CargaMovCtaAsiento Method")
End Function


Public Function CargaMovObjAsiento(ByVal psMovNro As String) As ADODB.Recordset
Dim psSql As String
   On Error GoTo CargaMovObjAsientoErr
      psSql = "SELECT cMovNro, mo.nMovItem, mo.nMovObjOrden, mo.cObjetoCod, o.cObjetoDesc " _
           & "FROM Mov m JOIN MovObj mo ON m.nMovNro = mo.nMovNro JOIN " & vsServerComunes & "Objeto O ON o.cObjetoCod = mo.cObjetoCod " _
           & "WHERE cMovNro = '" & psMovNro & "' "
      Set CargaMovObjAsiento = oConect.CargaRecordSet(psSql)
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
                                Optional ByVal pbEliminaMov As Boolean) As Integer
On Error GoTo ExtornaMovimientoErr
    Dim sql As String
    Dim lnMovNronew As Long
    Dim lsMovNroAnt As String
    Dim ldFechaAnt As Date
    Dim ldFechaExt As Date
    
    
    lsMovNroAnt = GetcMovNro(pnMovNroAnt)
    ldFechaAnt = CDate(Mid(lsMovNroAnt, 7, 2) & "/" & Mid(lsMovNroAnt, 5, 2) & "/" & Mid(lsMovNroAnt, 1, 4))
    ldFechaExt = CDate(Mid(psMovNroNew, 7, 2) & "/" & Mid(psMovNroNew, 5, 2) & "/" & Mid(psMovNroNew, 1, 4))
    
    
    ExtornaMovimiento = 1
    
    sql = " INSERT INTO MOV (cMovNro, cOpeCod, cMovDesc, nMovEstado, nMovFlag )" _
        & " SELECT '" & psMovNroNew & "'," & IIf(psOpeCod = "", "cOpeCod", "'" & psOpeCod & "'") & ", " _
        & IIf(psMovDesc = "", "cMovDesc", "'" & Replace(psMovDesc, "''", "'") & "'") & ",nMovEstado, '" & gMovFlagDeExtorno & "'  " _
        & " FROM Mov WHERE nMovNro ='" & pnMovNroAnt & "' "
   
    oConect.Ejecutar sql
    lnMovNronew = GetnMovNro(psMovNroNew)
    InsertaMovRef lnMovNronew, pnMovNroAnt
    If ldFechaAnt = ldFechaExt Or pbEliminaMov = True Then
        ActualizaMov pnMovNroAnt, , , gMovFlagEliminado
        ExtornaMovCtaIF pnMovNroAnt
    ElseIf ldFechaAnt <> ldFechaExt Or pbEliminaMov = False Then
        sql = ""
      
        sql = "INSERT MOVCONT (nMovNro, nMovMonto, nMovCorrela, cMovParalelo) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nMovMonto, nMovCorrela, cMovParalelo FROM MovCont Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
       
        sql = "INSERT MOVDOC (nMovNro, nDocTpo, cDocNro, dDocFecha) " _
            & " SELECT nMovNro = " & lnMovNronew & ", nDocTpo, cDocNro, dDocFecha FROM MovDoc Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVGASTO(nMovNro,cPersCod,cDestino) " _
            & " SELECT nMovNro = " & lnMovNronew & ",cPersCod,cDestino FROM MovGasto Where nMovNro =" & pnMovNroAnt & " "
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
        
        sql = "INSERT MOVOBJPers (nMovNro, nMovItem, nMovObjOrden, cPersCod) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cPersCod FROM MovObjPers Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJIF (nMovNro, nMovItem, nMovObjOrden, cPersCod, cCtaIfCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cPersCod, cCtaIfCod FROM MOVOBJIF Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVCTAIF (nMovNro, cPersCod, cIFTpo, cCtaIFCod, nConceptoCod, nMovImporte) " _
            & " SELECT nMovNro =" & lnMovNronew & ", cPersCod, cIFTpo, cCtaIFCod, nConceptoCod, nMovImporte FROM MOVCTAIF Where nMovNro =" & pnMovNroAnt & " "
        
        oConect.Ejecutar sql
                
        sql = "INSERT MOVOBJEfectivo (nMovNro, nMovItem, nMovObjOrden, cEfectivoCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cEfectivoCod FROM MOVOBJEfectivo Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJBS (nMovNro, nMovItem, nMovObjOrden, cBieSerCod, nMovCantidad ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cBieSerCod, nMovCantidad FROM MOVOBJBS Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT MOVOBJAreaAgencia (nMovNro, nMovItem, nMovObjOrden, cAreaCod, cAgeCod ) " _
            & " SELECT nMovNro =" & lnMovNronew & ", nMovItem, nMovObjOrden, cAreaCod, cAgeCod FROM MOVOBJAreaAgencia Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT INTO MOVHABILITACION(nMovNro, cTranspCod, nMontoTrans) " _
            & " SELECT nMovNro=" & lnMovNronew & ", cTranspCod, nMontoTrans FROM MOVHABILITACION Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "INSERT INTO MOVHABDEVCAJERO(nMovNro, cUser, nMovImporte)  " _
            & " SELECT nMovNro=" & lnMovNronew & ",cUser, nMovImporte FROM MOVHABDEVCAJERO Where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        sql = "UPDATE MOV SET nMovFlag = '" & gMovFlagExtornado & "' where nMovNro =" & pnMovNroAnt & " "
        oConect.Ejecutar sql
        
        ExtornaMovCtaIF pnMovNroAnt
    End If
    ExtornaMovimiento = 0
    Exit Function
ExtornaMovimientoErr:
    Call RaiseError(MyUnhandledError, "DMov:ExtornaMovimiento Method")
End Function
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
        & Format(pdPartida, vsFormatoFecha) & "'," & pnMovViaticosDias & ",'" & psDestinoDesc & "')"
     
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
    Call RaiseError(MyUnhandledError, "DMov:ActualizaCajaChica Method")
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
   Call RaiseError(MyUnhandledError, "DMov:CargaMovCta Method")
End Function
Public Sub ActualizaMovCta_Cuenta(psCtaCodant As String, psCtaCodNew As String, psFecha1 As String, psFecha2 As String)
    Dim sql As String
    On Error GoTo ActualizaMovCtaErr
    sql = "UPDATE MovCta SET cCtaContCod = '" & psCtaCodNew & "' WHERE cCtaContCod = '" & psCtaCodant & "' and LEFT(cMovNro,8) BETWEEN '" & psFecha1 & "' and '" & psFecha2 & "' "
    oConect.Ejecutar sql
    Exit Sub
ActualizaMovCtaErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovCta_Cuenta Method")

End Sub
Public Sub InsertaMovTpoCambio(ByVal pnMovNro As Long, ByVal pnMovTpoCambio As Currency)
    On Error GoTo InsertaMovTipCambioErr
    Dim sql As String
    
    sql = "INSERT INTO MovTpoCambio (nMovNro, nMovTpoCambio) " _
     & "VALUES (" & pnMovNro & ", " & pnMovTpoCambio & ")"
     
    oConect.Ejecutar sql
    Exit Sub
InsertaMovTipCambioErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovTipCambio Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaMov Method")
End Function

Public Function InsertaMovVenta(ByVal psMovNro As String, ByVal psMovItem As String, ByVal psMovNroSub As String, ByVal psMovNroIng As String) As Integer
    On Error GoTo InsertaMovErr
    Dim sql As String
    
    sql = " Insert MovVenta (nMovNro,nMovItem,nMovNroSub,nMovNroIng)" _
        & " Values(" & psMovNro & "," & psMovItem & "," & psMovNroSub & "," & psMovNroIng & ")"
     
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

Public Function GetCtaDeprecia(ByVal pnDepreciaTpo As Integer) As String
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

sql = "Select cCtaContCod From BSActivoFijoParam Where nDepre = " & pnDepreciaTpo
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetCtaDeprecia = rs!cCtaContCod
End If
rs.Close
Set rs = Nothing
End Function




'Devuelve un nMovNro Referenciado , con un codigo de operacion en particular
Public Function GetnMovNroRef(ByVal pnMovNro As Long, psOpeCod As String) As Long
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

sql = " Select M.nMovNro from MovRef MR " _
    & " Inner Join Mov M On MR.nMovNroRef = M.nMovNro" _
    & " Where m.cOpeCod = '" & psOpeCod & "' And MR.nMovNro = " & pnMovNro
Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    GetnMovNroRef = rs!nMovNro
Else
    GetnMovNroRef = 0
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
    Call RaiseError(MyUnhandledError, "DMov:EliminaMovRef Method")
End Sub
Public Function InsertaMovHabilitaciones(ByVal pnMovNro As Long, ByVal psCodTransp As String, ByVal psAreaOrig As String, ByVal psAgeOrig As String, _
                                        ByVal psAreaDest As String, ByVal psAgeDest As String, _
                                        ByVal pnMontoTrans As Currency, ByVal pnMovImporte As Currency) As Integer
    On Error GoTo InsertaMovHabilitacionesErr
    Dim sql As String
    InsertaMovHabilitaciones = 1
    sql = "INSERT INTO MovHabilitacion (nMovNro, cTranspCod, cAreaOrig, cAgeOrig, cAreaDest, cAgeDest, nMontoTrans, nMovImporte)" _
        & " VALUES(" & pnMovNro & ",'" & psCodTransp & "','" & psAreaOrig & "','" & psAgeOrig & "','" _
        & psAreaDest & "','" & psAgeDest & "'," & pnMontoTrans & "," & pnMovImporte & ")"
     
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
    sql = "Insert MovHabDevCajero(nMovNro, cUser, nMovImporte) Values(" & pnMovNro & ",'" & psUser & "'," & pnImporte & ")"
 
    oConect.Ejecutar sql
    InsertaMovHabDev = 0
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
Public Function GeneraMovNro(ByVal pdFecha As Date, Optional ByVal psCodAge As String = "07", Optional ByVal psUser As String = "SIST", Optional psMovNro As String = "") As String
    On Error GoTo GeneraMovNroErr
    Dim rs As ADODB.Recordset
    Dim sql As String
    Set rs = New ADODB.Recordset
    
    If psMovNro = "" Or Len(psMovNro) <> 25 Then
       sql = "sp_GeneraMovNro '" & Format(pdFecha & " " & oConect.GetHoraServer, gsFormatoFechaHora) & "','" & psCodAge & "','" & psUser & "'"
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
                                    ByVal pnMotivo As Integer, ByVal pnMonto As Currency, Optional psObjetoCod As String = "", Optional psobjeto As String = "", Optional pnMoneda As Moneda = gMonedaNacional) As Integer
    On Error GoTo InsertaNotaAbonoCargoErr
    Dim sql As String
    InsertaNotaAbonoCargo = 1
    
    sql = " INSERT INTO NotaAbonoCargo(nDocTpo, cDocNro, nEstado, nMotivoCod, nMonto, cObjetoCodPadre, cObjetoCod, nMoneda)" _
        & " VALUES(" & pnTpoDoc & ",'" & psNroDoc & "'," & pnEstado & "," & pnMotivo & "," & pnMonto & ",'" & psObjetoCod & "','" & psobjeto & "'," & pnMoneda & ") "
    
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
    Case TpoDocOrdenPago
        sSQL = "INSERT DocRecOP (nTpoDoc,cNroDoc,cCtaCod,nMonto,cIFCodPers,nEstado,cMovNro) " _
            & "VALUES (" & nTpoDoc & ",'" & sNroDoc & "','" & sCuenta & "'," & nMonto & ",'" & sCodIF & "'," & nEstadoOP & ",'" & sMovNro & "')"
    Case TpoDocNotaAbono, TpoDocNotaCargo
End Select
oConect.Ejecutar sSQL
sSQL = "INSERT MovDoc (nMovNro,nDocTpo,cDocNro,dDocFecha) " _
    & "VALUES (" & nMovNro & "," & nTpoDoc & ",'" & sNroDoc & "','" & sFecha & "')"
oConect.Ejecutar sSQL
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
Public Function GetMovExtorno(ByVal sDatoBus As String, Optional nTipoBus As Integer = 0, Optional pdFecSis As Date) As Recordset
Dim rsMov As Recordset
Dim sWhere As String
If nTipoBus = 0 Then
    sWhere = "M.cMovNro LIKE '%" & sDatoBus & "%'"
ElseIf nTipoBus = 1 Then
    sWhere = "M.cMovNro LIKE '" & Format$(pdFecSis, "yyyymmdd") & "%' And C.cCtaCod = '" & sDatoBus & "'"
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
Public Sub ActualizaEstadoMov(ByVal nMovNro As Long, ByVal nMovFlag As MovFlag)
sSQL = "Update Mov Set nMovFlag = " & nMovFlag & " Where nMovNro = " & nMovNro
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
                                    Optional psObjetoCod As String = "", Optional psobjeto As String = "") As Integer
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
    If psobjeto <> "" Then
        lsActualiza = lsActualiza + " cObjetoCod='" & psobjeto & "',"
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
Public Function InsertaCuentaIF(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                ByVal psCtaIfCod As String, ByVal psCtaIFDesc As String, _
                                ByVal pdCtaIFAper As Date, ByVal psCtaIFCap As String, ByVal pdCtaIFInt As String, _
                                ByVal pnCtaIFPlazo As Integer, ByVal pnCtaIFEstado As CGEstadoCtaIF, _
                                ByVal psUltAct As String) As Integer
    On Error GoTo InsertaCuentaIFErr
    Dim sql As String
    InsertaCuentaIF = 1
    
    sql = "INSERT INTO CTAIF( cPersCod, cIFTpo, cCtaIFCod, cCtaIFDesc, " _
        & "                dCtaIFAper, dCtaIFCap, nCtaIFPlazo, cCtaIFEstado, cUltimaActualizacion , dCtaIfInt ) " _
        & " VALUES('" & psPersCod & "','" & psIFTpo & "','" & psCtaIfCod & "','" & psCtaIFDesc & "','" _
        & Format(pdCtaIFAper, vsFormatoFecha) & "'," & IIf(psCtaIFCap = "", "Null", "'" & Format(psCtaIFCap, vsFormatoFecha) & "'") & "," _
        & pnCtaIFPlazo & "," & pnCtaIFEstado & ",'" & psUltAct & "'," & IIf(pdCtaIFInt = "", "Null", "'" & Format(pdCtaIFInt, vsFormatoFecha) & "'") & ")"
    
    oConect.Ejecutar sql
    InsertaCuentaIF = 0
    Exit Function
InsertaCuentaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaCuentaIF Method")
End Function
Public Function InsertaCuentaIFInteres(ByVal psPersCod As String, ByVal psIFTpo As String, _
                                ByVal psCtaIfCod As String, ByVal pdFechaInt As Date, ByVal pnPeriodo As Integer, _
                                ByVal pnInteres As Currency, ByVal psMovNro As String) As Integer
    On Error GoTo InsertaCuentaIFInteresErr
    Dim sql As String
    InsertaCuentaIFInteres = 1
    
    sql = "INSERT INTO CtaIFInteres( cPersCod, cIFTpo, cCtaIFCod, dCtaIFIntRegistro , nCtaIFIntPeriodo, nCtaIFIntValor, cUltimaActualizacion  ) " _
        & " VALUES('" & psPersCod & "','" & psIFTpo & "','" & psCtaIfCod & "','" _
        & Format(pdFechaInt, vsFormatoFecha) & "'," & pnPeriodo & "," & pnInteres & ",'" & psMovNro & "')"
    
    oConect.Ejecutar sql
    InsertaCuentaIFInteres = 0
    Exit Function
InsertaCuentaIFInteresErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaCuentaIFInteres Method")
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
        & Format(pdCtaIFIntRegistro, vsFormatoFecha) & "'," & pnCtaIFIntPeriodo & "," _
        & pnCtaIFIntValor & ",'" & psUltAct & "')"
    
    oConect.Ejecutar sql
    InsertaInteresCtaIF = 0
    Exit Function
InsertaInteresCtaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaInteresCtaIF Method")
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
    Call RaiseError(MyUnhandledError, "DMov:InsertaCuentaIFFiltro Method")
End Function
Public Function InsertaMovCtaIF(ByVal pnMovNro As Long, ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
                                 ByVal psCtaIfCod As String, ByVal pnConCod As CGCtaIFConceptos, ByVal pnImporte As Currency, Optional pnDiasTrans As Integer = 0) As Integer
                                    
    On Error GoTo InsertaMovCtaIFErr
    Dim sql As String
    InsertaMovCtaIF = 1
    
    sql = "INSERT INTO MovCtaIF(nMovNro, cPersCod, cIFTpo, cCtaIFCod, nConceptoCod, nMovImporte,nDiasTrans) " _
        & "     VALUES(" & pnMovNro & ",'" & psPersCod & "','" & Format(pnIFTpo, "00") & "','" _
        & psCtaIfCod & "'," & pnConCod & "," & pnImporte & "," & pnDiasTrans & ")"
    
    oConect.Ejecutar sql
    InsertaMovCtaIF = 0
    Exit Function
InsertaMovCtaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCtaIF Method")
End Function
Public Function ExtornaMovCtaIF(ByVal pnMovNro As Long) As Integer
Dim sql As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

sql = " SELECT  M.nMovNro,CONVERT(DATETIME , SUBSTRING(M.CMOVNRO,1,8)) AS FECHA,    " _
    & "         MC.cPersCod , MC.cIFTpo, MC.cCtaIFCod " _
    & " FROM MOV M JOIN MOVCTAIF  MC ON MC.NMOVNRO =M.NMOVNRO  " _
    & " WHERE M.NMOVNRO = " & pnMovNro & ""

Set rs = oConect.CargaRecordSet(sql)
If Not rs.EOF And Not rs.BOF Then
    Do While Not rs.EOF
        ActualizaSaldoCtaIF rs!fecha, rs!cPersCod, rs!cIFTpo, rs!cCtaIfCod
        rs.MoveNext
    Loop
End If
rs.Close
Set rs = Nothing
End Function
Public Function ActualizaSaldoCtaIF(ByVal pdFecha As Date, _
                                    ByVal psPersCod As String, ByVal pnIFTpo As CGTipoIF, _
                                    ByVal psCtaIfCod As String) As Integer
                                    
    On Error GoTo ActualizaSaldoCtaIFErr
    Dim sql As String
    ActualizaSaldoCtaIF = 1
    
    sql = "sp_ActualizaSaldoCtaIF '" & Format(pdFecha, gsFormatoFecha) & "','" _
                                    & psPersCod & "','" _
                                    & Format(pnIFTpo, "00") & "', '" & psCtaIfCod & "'"
    
    oConect.Ejecutar sql
    ActualizaSaldoCtaIF = 0
    Exit Function
ActualizaSaldoCtaIFErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaSaldoCtaIF Method")
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
        lsFiltro = lsFiltro + " dCtaIFAper= '" & Format(pdCtaIFAper, vsFormatoFecha) & "',"
    End If
    If pdCtaIFCap <> "" Then
        lsFiltro = lsFiltro + " dCtaIFCap= '" & Format(pdCtaIFCap, vsFormatoFecha) & "',"
    End If
    If pdCtaIFInt <> "" Then
        lsFiltro = lsFiltro + " dCtaIFInt= '" & Format(pdCtaIFInt, vsFormatoFecha) & "',"
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
    Call RaiseError(MyUnhandledError, "DMov:ActualizaCuentaIF Method")
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

Public Sub InsertaMovRefAnt(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long)
    Dim oCta As New DCtaCont
    Dim prs  As ADODB.Recordset
    Dim sql As String
    
    sql = " Insert MovRef (nMovNro,nMovNroRef) " _
        & " Select " & pnMovNro & ",nMovNroRef From MovRef Where nMovNro = " & pnMovNroRef
        
    oConect.Ejecutar sql
End Sub

Public Sub ModificaFlagDocumento(ByVal psDocNro As String, ByVal pnDocTpo As Long, ByVal pnMovFlag As Long)
    Dim sql As String
    
    sql = " Update Mov" _
        & " Set nMovFlag = " & pnMovFlag & "" _
        & " Where nMovNro In (Select nMovnro From Movdoc Where cDocNro like '" & psDocNro & "' and ndoctpo = " & pnDocTpo & ") And nMovFlag = " & gMovFlagVigente
    oConect.Ejecutar sql
End Sub

Public Function EliminaActivoFijo(psBSCod As String, psSerie As String, psPeriodo As String) As Boolean
    On Error GoTo EliminaAFErr
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    sql = "Delete BSActivoFijo Where nAnio = '" & psPeriodo & "' And cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
    oCon.Ejecutar sql
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
EliminaAFErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:AgregaAsistencia Method")
End Function

Public Function GetCtaDep(psBSCod As String, Optional bIntangible As Boolean = False) As String
    On Error GoTo EliminaAFErr
    Dim sql As String
    Dim oCon As DConecta
    Set oCon = New DConecta
    
    oCon.AbreConexion
    If Not bIntangible Then
        sql = " Select Top 1 cCtaContCodD from CtaBS CBS" _
            & " Inner Join OpeCtaCta OCC On CBS.cCtaContCod = OCC.cCtaContCodH And OCC.cOpeCod = '591201'" _
            & " Where '" & psBSCod & "' Like cObjetoCod  + '%' And CBS.cOpeCod = '" & gnAlmaSalXAtencion & "'"
    Else
        sql = " Select '" & psBSCod & "' "
    End If
            
    GetCtaDep = oCon.CargaRecordSet(sql).Fields(0)
    
    oCon.CierraConexion
    Set oCon = Nothing
    Exit Function
EliminaAFErr:
    Call RaiseError(MyUnhandledError, "DActualizaControlAsistencia:AgregaAsistencia Method")
End Function


Public Function GeneraAsientoRes(pnMovNro As Long, pnMovNroBase As Long)
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    sql = "Select Sum(nMovImporte) Monto, cCtaContCod from movcta where nmovnro = " & pnMovNroBase & " group by cCtaContCod"
    Set rs = oConect.CargaRecordSet(sql)
        
    While Not rs.EOF
        InsertaMovCta pnMovNro, rs.Bookmark, rs!cCtaContCod, rs!Monto
        rs.MoveNext
    Wend
    
    rs.Close
End Function

Public Sub IniActivoFijo(psAnioNuevo As String, psAnioAnterior As String, pdFechaProceso As Date, pnTipo As Integer)
    Dim sql As String
    
    'Sql = " Insert BSActivoFijo (nAnio, nMovNro, cBSCod, cSerie, nBSValor, nIGV, nBSDeprecia, bDepreciado, dActivacion, cAreCod, cAgeCod, nBSPerDeprecia, nBSPerDepAct, cDescripcion, bDeprecia, bIntangible, bForsado, cNroID, cNroPatrimonio, dCompra, dBase, Ban, bBaja, cPersCod, dBaja)" _
        & " Select " & psAnioNuevo & " nAnio, nMovNro,cBSCod,cSerie, round(dbo.GetFactorIPM (dActivacion, '" & Format(pdFechaProceso, gsFormatoFecha) & "') * nBSValor,2) nBSValor ,nIGV,nBSDeprecia,bDepreciado, '" & Format(pdFechaProceso, gsFormatoFecha) & "' dActivacion,cAreCod,cAgeCod,nBSPerDeprecia," _
        & " Case when nBSPerDepAct + datediff(month,dbase,'" & Format(pdFechaProceso, gsFormatoFecha) & "') >= nBSPerDeprecia then nBSPerDeprecia else nBSPerDepAct + datediff(month,dbase,'" & Format(pdFechaProceso, gsFormatoFecha) & "') end," _
        & " cDescripcion,bDeprecia,bIntangible,bForsado,cNroID,cNroPatrimonio,dCompra, '" & Format(pdFechaProceso, gsFormatoFecha) & "' dBase,Ban,bBaja,cPersCod,dBaja" _
        & " From BSActivoFijo where ban = " & pnTipo & "  And nAnio = " & psAnioAnterior & ""
    sql = " Insert BSActivoFijo (nAnio, nMovNro, cBSCod, cSerie, nBSValor, nIGV, nBSDeprecia, bDepreciado, dActivacion, cAreCod, cAgeCod, nBSPerDeprecia, nBSPerDepAct, cDescripcion, bDeprecia, bIntangible, bForsado, cNroID, cNroPatrimonio, dCompra, dBase, Ban, bBaja, cPersCod, dBaja)" _
        & " Select " & psAnioNuevo & " nAnio, nMovNro,cBSCod,cSerie, round(nBSValor,2) nBSValor ,nIGV,nBSDeprecia,bDepreciado, '" & Format(pdFechaProceso, gsFormatoFecha) & "' dActivacion,cAreCod,cAgeCod,nBSPerDeprecia," _
        & " Case when nBSPerDepAct + datediff(month,dbase,'" & Format(pdFechaProceso, gsFormatoFecha) & "') >= nBSPerDeprecia then nBSPerDeprecia else nBSPerDepAct + datediff(month,dbase,'" & Format(pdFechaProceso, gsFormatoFecha) & "') end," _
        & " cDescripcion,bDeprecia,bIntangible,bForsado,cNroID,cNroPatrimonio,dCompra, '" & Format(pdFechaProceso, gsFormatoFecha) & "' dBase,Ban,bBaja,cPersCod,dBaja" _
        & " From BSActivoFijo where ban = " & pnTipo & "  And nAnio = " & psAnioAnterior & ""
    oConect.Ejecutar sql

End Sub


Public Function GeneraDocNro(psDocTpo As TpoDoc, psMoneda As Moneda, Optional psDocSerie As String = "", Optional psTipo As String = "") As String
    'NAGL 20191212 Agregó Optional psTipo As String
    On Error GoTo GeneraMovNroErr
    Dim sSQL As String
    Dim rs As New ADODB.Recordset
    Dim oCaja As nCajaGeneral
    Set oCaja = New nCajaGeneral
      
    'If psDocSerie <> "" Then
'ALPA 20091112***********************************************************
'        sSql = "SELECT max(cDocNro) AS cDocCorrela FROM Movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
'            & "WHERE   m.nMovFlag NOT IN ('" & gMovFlagEliminado & "') and nDocTpo = '" & Format(psDocTpo, "00") & "' and substring(cDocNro,1," & Len(psDocSerie) & ") = '" & psDocSerie & "'"
        'sSQL = "SELECT substring(cDocNro,1,4) + '-' + convert(varchar(8),max(convert(int,right(cDocNro,7)))) AS cDocCorrela FROM Movdoc md JOIN Mov m ON m.nMovNro = md.nMovNro " _
            '& "WHERE   m.nMovFlag NOT IN ('" & gMovFlagEliminado & "') and nDocTpo = '" & Format(psDocTpo, "00") & "' and substring(cDocNro,1," & Len(psDocSerie) & ") = '" & psDocSerie & "' group by substring(cDocNro,1,4)"
'************************************************************************
    'Else
        'sSQL = "SELECT  cDocNro AS cDocCorrela , md.nMovnro " _
            '& " FROM    movdoc md " _
            '& " WHERE   nDocTpo = '" & Format(psDocTpo, "00") & "' " _
            '& "         and md.nmovnro = (  Select Max(MD1.nmovnro)  " _
            '& "                             From MovDoc MD1 JOIN Mov M on M.nMovNro= MD1.nMovnro " _
            '& "                             WHERE MD1.nDocTpo = '" & Format(psDocTpo, "00") & "' and Substring(M.cOpeCod,3,1) ='" & psMoneda & "') "
            
    'End If
    '***Comentado by NAGL 20191212
    
    '*******NAGL Según INC1912120012************
    sSQL = "Exec stp_sel_ObtieneGeneraDocNro '" & Format(psDocTpo, "00") & "', '" & psMoneda & "', '" & psDocSerie & "', '" & psTipo & "'"
    '*******Agregado by NAGL 20191212***********
    
    Set rs = oConect.CargaRecordSet(sSQL)
    If Not rs.EOF And Not rs.BOF Then
        If Not IsNull(rs!cDocCorrela) Then
            If psDocSerie <> "" Then
               GeneraDocNro = psDocSerie & "-" & Format(Val(Mid(rs!cDocCorrela, Len(psDocSerie) + 2, 20)) + 1, String(8, "0"))
            Else
               If Mid(rs!cDocCorrela, 4, 1) = "-" Then
                  GeneraDocNro = Mid(rs!cDocCorrela, 1, 3) & "-" & Format(Val(Mid(rs!cDocCorrela, 5, 20)) + 1, String(8, "0"))
               Else
                  GeneraDocNro = Format(Val(rs!cDocCorrela) + 1, String(8, "0"))
               End If
            End If
        Else
            GeneraDocNro = "00000001"
            If psDocSerie <> "" Then GeneraDocNro = psDocSerie & "-" & GeneraDocNro
        End If
    Else
        GeneraDocNro = "00000001"
        If psDocSerie <> "" Then GeneraDocNro = psDocSerie & "-" & GeneraDocNro
    End If
    rs.Close
    Set rs = Nothing
    If psDocTpo = TpoDocOrdenPago Then
        GeneraDocNro = oCaja.VerificaOPRango(GeneraDocNro, psMoneda)
    End If
    Set oCaja = Nothing
    
    Exit Function
GeneraMovNroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GeneraMovNro Method")
End Function

Public Sub InsertaAjusteInflacion(pcCtaContCod As String, pcAjusteCod As String, pcAjusteDescrip As String, pdAjusteFecha As Date, pnAjusteValor1 As Currency, pnAjusteValor2 As Currency, pnAjusteValor3 As Currency)
    Dim sql As String
    
    sql = " Insert AjusteInflacion (cCtaContCod,cAjusteCod,cAjusteDescrip,dAjusteFecha,nAjusteValor1,nAjusteValor2,nAjusteValor3)" _
        & " Values('" & pcCtaContCod & "','" & pcAjusteCod & "','" & pcAjusteDescrip & "','" & Format(pdAjusteFecha, gsFormatoFecha) & "'," & pnAjusteValor1 & "," & pnAjusteValor2 & "," & pnAjusteValor3 & ")"
    oConect.Ejecutar sql
End Sub

Public Function GetCodAjusteInflacion(pcCtaContCod As String) As String
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = " Select Top 1 cAjusteCod from AjusteInflacion where cCtaContCod = '" & pcCtaContCod & "' order by cAjusteCod desc"
    Set rs = oConect.CargaRecordSet(sql)
    
    If rs.EOF And rs.BOF Then
        GetCodAjusteInflacion = Format(1, "00")
    Else
        GetCodAjusteInflacion = Format(CInt(rs!cAjusteCod) + 1, "00")
    End If
        
    rs.Close
    Set rs = Nothing
End Function

Public Function ActualizaActivo(ByVal psBSCod As String, ByVal psSerie As String, ByVal psSerieMod As String, _
                                ByVal psDescripcion As String, ByVal pnYear As Double, _
                                Optional ByVal pnMontoIni As Double, Optional ByVal psFecha As String, Optional ByVal psMarca As String, Optional ByVal psModelo As String, Optional ByVal psSeriePlaca As String) As Integer
       '*** PEAC 20120327 se agrego monto, fecha , marcca, modelo, serie
    
    On Error GoTo ActualizaActivoErr
    Dim sql As String
    ActualizaActivo = 1
     
'    sql = " Update BSActivoFijo" _
        & " Set cSerie = '" & psSerieMod & "', cDescripcion = '" & psDescripcion & "'," _
        & " nAnio = " & pnYear & " " _
        & " Where cBSCod = '" & psBSCod & "' And cSerie = '" & psSerie & "'"
        
    '*** PEAC 20101129
    sql = " exec stp_upd_ActualizaActivo '" & psSerieMod & "','" & psDescripcion & "'," & pnYear & ",'" & psBSCod & "','" & psSerie & "'," & pnMontoIni & ",'" & psFecha & "','" & psMarca & "','" & psModelo & "','" & psSeriePlaca & "'"
    
    oConect.Ejecutar sql
    ActualizaActivo = 0
    Exit Function
ActualizaActivoErr:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaActivo Method")
End Function

Public Function InsertaKardexAlmacen(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal psBSCod As String, ByVal psNroDoc As String, ByVal pnCantidad As Currency, ByVal pnPreUni As Currency, ByVal pnStock As Currency, ByVal pnCosPro As Currency, ByVal pnCostoTotal As Currency, ByVal pnCodAlm As Integer, ByVal psTipMov As String, ByVal psMovNro As String) As Integer
    On Error GoTo InsertaKardexAlmacenErr
    Dim sql As String
    InsertaKardexAlmacen = 1
    
    sql = "INSERT INTO KardexAlmacen (nMovNro, nMovItem, cBSCod, cNroDoc, nCantidad, nPreuni, nStock, nCosProm, nCostoTotal, nAlmCod, cTipMov, cUltimaActualizacion) " _
        & "VALUES (" & pnMovNro & "," & pnMovItem & ",'" & psBSCod & "','" & psNroDoc & "'," & pnCantidad & "," & pnPreUni & "," & pnStock & "," & pnCosPro & "," & pnCostoTotal & "," & pnCodAlm & ",'" & psTipMov & "','" & psMovNro & "')"
     
    oConect.Ejecutar sql
    InsertaKardexAlmacen = 0
    Exit Function
InsertaKardexAlmacenErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaKardexAlmacen Method")
End Function
Public Function InsertaMasterAlmacen(ByVal psBSCod As String, ByVal pnCantidad As Currency, ByVal pnPreUni As Currency, ByVal pnCosProm As Currency, ByVal pnCostoTotal As Currency, ByVal psUniMed As String, ByVal pnAlmCod As Integer, ByVal psMoneda As String, ByVal psMovNro As String, ByVal pdUltimaCom As Date, ByVal psTipMov As String) As Integer
    On Error GoTo InsertaMasterAlmacenErr
    Dim sql As String
    InsertaMasterAlmacen = 1
    
    sql = "INSERT INTO MaestroAlmacen (cBSCod, nPreUni, nStock, nPreProm, nCostoTotal, cUniMed, nAlmCod, cMoneda, dUltimaCom, cUltimaActualizacion) " _
        & "VALUES ('" & psBSCod & "'," & pnPreUni & "," & pnCantidad & "," & pnCosProm & "," & pnCostoTotal & ",'" & psUniMed & "'," & pnAlmCod & ",'" & psMoneda & "','" & Format(pdUltimaCom, gsFormatoFecha) & "','" & psMovNro & "')"
     
    oConect.Ejecutar sql
    InsertaMasterAlmacen = 0
    Exit Function
InsertaMasterAlmacenErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaKardexAlmacen Method")
End Function
Public Function ActualizaMasterAlmacen(ByVal psBSCod As String, ByVal pnStock As Currency, ByVal pnPreUni As Currency, ByVal pnCosProm As Currency, ByVal pnCostoTotal As Currency, ByVal pnAlmCod As Integer, ByVal pdFecha As Date, ByVal pnTipoMov As String, ByVal psMovNro As String) As Integer
    On Error GoTo ActualizaMasterAlmacenErr
    Dim sql As String
    Dim sqlif As String
    
    ActualizaMasterAlmacen = 1
    
    If pnTipoMov = "I" Then
        sqlif = " dUltimaCom = '" & Format(pdFecha, gsFormatoFecha)
    ElseIf pnTipoMov = "S" Then
        sqlif = " dUltimaSal = '" & Format(pdFecha, gsFormatoFecha)
    Else
        sqlif = " dUltimaCom = '" & Format(pdFecha, gsFormatoFecha)
    End If
    
    sql = "UPDATE MaestroAlmacen Set " _
        & "nPreUni = " & pnPreUni & ", nStock = " & pnStock & ", nPreProm = " & pnCosProm & ", " _
        & "nCostoTotal = " & pnCostoTotal & ", " _
        & sqlif & "', cUltimaActualizacion = '" & psMovNro & "'" _
        & " WHERE cBSCod = '" & psBSCod & "' And nAlmCod = " & pnAlmCod & ""
     
    oConect.Ejecutar sql
    ActualizaMasterAlmacen = 0
    Exit Function
ActualizaMasterAlmacenErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaKardexAlmacen Method")
End Function
Public Function EliminaMasterAlmacen(ByVal psBSCod As String, ByVal pnAlmCod As Integer) As Integer
    On Error GoTo InsertaMovObjErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    EliminaMasterAlmacen = 1
    
    sql = " Delete MaestroAlmacen Where cBSCod = '" & psBSCod & "'" _
        & " and nAlmCod = " & pnAlmCod
    Set rs = oConect.CargaRecordSet(sql)
    
    oConect.Ejecutar sql
    
    EliminaMasterAlmacen = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
End Function
Public Function EliminaKardexAlmacen(ByVal psBSCod As String, ByVal pnMovNro As Long, ByVal pnAlmCod As Integer) As Integer
    On Error GoTo InsertaMovObjErr
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    EliminaKardexAlmacen = 1
    
    sql = " Delete KardexAlmacen Where cBSCod = '" & psBSCod & "'" _
        & " and nMovnro = " & pnMovNro & " and nAlmCod = " & pnAlmCod
    Set rs = oConect.CargaRecordSet(sql)
    
    oConect.Ejecutar sql
    
    EliminaKardexAlmacen = 0
    Exit Function
InsertaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovObj Method")
End Function

'*** PEAC 20120413
Public Function BuscaCtaPlantillaAF(ByVal psOpeCod As String, ByVal psCta As String, ByVal psCtaDH As String) As Recordset
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sql = "exec stp_sel_BuscaCtaPlantillaAF '" & psOpeCod & "','" & psCta & "','" & psCtaDH & "'"
    
    Set rs = oConect.CargaRecordSet(sql)
    Set BuscaCtaPlantillaAF = rs

End Function

'*** PEAC 20120511
Public Sub ModificaActivosFijos(ByVal psMovNro As String, ByVal psOpeCod As String)
Dim sql As String
   On Error GoTo ModificaActivosFijosErr

    sql = "exec stp_upd_ModificaActivosFijos '" & psMovNro & "','" & psOpeCod & "' "
    oConect.Ejecutar sql

Exit Sub

ModificaActivosFijosErr:
    Call RaiseError(MyUnhandledError, "DMov:ModificaActivosFijos Method")
End Sub

'*** PEAC 20120611
Public Function InsertaMovDepreAFTrib(ByVal pnMovNro As Long, ByVal pnMovItem As Long, ByVal pnMovImporte As Currency) As Integer
    On Error GoTo InsertaMovDepreAFTribErr
    Dim rsCta As ADODB.Recordset
    Set rsCta = New ADODB.Recordset
    Dim sql As String
    InsertaMovDepreAFTrib = 1
                    
            sql = "exec stp_ins_InsertaMovDepreAFTrib " & pnMovNro & "," & pnMovItem & "," & pnMovImporte

            oConect.Ejecutar sql
            InsertaMovDepreAFTrib = 0
            
    Exit Function

InsertaMovDepreAFTribErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovDepreAFTrib Method")
End Function
'EJVG20130531 ***
Public Sub InsertaInventarioAF(ByVal sCodInventario As String, ByVal sNombre As String, ByVal sAgencia As String, ByVal sMarca As String, ByVal sModelo As String, ByVal sSerie As String, ByVal sFIngreso As String, ByVal sMovNro As String, ByVal sBSCod As String, ByVal iMovItem As Integer, ByVal pnDocOrigen As Integer)
    Dim lsSQL As String
    On Error GoTo ErrInsertaInventarioAF
    lsSQL = "Exec stp_ins_ERS0592013_ActivoFijo '" & sCodInventario & "', '" & sNombre & "', '" & sAgencia & "', '" & sMarca & "', '" & sModelo & "', '" & sSerie & "', '" & sFIngreso & "', '" & sMovNro & "', '" & sBSCod & "', " & iMovItem & "," & pnDocOrigen
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaInventarioAF:
    Call RaiseError(MyUnhandledError, "DMov:InsertaInventarioAF Method")
End Sub
'END EJVG ***
'EJVG20131021 ***
Public Sub InsertaMovActaConformidad(ByVal pnMovNro As Long, ByVal psAreaAgeCod As String, ByVal psSubAreaDesc As String, ByVal pnMoneda As Moneda, ByVal pnDocOrigenTpo As Integer, ByVal psDocReferencia As String, ByVal psDescripcion As String, ByVal psObservacion As String, Optional ByVal pnMovNroRef As Long = 0)
    Dim lsSQL As String
    On Error GoTo ErrInsertaMovActaConformidad
    lsSQL = "Exec stp_ins_ERS0622013_MovActaConformidad " & pnMovNro & ",'" & psAreaAgeCod & "','" & psSubAreaDesc & "'," & pnMoneda & "," & pnDocOrigenTpo & ",'" & psDocReferencia & "','" & psDescripcion & "','" & psObservacion & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovActaConformidad:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovActaConformidad Method")
End Sub
Public Sub InsertaMovActaConformidadOrden(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long)
    Dim lsSQL As String
    On Error GoTo ErrInsertaMovActaConformidad
    lsSQL = "Exec stp_ins_ERS0622013_MovActaConformidadOrden " & pnMovNro & "," & pnMovNroRef
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovActaConformidad:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovActaConformidad Method")
End Sub
Public Sub InsertaMovActaConformidadOrdenDet(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal pnMovItemRef As Integer)
    Dim lsSQL As String
    On Error GoTo ErrInsertaMovActaConformidad
    lsSQL = "Exec stp_ins_ERS0622013_MovActaConformidadOrdenDet " & pnMovNro & "," & pnMovItem & "," & pnMovItemRef
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovActaConformidad:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovActaConformidad Method")
End Sub
Public Function GetCorrelativoActaConformidad(ByVal pnDocTpo As Integer, ByVal psAgeCod As String, ByVal psAnio As String) As String
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrGetCorrelativoActaConformidad
    lsSQL = "Exec stp_sel_ERS0622013_CorrelativoActaConformidad " & pnDocTpo & ",'" & psAgeCod & "','" & psAnio & "'"
    Set rs = oConect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        GetCorrelativoActaConformidad = rs!cCorrelativo
    End If
    Exit Function
ErrGetCorrelativoActaConformidad:
    Err.Raise Err.Number, "Error al recuperar correlativo Acta de Conformidad", Err.Description
End Function
Public Sub ActualizaSaldoMovimiento(sMovNro As String, sSimbolo As String)
    On Error GoTo ErrActualizaMovSaldo
    Dim sql As String
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
        sql = "Exec spActualizaMovSaldo " & lsMov & ", '" & Format(ldFecha, gsFormatoFecha) & "', '" & sSimbolo & "' "
        oConect.Ejecutar sql
        
        sql = "DELETE tmpMovActualiza"
        oConect.Ejecutar sql
    Else
        sql = "INSERT tmpMovActualiza (cMovNro) VALUES ('" & sMovNro & "')"
        oConect.Ejecutar sql
        
        ldFecha = GetFechaMov(Left(sMovNro, 8), True)
        sql = "spActualizaMovSaldo '" & sMovNro & "', '" & Format(ldFecha, gsFormatoFecha) & "', '" & sSimbolo & "' "
        oConect.Ejecutar sql
        
        sql = "DELETE tmpMovActualiza"
        oConect.Ejecutar sql
    End If
    Exit Sub
ErrActualizaMovSaldo:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub InsertaMovActaConformidadItem(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal psDescripcion As String)
    Dim lsSQL As String
    On Error GoTo ErrInsertaMovActaConformidadItem
    lsSQL = "Exec stp_ins_ERS0622013_MovActaConformidadItem " & pnMovNro & "," & pnMovItem & ",'" & psDescripcion & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovActaConformidadItem:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovActaConformidadItem Method")
End Sub
Public Sub InsertaMovComprobante(ByVal pnMovNro As Long, ByVal pnMoneda As Moneda, ByVal pnDocTpo As Integer, ByVal psDocNro As String, ByVal pdDocFecha As Date, ByVal pnTpoPago As Integer, ByVal psIFICod As String, ByVal psIFICtaCod As String)
    Dim lsSQL As String
    On Error GoTo ErrInsertaMovComprobante
    lsSQL = "Exec stp_ins_ERS0622013_MovComprobante " & pnMovNro & "," & pnMoneda & "," & pnDocTpo & ",'" & psDocNro & "','" & Format(pdDocFecha, "yyyymmdd") & "'," & pnTpoPago & ",'" & psIFICod & "','" & psIFICtaCod & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovComprobante:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovComprobante Method")
End Sub
Public Sub InsertaMovActaConformidadContrato(ByVal pnMovNro As Long, ByVal psncontrato As String)
    Dim lsSQL As String
    On Error GoTo ErrInsertaMovActaConformidadContrato
    lsSQL = "Exec stp_ins_ERS0622013_MovActaConformidadContrato " & pnMovNro & ",'" & psncontrato & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovActaConformidadContrato:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovActaConformidadContrato Method")
End Sub
Public Sub InsertaMovActaConformidadContratoDet(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal pnnpago As Integer)
    Dim lsSQL As String
    On Error GoTo ErrInsertaMovActaConformidadContratoDet
    lsSQL = "Exec stp_ins_ERS0622013_MovActaConformidadContratoDet " & pnMovNro & "," & pnMovItem & "," & pnnpago
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovActaConformidadContratoDet:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovActaConformidadContratoDet Method")
End Sub
Public Sub ActualizaEstadoCuotaCronograma(ByVal psncontrato As String, ByVal psNContRef As Integer, ByVal pnnpago As Integer, ByVal pnEstado As Integer)
    Dim lsSQL As String
    On Error GoTo ErrActualizaEstadoCuotaCronograma
    lsSQL = "Exec stp_ins_ERS0622013_ActualizaEstadoCuotaContrato '" & psncontrato & "'," _
                                                                    & psNContRef & "," _
                                                                    & pnnpago & "," _
                                                                    & pnEstado 'pnContRef Agregado PASI20140917 ERS0772014
    oConect.Ejecutar lsSQL
    Exit Sub
ErrActualizaEstadoCuotaCronograma:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotaCronograma Method")
End Sub
Public Sub ActualizaEstadoCuotasCronogramaxExtActaConformidad(ByVal pnMovNroAC As Long)
    Dim lsSQL As String
    On Error GoTo ErrActualizaEstadoCuotasCronogramaxExtActaConformidad
    lsSQL = "Exec stp_upd_ERS0622013_ActualizaEstadoCuotasCronogramaxExtActaConformidad " & pnMovNroAC
    oConect.Ejecutar lsSQL
    Exit Sub
ErrActualizaEstadoCuotasCronogramaxExtActaConformidad:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotasCronogramaxExtActaConformidad Method")
End Sub
Public Sub ActualizaEstadoCuotasCronogramaxRegComprobante(ByVal pnMovNroAC As Long, Optional ByVal pbExtorno As Boolean = False)
    Dim lsSQL As String
    On Error GoTo ErrActualizaEstadoCuotasCronogramaxRegComprobante
    lsSQL = "Exec stp_upd_ERS0622013_ActualizaEstadoCuotasCronogramaxRegComprobante " & pnMovNroAC & "," & IIf(pbExtorno, 1, 0)
    oConect.Ejecutar lsSQL
    Exit Sub
ErrActualizaEstadoCuotasCronogramaxRegComprobante:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoCuotasCronogramaxRegComprobante Method")
End Sub
'END EJVG *******
'EJVG20140319 ***
Public Sub InsertaMovAgencia(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, Optional psAgeCod As String)
    On Error GoTo ErrInsertaMovAgencia
    Dim lsSQL As String
    lsSQL = "Exec stp_ins_MovAgencia " & pnMovNro & "," & pnMovItem & ",'" & Format(psAgeCod, "00") & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovAgencia:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovAgencia Method")
End Sub
Public Function EsUnicaCtaCont(ByVal psBSCod As String) As Boolean
    On Error GoTo ErrEsUnicaCtaCont
    Dim rs As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "SELECT dbo.fnc_EsCtaContUnicaLog('" & psBSCod & "') bUnica"
    Set rs = oConect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        EsUnicaCtaCont = rs!bUnica
    End If
    Set rs = Nothing
    Exit Function
ErrEsUnicaCtaCont:
    Call RaiseError(MyUnhandledError, "DMov:EsUnicaCtaCont Method")
End Function
'END EJVG *******
'PASI20140818 ERS0772014
Public Sub InsertaMovComprobanteReg(ByVal pnMovNro As Long, _
                                    ByVal psAreaAgeCod As String, _
                                    ByVal pnMoneda As Integer, _
                                    ByVal pnDocOrigenTpo As Integer, _
                                    ByVal psObservacion As String, _
                                    ByVal pnTpoPago As Integer, _
                                    ByVal psIFICod As String, _
                                    ByVal psIFICtaCod As String, _
                                    ByVal pnDocTpo As TpoDoc, _
                                    ByVal psDocNro As String, _
                                    ByVal psDocFecha As String)
 On Error GoTo ERRORInsertaMovComprobanteReg
 sSQL = "stp_ins_ERS0772014_MovComprobanteReg " & pnMovNro & "," _
                                            & "'" & psAreaAgeCod & "'," _
                                            & pnMoneda & "," _
                                            & pnDocOrigenTpo & "," _
                                            & "'" & psObservacion & "'," _
                                            & pnTpoPago & "," _
                                            & "'" & psIFICod & "'," _
                                            & "'" & psIFICtaCod & "'," _
                                            & pnDocTpo & "," _
                                            & "'" & psDocNro & "'," _
                                            & "'" & psDocFecha & "'"
    oConect.Ejecutar sSQL
    Exit Sub
ERRORInsertaMovComprobanteReg:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovCompobanteReg Method")
End Sub
Public Sub InsertaMovComprobanteOrden(ByVal pnMovNro As Long, ByVal pnMovNroRef As Long)
    On Error GoTo ERRORInsertaMovComprobanteOrden
    sSQL = "stp_ins_ERS0772014_MovComprobanteOrden " & pnMovNro & "," _
                                & pnMovNroRef
    oConect.Ejecutar sSQL
    Exit Sub
ERRORInsertaMovComprobanteOrden:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovComprobanteOrden Method")
End Sub
Public Sub InsertaMovComprobanteOrdenDet(ByVal pnMovNro As Long, _
                                            ByVal pnMovItem As Integer, _
                                            ByVal pnMovItemRef As Integer)
    On Error GoTo ERRORInsertaMovComprobanteOrdenDet
    sSQL = "stp_ins_ERS0772014_MovComprobanteOrdenDet " & pnMovNro & "," _
                                                        & pnMovItem & "," _
                                                        & pnMovItemRef
    oConect.Ejecutar sSQL
    Exit Sub
ERRORInsertaMovComprobanteOrdenDet:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovComprobanteOrdenDet Method")
End Sub
Public Sub InsertaMovComprobanteItem(ByVal pnMovNro As Long, _
                                        ByVal pnMovItem As Long, _
                                        ByVal psDescripcion As String)
    On Error GoTo ERRORInsertaMovComprobanteItem
    sSQL = "stp_ins_ERS0772014_MovComprobanteItem " & pnMovNro & "," _
                                                    & pnMovItem & "," _
                                                    & "'" & psDescripcion & "'"
    oConect.Ejecutar sSQL
    Exit Sub
ERRORInsertaMovComprobanteItem:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovComprobanteItem Method")
End Sub
Public Sub InsertaMovComprobanteContrato(ByVal pnMovNro As Long, ByVal psncontrato As String, ByVal pncontref As Integer)
    On Error GoTo ERRORInsertaMovComprobanteContrato
    sSQL = "stp_ins_ERS0772014_MovComprobanteContrato " & pnMovNro & "," _
                                                        & "'" & psncontrato & "'," _
                                                        & pncontref
    oConect.Ejecutar sSQL
    Exit Sub
ERRORInsertaMovComprobanteContrato:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovComprobanteContrato Method")
End Sub
Public Sub InsertaMovComprobanteContratoCronDet(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal pnPago As Integer)
    On Error GoTo ERRORInsertaMovComprobanteContratoCronDet
    sSQL = "stp_ins_ERS0772014_MovComprobanteContratoCronDet " & pnMovNro & "," _
                                                            & pnMovItem & "," _
                                                            & pnPago
    oConect.Ejecutar sSQL
    Exit Sub
ERRORInsertaMovComprobanteContratoCronDet:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovComprobanteContratoCronDet Method")
End Sub
Public Sub InsertaMovComprobanteContratoItemDet(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal psBSCod As String)
    On Error GoTo ERRORInsertaMovComprobanteContratoItemDet
    sSQL = "stp_ins_ERS0772014_MovComprobanteContratoItemDet " & pnMovNro & "," _
                                                            & pnMovItem & "," _
                                                            & "'" & psBSCod & "'"
    oConect.Ejecutar sSQL
    Exit Sub
ERRORInsertaMovComprobanteContratoItemDet:
        Call RaiseError(MyUnhandledError, "DMov:InsertaMovComprobanteContratoItemDet Method")
End Sub
Public Sub InsertaMovComprobanteContratoBienDet(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal pnMovItemRef As String)
    On Error GoTo ERRORInsertaMovComprobanteContratoItemDet
    sSQL = "stp_ins_ERS0772014_MovComprobanteContratoBienDet " & pnMovNro & "," _
                                                            & pnMovItem & "," _
                                                            & "'" & pnMovItemRef & "'"
    oConect.Ejecutar sSQL
    Exit Sub
ERRORInsertaMovComprobanteContratoItemDet:
        Call RaiseError(MyUnhandledError, "DMov:InsertaMovComprobanteContratoItemDet Method")
End Sub
Public Sub ActualizaEstadoBienContrato(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnMovItem As Integer, ByVal pnEstado As Integer)
    On Error GoTo ERRORActualizaEstadoItemContrato
    
    sSQL = "stp_upd_ERS0772014_ActualizaEstadoBienContrato '" & psncontrato & "'," _
                                                            & pncontref & "," _
                                                            & pnMovItem & "," _
                                                            & pnEstado
    oConect.Ejecutar sSQL
    Exit Sub
ERRORActualizaEstadoItemContrato:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaEstadoItemContrato Method")
End Sub
Public Sub ActualizaSaldoContrato(ByVal psncontrato As String, ByVal pncontref As Integer, ByVal pnSaldo As Currency, Optional ByVal pnAdendaRef As Integer = 0)
    On Error GoTo ERRORActualizaSaldoContrato
    sSQL = "stp_upd_ERS0772014_ActualizaSaldoContrato '" & psncontrato & "'," _
                                                & pncontref & "," _
                                                & pnSaldo & "," _
                                                & pnAdendaRef
    oConect.Ejecutar sSQL
    Exit Sub
ERRORActualizaSaldoContrato:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaSaldoContrato Method")
End Sub
Public Sub InsertaMovActaConformidadReg(ByVal pnMovNro As Long, ByVal psAreaAgeCod As String, ByVal psSubAreaDesc As String, ByVal pnMoneda As Moneda, ByVal pnDocOrigenTpo As Integer, ByVal psDocReferencia As String, ByVal psDescripcion As String, ByVal psObservacion As String, ByVal psGuia As String)
    Dim lsSQL As String
    On Error GoTo ErrInsertaMovActaConformidadReg
    lsSQL = "Exec stp_ins_ERS0772014_MovActaConformidadReg " & pnMovNro & ",'" & psAreaAgeCod & "','" & psSubAreaDesc & "'," & pnMoneda & "," & pnDocOrigenTpo & ",'" & psDocReferencia & "','" & psDescripcion & "','" & psObservacion & "','" & psGuia & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaMovActaConformidadReg:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovActaConformidad Method")
End Sub
Public Sub InsertaContratoBienOrden(ByVal pncontref As Integer, ByVal psobjeto As String, ByVal pnCant As Integer, ByVal pnMonto As Currency, ByVal psMovNro As String)
    Dim lsSQL As String
    On Error GoTo ErrInsertaContratoBienOrden
    'lsSQL = "Exec stp_ins_ERS0772014_ContratoBienOrden " & pnContref & ",'" & psAgeDest & "','" & psobjeto & "'," & pnCant & "," & pnMonto & ",'" & psMovnro & "'"
    lsSQL = "Exec stp_ins_ERS0772014_ContratoBienOrden " & pncontref & ",'" & psobjeto & "'," & pnCant & "," & pnMonto & ",'" & psMovNro & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaContratoBienOrden:
    Call RaiseError(MyUnhandledError, "DMov:InsertaContratoBienOrden Method")
End Sub
Public Sub InsertaContratoBienOrdenEdit(ByVal pncontref As Integer, ByVal psobjeto As String, ByVal pnCant As Integer, ByVal pnMonto As Currency, ByVal psMovNro As String, ByVal psMovNroAnt As String)
    Dim lsSQL As String
    On Error GoTo ErrInsertaContratoBienOrden
    'lsSQL = "Exec stp_ins_ERS0772014_ContratoBienOrdenEdit " & pnContref & ",'" & psAgeDest & "','" & psobjeto & "'," & pnCant & "," & pnMonto & ",'" & psMovnro & "'," & psMovNroAnt & "'"
    lsSQL = "Exec stp_ins_ERS0772014_ContratoBienOrdenEdit " & pncontref & ",'" & psobjeto & "'," & pnCant & "," & pnMonto & ",'" & psMovNro & "'," & psMovNroAnt & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrInsertaContratoBienOrden:
    Call RaiseError(MyUnhandledError, "DMov:InsertaContratoBienOrden Method")
End Sub
Public Sub ActualizaActaPendComprobante(ByVal pnMovNro As Long, ByVal psAreaUsuaria As String, ByVal psMovNro As String, ByVal psDocNro As String)
    Dim lsSQL As String
    On Error GoTo ErrActualizaActaPendComprobante
    lsSQL = "Exec stp_upd_ERS0772014_ActaPendComprobante " & pnMovNro & ",'" & psAreaUsuaria & "','" & psMovNro & "','" & psDocNro & "'"
    oConect.Ejecutar lsSQL
    Exit Sub
ErrActualizaActaPendComprobante:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaActaPendComprobante Method")
End Sub
Public Function GetExisteNActaConformidadxContrato(ByVal pnMovNro As Long) As String
    Dim rs As ADODB.Recordset
    Dim lsSQL As String
    On Error GoTo ErrGetExisteNActaConformidadxContrato
    lsSQL = "stp_sel_ERS0772014_ExisteNroActaxConformidad " & pnMovNro
    Set rs = oConect.CargaRecordSet(lsSQL)
    If Not rs.EOF Then
        GetExisteNActaConformidadxContrato = rs!NActa
    End If
    Exit Function
ErrGetExisteNActaConformidadxContrato:
    Err.Raise Err.Number, "Error al recuperar correlativo Acta de Conformidad", Err.Description
End Function
Public Sub ActualizaContratoxExtActaConformidad(ByVal pnMovNroComp As Long)
    On Error GoTo ErrActualizaContratoxExtActaConformidad
    sSQL = "Exec stp_upd_ERS0772014_ActualizaContratoxExtActaConformidad " & pnMovNroComp
    oConect.Ejecutar sSQL
    Exit Sub
ErrActualizaContratoxExtActaConformidad:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaContratoxExtActaConformidad Method")
End Sub
Public Function ExisteMovContratoOrden(ByVal pnMovNro As Long) As Boolean
    On Error GoTo ExisteMovOrdenContrato
    Dim rs As ADODB.Recordset
    sSQL = "stp_sel_ERS0772014_ExisteMovContratoOrden " & pnMovNro
    Set rs = oConect.Ejecutar(sSQL)
    If Not (rs.EOF) Then
        If rs!CantBien = 0 Then
            ExisteMovContratoOrden = False
        Else
            ExisteMovContratoOrden = True
        End If
    Else
        ExisteMovContratoOrden = False
    End If
    oConect.Ejecutar sSQL
    Exit Function
ExisteMovOrdenContrato:
    Call RaiseError(MyUnhandledError, "DMov:ExisteMovOrdenContrato Method")
End Function
Public Sub ActualizaMovContratoOrden(ByVal pnMovNro As Long)
    On Error GoTo ErrActualizaMovContratoOrden
    sSQL = "stp_sel_ERS0772014_ActualizaMovContratoOrden " & pnMovNro
    oConect.Ejecutar sSQL
    Exit Sub
ErrActualizaMovContratoOrden:
    Call RaiseError(MyUnhandledError, "DMov:ActualizaMovContratoOrden Method")
End Sub
'end PASI
'PASI20150918 ERS0472015
Public Function obtieneExisteItemComprobanteOrden(ByVal pnItem As Integer, ByVal pnMovNro As Long) As Integer
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        sSQL = "stp_sel_ERS0472015_ObtieneExisteItemComprobantexOrden " & pnItem & "," & pnMovNro
        Set rs = oConect.CargaRecordSet(sSQL)
        If Not rs.EOF And Not rs.BOF Then
             obtieneExisteItemComprobanteOrden = rs!existe
        Else
             obtieneExisteItemComprobanteOrden = 0
        End If
End Function
Public Function ObtieneMovComprobantexOrden(ByVal pnMovNro As Long) As Long
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
        sSQL = "stp_sel_ERS0472015_ObtieneMovComprobantexOrden " & pnMovNro
        Set rs = oConect.CargaRecordSet(sSQL)
    ObtieneMovComprobantexOrden = CLng(rs!nMovNro)
End Function
Public Function RegularizaMovObj(ByVal pnMovNroAnt As Long, ByVal pnMovNroNew As Long, ByVal nMovItem As Integer, ByVal psCtaContCod As String) As Integer
    On Error GoTo RegularizaMovObjErr
    RegularizaMovObj = 1
    sSQL = "stp_sel_ERS0472015_RegularizaMovObj " & pnMovNroAnt & "," & pnMovNroNew & "," & nMovItem & ",'" & psCtaContCod & "'"
    oConect.Ejecutar sSQL
    RegularizaMovObj = 0
    Exit Function
RegularizaMovObjErr:
    Call RaiseError(MyUnhandledError, "DMov:RegularizaMovObj Method")
End Function
Public Function InsertaMovBSOrdPago(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal pnCantMillar As Integer, ByVal pnRangoIni As Long, ByVal pnRangoFin As Long, ByVal psGlosa As String) As Integer
    On Error GoTo InsertaMovBSOrdPagoErr
    InsertaMovBSOrdPago = 1
    sSQL = "stp_ins_ERS0782015_InsertaMovBSOrdPago " & pnMovNro & "," & pnMovItem & "," & pnCantMillar & "," & pnRangoIni & "," & pnRangoFin & ",'" & psGlosa & "'"
    oConect.Ejecutar sSQL
    InsertaMovBSOrdPago = 0
    Exit Function
InsertaMovBSOrdPagoErr:
    Call RaiseError(MyUnhandledError, "DMov:InsertaMovBSOrdPago Method")
End Function
'END PASI
