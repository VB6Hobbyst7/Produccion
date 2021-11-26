VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCajaGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim vsServerCom As String
Dim vsServerPers As String
Dim vsConexion As String
Dim vsFormatoFecha As String
Dim oConect As DConecta
Public Sub inicio(ByVal psFormatoFecha As String)
vsFormatoFecha = psFormatoFecha
End Sub
Public Function InsertaOpCajaGen(ByVal psRangoCod As String, ByVal pdFechaIng As Date, ByVal pnRangoEmiIni As Long, ByVal pnRangoEmiFin As Long, ByVal psEstadoRango As CGEstadosOp, ByVal psUltActualizacion As String, ByVal pnMoneda As Moneda, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo InsertaOpCajaGenErr

InsertaOpCajaGen = 1
sql = "INSERT INTO OPCAJAGEN(cRangoCod, dFechaIng, nRangoEmiIni, nRangoEmiFin, cEstadoRango, cUltActualizacion, cMoneda) " _
    & " VALUES ('" & psRangoCod & "','" & Format(pdFechaIng, vsFormatoFecha) & "'," & pnRangoEmiIni & "," & pnRangoEmiFin & ",'" & psEstadoRango & "','" & psUltActualizacion & "','" & pnMoneda & "')"

If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
InsertaOpCajaGen = 0
Exit Function
InsertaOpCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:InsertaOpCajaGen")

End Function
Public Function InsertaOpDetCajaGen(ByVal psRangoCod As String, ByVal pnRangoIni As Long, ByVal pnRangoFin As Long, ByVal psEstado As CGEstadosOp, ByVal psDescripcion As String, ByVal psUltActualizacion As String, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo InsertaOpDetCajaGenErr

InsertaOpDetCajaGen = 1
sql = "INSERT INTO OPCAJAGENDET(cRangoCod,nRangoIni,nRangoFin,cEstadoDet,cDescDet,cUltActualizacion) " _
    & " VALUES ('" & psRangoCod & "'," & pnRangoIni & "," & pnRangoFin & ",'" & psEstado & "','" & Replace(psDescripcion, "'", "''") & "','" & psUltActualizacion & "')"

If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
InsertaOpDetCajaGen = 0
Exit Function
InsertaOpDetCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:InsertaOpDetCajaGen")
End Function
Public Function EliminaOpCajaGen(ByVal psRangoCod As String, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo EliminaOpCajaGenErr
sql = "DELETE OPCAJAGEN WHERE cRangoCod= '" & psRangoCod & "'"
EliminaOpCajaGen = 1
If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
EliminaOpCajaGen = 0

Exit Function
EliminaOpCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:EliminaOpCajaGen")

End Function
Public Function EliminaOpDetCajaGen(ByVal psRangoCod As String, ByVal pnRangoIni As Long, ByVal pnRangoFin As Long, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo EliminaOpDetCajaGenErr

EliminaOpDetCajaGen = 1
sql = "DELETE FROM OPCAJAGENDET WHERE cRangoCod='" & psRangoCod & "' AND nRangoIni=" & pnRangoIni & " AND nRangoFin=" & pnRangoFin

If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
EliminaOpDetCajaGen = 0


Exit Function
EliminaOpDetCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:EliminaOpDetCajaGen ")

End Function
Public Function EjecutaBatch() As Integer
    EjecutaBatch = oConect.EjecutarBatch
End Function
Private Sub Class_Initialize()
Dim oIni As ClasIni
Set oIni = New ClasIni
vsServerCom = oIni.BaseComunes
vsServerPers = oIni.BasePersonas
Set oIni = Nothing

Set oConect = New DConecta
If oConect.AbreConexion = False Then
    Call RaiseError(MyUnhandledError, "DCajaGeneral:Initialize Method. Error en Conexion a Base de datos")
End If
End Sub
Private Sub Class_Terminate()
oConect.CierraConexion
Set oConect = Nothing
End Sub
Public Function ActualizaOpCajaGen(ByVal psRangoCod As String, ByVal pnRangoEmiIni As Long, ByVal pnRangoEmiFin As Long, ByVal psUltActualizacion As String, Optional pbEjectBacth As Boolean = False) As Integer
Dim sql As String
On Error GoTo ActualizaOpCajaGenErr

ActualizaOpCajaGen = 1
sql = "UPDATE OPCajaGen SET nRangoEmiIni=" & pnRangoEmiIni & ",nRangoEmiFin=" & pnRangoEmiFin & ", cUltActualizacion='" & psUltActualizacion & "' WHERE cRangoCod='" & psRangoCod & "'"
If pbEjectBacth Then
    oConect.AdicionaCmdBatch sql
Else
    oConect.Ejecutar sql
End If
ActualizaOpCajaGen = 0
Exit Function
ActualizaOpCajaGenErr:
    Call RaiseError(MyUnhandledError, "DCajaGeneral:ActualizaOpCajaGen ")
End Function

Public Function GetOrdPagFondoFijoEntregado(ByVal nMoneda As Moneda, ByVal sNumOrdPag As String) As adodb.Recordset
Dim sSQL As String, rstemp As adodb.Recordset

Set rstemp = New adodb.Recordset

sSQL = "SELECT A.nMovNro,B.dDocFecha, G.cDocAbrev, B.nDocTpo cDocTpo, B.cDocNro, E.cPersNombre cNomPers, A.cMovDesc, D.cPersCod cObjetoCod, " _
    & "A.cMovNro, ISNULL(ME.nMovMEImporte, C.nMovImporte) * -1 As nDocImporte " _
    & "FROM Mov A JOIN MovDoc B ON B.nMovNro = A.nMovNro " _
    & "     JOIN MovCta C ON c.nMovNro = a.nMovNro LEFT JOIN MovMe ME ON ME.nMovNro = C.nMovNro And ME.nMovItem = C.nMovItem " _
    & "     JOIN MovGasto D ON D.nMovNro = A.nMovNro " _
    & "     JOIN Persona E ON E.cPersCod = D.cPersCod " _
    & "     JOIN Documento G ON G.nDocTpo = B.nDocTpo " _
    & "WHERE A.nMovEstado = " & gMovEstContabMovContable & "  and A.nMovFlag <> " & gMovFlagEliminado & "  And B.nDoctpo = " & TpoDocOrdenPago & " "
    
    'A rendir
    sSQL = sSQL & " And (C.cCtaContCod LIKE '21_10502%' or c.cctacontcod like '29_80701%') and (copecod like  '40_132' or  copecod like  '42_113')   "
    sSQL = sSQL & " and c.cCtaContCod like '__" & nMoneda & "%' "
                '& "And C.cCtaContCod LIKE '21_10502%' and c.cCtaContCod like '__" & nMoneda & "%' "
    'Fin a rendir

sSQL = sSQL & "And C.nMovImporte < 0 And B.cDocNro = '" & sNumOrdPag & "'" _
    & "And EXISTS (SELECT H.nMovNro FROM MovRef H WHERE H.nMovNro = A.nMovNro) " _
    & "And NOT EXISTS (SELECT MP.nMovNro FROM MovRef Pag JOIN Mov MP on MP.nMovNro = Pag.nMovNro " _
    & "    WHERE Pag.nMovNroRef = A.nMovNro And MP.nMovFlag <> " & gMovFlagEliminado & "  And MP.cOpeCod IN ('" & gAhoRetFondoFijo & "','" & gAhoRetFondoFijoCanje & " ')) " _
    & "ORDER BY a.cMovNro, e.cPersNombre "
    
    rstemp.CursorLocation = adUseClient
    rstemp.Open sSQL, oConect.ConexionActiva, adOpenStatic, adLockOptimistic, adCmdText
    
    
Set GetOrdPagFondoFijoEntregado = rstemp
   Set rstemp = Nothing
   
End Function
