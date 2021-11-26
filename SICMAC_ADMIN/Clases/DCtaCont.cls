VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCtaCont"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3A7B5FCF038A"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Dim dbConec As DConecta
Dim oConec  As ClsIni.ClasIni
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sConexion As String
Dim sCentralCom As String

Public Function CargaCtaCont(Optional psFiltroCta As String = "", Optional psTabla As String = "CtaCont", Optional LockType As LockTypeEnum = adLockReadOnly, Optional plAlias As Boolean = False) As Recordset
On Error GoTo CargaCtaContErr
   If psTabla = "CtaCont" Then
      psTabla = sCentralCom & psTabla
   End If
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      If Not plAlias Then
         psSql = "SELECT cCtaContCod, cCtaContDesc FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      Else
         psSql = "SELECT cCtaContCod Código, cCtaContDesc Descripción FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      End If
      Set CargaCtaCont = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaCont Method")
End Function

Public Sub InsertaCtaCont(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont", Optional psMoneda As String = "")
   On Error GoTo InsertaCtaContErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      If psMoneda = "" Or psTabla = "CtaContBase" Then
         psSql = "INSERT " & psTabla & " (cCtaContCod, cCtaContDesc, cUltimaActualizacion) " _
               & "VALUES ('" & psCtaCod & "', '" & psCtaDesc & "','" & psMovNro & "')"
         dbConec.Ejecutar psSql
      Else
         If psTabla = sCentralCom & "CtaCont" Then
            psSql = sCentralCom & "sp_InsertaCtaCont '" & psCtaCod & "','" & psCtaDesc & "', '" & psMovNro & "', '" & psMoneda & "'"
            dbConec.Ejecutar psSql
         End If
      End If
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaCont Method")
End Sub

Public Sub ActualizaCtaCont(psCtaCod As String, psCtaDesc As String, psMovNro As String, Optional psTabla As String = "CtaCont")
   On Error GoTo ActualizaCtaContErr
      Set dbConec = New DConecta
      dbConec.AbreConexion
         psSql = "UPDATE " & psTabla & " SET cCtaContDesc = '" & psCtaDesc & "', cUltimaActualizacion = '" & psMovNro & "' WHERE cCtaContCod LIKE '" & psCtaCod & "'"
         dbConec.Ejecutar psSql
      dbConec.CierraConexion
      Set dbConec = Nothing
   Exit Sub
ActualizaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:ActualizaCtaCont Method")
End Sub

Public Sub EliminaCtaCont(psCtaCod As String, Optional psTabla As String = "CtaCont")
   On Error GoTo EliminaCtaContErr
      If psTabla = "CtaCont" Then
         psTabla = sCentralCom & psTabla
      End If
      Set dbConec = New DConecta
      dbConec.AbreConexion
         psSql = "DELETE " & psTabla & " WHERE cCtaContCod LIKE '" & psCtaCod & "'"
         dbConec.Ejecutar psSql
      dbConec.CierraConexion
      Set dbConec = Nothing
   Exit Sub
EliminaCtaContErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:EliminaCtaCont Method")
End Sub
Public Function CargaCtaObj(Optional psCtaCod As String = "", Optional psFiltro As String = "", Optional lObjDesc As Boolean = False, Optional pbLogAge As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaObjErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      If lObjDesc Then
            psSql = "SELECT co.cCtaContCod, co.nCtaObjOrden, co.cObjetoCod, o.cObjetoDesc, co.nCtaObjNiv, co.cCtaObjFiltro, co.cCtaObjImpre, o.nObjetoNiv FROM " & sCentralCom & "CtaObj co LEFT JOIN " & sCentralCom & "Objeto o ON o.cObjetoCod =  co.cObjetoCod " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, ""))
      Else
         psSql = "SELECT cCtaContCod, nCtaObjOrden, cObjetoCod, nCtaObjNiv, cCtaObjFiltro, cCtaObjImpre FROM " & sCentralCom & "CtaObj " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "'", IIf(psFiltro <> "", " WHERE " & psFiltro, ""))
      End If
      Set CargaCtaObj = dbConec.CargaRecordSet(psSql)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaCtaObjErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaObj Method")
End Function

Private Function ImpreFiltro(ObjCod As String, ObjFiltro As String) As String
Dim N As Integer
ImpreFiltro = ""
For N = 1 To Len(ObjFiltro)
    If Mid(ObjFiltro, N, 1) = "X" Then
       If Len(ImpreFiltro) = 0 Then
          ImpreFiltro = ""
       End If
       ImpreFiltro = ImpreFiltro & Mid(ObjCod, N, 1)
    End If
Next
End Function

Public Sub InsertaCtaObj(psCtaCod As String, psCtaObjOrden As String, psObjetoCod As String, pnCtaObjNiv As Integer, psCtaObjFiltro As String, psCtaObjImpre As String, psMovNro As String, Optional pnObjetoNiv As Integer = 0)
Dim lTrans As Boolean
   On Error GoTo InsertaCtaObjErr
   Set dbConec = New DConecta
   
   If dbConec.AbreConexion() Then
      dbConec.BeginTrans
      lTrans = True
      psSql = "INSERT " & sCentralCom & "CtaObj (cCtaContCod, cCtaObjOrden, cObjetoCod, nCtaObjNiv, cCtaObjFiltro, cCtaObjImpre, cUltimaActualizacion) " _
            & "VALUES ('" & psCtaCod & "', '" & psCtaObjOrden & "','" & psObjetoCod & "'," & pnCtaObjNiv + pnObjetoNiv & ",'" & psCtaObjFiltro & "','" & psCtaObjImpre & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      If psCtaObjImpre <> "" Then
         Set prs = dbConec.CargaRecordSet(sCentralCom & "spGetTreeObj '" & psObjetoCod & "'," & pnCtaObjNiv & ",'" & psCtaObjFiltro & "'", adLockReadOnly)
         Do While Not prs.EOF
            If prs!nObjetoNiv = pnCtaObjNiv + pnObjetoNiv Then
               InsertaCtaObjFiltro psCtaCod, psCtaObjOrden, prs!cObjetoCod, ImpreFiltro(prs!cObjetoCod, psCtaObjImpre), psMovNro
               dbConec.Ejecutar psSql
            End If
            prs.MoveNext
         Loop
         If prs.State = adStateOpen Then prs.Close: Set prs = Nothing
      End If
      dbConec.CommitTrans
      lTrans = False
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaCtaObjErr:
   If lTrans Then
      dbConec.RollBackTrans
   End If
   dbConec.CierraConexion
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaObj Method")
End Sub

Public Sub EliminaCtaObj(psCtaCod As String, psObjCod As String)
   On Error GoTo EliminaCtaObjErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod LIKE '" & psObjCod & "%'"
      dbConec.AdicionaCmdBatch psSql
      psSql = "DELETE " & sCentralCom & "CtaObj WHERE cCtaContCod = '" & psCtaCod & "' and cObjetoCod = '" & psObjCod & "'"
      dbConec.AdicionaCmdBatch psSql
      dbConec.EjecutarBatch
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaCtaObjErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:GrabaCtaObj Method")
End Sub

Public Function CargaCtaObjFiltro(Optional psCtaCod As String = "", Optional psObjetoCod As String = "", Optional psFiltro As String = "", Optional lObjDesc As Boolean = False) As ADODB.Recordset
   On Error GoTo CargaCtaObjFiltroErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      If lObjDesc Then
         psSql = "SELECT co.cCtaContCod, co.cCtaObjOrden, co.cObjetoCod, o.cObjetoDesc, co.cCtaObjSubCta cSubCtaCod FROM " & sCentralCom & "CtaObjFiltro co LEFT JOIN " & sCentralCom & "Objeto o ON o.cObjetoCod =  co.cObjetoCod " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "' ", "") & IIf(psFiltro <> "", IIf(psCtaCod = "", " WHERE ", " AND ") & psFiltro, "") & IIf(psObjetoCod = "", "", " and co.cObjetoCod LIKE '" & psObjetoCod & "%'")
      Else
         psSql = "SELECT cCtaContCod, cCtaObjOrden, cObjetoCod, cCtaObjSubCta cSubCtaCod FROM " & sCentralCom & "CtaObjFiltro co " & IIf(psCtaCod <> "", " WHERE cCtaContCod = '" & psCtaCod & "' ", "") & IIf(psFiltro <> "", IIf(psCtaCod = "", " WHERE ", " AND ") & psFiltro, "") & IIf(psObjetoCod = "", "", " and co.cObjetoCod LIKE '" & psObjetoCod & "%'")
      End If
      Set CargaCtaObjFiltro = dbConec.CargaRecordSet(psSql)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaObjFiltro Method")
End Function

Public Sub InsertaCtaObjFiltro(psCtaCod As String, psCtaObjOrden As String, psObjetoCod As String, psCtaObjSubCta As String, psMovNro As String)
   On Error GoTo InsertaCtaObjFiltroErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "INSERT " & sCentralCom & "CtaObjFiltro (cCtaContCod, cCtaObjOrden, cObjetoCod, cCtaObjSubCta, cUltimaActualizacion) " _
            & "VALUES ('" & psCtaCod & "', '" & psCtaObjOrden & "','" & psObjetoCod & "','" & psCtaObjSubCta & "','" & psMovNro & "')"
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
InsertaCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaCtaObjFiltro Method")
End Sub

Public Sub EliminaCtaObjFiltro(psCtaCod As String, Optional psCtaObjOrden As String = "", Optional psObjCod As String = "")
   On Error GoTo EliminaCtaObjFiltroErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "DELETE " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaCod & "' " & IIf(psCtaObjOrden = "", "", " and cCtaObjOrden = '" & psCtaObjOrden & "' ") & IIf(psObjCod = "", "", " and cObjetoCod = '" & psObjCod & "'")
      dbConec.Ejecutar psSql
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
EliminaCtaObjFiltroErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:EliminaCtaObjFiltro Method")
End Sub

Public Sub CargaConvertidor()
   On Error GoTo CargaConvertidorErr
   Exit Sub
CargaConvertidorErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaConvertidor Method")
End Sub

Public Sub InsertaConvertidor()
   On Error GoTo InsertaConvertidorErr
   Exit Sub
InsertaConvertidorErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:InsertaConvertidor Method")
End Sub

Public Sub ActualizaConvertidor()
   On Error GoTo ActualizaConvertidorErr
   Exit Sub
ActualizaConvertidorErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:ActualizaConvertidor Method")
End Sub

Public Sub EliminaConvertidor()
   On Error GoTo EliminaConvertidorErr
   Exit Sub
EliminaConvertidorErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:EliminaConvertidor Method")
End Sub

Public Function ExisteCuenta(psCtaContCod As String, Optional plEsInstancia As Boolean = False) As Boolean
   On Error GoTo ExisteCuentaErr
   Set dbConec = New DConecta
   ExisteCuenta = False
   If dbConec.AbreConexion() Then
      psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psCtaContCod & "%'"
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.EOF Then
         MsgBox "Cuenta Contable no existe...!", vbInformation, "!Aviso!"
      Else
         If prs!cCtaContCod <> psCtaContCod Then
            MsgBox "Cuenta Contable no existe...!", vbInformation, "!Aviso!"
         End If
      End If
      If plEsInstancia Then
         If prs.RecordCount > 1 Then
            MsgBox "Cuenta Contable no es ultima Instancia", vbInformation, "!AViso!"
            prs.Close: Set prs = Nothing
            Exit Function
         End If
      End If
      prs.Close: Set prs = Nothing
   End If
   ExisteCuenta = True
   Exit Function
ExisteCuentaErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:ExisteCuenta Method")
End Function

Public Function CtaInstancia(psCtaCod As String, Optional psTabla As String = "CtaCont") As Boolean
   On Error GoTo CtaInstancia
   CtaInstancia = True
   If psTabla = "CtaCont" Then
      psTabla = sCentralCom & psTabla
   End If
   psSql = "SELECT cCtaContCod FROM " & psTabla & " WHERE cCtaContCod LIKE '" & psCtaCod & "%'"
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.RecordCount > 1 Then
         CtaInstancia = False
      End If
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CtaInstancia:
   Call RaiseError(MyUnhandledError, "DCtaCont:CuentaInstancia Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As DImpresoras
    Set oImp = New DImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing

   Set oConec = New ClsIni.ClasIni
   sConexion = oConec.CadenaConexion
   sCentralCom = oConec.BaseComunes
   Set oConec = Nothing
End Sub

Public Function GetCtaObjFiltro(ByVal psCtaContCod As String, psObjetoCod As String, Optional lbMuestraCta As Boolean = True) As String
    On Error GoTo GetCtaObjFiltroErr
    Dim sql As String
    Dim rs   As New ADODB.Recordset
    Dim oconect As DConecta
    
    Set oconect = New DConecta
    If oconect.AbreConexion = False Then Exit Function
    GetCtaObjFiltro = ""
    sql = "SELECT cCtaObjSubCta FROM " & sCentralCom & "CtaObjFiltro WHERE cCtaContCod = '" & psCtaContCod & "' and cObjetoCod = '" & psObjetoCod & "'"
    Set rs = oconect.CargaRecordSet(sql)
    If Not rs.EOF Then
        GetCtaObjFiltro = IIf(lbMuestraCta, psCtaContCod, "") & rs!cCtaObjSubCta
    End If
    If GetCtaObjFiltro = "" Then
        GetCtaObjFiltro = IIf(lbMuestraCta, psCtaContCod, "")
    End If
    Exit Function
GetCtaObjFiltroErr:
    Call RaiseError(MyUnhandledError, "NContFunciones:GetCtaObjFiltro Method")
End Function
Public Function GetCtaObjetos(ByVal psCtaContCod As String, Optional psFiltro As String = "") As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim oRHAreas As DActualizaDatosArea
Dim oContFunct As NContFunciones
Dim oCtaIf As NCajaCtaIF
Dim oEfect As Defectivo


Set oCtaIf = New NCajaCtaIF
Set oEfect = New Defectivo
'Set oCtaCont = New DCtaCont
Set oContFunct = New NContFunciones
Set oRHAreas = New DActualizaDatosArea

Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set rs = CargaCtaObj(psCtaContCod, psFiltro, True)
If Not rs.EOF And Not rs.BOF Then
    Select Case Val(rs!cObjetoCod)
            Case ObjCMACAgencias
                Set rs1 = oRHAreas.GetAgencias(rs!cCtaObjFiltro)
            Case ObjCMACAgenciaArea
                Set rs1 = oRHAreas.GetAgenciasAreas(rs!cCtaObjFiltro)
            Case ObjCMACArea
                Set rs1 = oRHAreas.GetAreas(rs!cCtaObjFiltro)
            Case ObjEntidadesFinancieras
                Set rs1 = oCtaIf.GetCtasInstFinancieras(rs!cCtaObjFiltro, psCtaContCod)
            Case ObjDescomEfectivo
                Set rs1 = oEfect.GetBilletajes(rs!cCtaObjFiltro)
            Case ObjPersona
                Set rs1 = Nothing
            Case Else
                Set rs1 = oContFunct.GetObjetos(Val(IIf(rs!cCtaObjFiltro = "", rs!cObjetoCod, rs!cCtaObjFiltro)))
        End Select
End If
rs.Close
Set rs = Nothing
Set oContFunct = Nothing
Set oRHAreas = Nothing
Set oCtaIf = Nothing
Set oEfect = Nothing


Set GetCtaObjetos = rs1
End Function
Public Function GetCtaInstancia(psCtaCod As String, Optional psTabla As String = "CtaCont") As ADODB.Recordset
   On Error GoTo GetCtaInstancia
   If psTabla = "CtaCont" Then
      psTabla = sCentralCom & psTabla
   End If
   psSql = "SELECT c.cCtaContCod, c.cCtaContDesc FROM " & psTabla & " c, " & psTabla & " c1 " _
         & "WHERE c.cctacontcod LIKE '" & psCtaCod & "%' and c1.cctacontcod like c.cctacontcod + '%' " _
         & "GROUP BY c.cctacontcod , c.cctacontdesc " _
         & "HAVING count(*) = 1 "
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      Set GetCtaInstancia = dbConec.CargaRecordSet(psSql)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
GetCtaInstancia:
   Call RaiseError(MyUnhandledError, "DCtaCont:GetCuentaInstancia Method")
End Function

Public Function CargaCtaContClase(Optional psCtaCod As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
On Error GoTo CargaCtaContClaseErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
      psSql = "SELECT cCtaContCod, cCtaCaracter, nCtaMoneda FROM " & sCentralCom & "CtaContClase " & IIf(psCtaCod = "", "", " WHERE '" & psCtaCod & "' LIKE cCtaContCod + '%' ")
      Set CargaCtaContClase = dbConec.CargaRecordSet(psSql, LockType)
      dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Function
CargaCtaContClaseErr:
   Call RaiseError(MyUnhandledError, "DCtaCont:CargaCtaContClase Method")
End Function

'EJVG20130701 ***
Public Function VerificaExisteCuenta(psCtaContCod As String, Optional plEsInstancia As Boolean = False) As String
    Dim prs As New ADODB.Recordset
    On Error GoTo ExisteCuentaErr
    psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psCtaContCod & "%'"
    Set prs = dbConec.CargaRecordSet(psSql)
    If prs.EOF Then
        VerificaExisteCuenta = "Cuenta Contable " & psCtaContCod & " no existe...!" & vbCrLf & "Favor Consultar con Contabilidad o comunicarse con Sistemas"
        Exit Function
    Else
        If prs!cCtaContCod <> psCtaContCod Then
            VerificaExisteCuenta = "Cuenta Contable " & psCtaContCod & " no existe...!" & vbCrLf & "Favor Consultar con Contabilidad o de comunicarse con Sistemas"
            Exit Function
        End If
    End If
    If plEsInstancia Then
        If prs.RecordCount > 1 Then
            VerificaExisteCuenta = " Cuenta Contable " & psCtaContCod & " no es esta Configurada en el Ultimo Nivel " & vbCrLf & "Favor Consultar con Contabilidad o  de comunicarse con Sistemas"
            Exit Function
        End If
    End If
    Set prs = Nothing
    Exit Function
ExisteCuentaErr:
    Call RaiseError(MyUnhandledError, "DCtaCont:VerificaExisteCuenta Method")
End Function
'END EJVG *******
