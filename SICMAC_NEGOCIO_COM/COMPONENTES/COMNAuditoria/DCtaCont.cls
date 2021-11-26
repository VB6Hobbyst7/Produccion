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

Dim dbConec As COMConecta.DCOMConecta
Dim oConec  As COMConecta.DCOMClasIni
Dim psSql As String
Dim prs   As ADODB.Recordset

Dim sCentralCom As String

Public Function CargaCtaCont(Optional psFiltroCta As String = "", Optional psTabla As String = "CtaCont", Optional LockType As LockTypeEnum = adLockReadOnly, Optional plAlias As Boolean = False) As Recordset
On Error GoTo CargaCtaContErr
   If psTabla = "CtaCont" Then
      psTabla = sCentralCom & psTabla
   End If
      If Not plAlias Then
         psSql = "SELECT cCtaContCod, cCtaContDesc FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      Else
         psSql = "SELECT cCtaContCod Código, cCtaContDesc Descripción FROM " & psTabla & IIf(psFiltroCta = "", "", " WHERE " & psFiltroCta)
      End If
      Set CargaCtaCont = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaCtaContErr:
   MsgBox Err.Description, vbCritical
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

Public Function ExisteCuenta(psCtaContCod As String, Optional plEsInstancia As Boolean = False) As Boolean
   On Error GoTo ExisteCuentaErr
   ExisteCuenta = True
      psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE nCtaEstado =1 and cCtaContCod LIKE '" & psCtaContCod & "%'"
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.EOF Then
         ExisteCuenta = False
         Err.Raise 50001, "DCtaCont:ExisteCuenta", "Cuenta Contable no existe...!"
      Else
         If prs!cCtaContCod <> psCtaContCod Then
             ExisteCuenta = False
             Err.Raise 50001, "DCtaCont:ExisteCuenta", "Cuenta Contable no existe...!"
         End If
      End If
      If plEsInstancia Then
         If prs.RecordCount > 1 Then
             ExisteCuenta = False
             Err.Raise 50001, "DCtaCont:ExisteCuenta", "Cuenta Contable no es ultima Instancia"
         End If
      End If
   'ExisteCuenta = True
   Exit Function
ExisteCuentaErr:
    RSClose prs
    
End Function

Public Function ExisteCuentaCad(psCtaContCod As String, Optional plEsInstancia As Boolean = False) As String
   On Error GoTo ExisteCuentaErr
   ExisteCuentaCad = ""
      psSql = "SELECT cCtaContCod FROM " & sCentralCom & "CtaCont WHERE cCtaContCod LIKE '" & psCtaContCod & "%'"
      Set prs = dbConec.CargaRecordSet(psSql)
      If prs.EOF Then
         ExisteCuentaCad = "Cuenta Contable no existe...!"
      Else
         If prs!cCtaContCod <> psCtaContCod Then
             ExisteCuentaCad = "Cuenta Contable no existe...!"
         End If
      End If
      If plEsInstancia Then
         If prs.RecordCount > 1 Then
             ExisteCuentaCad = "Cuenta Contable no es ultima Instancia"
         End If
      End If
   Exit Function
ExisteCuentaErr:
   RSClose prs
   'Call RaiseError(MyUnhandledError, "DCtaCont:ExisteCuentaCad Method")
End Function

Private Sub Class_Initialize()
    Dim oImp As COMDConstSistema.DCOMImpresoras
    Set oImp = New COMDConstSistema.DCOMImpresoras
    
    oImpresora.Inicia oImp.GetImpreSetup(oImp.GetMaquina)
    
    Set oImp = Nothing
   Set oConec = New COMConecta.DCOMClasIni
   sCentralCom = oConec.BaseComunes
   Set oConec = Nothing
   Set dbConec = New COMConecta.DCOMConecta
   dbConec.AbreConexion
End Sub


Public Function CargaCtaContClase(Optional psCtaCod As String = "", Optional LockType As LockTypeEnum = adLockReadOnly) As Recordset
On Error GoTo CargaCtaContClaseErr
      psSql = "SELECT cCtaContCod, cCtaCaracter, nCtaMoneda FROM " & sCentralCom & "CtaContClase " & IIf(psCtaCod = "", "", " WHERE '" & psCtaCod & "' LIKE cCtaContCod + '%' ")
      Set CargaCtaContClase = dbConec.CargaRecordSet(psSql, LockType)
   Exit Function
CargaCtaContClaseErr:
   MsgBox Err.Description, vbCritical
End Function


Private Sub Class_Terminate()
If Not dbConec Is Nothing Then
    dbConec.CierraConexion
End If
Set dbConec = Nothing
End Sub
