VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XAcceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function GetMenusUsuario(ByVal psAppNumber As Integer, ByVal psDominio As String, Optional ByVal psUsuario As String = "", Optional ByVal sTipoUsu As String = "1")
Dim sSQL As String, oConn As DConecta
Dim ROpe As ADODB.Recordset
Dim RMenu As ADODB.Recordset, cUserString As String
Dim cMenu As String, nItem As Integer, i As Integer
Dim Ctrl As Control

On Error Resume Next

cUserString = GetStringUserGroup(psUsuario, psDominio)

Set oConn = New DConecta
If oConn.AbreConexion() Then

   Set ROpe = New ADODB.Recordset
   sSQL = "Select cOpeCod, cOpeDesc, cOpeGruCod, cOpeVisible, nOpeNiv " & _
          " from OpeTpo where cOpeCod in (select cName from Permiso where cGrupoUsu in (" & cUserString & ") ) " & _
          " Order by cOpeCod "
           
   Set ROpe = oConn.CargaRecordSet(sSQL)
   i = 0
   Do While Not ROpe.EOF
      MatOperac(i, 0) = ROpe!cOpeCod
      MatOperac(i, 1) = ROpe!cOpeDesc
      MatOperac(i, 2) = IIf(IsNull(ROpe!cOpeGruCod), "", ROpe!cOpeGruCod)
      MatOperac(i, 3) = ROpe!cOpeVisible
      MatOperac(i, 4) = ROpe!nOpeNiv
      NroRegOpe = i
      i = i + 1
      ROpe.MoveNext
   Loop
   Set ROpe = Nothing
   
   i = 0
   Set RMenu = New ADODB.Recordset
   
   sSQL = "select cName from menu where cAplicacion = '" & CStr(psAppNumber) & "' and " & _
          "  cName in (select cName from Permiso where cGrupoUsu in (" & cUserString & ")) Order by cName "
          
   Set RMenu = oConn.CargaRecordSet(sSQL)
   If Not RMenu.EOF Then
      Do While Not RMenu.EOF
         i = i + 1
         cMenu = Mid(RMenu!cname, 1, 11)
         nItem = CInt(Mid(RMenu!cname, 12, 2))
         For Each Ctrl In MDISicmact.Controls
             If TypeName(Ctrl) = "Menu" Then
                If Ctrl.Name = cMenu And Ctrl.Index = nItem Then
                   Ctrl.Visible = True
                Else
                   Ctrl.Visible = False
                End If
             End If
         Next
         RMenu.MoveNext
      Loop
      
'      For Each Ctrl In MDISicmact.Controls
'          If TypeName(Ctrl) = "Menu" Then
'             cMenu = Ctrl.Name
'
''            If Left(cMenu, 3) = "M17" Then
''               MsgBox "dd"
''            End If
'
'             i = Ctrl.Index
'             If Val(Mid(cMenu, 4, 8)) = 0 Then
'                RMenu.Find "cName like '" & cMenu & "%'", , adSearchForward, 1
'             Else
'                cMenu = Ctrl.Name + Format(i, "00")
'                RMenu.Find "cName = '" & cMenu & "'", , adSearchForward, 1
'             End If
'
'             If RMenu.EOF Then
'                Ctrl.Visible = False
'             End If
'          End If
'      Next
   End If
   oConn.CierraConexion
End If
Set oConn = Nothing
Exit Function

SalError:
   Set oConn = Nothing
   Err.Clear
End Function

Private Function GetStringUserGroup(ByVal sUsuario As String, ByVal psDominio As String)
Dim i As Integer, Group As IADsGroup, User As IADsUser
GetStringUserGroup = ""
Set User = GetObject("WinNT://" & psDominio & "/" & sUsuario & ",user")
For Each Group In User.Groups
    i = i + 1
    If i = 1 Then
       GetStringUserGroup = GetStringUserGroup + "'" & sUsuario & "'"
    End If
    If UCase(Left(Group.Name, 5)) = "GRUPO" Or UCase(Left(Group.Name, 3)) = "GG_" Then
       i = i + 1
       GetStringUserGroup = GetStringUserGroup + ",'" & Group.Name & "'"
    End If
Next
End Function

