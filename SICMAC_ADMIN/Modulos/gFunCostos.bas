Attribute VB_Name = "gFuncCostos"
Option Explicit

Global Const gsCtaGastoJudMN = "45130114"
Global Const gsCtaRecuperaMN = "4513012903"
Global Const gsCtaGastoJudME = "45230114"
Global Const gsCtaRecuperaME = "4523012903"

Public Enum tAnchoAjuste
    nAncho20 = 20
    nAncho30 = 30
    nAncho36 = 36
    nAncho40 = 40
    nAncho46 = 46
End Enum

Global nKeyAscii As Integer
Global HayTrans As Boolean

Public Sub InsRow(ByRef MSFlex As MSHFlexGrid, Index As Integer)
If Index = 1 And MSFlex.Rows > 1 Then
   MSFlex.RowHeight(1) = 260
Else
   MSFlex.AddItem ""
   MSFlex.RowHeight(Index) = 260
End If
End Sub

'*************************************************************************************
'Justificación a la IZQUIERDA
'Llena con espacios a la derecha de una cadena [vCadena] hasta completar una
'longitud de [nTam] caracteres

Public Function JIZQ(vCadena As String, nTam As Integer, Optional vChar As String) As String
Dim s As String, xChar As String

xChar = IIf(Len(Trim(vChar)) = 1, vChar, Space(1))
vCadena = Trim(vCadena)
If nTam > Len(Trim(vCadena)) Then
   s = String(nTam - Len(Trim(vCadena)), xChar)
   JIZQ = vCadena + s
Else
   JIZQ = Mid(vCadena, 1, nTam)
End If
End Function

'*************************************************************************************
'Justificación a la DERECHA
'Llena con espacios a la izquierda de una cadena [vCadena] hasta completar una
'longitud de [nTam] caracteres

Public Function JDER(vCadena As String, nTam As Integer, Optional vChar As String) As String
Dim s As String, xChar As String

vCadena = Trim(vCadena)
xChar = IIf(Len(Trim(vChar)) = 1, vChar, Space(1))
If nTam >= Len(Trim(vCadena)) Then
   s = String(nTam - Len(Trim(vCadena)), xChar)
   JDER = s + vCadena
Else
   JDER = String(nTam, "*")
End If
End Function

Public Function VNumero(vExpNumStr As String, Optional vNroDecimales As Integer) As Currency
Dim nDec As Integer
If Len(Trim(vExpNumStr)) = 0 Then
   VNumero = 0
Else
   nDec = IIf(vNroDecimales = 0, 2, vNroDecimales)
   VNumero = CCur(Format(vExpNumStr, "######0." + String(nDec, "0")))
End If
End Function

Public Function FNumero(vExpNumStr As Variant, Optional vNroDecimales As Integer) As String
Dim nDec As Integer
If Not IsNull(vExpNumStr) Then
   nDec = IIf(vNroDecimales = 0, 2, vNroDecimales)
   If InStr(",", vExpNumStr) = 0 Then
      FNumero = Format(vExpNumStr, "###,###,##0." + String(nDec, "0"))
   Else
      FNumero = Format(Val(vExpNumStr), "###,###,##0." + String(nDec, "0"))
   End If
Else
   FNumero = ""
End If
End Function


'***************************************************************
'Para validar ingreso de dígitos numéricos para un valor entero
'---------------------------------------------------------------
Public Function DigNumEnt(KeyAscii As Integer, Optional vOtrosChar As String) As Integer
Dim nPos As Integer
nPos = InStr("0123456789" + Trim(vOtrosChar), Chr(KeyAscii))
If nPos > 0 Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then
   DigNumEnt = KeyAscii
Else
   Beep
   DigNumEnt = 0
End If
End Function

'***************************************************************
'Valida el ingreso de números decimales en un TextBox
'---------------------------------------------------------------
Public Function DigNumDec(CTRLTextBox As TextBox, KeyAscii As Integer) As Integer
Dim nPos As Integer
nPos = InStr(".0123456789", Chr(KeyAscii))
If nPos > 0 Then
   If Chr(KeyAscii) = "." And InStr(CTRLTextBox, Chr(KeyAscii)) > 0 Then
      Beep
      DigNumDec = 0
   Else
      DigNumDec = KeyAscii
   End If
ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then
   DigNumDec = KeyAscii
Else
   Beep
   DigNumDec = 0
End If
End Function

'***************************************************************
'Para validar ingreso de dígitos de un campo tipo fecha
'---------------------------------------------------------------
Public Function DigFecha(TextBox As Control, KeyAscii As Integer) As Integer
Dim nPos As Integer
nPos = InStr("0123456789", Chr(KeyAscii))
If nPos > 0 Then
   If Len(Trim(TextBox)) = 2 Or Len(Trim(TextBox)) = 5 Then
      TextBox = TextBox & "/"
      TextBox.SelStart = Len(TextBox)
   End If
   DigFecha = KeyAscii
ElseIf KeyAscii = vbKeyBack Or KeyAscii = vbKeyTab Then
   DigFecha = KeyAscii
Else
   Beep
   DigFecha = 0
End If
End Function


Public Function HayPersonasCosto(ByVal pnMovNro As Long, ByRef cArchivo As String) As Boolean
   On Error GoTo CargaOpeGruErr
   Dim oCon As DConecta
   Set oCon = New DConecta
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Dim Sql As String
   
   cArchivo = ""
   HayPersonasCosto = False
   
   If oCon.AbreConexion() Then
   
      Sql = " select top 1 nMovNro from MovGastoPersonas where  nMovNro = " & pnMovNro
      Set rs = oCon.CargaRecordSet(Sql)
      If Not rs.EOF Then
         HayPersonasCosto = True
         cArchivo = "MovGastoPersonas"
         Exit Function
      End If
   
      Sql = " select top 1 nMovNro from MovGastoPersonasPorcent where  nMovNro = " & pnMovNro
      Set rs = oCon.CargaRecordSet(Sql)
      If Not rs.EOF Then
         HayPersonasCosto = True
         cArchivo = "MovGastoPersonasPorcent"
         Exit Function
      End If
      
      oCon.CierraConexion
   End If
   Set oCon = Nothing
   Exit Function
CargaOpeGruErr:
   MsgBox "Error : " + Err.Description
   'Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

'Ñique - funcion
Public Function GetPersonasMovGasto(ByVal pnMovNro As Long, ByVal pnMovItem As Integer, ByVal cArchivo As String) As ADODB.Recordset
Dim oCon As DConecta, rs As ADODB.Recordset, cSQL As String
On Error GoTo CargaOpeGruErr

Set oCon = New DConecta
Set rs = New ADODB.Recordset

Set GetPersonasMovGasto = Nothing

If oCon.AbreConexion() Then
   cSQL = "select cPersCod,'' as cAgeCod, '' as PrdCod ,nImporte as nMonto from " & cArchivo & " " & _
          " where  nMovNro = " & pnMovNro & " and nMovItem = " & pnMovItem & " "
   Set rs = oCon.CargaRecordSet(cSQL, adLockOptimistic)
   oCon.CierraConexion
End If

Set GetPersonasMovGasto = rs
Set rs = Nothing
Set oCon = Nothing
Exit Function

CargaOpeGruErr:
   MsgBox "Error : " + Err.Description
   'Call RaiseError(MyUnhandledError, "DOperacion:CargaOpeGru Method")
End Function

'Ñique Funcion
Public Function GetNombrePersona(psPersCod As String) As String
Dim Sql As String
Dim rs As New ADODB.Recordset
Dim oConect As New DConecta

GetNombrePersona = ""
If oConect.AbreConexion = False Then Exit Function

Sql = "Select cPersNombre from Persona where cPersCod = '" & psPersCod & "'"
Set rs = oConect.CargaRecordSet(Sql)
If Not rs.EOF Then
   GetNombrePersona = Replace(rs!cPersNombre, "/", " ")
End If
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function GetTipoCredito(psCodigo As String) As String
Dim sSQL As String, rs As New ADODB.Recordset
Dim cTipo As String, oConect As New DConecta
GetTipoCredito = ""
If oConect.AbreConexion = False Then Exit Function
cTipo = Mid(psCodigo, 3, 3)
sSQL = "Select cProductoDesc from DBCostos.dbo.Productos where cCodEQ = '" & cTipo & "'"
Set rs = oConect.CargaRecordSet(sSQL)
If Not rs.EOF Then
   GetTipoCredito = rs!cProductoDesc
End If
oConect.CierraConexion
Set oConect = Nothing
End Function

Public Function EliminaElemento(vMatriz() As ADODB.Recordset, vMaxNroElementos As Integer, vIndiceElemento As Integer) As Boolean
Dim i As Integer, rs As New ADODB.Recordset
EliminaElemento = False
For i = vIndiceElemento To vMaxNroElementos
    Set vMatriz(i) = vMatriz(i + 1)
Next i
Set vMatriz(vMaxNroElementos) = Nothing
EliminaElemento = True
End Function

Public Function GetBitCosto(vAgeCod As String) As Boolean
Dim sSQL As String, rs As New ADODB.Recordset, oConect As New DConecta

GetBitCosto = False
If oConect.AbreConexion = False Then Exit Function

sSQL = "Select nActivo from DBCmactAux.dbo.AgeCostoPermiso where cAgeCod = '" & vAgeCod & "'"

Set rs = oConect.CargaRecordSet(sSQL)
If Not rs.EOF Then
   GetBitCosto = IIf(rs!nActivo = 1, True, False)
End If
oConect.CierraConexion
Set oConect = Nothing
End Function


Public Function GetBitCostoArea(vAgeCod As String, vAreaCod As String) As Boolean
Dim sSQL As String, rs As New ADODB.Recordset, oConect As New DConecta

GetBitCostoArea = False
If oConect.AbreConexion = False Then Exit Function

sSQL = "Select nActivo from DBCmactAux.dbo.AgeAreaCostoPermiso " & _
       " where cAgeCod = '" & vAgeCod & "' and cAreaCod = '" & vAreaCod & "'"

Set rs = oConect.CargaRecordSet(sSQL)
If Not rs.EOF Then
   GetBitCostoArea = True
End If
oConect.CierraConexion
Set oConect = Nothing
End Function

'*********************************************************************************
'FUNCION PARA AJUSTAR EL TEXTO A UN ANCHO DETERMINADO
'
'versión del 06 de Abril del 2005
'E.S.Ñ.M.
'
'Devuelve una matriz con caracteres al ancho fijo indicado
'esta matriz se puede usar para la impresión de texto justificado
'
'NOTA: se ha probado solo para los anchos indicados,
'      otros anchos traen problemas (por depurar)
'*********************************************************************************

Public Function AjustaTexto(vTexto As String, vAncho As tAnchoAjuste) As String()
Dim v() As String
Dim i As Integer, t As Integer, n As Integer
Dim s As String, w As String, P As Integer
Dim k As Integer
Dim nRes As Integer

s = Trim(vTexto)
t = Len(Trim(vTexto))

n = t / vAncho
nRes = t Mod vAncho
If nRes < (vAncho / 2) Then
   n = n + 1
End If

ReDim v(1 To n)
For i = 1 To n
    v(i) = Mid(s, 1, vAncho)
    If Len(s) > vAncho Then
       s = Mid(s, vAncho + 1, Len(s) - vAncho)
    End If
Next

For i = 1 To n
    If Right(v(i), 1) <> " " And i < n Then
       If Left(v(i + 1), 1) <> " " Then
          Do While Right(v(i), 1) <> " "
             PasaLetra v(i), v(i + 1)
          Loop
       End If
    End If
    P = Len(Trim(v(i)))
    If Len(v(i)) > vAncho Then
       If i = n Then ReDim Preserve v(1 To i + 1)
       Do While Right(v(i), 1) <> " " Or Len(v(i)) > vAncho
          PasaLetra v(i), v(i + 1)
       Loop
    End If
    v(i) = Trim(v(i))
    P = Len(v(i))
    Do While Len(v(i)) < vAncho
       P = InStrRev(Trim(Mid(v(i), 1, P)), " ")
       If P = 0 Then
          If i = n Then Exit Do
          P = Len(v(i))
       Else
          v(i) = InsertaEspacio(v(i), P)
          If Len(v(i)) = vAncho Then Exit Do
       End If
    Loop
Next
AjustaTexto = v
End Function

Private Sub PasaLetra(ByRef vLineaOrig As String, ByRef vLineaDest As String)
vLineaDest = Right(vLineaOrig, 1) + vLineaDest
vLineaOrig = Mid(vLineaOrig, 1, Len(vLineaOrig) - 1)
End Sub

Private Function InsertaEspacio(vCadena As String, vPosIns As Integer) As String
Dim cResto As String
InsertaEspacio = ""
If vPosIns > 0 Then
   cResto = Mid(vCadena, vPosIns, Len(vCadena))
   vCadena = Trim(Mid(vCadena, 1, vPosIns - 1) + " " + cResto)
End If
InsertaEspacio = vCadena
End Function

Public Function UbigeoDescCompleto(vRaiz As String, vCodigo As String, Optional vOrden As Integer = 1) As String
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta

UbigeoDescCompleto = ""
If oConn.AbreConexion Then
   sSQL = "select cDpto=coalesce(rtrim(d.cUbigeoDescripcion),''),cProv=coalesce(rtrim(p.cUbigeoDescripcion),''),cDist=coalesce(rtrim(q.cUbigeoDescripcion),''),cZona = rtrim(u.cUbigeoDescripcion) " & _
   " from UbicacionGeografica u " & _
   " left join (select cCodigo='" & vRaiz & "'+substring(cUbigeoCod,2,2), cUbigeoDescripcion from UbicacionGeografica  where left(cUbigeoCod,1)='1') d on substring(u.cUbigeoCod,1,3)=d.cCodigo " & _
   " left join (select cCodigo='" & vRaiz & "'+substring(cUbigeoCod,2,4), cUbigeoDescripcion from UbicacionGeografica  where left(cUbigeoCod,1)='2') p on substring(u.cUbigeoCod,1,5)=p.cCodigo " & _
   " left join (select cCodigo='" & vRaiz & "'+substring(cUbigeoCod,2,6), cUbigeoDescripcion from UbicacionGeografica  where left(cUbigeoCod,1)='3') q on substring(u.cUbigeoCod,1,7)=q.cCodigo " & _
   " where u.cUbigeoCod = '" & vCodigo & "' "
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      If vOrden = 0 Then
         UbigeoDescCompleto = rs!cDpto + " - " + rs!cProv + " - " + rs!cDist + " - " + rs!cZona
      Else
         UbigeoDescCompleto = IIf(vRaiz = "4", rs!cZona, "")
         UbigeoDescCompleto = UbigeoDescCompleto + IIf(Len(rs!cDist) > 0, IIf(Len(UbigeoDescCompleto) > 0, " - ", "") + rs!cDist, "")
         UbigeoDescCompleto = UbigeoDescCompleto + IIf(Len(rs!cProv) > 0, IIf(Len(UbigeoDescCompleto) > 0, " - ", "") + rs!cProv, "")
         UbigeoDescCompleto = UbigeoDescCompleto + IIf(Len(rs!cDpto) > 0, IIf(Len(UbigeoDescCompleto) > 0, " - ", "") + rs!cDpto, "")
      End If
   End If
End If

End Function

Public Function GetUbigeoCorto(psCodigo As String) As String
Dim sSQL As String
Dim rs As New ADODB.Recordset
Dim oConn As New DConecta

GetUbigeoCorto = ""
sSQL = "select cUbigeoDescripcion from UbicacionGeografica where cUbigeoCod ='" & psCodigo & "'"
If oConn.AbreConexion Then
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      GetUbigeoCorto = rs!cUbigeoDescripcion
   End If
End If
End Function

Public Function DatosGrabacion() As String
DatosGrabacion = ""
DatosGrabacion = Format(gdFecSis, "YYYYMMDD") + Format(Time, "HHMMSS") + gsCodAge + gsCodUser
End Function

Public Function GetBSUnidad(psBSCod As String) As String
Dim oConn As New DConecta, sSQL As String, rs As New ADODB.Recordset
GetBSUnidad = ""
If oConn.AbreConexion Then
   'sSQL = "select cConsDescripcion as cUnidad from Constante where nConsCod = 1019 and nConsValor = " & psValor & " "
   sSQL = "select t.cConsDescripcion from BienesServicios b inner join " & _
          " (select nConsValor as nBSUnidad, cConsDescripcion from Constante where nConsCod = 1019) t " & _
          " on b.nBSUnidad = t.nBSUnidad where b.cBSCod = '" & psBSCod & "'"
   Set rs = oConn.CargaRecordSet(sSQL)
   If Not rs.EOF Then
      GetBSUnidad = rs!cUnidad
   End If
End If
End Function

