VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGrupoEco1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim oCon As DConecta
Dim sSQL As String

Private Sub Class_Initialize()

Set oCon = New DConecta

If oCon.AbreConexion = False Then
    Call RaiseError(MyUnhandledError, "DMov:Initialize Method. Error en Conexion a Base de datos")
End If

End Sub

Private Sub Class_Terminate()
oCon.CierraConexion
Set oCon = Nothing
End Sub


Public Function GetG1(ByRef nCantidad As Long) As ADODB.Recordset

sSQL = "Select PGE.cGECod, PGE.cGENom, PGE.nGETipo, C1.cConsDescripcion, PGE.cCodReporte, 'Reporte ' + PGE.cCodReporte as cDesReporte "
sSQL = sSQL & " From PersGrupoEcon PGE "
sSQL = sSQL & "      Inner Join Constante C1 "
sSQL = sSQL & "         On PGE.nGETipo=C1.nConsValor "
sSQL = sSQL & " Where C1.nConsCod = 4027 "
sSQL = sSQL & " Order By PGE.cGECod "

Set GetG1 = oCon.CargaRecordSet(sSQL)

nCantidad = GetG1.RecordCount

End Function

Public Function GetG2(ByRef nCantidad As Long, Optional cFiltrocGECod As String = "") As ADODB.Recordset

sSQL = "Select  PGE0.cGECod, PGE0.cGENom, PGE0.nGETipo, C2.cConsDescripcion, "
sSQL = sSQL & "    PGE.cGECod , PGE.cPersCodRel, P.cPersNombre, PGE.nPrdPersRelac, C1.cConsDescripcion as cConsDescripcion1, PGE.nRela1, PGE.nRela2, PGE.nRela3, PGE.cTexto "
sSQL = sSQL & " From PersGrupoEcon PGE0 "
sSQL = sSQL & "    Inner Join persgerelacion PGE "
sSQL = sSQL & "        On PGE.cGECod=PGE0.cGECod "
sSQL = sSQL & "    Inner Join Constante C2 "
sSQL = sSQL & "        On PGE0.nGETipo=C2.nConsValor "
sSQL = sSQL & "    Inner Join Constante C1 "
sSQL = sSQL & "        On PGE.nPrdPersRelac=C1.nConsValor "
sSQL = sSQL & "    Inner Join Persona P "
sSQL = sSQL & "        On PGE.cPersCodRel=P.cPersCod "
sSQL = sSQL & " Where C2.nConsCod = 4027 And C1.nConsCod = 4028 "

If Len(Trim(cFiltrocGECod)) > 0 Then
    sSQL = sSQL & " And PGE.cGECod='" & cFiltrocGECod & "' "
End If

sSQL = sSQL & " Order By PGE0.cGECod, PGE.cGECod "

Set GetG2 = oCon.CargaRecordSet(sSQL)

nCantidad = GetG2.RecordCount

End Function

Public Function GetG3(ByRef nCantidad As Long, Optional cFiltrocGECod As String = "", Optional cFiltroPersCodRel As String)

sSQL = "Select  PGE0.cGECod, PGE0.cGENom, PGE0.nGETipo, C2.cConsDescripcion, "
sSQL = sSQL & "     PGE.cGECod, PGE.cPersCodRel, P.cPersNombre, PGE.nPrdPersRelac, C1.cConsDescripcion as cConsDescripcion1, "
sSQL = sSQL & "     PGV.cPersCodRel , PGV.cPersCodVinc, P1.cPersNombre as cPersNombre1, PGV.nCargo, PGV.nParticip, C0.cConsDescripcion as cConsDescripcion2, PGV.nRela1, PGV.nRela2, PGV.nRela3 "
sSQL = sSQL & " From PersGrupoEcon PGE0 "
sSQL = sSQL & "     Inner Join persgerelacion PGE "
sSQL = sSQL & "         On PGE.cGECod=PGE0.cGECod "
sSQL = sSQL & "     Inner Join PersGEPersVinc PGV "
sSQL = sSQL & "         On PGE.cGECod=PGV.cGECod And PGE.cPersCodRel=PGV.cPersCodRel "
sSQL = sSQL & "     Inner Join Constante C2 "
sSQL = sSQL & "         On PGE0.nGETipo=C2.nConsValor "
sSQL = sSQL & "     Inner Join Constante C1 "
sSQL = sSQL & "         On PGE.nPrdPersRelac=C1.nConsValor "
sSQL = sSQL & "     Inner Join Constante C0 "
sSQL = sSQL & "         On PGV.nCargo =C0.nConsValor "
sSQL = sSQL & "     Inner Join Persona P "
sSQL = sSQL & "         On PGE.cPersCodRel=P.cPersCod "
sSQL = sSQL & "     Inner Join Persona P1 "
sSQL = sSQL & "         On PGV.cPersCodVinc=P1.cPersCod "
sSQL = sSQL & " Where C2.nConsCod = 4027 And C1.nConsCod = 4028 And C0.nConsCod = 4029 "

If Len(Trim(cFiltrocGECod)) > 0 Then
    sSQL = sSQL & " And PGE.cGECod='" & cFiltrocGECod & "' "
End If
If Len(Trim(cFiltroPersCodRel)) > 0 Then
    sSQL = sSQL & " And PGE.cPersCodRel='" & cFiltroPersCodRel & "' "
End If

sSQL = sSQL & " Order By PGV.nCargo, PGE0.cGECod, PGE.cGECod, P1.cPersNombre "

Set GetG3 = oCon.CargaRecordSet(sSQL)

nCantidad = GetG3.RecordCount

End Function

Public Sub GrabaPersGrupoEcon(ByVal cGEcod As String, ByVal cGENom As String, ByVal nGETipo As Integer, ByVal cCodReporte As String, bNuevoEditar As Integer)

Dim nCantidad As Long
Dim rs As New ADODB.Recordset
Dim cCodigo As String

    On Error GoTo ErrorGP
    
    sSQL = "select nMayor=isnull(max(convert(integer, cGECod)),0) from persgrupoecon"
    
    Set rs = oCon.CargaRecordSet(sSQL)
    If rs.BOF Then
        nCantidad = 0
    Else
        nCantidad = rs!nMayor
    End If
    Set rs = Nothing
    
    cCodigo = Right("000000" & Trim(Str(nCantidad + 1)), 6)
    
    If bNuevoEditar = 1 Then
        sSQL = "Insert Into persgrupoecon(cGECod, cGENom, nGETipo, cCodReporte) "
        sSQL = sSQL & " values('" & Trim(cCodigo) & "', '" & Trim(cGENom) & "', " & nGETipo & ", '" & cCodReporte & "')"
    ElseIf bNuevoEditar = 2 Then
        sSQL = "Update persgrupoecon Set cGENom ='" & Trim(cGENom) & "', nGETipo ='" & nGETipo & "' "
        sSQL = sSQL & "Where cGECod='" & Trim(cGEcod) & "' "
    End If
    
    oCon.Ejecutar sSQL
    
    Exit Sub

ErrorGP:
    Call RaiseError(MyUnhandledError, "DGrupoEco1:Metodo Grabar")
    
End Sub

Public Sub EliminaPersGrupoEcon(ByVal cGEcod As String)

    On Error GoTo ErrorGP
    
    sSQL = "Delete persgrupoecon "
    sSQL = sSQL & "Where cGECod='" & Trim(cGEcod) & "' "
    
    oCon.Ejecutar sSQL
    
    Exit Sub

ErrorGP:
    Call RaiseError(MyUnhandledError, "DGrupoEco1:Metodo Grabar")
    
End Sub

Public Function getNumHijospersgrupoecon(cGEcod As String) As Long

    Dim rs As New ADODB.Recordset

    sSQL = "Select isnull(count(*),0) as ncantidad from persgerelacion where cGECod='" & Trim(cGEcod) & "' "

    Set rs = oCon.CargaRecordSet(sSQL)
    If rs.BOF Then
        getNumHijospersgrupoecon = 0
    Else
        getNumHijospersgrupoecon = rs!nCantidad
    End If
    Set rs = Nothing

End Function

Public Function getExistePersonaEnPersGERelacion(cGEcod As String, cPersCodRel As String) As Long

    Dim rs As New ADODB.Recordset

    sSQL = "Select * from PersGERelacion Where cGEcod='" & cGEcod & "'  and cPersCodrel='" & cPersCodRel & "' "

    Set rs = oCon.CargaRecordSet(sSQL)
    If rs.BOF Then
        getExistePersonaEnPersGERelacion = 0
    Else
        getExistePersonaEnPersGERelacion = 1
    End If
    Set rs = Nothing

End Function

Public Function GetExisteReporte(cCodReporte As String) As Integer

    Dim rs As New ADODB.Recordset

    sSQL = "Select * from persgrupoecon Where cCodReporte='" & cCodReporte & "' "

    Set rs = oCon.CargaRecordSet(sSQL)
    If rs.BOF Then
        GetExisteReporte = 0
    Else
        GetExisteReporte = 1
    End If
    Set rs = Nothing

End Function

Public Sub GrabaPersGeRelacion(cGEcod As String, cPersCodRelActual As String, cPersCodRelNuevo As String, nPrdPersRelac As Integer, cTexto As String, nrela1 As Integer, nrela2 As Integer, nrela3 As Integer, bNuevoEditar As Integer)

Dim nCantidad As Long

    On Error GoTo ErrorGP
    
    If bNuevoEditar = 1 Then
        sSQL = "Insert Into PersGeRelacion(cGECod, cPersCodRel, nPrdPersRelac, cTexto, nRela1, nRela2, nRela3) "
        sSQL = sSQL & " values('" & Trim(cGEcod) & "', '" & Trim(cPersCodRelNuevo) & "', " & nPrdPersRelac & ", '" & cTexto & "', " & nrela1 & ", " & nrela2 & ", " & nrela3 & ")"
    ElseIf bNuevoEditar = 2 Then
        sSQL = "Update PersGeRelacion Set cPersCodRel ='" & Trim(cPersCodRelNuevo) & "', nPrdPersRelac =" & nPrdPersRelac & ", cTexto='" & cTexto & "', nRela1= " & nrela1 & ", nRela2=" & nrela2 & ", nRela3=" & nrela3 & " "
        sSQL = sSQL & "Where cGECod='" & Trim(cGEcod) & "' And cPersCodRel='" & cPersCodRelActual & "' "
    End If
    
    oCon.Ejecutar sSQL
    
    Exit Sub

ErrorGP:
    Call RaiseError(MyUnhandledError, "DGrupoEco1:Metodo Grabar")
    
End Sub


Public Function getNumHijosPersGEPersVinc(cGEcod As String, cPersCodVinc As String) As Long

    Dim rs As New ADODB.Recordset

    sSQL = "Select isnull(count(*),0) as ncantidad from PersGEPersVinc where cGEcod='" & cGEcod & "' and cPersCodVinc='" & Trim(cPersCodVinc) & "' "

    Set rs = oCon.CargaRecordSet(sSQL)
    If rs.BOF Then
        getNumHijosPersGEPersVinc = 0
    Else
        getNumHijosPersGEPersVinc = rs!nCantidad
    End If
    Set rs = Nothing

End Function

Public Sub EliminaPersGERelacion(ByVal cGEcod As String, ByVal cPersCodRel As String)

    On Error GoTo ErrorGP
    
    sSQL = "Delete PersGERelacion "
    sSQL = sSQL & " Where cGECod='" & Trim(cGEcod) & "' And cPersCodRel='" & cPersCodRel & "' "
    
    oCon.Ejecutar sSQL
    
    Exit Sub

ErrorGP:
    Call RaiseError(MyUnhandledError, "DGrupoEco1:Metodo Grabar")
    
End Sub

Public Sub GrabaPersGEPersVinc(cGEcod As String, cPersCodRel As String, cPersCodVincActual As String, cPersCodVincNuevo As String, nCargo As Long, nParticip As Double, nrela1 As Integer, nrela2 As Integer, nrela3 As Integer, bNuevoEditar As Integer)

Dim nCantidad As Long

    On Error GoTo ErrorGP
    
    If bNuevoEditar = 1 Then
        sSQL = "Insert Into PersGEPersVinc(cGEcod, cPersCodRel, cPersCodVinc, nCargo, nParticip, nRela1, nRela2, nRela3)"
        sSQL = sSQL & " values('" & Trim(cGEcod) & "', '" & Trim(cPersCodRel) & "', '" & Trim(cPersCodVincNuevo) & "', " & nCargo & ", " & nParticip & ", " & nrela1 & ", " & nrela2 & ", " & nrela3 & ")"
    ElseIf bNuevoEditar = 2 Then
        sSQL = "Update PersGEPersVinc Set cPersCodVinc ='" & Trim(cPersCodVincNuevo) & "', nCargo =" & nCargo & ", nParticip=" & nParticip & ", nRela1=" & nrela1 & ", nRela2=" & nrela2 & ", nRela3=" & nrela3 & " "
        sSQL = sSQL & "Where cGEcod='" & cGEcod & "' And cPersCodRel='" & Trim(cPersCodRel) & "' And cPersCodVinc='" & cPersCodVincActual & "' "
    End If
    
    oCon.Ejecutar sSQL
    
    Exit Sub

ErrorGP:
    Call RaiseError(MyUnhandledError, "DGrupoEco1:Metodo Grabar")
    
End Sub

Public Sub EliminaPersGEPersVinc(ByVal cGEcod As String, ByVal cPersCodRel As String, ByVal cPersCodVinc As String)
   

    On Error GoTo ErrorGP
    
    sSQL = "Delete PersGEPersVinc "
    sSQL = sSQL & " Where cGEcod='" & cGEcod & "' And cPersCodRel='" & Trim(cPersCodRel) & "' And cPersCodVinc='" & cPersCodVinc & "' "
    
    oCon.Ejecutar sSQL
    
    Exit Sub

ErrorGP:
    Call RaiseError(MyUnhandledError, "DGrupoEco1:Metodo Grabar")
    
End Sub

