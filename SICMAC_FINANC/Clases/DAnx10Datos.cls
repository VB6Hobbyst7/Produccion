VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DAnx10Datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim dbConec As DConecta
Public Sub ActualizaAnx10Datos(psAgeCod As String, pscPais As String, pscDepa As String, pscProv As String, pscDist As String, pscTipo As String, pscCodigo As String)
   Dim psSql As String
   Dim sExiste As Boolean
   On Error GoTo ActualizaAnx10DatosErr
   Set dbConec = New DConecta
   If dbConec.AbreConexion() Then
   
        sExiste = ExisteAnx10Datos(psAgeCod)
        
        If sExiste = True Then
            If Right(pscTipo, 1) = "*" Then
                psSql = "UPDATE Anx10Datos Set cPais = '" & pscPais & "', cDepa = '" & pscDepa & "', cProv='" & pscProv & "', cDist = '" & pscDist & "', cCodigo='" & pscCodigo & "' Where cAgeCod = '" & psAgeCod & "'"
            Else
                psSql = "UPDATE Anx10Datos Set cPais = '" & pscPais & "', cDepa = '" & pscDepa & "', cProv='" & pscProv & "', cDist = '" & pscDist & "', cTipo = '" & pscTipo & "' Where cAgeCod = '" & psAgeCod & "'"
            End If
        Else
            psSql = "INSERT INTO Anx10Datos (cAgeCod, cPais, cDepa, cProv, cDist, cTipo, cCodigo) VALUES('" & psAgeCod & "', '" & pscPais & "', '" & pscDepa & "', '" & pscProv & "', '" & pscDist & "', '" & pscTipo & "', '" & pscCodigo & "')"
        End If
        
        dbConec.Ejecutar psSql
        dbConec.CierraConexion
   End If
   Set dbConec = Nothing
   Exit Sub
ActualizaAnx10DatosErr:
   Call RaiseError(MyUnhandledError, "DAnx10Datos:ActualizaAnx10Datos Method")
End Sub

Public Function ExisteAnx10Datos(psAgeCod As String) As Boolean
   Dim psSql As String
   Dim rs As ADODB.Recordset
   psSql = "select cpais from Anx10Datos where cAgeCod='" & psAgeCod & "'"
   
   On Error GoTo ActualizaAnx10DatosErr
                 
        Set rs = dbConec.CargaRecordSet(psSql)
        If rs.BOF Then
            ExisteAnx10Datos = False
        Else
            ExisteAnx10Datos = True
        End If
         
    
   Exit Function
ActualizaAnx10DatosErr:
   Call RaiseError(MyUnhandledError, "DAnx10Datos:ActualizaAnx10Datos Method")
End Function

 
