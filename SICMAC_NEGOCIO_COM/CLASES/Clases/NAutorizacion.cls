VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Aucperscod As String
Dim Aucpersnombre As String
Dim Audpersnaccreac As String
Dim Aucnivel As String
Dim Aucrhcargodescripcion As String
Dim Aucrhcargocod As String
Dim AunMontoIniSol As Double
Dim AunMontoFinSol As Double
Dim AunMontoIniDol As Double
Dim AunMontoFinDol As Double


'Public Sub DevInfoAutorizacionNo(ByVal Cuser As String, ByVal copecod As String)
' Dim aut As DAutorizacion
' Dim rs As ADODB.Recordset
'
'     Set aut = New DAutorizacion
'     Set rs = aut.InfoAutorizacion(Cuser, copecod)
'   If Not (rs.EOF Or rs.BOF) Then
'            Aucperscod = rs!cPersCod
'            Aucpersnombre = rs!cpersnombre
'            Audpersnaccreac = rs!dPersNacCreac
'            Aucnivel = rs!cNivel
'            Aucrhcargodescripcion = rs!crhcargodescripcion
'            Aucrhcargocod = rs!cRHCargoCod
'            AunMontoIniSol = rs!nMontoIniSol
'            AunMontoFinSol = rs!nMontoFinSol
'            AunMontoIniDol = rs!nMontoIniDol
'            AunMontoFinDol = rs!nMontoFinDol
'    End If
'
'
'End Sub

 Public Sub DevInfoAutorizacion(ByVal Cuser As String, ByVal copecod As String)
  Dim oconecta As DConecta, rs As ADODB.Recordset
  Dim ssql As String
     Set oconecta = New DConecta
     Set rs = New ADODB.Recordset
     
      ssql = "Select p.cperscod,p.cpersnombre,p.dpersnaccreac,c.cnivel,rhct.crhcargodescripcion,rhc.crhcargocod,nMontoIniSol,nMontoFinSol, nMontoIniDol,nMontoFinDol  "
      ssql = ssql & "  from persona p"
      ssql = ssql & "  inner join rrhh  rh on rh.cperscod=p.cperscod"
      ssql = ssql & "  inner join (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod "
      ssql = ssql & "  where rh1.cuser='" & Cuser & "' order by drhcargoFecha desc) rhc on rhc.cperscod=p.cperscod"
      ssql = ssql & "  inner join rhcargostabla rhct on rhct.crhcargocod=rhc.crhcargocod "
      ssql = ssql & "  inner join capautorizacionrango c on c.crhcargocod=rhc.crhcargocod "
      ssql = ssql & "  where  rh.nrhestado='201' and rh.cUser='" & Cuser & "' and c.copecod='" & copecod & "'"
On Error GoTo MensaError
      Set oconecta = New DConecta
      oconecta.AbreConexion
      Set rs = oconecta.Ejecutar(ssql)  '  oConecta.ConexionActiva.Execute(sSql)
        If rs.State = 1 Then
          If Not (rs.EOF Or rs.BOF) Then
            Aucperscod = rs!cPersCod
            Aucpersnombre = rs!cpersnombre
            Audpersnaccreac = rs!dPersNacCreac
            Aucnivel = rs!cNivel
            Aucrhcargodescripcion = rs!crhcargodescripcion
            Aucrhcargocod = rs!cRHCargoCod
            AunMontoIniSol = rs!nMontoIniSol
            AunMontoFinSol = rs!nMontoFinSol
            AunMontoIniDol = rs!nMontoIniDol
            AunMontoFinDol = rs!nMontoFinDol
           End If
           rs.Close
         End If
         Set rs = Nothing
      oconecta.CierraConexion
      Set oconecta = Nothing
    Exit Sub
MensaError:
        Call RaiseError(MyUnhandledError, "DAutorizacion:IAutRango Method")
     
End Sub


Property Get DAucPersCod() As String
         DAucPersCod = Aucperscod
End Property

Property Get DAucpersnombre() As String
         DAucpersnombre = Aucpersnombre
End Property
Property Get DAudpersnaccreac() As String
         DAudpersnaccreac = Audpersnaccreac
End Property
Property Get DAucnivel() As String
         DAucnivel = Aucnivel
End Property
Property Get DAucrhcargodescripcion() As String
         DAucrhcargodescripcion = Aucrhcargodescripcion
End Property

Property Get DAucrhcargocod() As String
         DAucrhcargocod = Aucrhcargocod
End Property

Property Get DAunMontoIniSol() As Double
    DAunMontoIniSol = AunMontoIniSol
End Property
Property Get DAunMontoFinSol() As Double
    DAunMontoFinSol = AunMontoFinSol
End Property
Property Get DAunMontoIniDol() As Double
    DAunMontoIniDol = AunMontoIniDol
End Property
Property Get DAunMontoFinDol() As Double
    DAunMontoFinDol = AunMontoFinDol
End Property

 
             
            
             
            

