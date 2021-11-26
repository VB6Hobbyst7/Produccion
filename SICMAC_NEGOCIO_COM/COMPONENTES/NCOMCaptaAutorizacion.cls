VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NCOMCaptAutorizacion"
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


 Public Sub DevInfoAutorizacion(ByVal cUser As String, ByVal copecod As String)
  Dim oConecta As COMConecta.DCOMConecta, rs As ADODB.Recordset
  Dim oError As New COMConecta.COMErrorHandling
  
  Dim sSQL As String
     Set oConecta = New COMConecta.DCOMConecta
     Set rs = New ADODB.Recordset
     
      sSQL = "Select p.cperscod,p.cpersnombre,p.dpersnaccreac,c.cnivel,rhct.crhcargodescripcion,rhc.crhcargocod,nMontoIniSol,nMontoFinSol, nMontoIniDol,nMontoFinDol  "
      sSQL = sSQL & "  from persona p"
      sSQL = sSQL & "  inner join rrhh  rh on rh.cperscod=p.cperscod"
      sSQL = sSQL & "  inner join (Select top 1 rhc1.crhcargocod,rhc1.cperscod from rhcargos rhc1 inner join rrhh rh1 on rh1.cperscod=rhc1.cperscod "
      sSQL = sSQL & "  where rh1.cuser='" & cUser & "' order by drhcargoFecha desc) rhc on rhc.cperscod=p.cperscod"
      sSQL = sSQL & "  inner join rhcargostabla rhct on rhct.crhcargocod=rhc.crhcargocod "
      sSQL = sSQL & "  inner join capautorizacionrango c on c.crhcargocod=rhc.crhcargocod "
      sSQL = sSQL & "  where  rh.nrhestado='201' and rh.cUser='" & cUser & "' and c.copecod='" & copecod & "'"
On Error GoTo MensaError
      Set oConecta = New COMConecta.DCOMConecta
      oConecta.AbreConexion
      Set rs = oConecta.Ejecutar(sSQL)  '  oConecta.ConexionActiva.Execute(sSql)
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
      oConecta.CierraConexion
      Set oConecta = Nothing
    Exit Sub
MensaError:
        Call oError.RaiseError(oError.MyUnhandledError, "DAutorizacion:IAutRango Method")
     
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

