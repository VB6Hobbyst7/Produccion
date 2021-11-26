VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variables locales para almacenar los valores de las propiedades
Public sPersNombre As String
Public sPersCod As String
Public dPersNacCreac As Date
Public sPersDireccDomicilio As String
Public sPersTelefono As String
Public sPersPersoneria As String
Public sPersPersoneriaDesc As String
Public sPersIdnroDNI As String
Public sPersIdnroRUC As String
Public sPersIdnroOtro As String
Public sPersNatSexo As String
Public DocsPers As ADODB.Recordset

Public Sub CargaDatos(ByVal psPersCod As String, psPersNombre As String, _
    ByVal pdPersNacCreac As Date, ByVal psPersDireccDomicilio As String, _
    ByVal psPersTelefono As String, ByVal psPersPersoneria As String, _
    ByVal psPersIdnroDNI As String, ByVal psPersIdnroRUC As String, _
    ByVal psPesrIDNroOtro As String, ByVal psPersNatSexo As String)
Dim Conn As DConecta
Dim R1 As ADODB.Recordset
Dim sSQL As String
Dim sPersoneriaDesc As String
   Set Conn = New DConecta
   If Not Conn.AbreConexion Then
        Set Conn = Nothing
        Exit Sub
   End If
   
   sSQL = "Select * from Constante where nConsCod = " & Trim(Str(gPersPersoneria)) & " And nConsValor = " & Trim(psPersPersoneria)
   Set R1 = Conn.CargaRecordSet(sSQL)
   If Not R1.BOF And Not R1.EOF Then
        sPersoneriaDesc = Trim(R1!cConsDescripcion)
   Else
        sPersoneriaDesc = ""
   End If
   R1.Close
   Set R1 = Nothing
   
   sSQL = "Select C.cConsDescripcion  as Tipo, PID.cPersIDNro,PID.cPersIDTpo from PersID PID left join Constante C on convert(int,PID.cPersIDTpo) = C.nConsValor where PID.cPersCod = '" & Trim(psPersCod) & "' And C.nConsCod = " & Trim(Str(gPersIdTipo))
   Set R1 = Conn.CargaRecordSet(sSQL)
   Conn.CierraConexion
   Set Conn = Nothing
   
    sPersNombre = PstaNombre(psPersNombre)
    sPersCod = psPersCod
    dPersNacCreac = pdPersNacCreac
    sPersDireccDomicilio = psPersDireccDomicilio
    sPersTelefono = psPersTelefono
    sPersPersoneria = psPersPersoneria
    sPersPersoneriaDesc = sPersoneriaDesc
    sPersIdnroDNI = psPersIdnroDNI
    sPersIdnroRUC = psPersIdnroRUC
    sPersIdnroOtro = psPesrIDNroOtro
    sPersNatSexo = psPersNatSexo
    Set DocsPers = R1
End Sub

Public Sub ObtieneClientexCodigo(ByVal psPersCod As String)
Dim oPersona As DPersonas
Dim R As ADODB.Recordset
    
    On Error GoTo ErrorObtieneClientexCodigo
    Set oPersona = New DPersonas
    Set R = oPersona.BuscaCliente(psPersCod, BusquedaCodigo)
    Call CargaDatos(R!cPerscod, R!cpersNombre, R!dPersNacCreac, R!cPersDireccDomicilio, R!cPersTelefono, R!nperspersoneria, Trim(R!cPersIDNroDNI), IIf(IsNull(R!cPersIDNroRuc), "", R!cPersIDNroRuc), IIf(IsNull(R!cPersIDNro), "", R!cPersIDNro), IIf(IsNull(R!cPersnatSexo), "", R!cPersnatSexo))
    R.Close
    Set R = Nothing
    Exit Sub
    
ErrorObtieneClientexCodigo:
    Err.Raise Err.Number, "RecuperaPersona", Err.Description
End Sub


