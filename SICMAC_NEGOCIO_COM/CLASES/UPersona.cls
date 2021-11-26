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
Public sPerEstado As String
Public DocsPers As ADODB.Recordset

Public Sub CargaDatos(ByVal psPersCod As String, psPersNombre As String, _
    ByVal pdPersNacCreac As Date, ByVal psPersDireccDomicilio As String, _
    ByVal psPersTelefono As String, ByVal psPersPersoneria As String, _
    ByVal psPersIdnroDNI As String, ByVal psPersIdnroRUC As String, _
    ByVal psPesrIDNroOtro As String, ByVal psPersNatSexo As String)
Dim Conn As DConecta
Dim R1 As ADODB.Recordset
Dim sSql As String
Dim sPersoneriaDesc As String
   Set Conn = New DConecta
   If Not Conn.AbreConexion Then
        Set Conn = Nothing
        Exit Sub
   End If
   
   sSql = "Select * from Constante where nConsCod = " & Trim(Str(gPersPersoneria)) & " And nConsValor = " & Trim(psPersPersoneria)
   Set R1 = Conn.CargaRecordSet(sSql)
   If Not R1.BOF And Not R1.EOF Then
        sPersoneriaDesc = Trim(R1!cConsDescripcion)
   Else
        sPersoneriaDesc = ""
   End If
   R1.Close
   Set R1 = Nothing
   
   sSql = "Select C.cConsDescripcion  as Tipo, PID.cPersIDNro,PID.cPersIDTpo from PersID PID left join Constante C on convert(int,PID.cPersIDTpo) = C.nConsValor where PID.cPersCod = '" & Trim(psPersCod) & "' And C.nConsCod = " & Trim(Str(gPersIdTipo))
   Set R1 = Conn.CargaRecordSet(sSql)
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
    
    If IsNull(R!cPersIDNroDNI) And IsNull(R!cPersIDNroRuc) Then
        MsgBox "El cliente no tiene documento de identidad ni RUC.", vbInformation, "Aviso"
    End If
    
    Call CargaDatos(R!cPersCod, R!cPersNombre, R!dPersNacCreac, R!cPersDireccDomicilio, IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono), R!nPersPersoneria, Trim(IIf(IsNull(R!cPersIDNroDNI), "", R!cPersIDNroDNI)), IIf(IsNull(R!cPersIDNroRuc), "", R!cPersIDNroRuc), IIf(IsNull(R!cPersIDNro), "", R!cPersIDNro), IIf(IsNull(R!cPersnatSexo), "", R!cPersnatSexo))
    R.Close
    Set R = Nothing
    Exit Sub

ErrorObtieneClientexCodigo:
    Err.Raise Err.Number, "RecuperaPersona", Err.Description
End Sub

Public Sub ObtieneConyugeDePersona(ByVal psPersCod As String)
Dim oPersona As DPersonas
Dim R As ADODB.Recordset
Dim sPersCodCony As String


    Set oPersona = New DPersonas
    sPersCodCony = oPersona.BuscaPersCodConyuge(psPersCod)
    sPersCodCony = Trim(sPersCodCony)
    If sPersCodCony <> "" Then
        Set R = oPersona.BuscaCliente(sPersCodCony, BusquedaCodigo)
        Call CargaDatos(R!cPersCod, R!cPersNombre, R!dPersNacCreac, R!cPersDireccDomicilio, IIf(IsNull(R!cPersTelefono), "", R!cPersTelefono), R!nPersPersoneria, Trim(R!cPersIDNroDNI), IIf(IsNull(R!cPersIDNroRuc), "", R!cPersIDNroRuc), IIf(IsNull(R!cPersIDNro), "", R!cPersIDNro), IIf(IsNull(R!cPersnatSexo), "", R!cPersnatSexo))
        R.Close
        Set R = Nothing
    Else
        sPersCod = ""
    End If
    Exit Sub
    
ErrorObtieneClientexCodigo:
    Err.Raise Err.Number, "RecuperaPersona", Err.Description
End Sub

Public Function ObtieneComentario(ByVal psPersCod As String) As ADODB.Recordset
Dim sql As String
Dim Rs As ADODB.Recordset
Dim Co As DConecta
Set Co = New DConecta
sql = "Select * from PersComentario where cPerscod='" & psPersCod & "'"
Co.AbreConexion
Set Rs = Co.CargaRecordSet(sql)
Co.CierraConexion
Set ObtieneComentario = Rs
Set Co = Nothing
Set Rs = Nothing
End Function

Public Sub ActComentario(ByVal psPersCod As String, ByVal psComentario As String, _
ByVal psFecha As String, ByVal opt As Integer)
Dim sql As String
Dim Co As DConecta
Set Co = New DConecta
If opt = 0 Then
    sql = "Update PersComentario " & _
          " Set cComentario='" & psComentario & "' , dFecha='" & psFecha & "'" & _
          " where cPerscod='" & psPersCod & "'"
Else
    sql = " Insert PersComentario  " & _
          " (cPersCod,cComentario,dFecha) values " & _
          " ('" & psPersCod & "','" & psComentario & "','" & psFecha & "')"
End If
Co.AbreConexion
Co.Ejecutar (sql)
Co.CierraConexion
Set Co = Nothing
End Sub
