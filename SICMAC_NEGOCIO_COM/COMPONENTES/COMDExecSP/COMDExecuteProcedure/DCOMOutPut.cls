VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DCOMOutPut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'variable local para contener colección
'20060324
'creado para contener los valores de retorno de un procedimiento
'almacenado.
'No importa el tipo de dato de valor de retorno. Siempre y cuando sea
'este un tipo de dato predeterminado.
Private mCol As Collection

Public Function Add(Valor As Variant, Optional sKey As String) As Variant
    'crear un nuevo objeto
    Dim objNewMember As Variant
    'Set objNewMember = New clsConecta


    'establecer las propiedades que se transfieren al método
    objNewMember = Valor
'
'    If IsObject(OutPuts) Then
'        Set objNewMember.OutPuts = OutPuts
'    Else
'        objNewMember.OutPuts = OutPuts
'    End If
'    If IsObject(RecordSet) Then
'        Set objNewMember.RecordSet = RecordSet
'    Else
'        objNewMember.RecordSet = RecordSet
'    End If
'    If IsObject(Conexion) Then
'        Set objNewMember.Conexion = Conexion
'    Else
'        objNewMember.Conexion = Conexion
'    End If
'    If IsObject(Comando) Then
'        Set objNewMember.Comando = Comando
'    Else
'        objNewMember.Comando = Comando
'    End If
'    Set objNewMember.cOutPuts = cOutPuts
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'devolver el objeto creado
    Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Integer) As Variant
    'se usa al hacer referencia a un elemento de la colección
    'vntIndexKey contiene el índice o la clave de la colección,
    'por lo que se declara como un Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
   Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'se usa al obtener el número de elementos de la
    'colección. Sintaxis: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Integer)
    'se usa al quitar un elemento de la colección
    'vntIndexKey contiene el índice o la clave, por lo que se
    'declara como un Variant
    'Sintaxis: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'crea la colección cuando se crea la clase
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destruye la colección cuando se termina la clase
    Set mCol = Nothing
End Sub



