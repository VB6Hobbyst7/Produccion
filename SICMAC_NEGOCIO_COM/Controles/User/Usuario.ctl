VERSION 5.00
Begin VB.UserControl Usuario 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   ClipControls    =   0   'False
   Enabled         =   0   'False
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   450
   ScaleWidth      =   480
   ToolboxBitmap   =   "Usuario.ctx":0000
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   435
      Left            =   -15
      Picture         =   "Usuario.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
Const m_def_DireccionUser = ""
Const m_def_DescAgeAct = ""
Const m_def_CodAgeAct = ""
Const m_def_CodAgeAsig = ""
Const m_def_DescAgeAsig = ""
Const m_def_NroDNIUser = ""
Const m_def_NroRucUser = ""
Const m_def_PersCod = ""
'Const m_def_DescAgeAct = 0
'Const m_def_CodAgeAct = 0
'Const m_def_CodAgeAsig = 0
'Const m_def_DescAgeAsig = 0
'Const m_def_NroDNIUser = 0
'Const m_def_NroRucUser = 0
'Const m_def_PersCod = 0
Const m_def_AreaStru = ""
Const m_def_AreaCod = ""
Const m_def_UserCod = ""
Const m_def_UserNom = ""
Const m_def_AreaNom = ""
'Property Variables:
Dim m_DireccionUser As String
Dim m_DescAgeAct As String
Dim m_CodAgeAct As String
Dim m_CodAgeAsig As String
Dim m_DescAgeAsig As String
Dim m_NroDNIUser As String
Dim m_NroRucUser As String
Dim m_PersCod As String
Dim m_AreaStru As String
Dim m_AreaCod As String
Dim m_UserCod As String
Dim m_UserNom As String
Dim m_AreaNom As String

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14
Public Sub Inicio(ByVal psCodUser As String)
Dim oGen As DGeneral

Set oGen = New DGeneral
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = oGen.GetDataUser(psCodUser)
If Not rs.EOF And Not rs.BOF Then
    m_DescAgeAct = Trim(rs!cDescAgActual)
    m_CodAgeAct = Trim(rs!cAgenciaActual)
    m_CodAgeAsig = Trim(rs!cAgenciaAsig)
    m_DescAgeAsig = Trim(rs!cDescAgAsig)
    m_NroDNIUser = Trim(rs!DNI)
    m_NroRucUser = Trim(IIf(IsNull(rs!RUC), "", rs!RUC))
    m_PersCod = Trim(rs!cPersCod)
    m_AreaStru = Trim(rs!cAreaEstruc)
    m_AreaCod = Trim(rs!cAreaCod)
    m_UserCod = psCodUser
    m_UserNom = Trim(rs!cPersNombre)
    m_AreaNom = Trim(rs!cAreaDescripcion)
    m_DireccionUser = Trim(rs!cPersDireccDomicilio)
End If
rs.Close
Set rs = Nothing
Set oGen = Nothing
End Sub
Public Sub DatosPers(ByVal psPersCod As String)
Dim oGen As DGeneral

Set oGen = New DGeneral
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = oGen.GetDatosPersona(psPersCod)
If Not rs.EOF And Not rs.BOF Then
    m_DescAgeAct = Trim(rs!cDescAgActual)
    m_CodAgeAct = Trim(rs!cAgenciaActual)
    m_CodAgeAsig = Trim(rs!cAgenciaAsig)
    m_DescAgeAsig = Trim(rs!cDescAgAsig)
    m_NroDNIUser = Trim(rs!DNI)
    m_NroRucUser = Trim(rs!RUC)
    m_PersCod = Trim(rs!cPersCod)
    m_AreaStru = Trim(rs!cAreaEstruc)
    m_AreaCod = Trim(rs!cAreaCod)
    m_UserCod = psCodUser
    m_UserNom = Trim(rs!cPersNombre)
    m_AreaNom = Trim(rs!cAreaDescripcion)
    m_DireccionUser = Trim(rs!cPersDireccDomicilio)
Else
    m_DescAgeAct = ""
    m_CodAgeAct = ""
    m_CodAgeAsig = ""
    m_DescAgeAsig = ""
    m_NroDNIUser = ""
    m_NroRucUser = ""
    m_PersCod = ""
    m_AreaStru = ""
    m_AreaCod = ""
    m_UserCod = ""
    m_UserNom = ""
    m_AreaNom = ""
    m_DireccionUser = ""
End If
rs.Close
Set rs = Nothing
Set oGen = Nothing
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get AreaCod() As String
Attribute AreaCod.VB_MemberFlags = "400"
    AreaCod = m_AreaCod
End Property

Public Property Let AreaCod(ByVal New_AreaCod As String)
    m_AreaCod = New_AreaCod
    PropertyChanged "AreaCod"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get UserCod() As String
    UserCod = m_UserCod
End Property

Public Property Let UserCod(ByVal New_UserCod As String)
    m_UserCod = New_UserCod
    PropertyChanged "UserCod"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get UserNom() As String
Attribute UserNom.VB_MemberFlags = "400"
    UserNom = m_UserNom
End Property

Public Property Let UserNom(ByVal New_UserNom As String)
    m_UserNom = New_UserNom
    PropertyChanged "UserNom"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get AreaNom() As String
Attribute AreaNom.VB_MemberFlags = "400"
    AreaNom = m_AreaNom
End Property
Public Property Let AreaNom(ByVal New_AreaNom As String)
    m_AreaNom = New_AreaNom
    PropertyChanged "AreaNom"
End Property
'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_AreaCod = m_def_AreaCod
    m_UserCod = m_def_UserCod
    m_UserNom = m_def_UserNom
    m_AreaNom = m_def_AreaNom
    m_AreaStru = m_def_AreaStru
'    m_DescAgeAct = m_def_DescAgeAct
'    m_CodAgeAct = m_def_CodAgeAct
'    m_CodAgeAsig = m_def_CodAgeAsig
'    m_DescAgeAsig = m_def_DescAgeAsig
'    m_NroDNIUser = m_def_NroDNIUser
'    m_NroRucUser = m_def_NroRucUser
'    m_PersCod = m_def_PersCod
    m_DescAgeAct = m_def_DescAgeAct
    m_CodAgeAct = m_def_CodAgeAct
    m_CodAgeAsig = m_def_CodAgeAsig
    m_DescAgeAsig = m_def_DescAgeAsig
    m_NroDNIUser = m_def_NroDNIUser
    m_NroRucUser = m_def_NroRucUser
    m_PersCod = m_def_PersCod
    m_DireccionUser = m_def_DireccionUser
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_AreaCod = PropBag.ReadProperty("AreaCod", m_def_AreaCod)
    m_UserCod = PropBag.ReadProperty("UserCod", m_def_UserCod)
    m_UserNom = PropBag.ReadProperty("UserNom", m_def_UserNom)
    m_AreaNom = PropBag.ReadProperty("AreaNom", m_def_AreaNom)
    m_AreaStru = PropBag.ReadProperty("AreaStru", m_def_AreaStru)
'    m_DescAgeAct = PropBag.ReadProperty("DescAgeAct", m_def_DescAgeAct)
'    m_CodAgeAct = PropBag.ReadProperty("CodAgeAct", m_def_CodAgeAct)
'    m_CodAgeAsig = PropBag.ReadProperty("CodAgeAsig", m_def_CodAgeAsig)
'    m_DescAgeAsig = PropBag.ReadProperty("DescAgeAsig", m_def_DescAgeAsig)
'    m_NroDNIUser = PropBag.ReadProperty("NroDNIUser", m_def_NroDNIUser)
'    m_NroRucUser = PropBag.ReadProperty("NroRucUser", m_def_NroRucUser)
'    m_PersCod = PropBag.ReadProperty("PersCod", m_def_PersCod)
    m_DescAgeAct = PropBag.ReadProperty("DescAgeAct", m_def_DescAgeAct)
    m_CodAgeAct = PropBag.ReadProperty("CodAgeAct", m_def_CodAgeAct)
    m_CodAgeAsig = PropBag.ReadProperty("CodAgeAsig", m_def_CodAgeAsig)
    m_DescAgeAsig = PropBag.ReadProperty("DescAgeAsig", m_def_DescAgeAsig)
    m_NroDNIUser = PropBag.ReadProperty("NroDNIUser", m_def_NroDNIUser)
    m_NroRucUser = PropBag.ReadProperty("NroRucUser", m_def_NroRucUser)
    m_PersCod = PropBag.ReadProperty("PersCod", m_def_PersCod)
    m_DireccionUser = PropBag.ReadProperty("DireccionUser", m_def_DireccionUser)
End Sub

Private Sub UserControl_Resize()
UserControl.Width = Image1.Width + 50
UserControl.Height = Image1.Height + 30

End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("AreaCod", m_AreaCod, m_def_AreaCod)
    Call PropBag.WriteProperty("UserCod", m_UserCod, m_def_UserCod)
    Call PropBag.WriteProperty("UserNom", m_UserNom, m_def_UserNom)
    Call PropBag.WriteProperty("AreaNom", m_AreaNom, m_def_AreaNom)
    Call PropBag.WriteProperty("AreaStru", m_AreaStru, m_def_AreaStru)
'    Call PropBag.WriteProperty("DescAgeAct", m_DescAgeAct, m_def_DescAgeAct)
'    Call PropBag.WriteProperty("CodAgeAct", m_CodAgeAct, m_def_CodAgeAct)
'    Call PropBag.WriteProperty("CodAgeAsig", m_CodAgeAsig, m_def_CodAgeAsig)
'    Call PropBag.WriteProperty("DescAgeAsig", m_DescAgeAsig, m_def_DescAgeAsig)
'    Call PropBag.WriteProperty("NroDNIUser", m_NroDNIUser, m_def_NroDNIUser)
'    Call PropBag.WriteProperty("NroRucUser", m_NroRucUser, m_def_NroRucUser)
'    Call PropBag.WriteProperty("PersCod", m_PersCod, m_def_PersCod)
    Call PropBag.WriteProperty("DescAgeAct", m_DescAgeAct, m_def_DescAgeAct)
    Call PropBag.WriteProperty("CodAgeAct", m_CodAgeAct, m_def_CodAgeAct)
    Call PropBag.WriteProperty("CodAgeAsig", m_CodAgeAsig, m_def_CodAgeAsig)
    Call PropBag.WriteProperty("DescAgeAsig", m_DescAgeAsig, m_def_DescAgeAsig)
    Call PropBag.WriteProperty("NroDNIUser", m_NroDNIUser, m_def_NroDNIUser)
    Call PropBag.WriteProperty("NroRucUser", m_NroRucUser, m_def_NroRucUser)
    Call PropBag.WriteProperty("PersCod", m_PersCod, m_def_PersCod)
    Call PropBag.WriteProperty("DireccionUser", m_DireccionUser, m_def_DireccionUser)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get AreaStru() As String
Attribute AreaStru.VB_MemberFlags = "400"
    AreaStru = m_AreaStru
End Property

Public Property Let AreaStru(ByVal New_AreaStru As String)
    m_AreaStru = New_AreaStru
    PropertyChanged "AreaStru"
End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=14,0,0,0
'Public Property Get DescAgeAct() As Variant
'    DescAgeAct = m_DescAgeAct
'End Property
'
'Public Property Let DescAgeAct(ByVal New_DescAgeAct As Variant)
'    m_DescAgeAct = New_DescAgeAct
'    PropertyChanged "DescAgeAct"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=14,0,0,0
'Public Property Get CodAgeAct() As Variant
'    CodAgeAct = m_CodAgeAct
'End Property
'
'Public Property Let CodAgeAct(ByVal New_CodAgeAct As Variant)
'    m_CodAgeAct = New_CodAgeAct
'    PropertyChanged "CodAgeAct"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=14,0,0,0
'Public Property Get CodAgeAsig() As Variant
'    CodAgeAsig = m_CodAgeAsig
'End Property
'
'Public Property Let CodAgeAsig(ByVal New_CodAgeAsig As Variant)
'    m_CodAgeAsig = New_CodAgeAsig
'    PropertyChanged "CodAgeAsig"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=14,0,0,0
'Public Property Get DescAgeAsig() As Variant
'    DescAgeAsig = m_DescAgeAsig
'End Property
'
'Public Property Let DescAgeAsig(ByVal New_DescAgeAsig As Variant)
'    m_DescAgeAsig = New_DescAgeAsig
'    PropertyChanged "DescAgeAsig"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=14,0,0,0
'Public Property Get NroDNIUser() As Variant
'    NroDNIUser = m_NroDNIUser
'End Property
'
'Public Property Let NroDNIUser(ByVal New_NroDNIUser As Variant)
'    m_NroDNIUser = New_NroDNIUser
'    PropertyChanged "NroDNIUser"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=14,0,0,0
'Public Property Get NroRucUser() As Variant
'    NroRucUser = m_NroRucUser
'End Property
'
'Public Property Let NroRucUser(ByVal New_NroRucUser As Variant)
'    m_NroRucUser = New_NroRucUser
'    PropertyChanged "NroRucUser"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=14,0,0,0
'Public Property Get PersCod() As Variant
'    PersCod = m_PersCod
'End Property
'
'Public Property Let PersCod(ByVal New_PersCod As Variant)
'    m_PersCod = New_PersCod
'    PropertyChanged "PersCod"
'End Property
'
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get DescAgeAct() As String
Attribute DescAgeAct.VB_MemberFlags = "400"
    DescAgeAct = m_DescAgeAct
End Property

Public Property Let DescAgeAct(ByVal New_DescAgeAct As String)
    m_DescAgeAct = New_DescAgeAct
    PropertyChanged "DescAgeAct"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get CodAgeAct() As String
Attribute CodAgeAct.VB_MemberFlags = "400"
    CodAgeAct = m_CodAgeAct
End Property

Public Property Let CodAgeAct(ByVal New_CodAgeAct As String)
    m_CodAgeAct = New_CodAgeAct
    PropertyChanged "CodAgeAct"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get CodAgeAsig() As String
Attribute CodAgeAsig.VB_MemberFlags = "400"
    CodAgeAsig = m_CodAgeAsig
End Property

Public Property Let CodAgeAsig(ByVal New_CodAgeAsig As String)
    m_CodAgeAsig = New_CodAgeAsig
    PropertyChanged "CodAgeAsig"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get DescAgeAsig() As String
Attribute DescAgeAsig.VB_MemberFlags = "400"
    DescAgeAsig = m_DescAgeAsig
End Property

Public Property Let DescAgeAsig(ByVal New_DescAgeAsig As String)
    m_DescAgeAsig = New_DescAgeAsig
    PropertyChanged "DescAgeAsig"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get NroDNIUser() As String
Attribute NroDNIUser.VB_MemberFlags = "400"
    NroDNIUser = m_NroDNIUser
End Property

Public Property Let NroDNIUser(ByVal New_NroDNIUser As String)
    m_NroDNIUser = New_NroDNIUser
    PropertyChanged "NroDNIUser"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get NroRucUser() As String
Attribute NroRucUser.VB_MemberFlags = "400"
    NroRucUser = m_NroRucUser
End Property

Public Property Let NroRucUser(ByVal New_NroRucUser As String)
    m_NroRucUser = New_NroRucUser
    PropertyChanged "NroRucUser"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get PersCod() As String
Attribute PersCod.VB_MemberFlags = "400"
    PersCod = m_PersCod
End Property

Public Property Let PersCod(ByVal New_PersCod As String)
    m_PersCod = New_PersCod
    PropertyChanged "PersCod"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get DireccionUser() As String
Attribute DireccionUser.VB_MemberFlags = "400"
    DireccionUser = m_DireccionUser
End Property

Public Property Let DireccionUser(ByVal New_DireccionUser As String)
    m_DireccionUser = New_DireccionUser
    PropertyChanged "DireccionUser"
End Property

