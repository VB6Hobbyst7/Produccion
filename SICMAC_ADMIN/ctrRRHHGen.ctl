VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.UserControl ctrRRHHGen 
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   ScaleHeight     =   1245
   ScaleWidth      =   7170
   Begin VB.Frame fraEmpleado 
      Caption         =   "Recurso Humano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   7095
      Begin Sicmact.TxtBuscar txtPersCod 
         Height          =   270
         Left            =   840
         TabIndex        =   7
         Top             =   247
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   476
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   6
         TipoBusPers     =   1
         EnabledText     =   0   'False
      End
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   555
         Width           =   6135
      End
      Begin VB.TextBox txtRRHH 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   6
         TabIndex        =   2
         Top             =   855
         Width           =   1920
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   300
         Left            =   2760
         TabIndex        =   1
         Top             =   840
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Max             =   1000
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCodPers 
         Caption         =   "Codigo :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   255
         Width           =   735
      End
      Begin VB.Label lblNomPers 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   570
         Width           =   735
      End
      Begin VB.Label lblCodEmp 
         Caption         =   "RRHH"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   870
         Width           =   735
      End
   End
End
Attribute VB_Name = "ctrRRHHGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Default Property Values:
'Const m_def_psCodigoPersona = ""
'Const m_def_SpinnerValor = 0
Const m_def_BackColor = 0
'Const m_def_ForeColor = 0
'Const m_def_Enabled = 0
'Const m_def_Appearance = 0
Const m_def_BackStyle = 0
'Const m_def_BorderStyle = 0
Const m_def_AutoSize = 0
'Property Variables:
'Dim m_psCodigoPersona As String
'Dim m_SpinnerValor As Variant
Dim m_BackColor As Long
'Dim m_ForeColor As Long
'Dim m_Enabled As Boolean
'Dim m_Font As Font
'Dim m_Appearance As Integer
Dim m_BackStyle As Integer
'Dim m_BorderStyle As Integer
Dim m_AutoSize As Boolean
'Event Declarations:
Event EmiteDatos() 'MappingInfo=txtPersCod,txtPersCod,-1,EmiteDatos
Event Click(psCodigo As String, psDescripcion As String) 'MappingInfo=txtPersCod,txtPersCod,-1,Click
'Event cmdRecodatorioClick() 'MappingInfo=cmdRecordatorio,cmdRecordatorio,-1,Click
'Event Click() 'MappingInfo=cmdBuscar,cmdBuscar,-1,Click
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtRRHH,txtRRHH,-1,KeyPress
'Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
'Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=8,0,0,0
'Public Property Get ForeColor() As Long
'    ForeColor = m_ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As Long)
'    m_ForeColor = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=0,0,0,0
'Public Property Get Enabled() As Boolean
'    Enabled = m_Enabled
'End Property
'
'Public Property Let Enabled(ByVal New_Enabled As Boolean)
'    m_Enabled = New_Enabled
'    PropertyChanged "Enabled"
'End Property
''
'''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'''MemberInfo=6,0,0,0
''Public Property Get Font() As Font
''    Set Font = m_Font
''End Property
''
''Public Property Set Font(ByVal New_Font As Font)
''    Set m_Font = New_Font
''    PropertyChanged "Font"
''End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=7,0,0,0
'Public Property Get Appearance() As Integer
'    Appearance = m_Appearance
'End Property
'
'Public Property Let Appearance(ByVal New_Appearance As Integer)
'    m_Appearance = New_Appearance
'    PropertyChanged "Appearance"
'End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=7,0,0,0
'Public Property Get BorderStyle() As Integer
'    BorderStyle = m_BorderStyle
'End Property
'
'Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
'    m_BorderStyle = New_BorderStyle
'    PropertyChanged "BorderStyle"
'End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=5
Public Sub Refresh()
     
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
End Property

Private Sub txtPersCod_EmiteDatos()
    RaiseEvent EmiteDatos
    'txtNombre.Text = txtPersCod.psDescripcion
End Sub

Private Sub txtPersCod_KeyPress(KeyAscii As Integer)
    Dim oRH As DActualizaDatosRRHH
    Set oRH = New DActualizaDatosRRHH
    
    txtRRHH.Text = oRH.GetCodigoEmpleado(txtPersCod.Text)
End Sub

Private Sub txtRRHH_GotFocus()
    txtRRHH.SelStart = 0
    txtRRHH.SelLength = 6
End Sub

Private Sub UpDown_DownClick()
    If txtRRHH.Text <> "" Then
        txtRRHH.Text = Left(txtRRHH.Text, 1) & FillNum(Trim(Str(UpDown.Value)), 5, "0")
        txtRRHH_KeyPress 13
    End If
End Sub

Private Sub UpDown_UpClick()
    If txtRRHH.Text <> "" Then
        txtRRHH.Text = Left(txtRRHH.Text, 1) & FillNum(Trim(Str(UpDown.Value)), 5, "0")
        txtRRHH_KeyPress 13
    End If
End Sub

Private Sub UserControl_GotFocus()
    If txtRRHH.Enabled Then txtRRHH.SetFocus
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
    Dim lnIncremento As Long
    
    If UserControl.Width < 7095 Then
        fraEmpleado.Width = 7095
        UserControl.Width = 7095
        txtNombre.Width = 5175
    Else
        lnIncremento = UserControl.Width - 7095
        fraEmpleado.Width = UserControl.Width
        txtNombre.Width = 5175 + lnIncremento
    End If
    
    If UserControl.Height < 1205 Then
        fraEmpleado.Height = 1205
        UserControl.Height = 1205
    Else
        fraEmpleado.Height = UserControl.Height
    End If
    
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
'    m_ForeColor = m_def_ForeColor
'    m_Enabled = m_def_Enabled
'    Set m_Font = Ambient.Font
'    m_Appearance = m_def_Appearance
    m_BackStyle = m_def_BackStyle
'    m_BorderStyle = m_def_BorderStyle
    m_AutoSize = m_def_AutoSize
    Set UserControl.Font = Ambient.Font
'    m_SpinnerValor = m_def_SpinnerValor
'    m_psCodigoPersona = m_def_psCodigoPersona
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
'    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
'    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
'    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
'    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
'    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    fraEmpleado.ForeColor = PropBag.ReadProperty("ForeColor", &H800000)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    fraEmpleado.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    txtNombre.Text = PropBag.ReadProperty("psNombreEmpledo", "")
    txtRRHH.Text = PropBag.ReadProperty("psCodigoEmpleado", "")
    fraEmpleado.Caption = PropBag.ReadProperty("Caption", "Recurso Humano")
    fraEmpleado.Enabled = PropBag.ReadProperty("Enabled", True)
    UpDown.Value = PropBag.ReadProperty("Appearance", 0)
'    m_SpinnerValor = PropBag.ReadProperty("SpinnerValor", m_def_SpinnerValor)
    UpDown.Value = PropBag.ReadProperty("SpinnerValor", 0)
'    m_psCodigoPersona = PropBag.ReadProperty("psCodigoPersona", m_def_psCodigoPersona)
    txtPersCod.Text = PropBag.ReadProperty("psCodigoPersona", "")
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
'    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
'    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
'    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
'    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
'    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("ForeColor", fraEmpleado.ForeColor, &H800000)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", fraEmpleado.BorderStyle, 1)
    Call PropBag.WriteProperty("psNombreEmpledo", txtNombre.Text, "")
    Call PropBag.WriteProperty("psCodigoEmpleado", txtRRHH.Text, "")
    Call PropBag.WriteProperty("Caption", fraEmpleado.Caption, "Recurso Humano")
    Call PropBag.WriteProperty("Enabled", fraEmpleado.Enabled, True)
    Call PropBag.WriteProperty("Appearance", UpDown.Value, 0)
'    Call PropBag.WriteProperty("SpinnerValor", m_SpinnerValor, m_def_SpinnerValor)
    Call PropBag.WriteProperty("SpinnerValor", UpDown.Value, 0)
'    Call PropBag.WriteProperty("psCodigoPersona", m_psCodigoPersona, m_def_psCodigoPersona)
    Call PropBag.WriteProperty("psCodigoPersona", txtPersCod.Text, "")
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=fraEmpleado,fraEmpleado,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = fraEmpleado.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    fraEmpleado.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=fraEmpleado,fraEmpleado,-1,BorderStyle
Public Property Get BorderStyle() As Integer
    BorderStyle = fraEmpleado.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    fraEmpleado.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Private Sub txtRRHH_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii)
    RaiseEvent KeyPress(KeyAscii)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtNombre,txtNombre,-1,Text
Public Property Get psNombreEmpledo() As String
    psNombreEmpledo = txtNombre.Text
End Property

Public Property Let psNombreEmpledo(ByVal New_psNombreEmpledo As String)
    txtNombre.Text() = New_psNombreEmpledo
    PropertyChanged "psNombreEmpledo"
End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MappingInfo=txtCodPers,txtCodPers,-1,Text
'Public Property Get psCodigoPersona() As String
'    psCodigoPersona = txtPersCod.Text
'End Property
'
'Public Property Let psCodigoPersona(ByVal New_psCodigoPersona As String)
'    txtCodPers.Text() = New_psCodigoPersona
'    PropertyChanged "psCodigoPersona"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MappingInfo=txtSueldo,txtSueldo,-1,Text
'Public Property Get psSueldoContrato() As String
'
'End Property
'
'Public Property Let psSueldoContrato(ByVal New_psSueldoContrato As String)
'    txtSueldo.Text() = New_psSueldoContrato
'    PropertyChanged "psSueldoContrato"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MappingInfo=txtFecNac,txtFecNac,-1,Text
'Public Property Get psFechaNacimiento() As String
'End Property
'
'Public Property Let psFechaNacimiento(ByVal New_psFechaNacimiento As String)
'    txtFecNac.Text() = New_psFechaNacimiento
'    PropertyChanged "psFechaNacimiento"
'End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MappingInfo=txtDNI,txtDNI,-1,Text
'Public Property Get psDNIPersona() As String
'End Property
'
'Public Property Let psDNIPersona(ByVal New_psDNIPersona As String)
'    txtDNI.Text() = New_psDNIPersona
'    PropertyChanged "psDNIPersona"
'End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtRRHH,txtRRHH,-1,Text
Public Property Get psCodigoEmpleado() As String
    psCodigoEmpleado = txtRRHH.Text
End Property

Public Property Let psCodigoEmpleado(ByVal New_psCodigoEmpleado As String)
    txtRRHH.Text() = New_psCodigoEmpleado
    PropertyChanged "psCodigoEmpleado"
End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MappingInfo=txtDir,txtDir,-1,Text
'Public Property Get psDireccionPersona() As String
'End Property
'
'Public Property Let psDireccionPersona(ByVal New_psDireccionPersona As String)
'    txtDir.Text() = New_psDireccionPersona
'    PropertyChanged "psDireccionPersona"
'End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14
Public Function ClearScreen() As Variant
    
    txtPersCod.Text = ""
    txtNombre.Text = ""
    txtRRHH.Text = ""
    'If txtRRHH.Enabled Then txtRRHH.SetFocus
    
End Function

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=fraEmpleado,fraEmpleado,-1,Caption
Public Property Get Caption() As String
    Caption = fraEmpleado.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    fraEmpleado.Caption() = New_Caption
    PropertyChanged "Caption"
End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MappingInfo=txtSueldo,txtSueldo,-1,Enabled
'Public Property Get EsEmpleado() As Boolean
'    'EsEmpleado = txtSueldo.Visible
'End Property
'
'Public Property Let EsEmpleado(ByVal New_EsEmpleado As Boolean)
'    txtSueldo.Visible() = New_EsEmpleado
'    txtRRHH.Visible() = New_EsEmpleado
'    lblSueldo.Visible() = New_EsEmpleado
'    lblCodEmp.Visible() = New_EsEmpleado
'    UpDown.Visible() = New_EsEmpleado
'
'    If New_EsEmpleado Then
'        fraEmpleado.Height = 1605
'        UserControl.Height = 1605
'    Else
'        fraEmpleado.Height = 1905
'        UserControl.Height = 1905
'    End If
'
'    PropertyChanged "EsEmpleado"
'End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=fraEmpleado,fraEmpleado,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = fraEmpleado.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    fraEmpleado.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property



'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UpDown,UpDown,-1,Value
Public Property Get Appearance() As Long
    Appearance = UpDown.Value
End Property

Public Property Let Appearance(ByVal New_Appearance As Long)
    UpDown.Value() = New_Appearance
    PropertyChanged "Appearance"
End Property
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=14,0,0,0
'Public Property Get SpinnerValor() As Variant
'    SpinnerValor = m_SpinnerValor
'End Property
'
'Public Property Let SpinnerValor(ByVal New_SpinnerValor As Variant)
'    m_SpinnerValor = New_SpinnerValor
'    PropertyChanged "SpinnerValor"
'End Property
'
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UpDown,UpDown,-1,Value

Public Property Get SpinnerValor() As Long
    'SpinnerValor = UpDown.Value
End Property

Public Property Let SpinnerValor(ByVal New_SpinnerValor As Long)
    UpDown.Value() = New_SpinnerValor
    PropertyChanged "SpinnerValor"
End Property

Private Sub txtPersCod_Click(psCodigo As String, psDescripcion As String)
    RaiseEvent Click(psCodigo, psDescripcion)
End Sub
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MemberInfo=13,0,0,
'Public Property Get psCodigoPersona() As String
'    psCodigoPersona = m_psCodigoPersona
'End Property
'
'Public Property Let psCodigoPersona(ByVal New_psCodigoPersona As String)
'    m_psCodigoPersona = New_psCodigoPersona
'    PropertyChanged "psCodigoPersona"
'End Property
'
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtPersCod,txtPersCod,-1,Text
Public Property Get psCodigoPersona() As String
Attribute psCodigoPersona.VB_Description = "Devuelve o establece el texto contenido en el control."
    psCodigoPersona = txtPersCod.Text
End Property

Public Property Let psCodigoPersona(ByVal New_psCodigoPersona As String)
    txtPersCod.Text() = New_psCodigoPersona
    PropertyChanged "psCodigoPersona"
End Property

