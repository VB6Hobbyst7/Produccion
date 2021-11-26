VERSION 5.00
Begin VB.UserControl TxtBuscarGeneral 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   LockControls    =   -1  'True
   ScaleHeight     =   390
   ScaleWidth      =   1890
   ToolboxBitmap   =   "ActXTextBuscarGeneral.ctx":0000
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "..."
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1485
      TabIndex        =   1
      Top             =   30
      Width           =   375
   End
   Begin VB.TextBox txtExaminar 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1500
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtCodigo 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1530
   End
End
Attribute VB_Name = "TxtBuscarGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim sDescripcion As String
'Default Property Values:
Const m_def_dFecNac = 0
Const m_def_ColCodigo = 0
Const m_def_ColDesc = 1
Const m_def_PersPersoneria = gPersonaNat
Const m_def_Ok = False
Const m_def_TipoBusPers = 0
Const m_def_rsDebe = 0
Const m_def_rsHaber = 0
Const m_def_psDH = ""
Const m_def_lbUltimaInstancia = True
Const m_def_sPersDireccion = ""
Const m_def_sPersOcupa = 99 'madm 20100723
Const m_def_sPersNroDoc = ""
Const m_def_TipoBusqueda = 1
Const m_def_EditFlex = False
Const m_def_psRaiz = ""
Const m_def_sTitulo = "Datos"

'Seg 2
Dim m_dFecNac As Date
Dim m_rsDocPers As ADODB.Recordset
Dim m_ColCodigo As Long
Dim m_ColDesc As Long
Dim m_PersPersoneria As PersPersoneria
Dim m_Ok As Boolean
Dim m_TipoBusPers As TipoBusqueda
Dim m_rsDebe As ADODB.Recordset
Dim m_rsHaber As ADODB.Recordset
Dim m_psDH As String
Dim m_lbUltimaInstancia As Boolean
Dim m_sPersDireccion As String
Dim m_sPersNroDoc As String
Dim m_sPersOcupa As Integer 'madm 20100723
Dim m_TipoBusqueda As Integer
Dim m_EditFlex As Boolean
Dim m_psRaiz As String
Dim rs1 As ADODB.Recordset
Dim m_sTitulo As String
Dim lbEnabled As Boolean

'Event Declarations:
Event Click(psCodigo As String, psDescripcion As String)    'MappingInfo=cmdExaminar,cmdExaminar,-1,Click
Event OnValidaClick(Vacio As Boolean)
Event EmiteDatos()
Event Change() 'MappingInfo=txtCodigo,txtCodigo,-1,Change
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtCodigo,txtCodigo,-1,KeyPress
Dim lsOption1 As Variant
Dim lsOption2 As Variant
Dim lsOption3 As Variant
Dim lsCodigo As String
Dim lsDescripcion As String
Dim lsSQL As String
Dim lsTituloForm As String
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,Text
Public Property Get Text() As String
   Text = txtCodigo.Text
End Property

Public Property Let Text(ByVal New_Text As String)
   txtCodigo.Text() = New_Text
   PropertyChanged "Text"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
   BackColor = txtCodigo.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   txtCodigo.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,Font
Public Property Get Font() As Font
   Set Font = txtCodigo.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set txtCodigo.Font = New_Font
   PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,SelLength
Public Property Get SelLength() As Long
   SelLength = txtCodigo.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
   txtCodigo.SelLength() = New_SelLength
   PropertyChanged "SelLength"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,SelStart
Public Property Get SelStart() As Long
   SelStart = txtCodigo.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
   txtCodigo.SelStart() = New_SelStart
   PropertyChanged "SelStart"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,SelText
Public Property Get SelText() As String
   SelText = txtCodigo.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
   txtCodigo.SelText() = New_SelText
   PropertyChanged "SelText"
End Property


'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get psRaiz() As String
    psRaiz = m_psRaiz
End Property

Public Property Let psRaiz(ByVal New_psRaiz As String)
    m_psRaiz = New_psRaiz
    PropertyChanged "psRaiz"
End Property
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtCodigo.Appearance = PropBag.ReadProperty("Appearance", Apariencia.flat)
    txtExaminar.Appearance = PropBag.ReadProperty("Appearance", Apariencia.flat)
    txtCodigo.Text = PropBag.ReadProperty("Text", "")
    txtCodigo.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set txtCodigo.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtCodigo.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtCodigo.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtCodigo.SelText = PropBag.ReadProperty("SelText", "")
'    m_psRaiz = PropBag.ReadProperty("psRaiz", m_def_psRaiz)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
'    lbEnabled = UserControl.Enabled
    
    'cmdExaminar.Enabled = PropBag.ReadProperty("Enabled", True)
    txtCodigo.Appearance = PropBag.ReadProperty("Appearance", 1)
    txtCodigo.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_EditFlex = PropBag.ReadProperty("EditFlex", m_def_EditFlex)
    m_TipoBusqueda = PropBag.ReadProperty("TipoBusqueda", m_def_TipoBusqueda)
    m_sPersDireccion = PropBag.ReadProperty("sPersDireccion", m_def_sPersDireccion)
    m_sPersOcupa = PropBag.ReadProperty("sPersOcupa", m_def_sPersOcupa) 'madm 20100723
    m_sPersNroDoc = PropBag.ReadProperty("sPersNroDoc", m_def_sPersNroDoc)
    m_sTitulo = PropBag.ReadProperty("sTitulo", m_def_sTitulo)
    m_lbUltimaInstancia = PropBag.ReadProperty("lbUltimaInstancia", m_def_lbUltimaInstancia)
    Set m_rsDebe = PropBag.ReadProperty("rsDebe", Nothing)
    Set m_rsHaber = PropBag.ReadProperty("rsHaber", Nothing)
    m_psDH = PropBag.ReadProperty("psDH", m_def_psDH)
    m_TipoBusPers = PropBag.ReadProperty("TipoBusPers", m_def_TipoBusPers)
    txtCodigo.Enabled = PropBag.ReadProperty("EnabledText", True)
    m_Ok = PropBag.ReadProperty("Ok", m_def_Ok)
    m_PersPersoneria = PropBag.ReadProperty("PersPersoneria", m_def_PersPersoneria)
    m_ColCodigo = PropBag.ReadProperty("ColCodigo", m_def_ColCodigo)
    m_ColDesc = PropBag.ReadProperty("ColDesc", m_def_ColDesc)
    txtCodigo.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    Set m_rsDocPers = PropBag.ReadProperty("rsDocPers", Nothing)
    m_dFecNac = PropBag.ReadProperty("dFecNac", m_def_dFecNac)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Appearance", txtCodigo.Appearance
    Call PropBag.WriteProperty("Text", txtCodigo.Text, "")
    Call PropBag.WriteProperty("BackColor", txtCodigo.BackColor, &H80000005)
    Call PropBag.WriteProperty("Font", txtCodigo.Font, Ambient.Font)
    Call PropBag.WriteProperty("SelLength", txtCodigo.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtCodigo.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtCodigo.SelText, "")
    Call PropBag.WriteProperty("psRaiz", m_psRaiz, m_def_psRaiz)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Enabled", cmdexaminar.Enabled, True)
    Call PropBag.WriteProperty("Appearance", txtCodigo.Appearance, 1)
    Call PropBag.WriteProperty("Alignment", txtCodigo.Alignment, 0)
    Call PropBag.WriteProperty("EditFlex", m_EditFlex, m_def_EditFlex)
    Call PropBag.WriteProperty("TipoBusqueda", m_TipoBusqueda, m_def_TipoBusqueda)
    Call PropBag.WriteProperty("sPersDireccion", m_sPersDireccion, m_def_sPersDireccion)
    Call PropBag.WriteProperty("sPersOcupa", m_sPersOcupa, m_def_sPersOcupa) 'madm 20100723
    Call PropBag.WriteProperty("sPersNroDoc", m_sPersNroDoc, m_def_sPersNroDoc)
    Call PropBag.WriteProperty("sTitulo", m_sTitulo, m_def_sTitulo)
    Call PropBag.WriteProperty("lbUltimaInstancia", m_lbUltimaInstancia, m_def_lbUltimaInstancia)
    Call PropBag.WriteProperty("rsDebe", m_rsDebe, Nothing)
    Call PropBag.WriteProperty("rsHaber", m_rsHaber, Nothing)
    Call PropBag.WriteProperty("psDH", m_psDH, m_def_psDH)
    Call PropBag.WriteProperty("TipoBusPers", m_TipoBusPers, m_def_TipoBusPers)
    Call PropBag.WriteProperty("EnabledText", txtCodigo.Enabled, True)
    Call PropBag.WriteProperty("Ok", m_Ok, m_def_Ok)
    Call PropBag.WriteProperty("PersPersoneria", m_PersPersoneria, m_def_PersPersoneria)
    Call PropBag.WriteProperty("ColCodigo", m_ColCodigo, m_def_ColCodigo)
    Call PropBag.WriteProperty("ColDesc", m_ColDesc, m_def_ColDesc)
    Call PropBag.WriteProperty("ForeColor", txtCodigo.ForeColor, &H80000008)
    Call PropBag.WriteProperty("rsDocPers", m_rsDocPers, Nothing)
    Call PropBag.WriteProperty("dFecNac", m_dFecNac, m_def_dFecNac)
End Sub

Private Sub UserControl_Resize()
ResizeControl IIf(txtCodigo.Appearance = flat, 15, 0)
End Sub
'
Private Sub ResizeControl(pnValor As Integer)
If UserControl.Height > 10 And UserControl.Width > cmdexaminar.Width Then
    txtCodigo.Width = UserControl.Width - cmdexaminar.Width + 35
    txtCodigo.Height = UserControl.Height - 10
    txtExaminar.Height = txtCodigo.Height
    cmdexaminar.Top = txtCodigo.Top + 30 - pnValor
    cmdexaminar.Height = txtCodigo.Height - 45 + pnValor
    txtExaminar.Left = txtCodigo.Left + txtCodigo.Width - 45
    cmdexaminar.Left = txtExaminar.Left - 15 + pnValor
End If
End Sub

Public Sub Inicio(ByVal psSQL As String, ByVal psOption1 As Variant, ByVal psOption2 As Variant, ByVal psOption3 As Variant, ByRef psCodigo As String, ByRef psDescripcion As String, ByVal psTituloForm As String)
    lsSQL = psSQL
    lsOption1 = psOption1
    lsOption2 = psOption2
    lsOption3 = psOption3
    lsCodigo = psCodigo
    lsDescripcion = psDescripcion
    lsTituloForm = psTituloForm
End Sub

Private Sub cmdExaminar_Click()
Dim lbVacio As Boolean
    Call frmBuscarObjetoDB.Inicio(lsSQL, lsOption1, lsOption2, lsOption3, lsCodigo, lsDescripcion, lsTituloForm) 'TORE 20210510: comentado por mensaje de memoria insuficiente
    txtCodigo = lsCodigo
    sDescripcion = lsDescripcion
RaiseEvent OnValidaClick(lbVacio)
RaiseEvent EmiteDatos
End Sub

Public Property Get psDescripcion() As String
    ValidaDato
    psDescripcion = sDescripcion
End Property

Public Property Let psDescripcion(cCodigo As String)
    sDescripcion = cCodigo
End Property
Private Function ValidaDato() As Boolean
 On Error GoTo ErrorValidaDato
   ValidaDato = False
   If txtCodigo = "" Then
        sDescripcion = ""
        Exit Function
   End If
Exit Function
ErrorValidaDato:
    MsgBox Err.Number & " " & Err.Description, vbInformation, "Aviso"
End Function
Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    
    If ValidaCodigo(lsSQL, lsOption1(2), Trim(txtCodigo)) = False Then
        m_Ok = False
        If txtCodigo.Visible And txtCodigo.Enabled = True Then
            txtCodigo.SetFocus
        End If
    End If
        RaiseEvent EmiteDatos
    End If
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Function ValidaCodigo(ByVal psSQL As String, ByVal psWheDato As String, ByVal psDato As String) As Boolean
Dim DP As COMDCredito.DCOMleasing
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
ValidaCodigo = False
If psDato = "" Then
    txtCodigo = ""
    sDescripcion = ""
    Exit Function
End If
Set DP = New COMDCredito.DCOMleasing
Set rs = DP.ObtenerDatoText(psSQL, psWheDato, psDato)

If Not rs.EOF And Not rs.BOF Then
    ValidaCodigo = True
    txtCodigo = Trim(rs.Fields(0))
    sDescripcion = Trim(rs.Fields(1))
Else
    txtCodigo = ""
    sDescripcion = ""
End If
rs.Close
Set rs = Nothing
Set DP = Nothing
End Function
