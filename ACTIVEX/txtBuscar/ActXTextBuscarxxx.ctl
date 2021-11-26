VERSION 5.00
Begin VB.UserControl TxtBuscar 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1890
   LockControls    =   -1  'True
   ScaleHeight     =   390
   ScaleWidth      =   1890
   ToolboxBitmap   =   "ActXTextBuscar.ctx":0000
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
Attribute VB_Name = "TxtBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim sCond        As String
Dim sDescripcion As String
Dim sCodPersona  As String
Dim sSql         As String

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
Const m_def_sPersNroDoc = ""
Const m_def_TipoBusqueda = 1
Const m_def_EditFlex = False
Const m_def_psRaiz = ""
Const m_def_sTitulo = "Datos"

Public Enum Apariencia
    T3D = 1
    Flat = 0
End Enum
Private Enum JustificaTexto
    Izquierda = 0
    Derecha = 1
    Centro = 2
End Enum
Public Enum TipoBusqueda
   BuscaArbol = 1
   BuscaGrid = 2
   BuscaPersona = 3
   BuscaDatoEnGrid = 4
   BuscaSeleCuentas = 5
   BuscaLibre = 6
   buscaempleado = 7
End Enum

Public Enum TipoBusquedaPersona
    BusPersCodigo = 0
    BusPersDocumento = 1
    BusPersDocumentoRuc = 2
End Enum
'Property Variables:
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
Dim m_TipoBusqueda As Integer
Dim m_EditFlex As Boolean
Dim m_psRaiz As String
Dim rs1 As ADODB.Recordset
Dim m_sTitulo As String
Dim lbEnabled As Boolean

'datos presupuesto
Dim lsFrmTpo As String
Dim lsPeriodo As String
Dim lsReqNro As String
Dim lsReqTraNro As String
Dim lsBsCod As String
Dim lsCtaCod As String
Dim lbLectura As Boolean


'Event Declarations:
Event Click(psCodigo As String, psDescripcion As String)    'MappingInfo=cmdExaminar,cmdExaminar,-1,Click
Event OnValidaClick(Vacio As Boolean)
Event EmiteDatos()
Attribute EmiteDatos.VB_MemberFlags = "200"
Event Change() 'MappingInfo=txtCodigo,txtCodigo,-1,Change
Attribute Change.VB_Description = "Ocurre cuando cambia el contenido de un control."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtCodigo,txtCodigo,-1,KeyPress
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."

Public Property Let rs(ByVal vNewValue As ADODB.Recordset)
   Set rs1 = vNewValue
    If m_EditFlex = False Then
         If Not rs1 Is Nothing Then
              If rs1.State = adStateOpen Then
                  If rs1.RecordCount = 1 Then
                      txtCodigo.Text = rs1(0)
                      psDescripcion = rs1(1) & ""
                      Enabled = False
                      RaiseEvent EmiteDatos
                  Else
                      Enabled = lbEnabled
                  End If
              End If
          End If
     End If
End Property

Public Property Get rs() As ADODB.Recordset
   Set rs = rs1
End Property

Public Property Let QuerySeek(ByVal vNewValue As String)
   sSql = vNewValue
End Property

Public Property Get QuerySeek() As String
   QuerySeek = sSql
End Property

Public Property Get psDescripcion() As String
    Select Case TipoBusqueda
        Case BuscaPersona
            'ValidaPersona Trim(txtCodigo)
        Case Else
            ValidaDato
    End Select
    psDescripcion = sDescripcion
End Property

Public Property Let psDescripcion(cCodigo As String)
    sDescripcion = cCodigo
End Property

Private Sub txtCodigo_LostFocus()
    If txtCodigo = "" Then Exit Sub
    If EditFlex = True Then Exit Sub
    Select Case TipoBusqueda
        Case BuscaPersona
        Case buscaempleado
        Case Else
            If ValidaDato = False Then
                If txtCodigo.Visible And txtCodigo.Enabled = True Then
                    txtCodigo.SetFocus
                End If
            End If
    End Select
End Sub

Private Sub txtCodigo_Validate(Cancel As Boolean)
Select Case TipoBusqueda
    Case BuscaPersona, buscaempleado
        If ValidaPersona(Trim(txtCodigo)) = False Then
            m_Ok = False
        End If
    Case BuscaLibre
        Cancel = False
        m_Ok = True
    Case Else
        If ValidaDato = False Then
            m_Ok = False
            If txtCodigo.Visible And txtCodigo.Enabled = True Then
                RaiseEvent EmiteDatos
                Exit Sub
            End If
        End If
End Select
RaiseEvent EmiteDatos
End Sub


Private Sub UserControl_GotFocus()
If txtCodigo.Enabled Then
   txtCodigo.SetFocus
End If
End Sub
Private Sub UserControl_LostFocus()
    If EditFlex = True Then
        If TipoBusqueda = 1 Then
            ValidaDato
        End If
    End If
End Sub
Private Function ValidaDato() As Boolean
Dim rsObj As New ADODB.Recordset
 On Error GoTo ErrorValidaDato
   ValidaDato = False
   If sSql <> "" Then
      Dim oConec As New COMConecta.DCOMConecta
      
   Else
      If rs Is Nothing Then Exit Function
      Set rsObj = rs
      If rs.State = adStateClosed Then Exit Function
      If rs.RecordCount = 0 Then Exit Function
      rsObj.MoveFirst
   End If
   If txtCodigo = "" Then
        sDescripcion = ""
        Exit Function
   End If
   If Not rsObj.EOF Then
      rsObj.Find rsObj(0).Name & " = '" & txtCodigo.Text & "'", , , 1
      If Not rsObj.EOF Then
         ValidaDato = True
         txtCodigo = rsObj(0)
         sDescripcion = rsObj(1) & ""
         If m_lbUltimaInstancia Then
            rsObj.MoveNext
            If Not rsObj.EOF Then
                 If Mid(rsObj(0), 1, Len(txtCodigo)) = txtCodigo Then
                     txtCodigo = ""
                    sDescripcion = ""
                End If
            End If
        End If
      Else
         txtCodigo = ""
         sDescripcion = ""
      End If
   Else
      txtCodigo = ""
      sDescripcion = ""
   End If
   If rsObj.EOF Then
        rsObj.MoveFirst
   End If
   If sSql <> "" Then
      rsObj.Close: Set rsObj = Nothing
   End If
   
Exit Function
ErrorValidaDato:
    MsgBox Err.Number & " " & Err.Description, vbInformation, "Aviso"
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtCodigo.Appearance = PropBag.ReadProperty("Appearance", Apariencia.Flat)
    txtExaminar.Appearance = PropBag.ReadProperty("Appearance", Apariencia.Flat)
    txtCodigo.Text = PropBag.ReadProperty("Text", "")
    txtCodigo.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set txtCodigo.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtCodigo.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtCodigo.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtCodigo.SelText = PropBag.ReadProperty("SelText", "")
    m_psRaiz = PropBag.ReadProperty("psRaiz", m_def_psRaiz)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
    lbEnabled = UserControl.Enabled
    
    'cmdExaminar.Enabled = PropBag.ReadProperty("Enabled", True)
    txtCodigo.Appearance = PropBag.ReadProperty("Appearance", 1)
    txtCodigo.Alignment = PropBag.ReadProperty("Alignment", 0)
    m_EditFlex = PropBag.ReadProperty("EditFlex", m_def_EditFlex)
    m_TipoBusqueda = PropBag.ReadProperty("TipoBusqueda", m_def_TipoBusqueda)
    m_sPersDireccion = PropBag.ReadProperty("sPersDireccion", m_def_sPersDireccion)
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
    Call PropBag.WriteProperty("Enabled", cmdExaminar.Enabled, True)
    Call PropBag.WriteProperty("Appearance", txtCodigo.Appearance, 1)
    Call PropBag.WriteProperty("Alignment", txtCodigo.Alignment, 0)
    Call PropBag.WriteProperty("EditFlex", m_EditFlex, m_def_EditFlex)
    Call PropBag.WriteProperty("TipoBusqueda", m_TipoBusqueda, m_def_TipoBusqueda)
    Call PropBag.WriteProperty("sPersDireccion", m_sPersDireccion, m_def_sPersDireccion)
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
ResizeControl IIf(txtCodigo.Appearance = Flat, 15, 0)
End Sub
'
Private Sub ResizeControl(pnValor As Integer)
If UserControl.Height > 10 And UserControl.Width > cmdExaminar.Width Then
    txtCodigo.Width = UserControl.Width - cmdExaminar.Width + 35
    txtCodigo.Height = UserControl.Height - 10
    txtExaminar.Height = txtCodigo.Height
    cmdExaminar.Top = txtCodigo.Top + 30 - pnValor
    cmdExaminar.Height = txtCodigo.Height - 45 + pnValor
    txtExaminar.Left = txtCodigo.Left + txtCodigo.Width - 45
    cmdExaminar.Left = txtExaminar.Left - 15 + pnValor
End If
End Sub
Private Sub cmdExaminar_Click()
Dim lbVacio As Boolean
Select Case m_TipoBusqueda
   Case 1:
        lbVacio = CargaObjetos
   Case 2:
        lbVacio = CargaBuscaDiversa
   Case 3:
        lbVacio = CargaBuscaPers
   Case 4:
        lbVacio = CargaBuscaDato
   Case 5:
        lbVacio = CargaSeleCuentas
   Case 6
        Dim lsCodigo As String
        Dim lsDescripcion As String
        RaiseEvent Click(lsCodigo, lsDescripcion)
        txtCodigo = lsCodigo
        sDescripcion = lsDescripcion
   Case 7:
        lbVacio = CargaBuscaPersEmpleado
End Select
RaiseEvent OnValidaClick(lbVacio)
RaiseEvent EmiteDatos
End Sub

Private Function CargaBuscaPers() As Boolean
Dim UP As comdpersona.UCOMPersona
CargaBuscaPers = False
Set UP = frmBuscaPersona.Inicio()
If Not UP Is Nothing Then
    
    Select Case m_TipoBusPers
    Case BusPersDocumentoRuc
        If UP Is Nothing Then
            txtCodigo = UP.sPersCod
        ElseIf IsNull(UP.sPersIdnroRUC) Then
            txtCodigo = UP.sPersCod
        Else
            If Trim(UP.sPersIdnroRUC) = "" Then
                txtCodigo = String(11, "0")
            Else
                txtCodigo = UP.sPersIdnroRUC
            End If
        End If
    Case Else
        txtCodigo = UP.sPersCod
    End Select
    
    sCodPersona = UP.sPersCod
    
    sDescripcion = PstaNombre(UP.sPersNombre)
    m_sPersDireccion = UP.sPersDireccDomicilio
    m_PersPersoneria = Int(IIf(UP.sPersPersoneria = "", 0, UP.sPersPersoneria))
    m_dFecNac = UP.dPersNacCreac
    If Val(UP.sPersPersoneria) = gPersonaNat Then
      ' Call UP.ObtieneClientexCodigo(sCodPersona)
       m_sPersNroDoc = UP.sPersIdnroDNI & ""
      ' m_sPersNroDoc = UP.DocsPers!cPersIDNro
    Else
        m_sPersNroDoc = UP.sPersIdnroRUC & ""
    End If
    Set m_rsDocPers = UP.DocsPers
    CargaBuscaPers = True
Else
    txtCodigo = ""
    sDescripcion = ""
    sPersDireccion = ""
    sPersNroDoc = ""
End If
Set UP = Nothing
End Function

Private Function CargaBuscaPersEmpleado() As Boolean
Dim UP As comdpersona.UCOMPersona
CargaBuscaPersEmpleado = False
Set UP = frmBuscaPersona.Inicio(True)
If Not UP Is Nothing Then
    
    Select Case m_TipoBusPers
    Case BusPersDocumentoRuc
        If UP Is Nothing Then
            txtCodigo = UP.sPersCod
        ElseIf IsNull(UP.sPersIdnroRUC) Then
            txtCodigo = UP.sPersCod
        Else
            txtCodigo = UP.sPersIdnroRUC
        End If
    Case Else
        txtCodigo = UP.sPersCod
    End Select
    
    sCodPersona = UP.sPersCod
    
    sDescripcion = PstaNombre(UP.sPersNombre)
    m_sPersDireccion = UP.sPersDireccDomicilio
    m_PersPersoneria = Int(IIf(UP.sPersPersoneria = "", 0, UP.sPersPersoneria))
    m_dFecNac = UP.dPersNacCreac
    If Val(UP.sPersPersoneria) = gPersonaNat Then
        m_sPersNroDoc = UP.sPersIdnroDNI & ""
    Else
        m_sPersNroDoc = UP.sPersIdnroRUC & ""
    End If
    Set m_rsDocPers = UP.DocsPers
    CargaBuscaPersEmpleado = True
Else
    txtCodigo = ""
    sDescripcion = ""
    sPersDireccion = ""
    sPersNroDoc = ""
End If
Set UP = Nothing
End Function


Private Function ValidaPersona(ByVal psDato As String) As Boolean
Dim DP As comdpersona.DCOMPersonas
Dim UP As comdpersona.UCOMPersona
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
ValidaPersona = False
If psDato = "" Then
    txtCodigo = ""
    sCodPersona = ""
    sDescripcion = ""
    sPersDireccion = ""
    sPersNroDoc = ""
    Set m_rsDocPers = Nothing
    Exit Function
End If
Set DP = New comdpersona.DCOMPersonas
Set UP = New comdpersona.UCOMPersona
Select Case m_TipoBusPers
    Case BusPersCodigo
            Set rs = DP.BuscaCliente(psDato, BusquedaCodigo)
    Case BusPersDocumento
            Set rs = DP.BuscaCliente(psDato, BusquedaDocumento)
    Case BusPersDocumentoRuc
            Set rs = DP.BuscaCliente(psDato, BusquedaDocumento)
End Select
If Not rs.EOF And Not rs.BOF Then
    ValidaPersona = True
    
    Select Case m_TipoBusPers
    Case BusPersDocumentoRuc
        If rs.Fields(9) Is Nothing Then
            txtCodigo = rs!cPersCod
        ElseIf IsNull(rs.Fields(9)) Then
            txtCodigo = rs!cPersCod
        Else
            txtCodigo = rs.Fields(9)
        End If
    Case Else
        txtCodigo = rs!cPersCod
    End Select
    
    sCodPersona = rs!cPersCod
    sDescripcion = PstaNombre(rs!cPersNombre)
    m_sPersDireccion = rs!cPersDireccDomicilio
    m_PersPersoneria = Int(IIf(rs!nPersPersoneria = "", 0, rs!nPersPersoneria))
    m_dFecNac = rs("dPersNacCreac")
    If rs!nPersPersoneria = gPersonaNat Then
        m_sPersNroDoc = rs!cPersIDnroDNI & ""
    Else
        m_sPersNroDoc = rs!cPersIDnroRUC & ""
    End If
    UP.CargaDatos rs!cPersCod, sDescripcion, Date, "", "", m_PersPersoneria, "", "", "", IIf(IsNull(rs!cPersNatSexo), "", rs!cPersNatSexo)
    Set m_rsDocPers = UP.DocsPers
Else
    txtCodigo = ""
    sDescripcion = ""
    sPersDireccion = ""
    sPersNroDoc = ""
End If
rs.Close
Set rs = Nothing
Set DP = Nothing
Set UP = Nothing
End Function
Private Function CargaObjetos() As Boolean
Dim rsObj As New ADODB.Recordset
   Dim oConec As New COMConecta.DCOMConecta
   CargaObjetos = False
    If rs Is Nothing Then
        Exit Function
    Else
        If rs.State = adStateClosed Then Exit Function
        If rs.RecordCount = 0 Then Exit Function
    End If
    Set rsObj = rs
    rsObj.MoveFirst
   If Not rsObj.EOF Then
      Dim oDescObj As ClassDescObjeto
      Set oDescObj = New ClassDescObjeto
      
      oDescObj.ColCod = m_ColCodigo
      oDescObj.ColDesc = m_ColDesc
      oDescObj.lbUltNivel = m_lbUltimaInstancia
      oDescObj.Show rsObj, txtCodigo, m_psRaiz
      m_Ok = oDescObj.lbOk
      If oDescObj.lbOk Then
         txtCodigo = oDescObj.gsSelecCod
         sDescripcion = oDescObj.gsSelecDesc
         CargaObjetos = True
      Else
        If txtCodigo.Enabled And txtCodigo.Visible Then
            txtCodigo.SetFocus
        End If
      End If
      Set oDescObj = Nothing
   End If
End Function
Private Function CargaSeleCuentas() As Boolean
Dim rsDebe1 As New ADODB.Recordset
Dim rsHaber1 As New ADODB.Recordset
Dim oDescObj As New ClassDescObjeto
    CargaSeleCuentas = False
    If rsDebe Is Nothing And rsHaber Is Nothing Then
         Exit Function
    End If
    Set rsDebe1 = rsDebe
    Set rsHaber1 = rsHaber
    oDescObj.lbUltNivel = m_lbUltimaInstancia
    oDescObj.ShowSeleCuentas rsDebe, rsHaber, "Seleccion de Cuentas Contables"
    m_Ok = oDescObj.lbOk
    If oDescObj.lbOk Then
        txtCodigo = oDescObj.gsSelecCod
        sDescripcion = oDescObj.gsSelecDesc
        psDH = oDescObj.gsSeleCtasDH
        CargaSeleCuentas = True
    Else
        If txtCodigo.Visible And txtCodigo.Enabled Then
            txtCodigo.SetFocus
        End If
    End If
    Set oDescObj = Nothing
End Function
Private Function CargaBuscaDiversa() As Boolean
Dim rsObj As New ADODB.Recordset
Dim oConec As New COMConecta.DCOMConecta
   CargaBuscaDiversa = False
   If sSql <> "" Then
        oConec.AbreConexion
        Set rsObj = oConec.CargaRecordSet(sSql)
        Set oConec = Nothing
   Else
        If rs Is Nothing Then
            Exit Function
        Else
            If rs.EOF Then Exit Function
        End If
        Set rsObj = rs
        rsObj.MoveFirst
   End If
   If Not rsObj.EOF Then
      Dim oDescObj As New ClassDescObjeto
      oDescObj.lbUltNivel = m_lbUltimaInstancia
      oDescObj.ShowGrid rsObj, m_sTitulo
      m_Ok = oDescObj.lbOk
      If oDescObj.lbOk Then
         txtCodigo = oDescObj.gsSelecCod
         sDescripcion = oDescObj.gsSelecDesc
         CargaBuscaDiversa = True
      Else
         If txtCodigo.Visible And txtCodigo.Enabled Then
            txtCodigo.SetFocus
         End If
      End If
   End If
   If sSql <> "" Then
        Set oDescObj = Nothing
        rsObj.Close: Set rsObj = Nothing
    End If
End Function
Private Function CargaBuscaDato() As Boolean
Dim rsObj As New ADODB.Recordset
   CargaBuscaDato = False
   Dim oConec As New COMConecta.DCOMConecta
   If sSql <> "" Then
        oConec.AbreConexion
        Set rsObj = oConec.CargaRecordSet(sSql)
        Set oConec = Nothing
   Else
        If rs Is Nothing Then
            Exit Function
        Else
            If rs.EOF Then Exit Function
        End If
        Set rsObj = rs
        rsObj.MoveFirst
   End If
   If Not rsObj.EOF Then
      Dim oDescObj As New ClassDescObjeto
      oDescObj.BuscarDato rsObj, 0, m_sTitulo
      If oDescObj.lbOk Then
         txtCodigo = oDescObj.gsSelecCod
         sDescripcion = oDescObj.gsSelecDesc
         CargaBuscaDato = True
      Else
         If txtCodigo.Visible And txtCodigo.Enabled Then
            txtCodigo.SetFocus
         End If
      End If

      Set oDescObj = Nothing
   End If
   If sSql <> "" Then
        Set oDescObj = Nothing
        rsObj.Close: Set rsObj = Nothing
    End If
End Function
Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case TipoBusqueda
            Case BuscaPersona, buscaempleado
                If ValidaPersona(Trim(txtCodigo)) = False Then
                    m_Ok = False
                    If txtCodigo.Visible And txtCodigo.Enabled = True Then
                        txtCodigo.SetFocus
                    End If
                End If
            Case Else
                If ValidaDato = False Then
                    m_Ok = False
                    If txtCodigo.Visible And txtCodigo.Enabled = True Then
                        txtCodigo.SetFocus
                        RaiseEvent EmiteDatos
                        Exit Sub
                    End If
                End If
        End Select
        RaiseEvent EmiteDatos
    End If
    RaiseEvent KeyPress(KeyAscii)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Devuelve o establece el texto contenido en el control."
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "200"
   Text = txtCodigo.Text
End Property

Public Property Let Text(ByVal New_Text As String)
   txtCodigo.Text() = New_Text
   PropertyChanged "Text"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
   BackColor = txtCodigo.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
   txtCodigo.BackColor() = New_BackColor
   PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
   Set Font = txtCodigo.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
   Set txtCodigo.Font = New_Font
   PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Devuelve o establece el número de caracteres seleccionados."
   SelLength = txtCodigo.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
   txtCodigo.SelLength() = New_SelLength
   PropertyChanged "SelLength"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Devuelve o establece el punto inicial del texto seleccionado."
   SelStart = txtCodigo.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
   txtCodigo.SelStart() = New_SelStart
   PropertyChanged "SelStart"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Devuelve o establece la cadena que contiene el texto seleccionado actualmente."
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

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_psRaiz = m_def_psRaiz
    m_EditFlex = m_def_EditFlex
    m_TipoBusqueda = m_def_TipoBusqueda
    m_sPersDireccion = m_def_sPersDireccion
    m_sPersNroDoc = m_def_sPersNroDoc
    m_lbUltimaInstancia = m_def_lbUltimaInstancia
    'm_rsDebe = m_def_rsDebe
    'm_rsHaber = m_def_rsHaber
    m_psDH = m_def_psDH
    m_TipoBusPers = m_def_TipoBusPers
    m_Ok = m_def_Ok
    m_PersPersoneria = m_def_PersPersoneria
    m_ColCodigo = m_def_ColCodigo
    m_ColDesc = m_def_ColDesc
    m_dFecNac = m_def_dFecNac
End Sub

Private Sub txtCodigo_Change()
    RaiseEvent Change
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
    'txtCodigo.Enabled = New_Enabled
    'PropertyChanged "Enabled"
    'cmdExaminar.Enabled = New_Enabled
    'PropertyChanged "Enabled"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,Appearance
Public Property Get Appearance() As Apariencia
Attribute Appearance.VB_Description = "Devuelve o establece si los objetos se dibujan en tiempo de ejecución con efectos 3D."
    Appearance = txtCodigo.Appearance
End Property
Public Property Let Appearance(ByVal New_Appearance As Apariencia)
If txtCodigo.Appearance = Flat And New_Appearance = T3D Then
   ResizeControl 0
End If
If txtCodigo.Appearance = T3D And New_Appearance = Flat Then
   ResizeControl 15
End If
txtCodigo.Appearance = New_Appearance
PropertyChanged "Appearance"
txtExaminar.Appearance = New_Appearance
PropertyChanged "Appearance"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,Alignment
'Public Property Get Alignment() As JustificaTexto
'    Alignment = txtCodigo.Alignment
'End Property
'
'Public Property Let Alignment(ByVal New_Alignment As JustificaTexto)
'    txtCodigo.Alignment() = New_Alignment
'    PropertyChanged "Alignment"
'End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get EditFlex() As Boolean
    EditFlex = m_EditFlex
End Property

Public Property Let EditFlex(ByVal New_EditFlex As Boolean)
    m_EditFlex = New_EditFlex
    PropertyChanged "EditFlex"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,false
Public Property Get TipoBusqueda() As TipoBusqueda
    TipoBusqueda = m_TipoBusqueda
End Property

Public Property Let TipoBusqueda(ByVal New_TipoBusqueda As TipoBusqueda)
    m_TipoBusqueda = New_TipoBusqueda
    PropertyChanged "TipoBusqueda"
    If m_TipoBusqueda = BuscaLibre Then
        txtCodigo.Enabled = False
    End If
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get sPersDireccion() As String
    sPersDireccion = m_sPersDireccion
End Property

Public Property Let sPersDireccion(ByVal New_sPersDireccion As String)
    m_sPersDireccion = New_sPersDireccion
    PropertyChanged "sPersDireccion"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get sPersNroDoc() As String
    sPersNroDoc = m_sPersNroDoc
End Property

Public Property Let sPersNroDoc(ByVal New_sPersNroDoc As String)
    m_sPersNroDoc = New_sPersNroDoc
    PropertyChanged "sPersNroDoc"
End Property

Public Property Get sTitulo() As String
sTitulo = m_sTitulo
End Property

Public Property Let sTitulo(ByVal vNewValue As String)
m_sTitulo = vNewValue
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,True
Public Property Get lbUltimaInstancia() As Boolean
    lbUltimaInstancia = m_lbUltimaInstancia
End Property

Public Property Let lbUltimaInstancia(ByVal New_lbUltimaInstancia As Boolean)
    m_lbUltimaInstancia = New_lbUltimaInstancia
    PropertyChanged "lbUltimaInstancia"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=23,0,0,0
Public Property Get rsDebe() As ADODB.Recordset
    Set rsDebe = m_rsDebe
End Property
Public Property Set rsDebe(ByVal New_rsDebe As ADODB.Recordset)
    Set m_rsDebe = New_rsDebe
    PropertyChanged "rsDebe"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=23,0,0,0
Public Property Get rsHaber() As ADODB.Recordset
    Set rsHaber = m_rsHaber
End Property

Public Property Set rsHaber(ByVal New_rsHaber As ADODB.Recordset)
    Set m_rsHaber = New_rsHaber
    PropertyChanged "rsHaber"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get psDH() As String
    psDH = m_psDH
End Property

Public Property Let psDH(ByVal New_psDH As String)
    m_psDH = New_psDH
    PropertyChanged "psDH"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get TipoBusPers() As TipoBusquedaPersona
    TipoBusPers = m_TipoBusPers
End Property
Public Property Let TipoBusPers(ByVal New_TipoBusPers As TipoBusquedaPersona)
    m_TipoBusPers = New_TipoBusPers
    PropertyChanged "TipoBusPers"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,Enabled
Public Property Get EnabledText() As Boolean
Attribute EnabledText.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    EnabledText = txtCodigo.Enabled
End Property

Public Property Let EnabledText(ByVal New_EnabledText As Boolean)
    txtCodigo.Enabled() = New_EnabledText
    PropertyChanged "EnabledText"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get Ok() As Boolean
    Ok = m_Ok
End Property

Public Property Let Ok(ByVal New_Ok As Boolean)
    m_Ok = New_Ok
    PropertyChanged "Ok"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,
Public Property Get PersPersoneria() As PersPersoneria
Attribute PersPersoneria.VB_MemberFlags = "400"
    PersPersoneria = m_PersPersoneria
End Property

Public Property Let PersPersoneria(ByVal New_PersPersoneria As PersPersoneria)
    m_PersPersoneria = New_PersPersoneria
    PropertyChanged "PersPersoneria"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get ColCodigo() As Long
    ColCodigo = m_ColCodigo
End Property

Public Property Let ColCodigo(ByVal New_ColCodigo As Long)
    m_ColCodigo = New_ColCodigo
    PropertyChanged "ColCodigo"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,1
Public Property Get ColDesc() As Long
    ColDesc = m_ColDesc
End Property

Public Property Let ColDesc(ByVal New_ColDesc As Long)
    m_ColDesc = New_ColDesc
    PropertyChanged "ColDesc"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=txtCodigo,txtCodigo,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
    ForeColor = txtCodigo.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    txtCodigo.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=23,0,0,0
Public Property Get rsDocPers() As ADODB.Recordset
    Set rsDocPers = m_rsDocPers
End Property

Public Property Set rsDocPers(ByVal New_rsDocPers As ADODB.Recordset)
    Set m_rsDocPers = New_rsDocPers
    PropertyChanged "rsDocPers"
End Property

Public Property Get psCodigoPersona() As String
    psCodigoPersona = sCodPersona
End Property

Public Property Let psCodigoPersona(ByVal vNewValue As String)
    sCodPersona = vNewValue
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=3,0,0,0
Public Property Get dFecNac() As Date
    dFecNac = m_dFecNac
End Property

Public Property Let dFecNac(ByVal New_dFecNac As Date)
    m_dFecNac = New_dFecNac
    PropertyChanged "dFecNac"
End Property

