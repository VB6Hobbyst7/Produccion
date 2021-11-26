VERSION 5.00
Begin VB.UserControl ComboCheck 
   BackStyle       =   0  'Transparent
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1785
   ForeColor       =   &H8000000D&
   ScaleHeight     =   660
   ScaleWidth      =   1785
   Begin VB.Frame fr_Data 
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1785
      Begin VB.ComboBox cmbData 
         Height          =   315
         Left            =   80
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.CheckBox chkData 
         Caption         =   "Fondo Crecer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   70
         TabIndex        =   1
         Top             =   10
         Width           =   1575
      End
   End
End
Attribute VB_Name = "ComboCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Event Click()
Dim Modulo As Integer

Public Sub TodoControlEnable(ByVal pbValor As Boolean)
    fr_Data.Enabled = pbValor
    chkData.Enabled = pbValor
    cmbData.Enabled = pbValor
End Sub

'=============== Frame ================================
Public Sub frameVisual(ByVal bValor As Boolean)
    fr_Data.Visible = bValor
End Sub

Public Property Get frameGetVisual() As Boolean
    frameGetVisual = fr_Data.Visible
End Property

Public Sub frameEnable(ByVal pbValor As Boolean)
    fr_Data.Enabled = pbValor
End Sub
'=============== Frame ================================

'=============== CheckBox ================================
Private Sub chkData_Click()
Dim cProducto As String
cProducto = frmCredSugerencia_NEW.FondoCrecer_GetProductoSuger

Call frmCredSolicitud.FondoCrecerQuitarBotones

    If chkData = Checked Then
        If cProducto = "" Then
            Call frmCredSolicitud.FondoCrecer_CargaDestino
            cmbData.Enabled = True
        Else
            cmbData.Enabled = True
        End If
    Else
        If cProducto = "" Then
            cmbData.ListIndex = -1
            Call frmCredSolicitud.FondoCrecerCatalogoLlenaCombox
            cmbData.Enabled = False
        Else
            cmbData.Enabled = False
            cmbData.ListIndex = -1
        End If
    End If
Modulo = 0
End Sub

Public Property Let Enabled(ByVal vNewEnabled As Boolean)
    UserControl.Enabled = vNewEnabled
    PropertyChanged "Enabled"
End Property

Public Property Let CheckEnabled(ByVal vNewEnabled As Boolean)
    UserControl.chkData.Enabled = vNewEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get CheckValue() As Boolean
    CheckValue = chkData.value
End Property

Public Sub ChckAsignaValue(ByVal pnValor As Integer)
    chkData.value = pnValor
End Sub
'=============== CheckBox ================================

'=============== ComboBox ================================
Public Sub CargaCombo(ByVal rs As ADODB.Recordset)
On Error Resume Next
cmbData.Clear
Do While Not rs.EOF
    cmbData.AddItem Trim(rs!cConsDescripcion) & Space(100) & Trim((rs!nConsValor))
    rs.MoveNext
Loop
rs.Close
End Sub

Public Property Let ComboEnabled(ByVal vNewEnabled As Boolean)
    UserControl.cmbData.Enabled = vNewEnabled
    PropertyChanged "Enabled"
End Property

Public Property Get ComboValue() As Boolean
    ComboValue = IIf(cmbData = "", 0, 1)
End Property

Public Function ComboObtieneValor(ByVal pnValor As Integer) As String
    ComboObtieneValor = Right(cmbData.Text, pnValor)
End Function

Public Function ComboIndiceLista(ByVal psValor As String, Optional ByVal pnDigitoDerecha As Integer = 15) As Long
Dim i As Integer
    ComboIndiceLista = -1
    For i = 0 To (cmbData.ListCount - 1)
        If Trim(Right(cmbData.List(i), pnDigitoDerecha)) = Trim(psValor) Then
            ComboIndiceLista = i
            Exit For
        End If
    Next i
    cmbData.ListIndex = ComboIndiceLista
End Function

Public Sub comboSetFocus()
    If cmbData.Enabled = True Then
        cmbData.SetFocus
    End If
End Sub

Public Sub comboAsignaValue(ByVal pnValor As Integer)
    cmbData.ListIndex = pnValor
End Sub

Public Function comboCantidadValor() As Integer
    comboCantidadValor = cmbData.ListCount
End Function
'=============== ComboBox ================================

Public Sub FondoCrecerDesHabilita(ByVal pnModulo As Integer)
    'If (fr_Data.Visible) = True Then
        Select Case pnModulo
            Case 1 'Solicitud
                If (cmbData.Text) <> "" Then
                    chkData.value = 1
                    cmbData.Enabled = True
                Else
                   chkData.value = 0
                   cmbData.Enabled = False
                End If
            Case 2 'Sugerencia
                If (cmbData.Text) <> "" Then
                    chkData.Enabled = True
                    cmbData.Enabled = True
                    chkData.value = 1
                Else
                    chkData.Enabled = True
                    cmbData.Enabled = False
                    chkData.value = 0
                End If
            Case 5 'Aprobacion
                If (cmbData.Text) <> "" Then
                    chkData.Enabled = False
                    cmbData.Enabled = False
                    chkData.value = 1
                Else
                    chkData.Enabled = False
                    cmbData.Enabled = False
                    chkData.value = 0
                End If
        End Select
    'End If
End Sub

'Public Sub FondoCrecerVisibilidad(ByVal pcCategoria As String, ByVal pcProducto As String, ByVal pnCondicion As Integer, Optional ByVal cPersCod As String)
'Dim obj As COMDCredito.DCOMCredito
'Dim rs As ADODB.Recordset
'Set obj = New COMDCredito.DCOMCredito
'
'Set rs = obj.FondoCrecerVisibilidad(pcCategoria, pcProducto, pnCondicion, cPersCod)
'If Not (rs.BOF And rs.EOF) Then
'    If rs!bAplica = 0 Then
'        fr_Data.Visible = False
'        chkData.value = 0
'        cmbData.ListIndex = -1
'    ElseIf rs!bAplica = 1 Then
'        fr_Data.Visible = True
'        If cmbData.Text <> "" Then
'            chkData.value = 1
'        End If
'    End If
'End If
'
'Set obj = Nothing
'RSClose rs
'End Sub

Public Sub FondoCrecerGetOpcion(Optional ByVal pcCategoria As String = "", Optional ByVal pcProducto As String = "", Optional ByVal pnCondicion As Integer = -1, Optional ByVal pcPersCod As String = "", Optional ByVal pnPlazo As Integer = -1, Optional ByVal pnFormLoad As Integer = 0)
Dim obj As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Set obj = New COMDCredito.DCOMCredito

Set rs = obj.FondoCrecerGetOpcion(pcCategoria, pcProducto, pnCondicion, pcPersCod, pnPlazo, pnFormLoad)
If Not (rs.BOF And rs.EOF) Then
    If rs!bVisible = 0 Then
        fr_Data.Visible = False
        chkData.value = 0
        cmbData.ListIndex = -1
    Else
        fr_Data.Visible = True
        If rs!nOpcionesFondoCrecer <> 0 Then
            If cmbData.ListCount <> 1 Then
                cmbData.RemoveItem (IIf(rs!nOpcionesFondoCrecer = 1, 1, 0))
            End If
        End If
    End If
End If

Set obj = Nothing
RSClose rs
End Sub

Public Sub FondoCrecerGetOpcionSugerencia(ByVal pcCtaCod As String, Optional ByVal pdFechaPago As String = "", Optional ByVal pnFormLoad As Integer = 0)
Dim obj As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Set obj = New COMDCredito.DCOMCredito

Set rs = obj.FondoCrecerGetOpcionSeger(pcCtaCod, pdFechaPago, pnFormLoad)
If Not (rs.BOF And rs.EOF) Then
    If rs!bVisible = 0 Then
        If rs!bAplicaVisibilidad = 1 And fr_Data.Visible = True And chkData.value = 1 Then
            MsgBox "Se modifico el plazo, se quitara el beneficio de Fondo Crecer", vbInformation, "Aviso"
        End If
    
        fr_Data.Visible = False
        chkData.value = 0
        cmbData.ListIndex = -1
    Else
        fr_Data.Visible = True
        If rs!nOpcionesFondoCrecer <> 0 Then
            If cmbData.ListCount <> 1 Then
                cmbData.RemoveItem (IIf(rs!nOpcionesFondoCrecer = 1, 1, 0))
            End If
        End If
        If cmbData.Text <> "" Then
            chkData.value = 1
        Else
            chkData.value = 0
        End If
    End If
End If

Set obj = Nothing
RSClose rs
End Sub
