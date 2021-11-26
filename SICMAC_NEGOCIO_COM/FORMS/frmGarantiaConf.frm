VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGarantiaConf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Garantías"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9225
   Icon            =   "frmGarantiaConf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "Examinar"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   80
      Width           =   1095
   End
   Begin TabDlg.SSTab sstGeneral 
      Height          =   6615
      Left            =   120
      TabIndex        =   21
      Top             =   195
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Items Garantía"
      TabPicture(0)   =   "frmGarantiaConf.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraClasificacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdNuevo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdEditar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdGrabar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdSalir"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdCancelar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraDocPropiedad"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraTipoValor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.Frame fraTipoValor 
         Caption         =   "Tipo de Valorización"
         Enabled         =   0   'False
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
         Height          =   3015
         Left            =   4560
         TabIndex        =   26
         Top             =   3000
         Width           =   4335
         Begin VB.ListBox lstTipoVal 
            Height          =   2310
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   15
            Top             =   600
            Width           =   4095
         End
         Begin VB.CheckBox chkTodosTipo 
            Caption         =   "Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraDocPropiedad 
         Caption         =   "Doc. Propiedad"
         Enabled         =   0   'False
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
         Height          =   3015
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   4335
         Begin VB.ListBox lstDocProp 
            Height          =   2310
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   13
            Top             =   600
            Width           =   4095
         End
         Begin VB.CheckBox chkTodosDocProp 
            Caption         =   "Todos"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6675
         TabIndex        =   19
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   7800
         TabIndex        =   20
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2370
         TabIndex        =   18
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1240
         TabIndex        =   17
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   6120
         Width           =   1095
      End
      Begin VB.Frame fraClasificacion 
         Caption         =   "Clasificación de Garantía"
         Enabled         =   0   'False
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
         Height          =   2415
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   8775
         Begin VB.CheckBox chkActivado 
            Caption         =   "Activado"
            Height          =   255
            Left            =   5040
            TabIndex        =   11
            Top             =   1950
            Width           =   975
         End
         Begin VB.TextBox txtDescLeasing 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2520
            TabIndex        =   9
            Top             =   1560
            Width           =   5055
         End
         Begin VB.TextBox txtCodLeasing 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   8
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtDescConstitucion 
            Enabled         =   0   'False
            Height          =   285
            Left            =   2520
            TabIndex        =   6
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox txtCodConstitucion 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   5
            Top             =   1200
            Width           =   495
         End
         Begin VB.TextBox txtDescObjGarantia 
            Height          =   285
            Left            =   2520
            TabIndex        =   3
            Top             =   840
            Width           =   5055
         End
         Begin VB.TextBox txtCodObjGarantia 
            Height          =   285
            Left            =   1920
            MaxLength       =   4
            TabIndex        =   2
            Top             =   840
            Width           =   495
         End
         Begin VB.CheckBox chkConstitucion 
            Caption         =   "Constitución"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   1200
            Width           =   1335
         End
         Begin VB.CheckBox chkLeasing 
            Caption         =   "Leasing"
            Enabled         =   0   'False
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   1560
            Width           =   975
         End
         Begin VB.ComboBox cmbSinTramite 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1920
            Width           =   2895
         End
         Begin VB.ComboBox cmbClasificacion 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   360
            Width           =   2895
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Objeto de la Garantía:"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Sin Trámite:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   1960
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Clasificación:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Width           =   930
         End
      End
   End
End
Attribute VB_Name = "frmGarantiaConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre      : frmGarantiaConf
'** Descripción : Formulario para guardar las configuraciones de tipos de garantias
'** Creación    : WIOR, 20140708 10:00:00 AM
'*****************************************************************************************************
Option Explicit

Private fbConstitucion As Boolean
Private fbLeasing As Boolean
Private fbEditar As Boolean
Private pnCod As Long

Private Sub chkTodosDos_Click()
    Call CheckLista(IIf(chkTodosDocProp.value = 1, True, False), lstDocProp)
End Sub

Private Sub chkActivado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl chkTodosDocProp
    End If
End Sub

Private Sub chkConstitucion_Click()
    If chkConstitucion.value = 1 Then
        txtCodConstitucion.Enabled = True
        txtDescConstitucion.Enabled = True
    Else
        txtCodConstitucion.Enabled = False
        txtDescConstitucion.Enabled = False
        txtCodConstitucion.Text = ""
        txtDescConstitucion.Text = ""
    End If
End Sub

Private Sub chkConstitucion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkConstitucion.value Then
            EnfocaControl txtCodConstitucion
        Else
            EnfocaControl chkLeasing
        End If
    End If
End Sub

Private Sub chkLeasing_Click()
    If chkLeasing.value Then
        txtCodLeasing.Enabled = True
        txtDescLeasing.Enabled = True
    Else
        txtCodLeasing.Enabled = False
        txtDescLeasing.Enabled = False
        txtCodLeasing.Text = ""
        txtDescLeasing.Text = ""
    End If
End Sub

Private Sub chkLeasing_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chkLeasing.value = 1 Then
            EnfocaControl txtCodLeasing
        Else
            EnfocaControl cmbSinTramite
        End If
    End If
End Sub

Private Sub chkTodosDocProp_Click()
    Call CheckLista(IIf(chkTodosDocProp.value = 1, True, False), lstDocProp)
End Sub

Private Sub chkTodosDocProp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl lstDocProp
    End If
End Sub

Private Sub chkTodosTipo_Click()
    Call CheckLista(IIf(chkTodosTipo.value = 1, True, False), lstTipoVal)
End Sub

Private Sub chkTodosTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl lstTipoVal
    End If
End Sub

Private Sub cmbClasificacion_Click()
    Dim oGarant As COMNCredito.NCOMGarantia
    Dim rsGarant As ADODB.Recordset
    
    Dim id As Long
    Set oGarant = New COMNCredito.NCOMGarantia
    
    On Error GoTo ErrClasificacion
    
    Screen.MousePointer = 11
    If Trim(cmbClasificacion.Text) = "" Then
        id = 0
    Else
        id = CLng(Trim(Right(cmbClasificacion.Text, 2)))
    End If
    
    Set rsGarant = oGarant.ObtenerConfigGarantParam(id)
    fbConstitucion = False
    fbLeasing = False
    chkConstitucion.value = False
    chkLeasing.value = False
    chkTodosDocProp.value = False
    chkTodosTipo.value = False
    
    If Not (rsGarant.EOF And rsGarant.BOF) Then
        fbConstitucion = CBool(rsGarant!bConstitucion)
        fbLeasing = CBool(rsGarant!bLeasing)
    End If
    
    If Not fbConstitucion Then
        txtCodConstitucion.Text = ""
        txtDescConstitucion.Text = ""
        txtCodConstitucion.Enabled = fbConstitucion
        txtDescConstitucion.Enabled = fbConstitucion
        chkConstitucion.Enabled = False
    End If
    
    If Not fbLeasing Then
        txtCodLeasing.Text = ""
        txtDescLeasing.Text = ""
        txtCodLeasing.Enabled = fbLeasing
        txtDescLeasing.Enabled = fbLeasing
        chkLeasing.Enabled = False
    End If
    
    chkConstitucion.Enabled = fbConstitucion
    chkLeasing.Enabled = fbLeasing
    
    Call LlenaListas(lstDocProp, 1, 0, , id) 'EJVG20150820
    Call LlenaListas(lstTipoVal, 2, id)
    Set rsGarant = Nothing
    Screen.MousePointer = 0
    Exit Sub
ErrClasificacion:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmbClasificacion_KeyPress(KeyAscii As Integer)
    EnfocaControl txtCodObjGarantia
End Sub

Private Sub cmbSinTramite_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl chkActivado
    End If
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
End Sub

Private Sub cmdEditar_Click()
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdEditar.Enabled = False
    CmdExaminar.Enabled = False
    HabilitarControles True
End Sub

Private Sub cmdexaminar_Click()
    Dim nCod As Long
    nCod = frmGarantiaConfSel.Inicio
    
    If nCod = 0 Then
        MsgBox "No ha seleccionado ninguna configuración", vbInformation, "Aviso"
    Else
        pnCod = nCod
        cmdNuevo.Enabled = False
        cmdEditar.Enabled = True
        cmdCancelar.Enabled = True
        CargaDatos nCod
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim oGarant As COMNCredito.NCOMGarantia
    Dim MatDocs() As Variant
    Dim MatTpoValor() As Variant
    Dim sError As String
    Dim i As Long
    Dim nCant As Long
    
    If Not ValidaDatos Then Exit Sub
    If MsgBox("¿Esta seguro de grabar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    nCant = 0
    For i = 0 To lstDocProp.ListCount - 1
        If lstDocProp.Selected(i) Then
            ReDim Preserve MatDocs(nCant)
            MatDocs(nCant) = CLng(Trim(Left(lstDocProp.List(i), 3)))
            nCant = nCant + 1
        End If
    Next i
    
    nCant = 0
    For i = 0 To lstTipoVal.ListCount - 1
        If lstTipoVal.Selected(i) Then
            ReDim Preserve MatTpoValor(nCant)
            MatTpoValor(nCant) = CLng(Trim(Left(lstTipoVal.List(i), 3)))
            nCant = nCant + 1
        End If
    Next i
    
    Set oGarant = New COMNCredito.NCOMGarantia
    sError = oGarant.RegistrarConfigGarant(pnCod, CLng(Trim(Right(cmbClasificacion.Text, 4))), CLng(Trim(txtCodObjGarantia.Text)), UCase(Trim(txtDescObjGarantia.Text)), _
                                            CLng(IIf(Trim(txtCodConstitucion.Text) = "", 0, Trim(txtCodConstitucion.Text))), UCase(Trim(txtDescConstitucion.Text)), _
                                            CLng(IIf(Trim(txtCodLeasing.Text) = "", 0, Trim(txtCodLeasing.Text))), UCase(Trim(txtDescLeasing.Text)), _
                                            CLng(Trim(Right(cmbSinTramite.Text, 4))), IIf(chkActivado.value = 1, True, False), MatDocs, MatTpoValor)
    Set oGarant = Nothing
    If Trim(sError) <> "" Then
         MsgBox sError, vbCritical, "Error"
    Else
        cmdGrabar.Enabled = False
        CmdExaminar.Enabled = True
        cmdCancelar.Enabled = True
        HabilitarControles False
        MsgBox "Se registro correctamente los datos", vbInformation, "Aviso"
    End If
    Screen.MousePointer = 0
End Sub

Private Sub cmdNuevo_Click()
    CmdExaminar.Enabled = False
    cmdGrabar.Enabled = True
    cmdCancelar.Enabled = True
    cmdNuevo.Enabled = False
    HabilitarControles True
    pnCod = 0
    chkActivado.value = 1
    EnfocaControl cmbClasificacion
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cmdNuevo.Enabled = True
    CargaControles
    pnCod = 0
End Sub

Private Sub lstDocProp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl chkTodosTipo
    End If
End Sub

Private Sub lstTipoVal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdGrabar
    End If
End Sub

Private Sub txtCodConstitucion_KeyDown(KeyCode As Integer, Shift As Integer)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub txtCodConstitucion_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCodConstitucion, KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtDescConstitucion
    End If
End Sub

Private Sub txtCodConstitucion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub txtCodLeasing_KeyDown(KeyCode As Integer, Shift As Integer)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub txtCodLeasing_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCodLeasing, KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtDescLeasing
    End If
End Sub

Private Sub txtCodLeasing_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub txtCodObjGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub txtCodObjGarantia_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCodObjGarantia, KeyAscii)
    If KeyAscii = 13 Then
        EnfocaControl txtDescObjGarantia
    End If
End Sub

Private Sub CargaControles()
    Dim oCons As COMDConstantes.DCOMConstantes
    Dim rsCons As ADODB.Recordset
    Set oCons = New COMDConstantes.DCOMConstantes
        
    Set rsCons = oCons.RecuperaConstantes(gGarantiaTpoClasif)
     'Carga Clasificaciones
    Call Llenar_Combo_con_Recordset(rsCons, cmbClasificacion)
    Set rsCons = Nothing
    
    Set rsCons = oCons.RecuperaConstantes(gGarantiaTpoSinTramite)
     'Carga Clasificacione
    Call Llenar_Combo_con_Recordset(rsCons, cmbSinTramite)
    Set rsCons = Nothing
    
    Call LlenaListas(lstDocProp, 1)
    Call LlenaListas(lstTipoVal, 2)
    Set oCons = Nothing
End Sub
Private Sub CheckLista(ByVal bCheck As Boolean, ByVal lstLista As ListBox)
    Dim i As Integer
    For i = 0 To lstLista.ListCount - 1
        lstLista.Selected(i) = bCheck
    Next i
End Sub

Private Sub LlenaListas(ByRef pLista As ListBox, ByVal pnTipo As Integer, Optional ByVal pnClaConConfig As Long = 0, Optional ByVal prsDatos As ADODB.Recordset = Nothing, Optional ByVal pnClasificacion As Long = 0)
    Dim rsLista As ADODB.Recordset
    Dim i As Integer, j As Integer
    Dim oGarant As COMNCredito.NCOMGarantia
    Set oGarant = New COMNCredito.NCOMGarantia
        
    Select Case pnTipo
        Case 1:
                Set rsLista = oGarant.ObtenerDocsGarantias(pnClaConConfig, pnClasificacion)
        Case 2:
                Set rsLista = oGarant.ObtenerConfigGarantParamTpoValor(pnClaConConfig)
    End Select
    
    pLista.Clear
    If Not (rsLista Is Nothing) Then
        If Not (rsLista.EOF And rsLista.BOF) Then
            For i = 0 To rsLista.RecordCount - 1
                pLista.AddItem Format(rsLista!codigo, "000") & " " & Trim(rsLista!Descripcion)
                pLista.Selected(i) = CBool(rsLista!valor)
                
                If Not (prsDatos Is Nothing) Then
                    If Not (prsDatos.EOF And prsDatos.BOF) Then
                        prsDatos.MoveFirst
                        For j = 0 To prsDatos.RecordCount - 1
                            If CLng(Trim(rsLista!codigo)) = CLng(Trim(prsDatos!nConsValor)) Then
                                pLista.Selected(i) = True
                                Exit For
                            End If
                            prsDatos.MoveNext
                        Next j
                    End If
                End If
                rsLista.MoveNext
            Next i
        End If
    End If
    
    AplicarScrollBarListBox pLista
    
    Set rsLista = Nothing
    Set prsDatos = Nothing
    Set oGarant = Nothing
End Sub

Private Sub HabilitarControles(ByVal pbValor As Boolean)
    fraClasificacion.Enabled = pbValor
    fraDocPropiedad.Enabled = pbValor
    fraTipoValor.Enabled = pbValor
End Sub

Private Sub Limpiar()
    HabilitarControles False
    CmdExaminar.Enabled = True
    cmdEditar.Enabled = False
    cmdCancelar.Enabled = False
    cmdNuevo.Enabled = True
    cmdGrabar.Enabled = False
    txtCodConstitucion.Text = ""
    txtCodConstitucion.Enabled = False
    txtDescConstitucion.Text = ""
    txtDescConstitucion.Enabled = False
    txtCodLeasing.Text = ""
    txtCodLeasing.Enabled = False
    txtDescLeasing.Text = ""
    txtDescLeasing.Enabled = False
    txtDescObjGarantia.Text = ""
    txtCodObjGarantia.Text = ""
    cmbClasificacion.ListIndex = -1
    cmbSinTramite.ListIndex = -1
    chkConstitucion.value = 0
    chkLeasing.value = 0
    chkActivado.value = 0
    chkTodosDocProp.value = 0
    chkTodosTipo.value = 0
    pnCod = 0
    CargaControles
End Sub

Private Sub CargaDatos(ByVal pnCodigo As Long)
    Dim oGarant As COMNCredito.NCOMGarantia
    Dim rsGarant As ADODB.Recordset
    
    Dim rsDatos As ADODB.Recordset
    Set oGarant = New COMNCredito.NCOMGarantia
    
    Set rsGarant = oGarant.ObtenerConfigGarantXID(pnCodigo)
    
    If Not (rsGarant.EOF And rsGarant.BOF) Then
        cmbClasificacion.ListIndex = IndiceListaCombo(cmbClasificacion, rsGarant!nClasificacion)
        txtCodObjGarantia.Text = Trim(rsGarant!nObjGarantCod)
        txtDescObjGarantia.Text = Trim(rsGarant!cObjGarantDesc)
        txtCodConstitucion.Text = Trim(rsGarant!nConstitucionCod)
        txtDescConstitucion.Text = Trim(rsGarant!cConstitucionDesc)
        txtCodLeasing.Text = Trim(rsGarant!nLeasingCod)
        txtDescLeasing.Text = Trim(rsGarant!cLeasingDesc)
        cmbSinTramite.ListIndex = IndiceListaCombo(cmbSinTramite, rsGarant!nSinTramite)
        chkActivado.value = IIf(rsGarant!bActivado, 1, 0)
        
        If Trim(txtCodConstitucion.Text) <> "" Then
            chkConstitucion.value = 1
        End If
        
        If Trim(txtCodLeasing.Text) <> "" Then
            chkLeasing.value = 1
        End If
        
        Set rsDatos = oGarant.ObtenerConfigGarantTpoValor(pnCodigo)
        
        Call LlenaListas(lstDocProp, 1, pnCodigo)
        Call LlenaListas(lstTipoVal, 2, CLng(rsGarant!nClasificacion), rsDatos)
    End If
    Set rsGarant = Nothing
    Set rsDatos = Nothing
    Set oGarant = Nothing
End Sub

Private Function ValidaLista(ByVal lstLista As ListBox) As Boolean
    Dim i As Integer
    Dim valor As Boolean
    valor = False
        For i = 0 To lstLista.ListCount - 1
            If lstLista.Selected(i) Then
                valor = True
                Exit For
            End If
        Next i
    ValidaLista = valor
End Function
Private Function ValidaCodigosGarant(ByVal pnCod As Long, ByVal pnTipo As Long, Optional ByVal pnCodConfig As Long = 0) As Boolean
    Dim rsGarant As ADODB.Recordset
    Dim oGarant As COMNCredito.NCOMGarantia
    
    Set oGarant = New COMNCredito.NCOMGarantia
    Set rsGarant = oGarant.GarantValidaCod(pnCod, pnTipo)
    
    If Not (rsGarant.EOF And rsGarant.BOF) Then
        If pnCodConfig = 0 Then
            ValidaCodigosGarant = True
        Else
             If pnCodConfig = CLng(rsGarant!nGarantConfigID) Then
                ValidaCodigosGarant = False
             Else
                ValidaCodigosGarant = True
             End If
        End If
    Else
        ValidaCodigosGarant = False
    End If
End Function

Private Function ValidaDatos() As Boolean
    Dim sExclusion As String
    Dim sExclusionMat() As String
    Dim bExclusion As Boolean
    Dim oConsSist As COMDConstSistema.NCOMConstSistema
    Dim i As Long
    
    ValidaDatos = True
    
    If Trim(cmbClasificacion.Text) = "" Then
        MsgBox "Ud. debe seleccionar la Clasificación", vbInformation, "Aviso"
        ValidaDatos = False
        EnfocaControl cmbClasificacion
        Exit Function
    End If
    
    If Not IsNumeric(Trim(txtCodObjGarantia.Text)) Then
        MsgBox "Ud. debe ingresar el Código del Objeto de la Garantía", vbInformation, "Aviso"
        ValidaDatos = False
        EnfocaControl txtCodObjGarantia
        Exit Function
    End If
    
    If Trim(txtCodObjGarantia.Text) = "" Or CLng(Trim(txtCodObjGarantia.Text)) = 0 Then
        MsgBox "Ud. debe ingresar el Código del Objeto de la Garantía", vbInformation, "Aviso"
        ValidaDatos = False
        EnfocaControl txtCodObjGarantia
        Exit Function
    End If
    
    If Trim(txtDescObjGarantia.Text) = "" Then
        MsgBox "Ud. debe ingresar la decripción del Objeto de la Garantía", vbInformation, "Aviso"
        ValidaDatos = False
        EnfocaControl txtDescObjGarantia
        Exit Function
    End If
    
    If ValidaCodigosGarant(CLng(Trim(txtCodObjGarantia.Text)), 1, pnCod) Then
        MsgBox "El Código de Objeto de la Garantía ya existe", vbInformation, "Aviso"
        ValidaDatos = False
        EnfocaControl txtCodObjGarantia
        Exit Function
    End If
    
    If chkConstitucion.Enabled And chkConstitucion.value Then
        If Not IsNumeric(Trim(txtCodConstitucion.Text)) Then
            MsgBox "Ud. debe ingresar el Código de la Constitución", vbInformation, "Aviso"
            ValidaDatos = False
            EnfocaControl txtCodConstitucion
            Exit Function
        End If
        
        If Trim(txtCodConstitucion.Text) = "" Or CLng(Trim(txtCodConstitucion.Text)) = 0 Then
            MsgBox "Ud. debe ingresar el Código de la Constitución", vbInformation, "Aviso"
            ValidaDatos = False
            EnfocaControl txtCodConstitucion
            Exit Function
        End If
        
        If Trim(txtDescConstitucion.Text) = "" Then
            MsgBox "Ingrese la Descripción de la Constitución.", vbInformation, "Aviso"
            ValidaDatos = False
            EnfocaControl txtDescConstitucion
            Exit Function
        End If
        
        If ValidaCodigosGarant(CLng(Trim(txtCodConstitucion.Text)), 2, pnCod) Then
            MsgBox "El Código de Constitución de la Garantía ya existe", vbInformation, "Aviso"
            ValidaDatos = False
            EnfocaControl txtCodConstitucion
            Exit Function
        End If
    End If
    
    If chkLeasing.Enabled And chkLeasing.value Then
        If Not IsNumeric(Trim(txtCodLeasing.Text)) Then
            MsgBox "Ud. debe ingresar el Código del Leasing", vbInformation, "Aviso"
            ValidaDatos = False
            EnfocaControl txtCodLeasing
            Exit Function
        End If
        
        If Trim(txtCodLeasing.Text) = "" Or CLng(Trim(txtCodLeasing.Text)) = 0 Then
            MsgBox "Ud. debe ingresar el Código del Leasing", vbInformation, "Aviso"
            ValidaDatos = False
            EnfocaControl txtCodLeasing
            Exit Function
        End If
        
        If Trim(txtDescLeasing.Text) = "" Then
            MsgBox "Ud. debe ingresar la Decripción del Leasing", vbInformation, "Aviso"
            ValidaDatos = False
            EnfocaControl txtDescLeasing
            Exit Function
        End If
        
        Set oConsSist = New COMDConstSistema.NCOMConstSistema
        sExclusion = oConsSist.LeeConstSistema(398)
        Set oConsSist = Nothing
        
        sExclusionMat = Split(sExclusion, ",")
        bExclusion = False
        For i = 0 To UBound(sExclusionMat)
            If CLng(sExclusionMat(i)) = CLng(Trim(Right(cmbClasificacion.Text, 5))) Then
                bExclusion = True
                Exit For
            End If
        Next i
            
        If Not bExclusion Then
            If ValidaCodigosGarant(CLng(Trim(txtCodLeasing.Text)), 3, pnCod) Then
                MsgBox "El Codigo del Leasing de la Garantía ya existe", vbInformation, "Aviso"
                ValidaDatos = False
                EnfocaControl txtCodLeasing
                Exit Function
            End If
        End If
    End If
    
    If chkConstitucion.Enabled And chkConstitucion.value Then
        If chkLeasing.Enabled And chkLeasing.value Then
            If Trim(txtCodConstitucion.Text) = Trim(txtCodLeasing.Text) Then
                MsgBox "Codigo de Constitución y Codigo de Leasing no pueden ser iguales", vbInformation, "Aviso"
                ValidaDatos = False
                EnfocaControl txtCodConstitucion
                Exit Function
            End If
        End If
    End If
    
    If Trim(cmbSinTramite.Text) = "" Then
        MsgBox "Ud. debe seleccionar la opción Sin Trámite", vbInformation, "Aviso"
        ValidaDatos = False
        EnfocaControl cmbSinTramite
        Exit Function
    End If
    
    If Not ValidaLista(lstDocProp) Then
        MsgBox "Ud. debe seleccionar por lo menos un Documento.", vbInformation, "Aviso"
        ValidaDatos = False
        EnfocaControl lstDocProp
        Exit Function
    End If
    
    If Not ValidaLista(lstTipoVal) Then
        MsgBox "Ud. debe seleccionar por lo menos un Tipo de Valorización", vbInformation, "Aviso"
        ValidaDatos = False
        EnfocaControl lstDocProp
        Exit Function
    End If
End Function

Private Sub txtCodObjGarantia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
    Clipboard.SetText ""
End Sub

Private Sub txtDescConstitucion_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl chkLeasing
    End If
End Sub

Private Sub txtDescLeasing_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        EnfocaControl cmbSinTramite
    End If
End Sub

Private Sub txtDescObjGarantia_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        If chkConstitucion.Enabled Then
            EnfocaControl chkConstitucion
        ElseIf chkLeasing.Enabled Then
            EnfocaControl chkLeasing
        Else
            EnfocaControl cmbSinTramite
        End If
    End If
End Sub
