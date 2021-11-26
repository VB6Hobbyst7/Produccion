VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCapRegNivelAutRetCan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Niveles de Autorización"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "FrmCapRegNivelAutRetCan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraLista 
      Height          =   2655
      Left            =   60
      TabIndex        =   16
      Top             =   720
      Width           =   7095
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   2280
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4022
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Operación"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Agencia"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripción"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Monto Dolares"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Monto Soles"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame FraDatos 
      Height          =   735
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   7095
      Begin VB.ComboBox CboOperacion 
         Height          =   315
         ItemData        =   "FrmCapRegNivelAutRetCan.frx":030A
         Left            =   960
         List            =   "FrmCapRegNivelAutRetCan.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         ItemData        =   "FrmCapRegNivelAutRetCan.frx":030E
         Left            =   4320
         List            =   "FrmCapRegNivelAutRetCan.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Operacion:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   315
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   3600
         TabIndex        =   18
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   975
   End
   Begin VB.Frame FraNivel 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   3360
      Width           =   7095
      Begin VB.TextBox TxtNivel 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtMontoD 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   5
         Text            =   "0"
         Top             =   1050
         Width           =   2055
      End
      Begin VB.TextBox TxtMontoS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5040
         TabIndex        =   6
         Text            =   "0"
         Top             =   1050
         Width           =   1935
      End
      Begin VB.TextBox TxtDescripcion 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   650
         Width           =   5175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel Autorización: "
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
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1680
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Monto Dolrares:"
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Monto Soles:"
         Height          =   195
         Left            =   3960
         TabIndex        =   13
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Descripción:"
         Height          =   195
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   885
      End
   End
End
Attribute VB_Name = "FrmCapRegNivelAutRetCan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gnOpcion As Integer '1=nuevo /2=modificar
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by

 
Public Sub Limpiar()
    'CboAgencia.ListIndex = -1
    'CboOperacion.ListIndex = -1
    txtDescripcion.Text = ""
    TxtMontoD.Text = "0.00"
    TxtMontoS.Text = "0.00"
End Sub

Public Sub EstadoBotones(ByVal bnuevo As Boolean, ByVal beditar As Boolean, ByVal bcancelar As Boolean, ByVal bgrabar As Boolean)
    CmdNuevo.Enabled = bnuevo
    cmdEditar.Enabled = beditar
    cmdCancelar.Enabled = bcancelar
    cmdGrabar.Enabled = bgrabar
End Sub

Public Sub EstadoControles(ByVal bNivel As Boolean, ByVal bDatos As Boolean)
    fraNivel.Enabled = bNivel
    FraDatos.Enabled = bDatos
End Sub

Public Function ValidaControles() As Boolean
    Dim i As Integer
    Dim lsNivel As String
    Dim lsOpe As String
    Dim lsAge As String
    Dim lnNum As Integer
    Dim lsGrupo As String
    
    ValidaControles = True
    If Trim(TxtNivel.Text) = "" Or Len(TxtNivel.Text) <> 3 Then
        ValidaControles = False
        MsgBox "Ingrese un Nivel", vbInformation, "Aviso"
        Exit Function
    End If
    
    If Trim(TxtMontoS) = "" Or IsNumeric(TxtMontoS) = False Then
        ValidaControles = False
        MsgBox "Monto en Soles incorrecto", vbInformation, "Aviso"
        Exit Function
    End If
    
    If Trim(TxtMontoD) = "" Or IsNumeric(TxtMontoD) = False Then
        ValidaControles = False
        MsgBox "Monto en Dolares incorrecto", vbInformation, "Aviso"
        Exit Function
    End If
    
    If CboOperacion.ListIndex = -1 Then
        ValidaControles = False
        MsgBox "Seleccione una Operación", vbInformation, "Aviso"
        Exit Function
    End If
    
    If cboagencia.ListIndex = -1 Then
        ValidaControles = False
        MsgBox "Seleccione una Agencia", vbInformation, "Aviso"
        Exit Function
    End If

    If gnOpcion = 1 Then
        'Datos no se dupliquen
        For i = 1 To lvwNiveles.ListItems.Count
            lsNivel = lvwNiveles.ListItems.iTem(i).Text
            lsOpe = Trim(Right(lvwNiveles.ListItems.iTem(1).SubItems(1), 5))
            lsAge = Trim(Right(lvwNiveles.ListItems.iTem(1).SubItems(2), 5))
            If Trim(TxtNivel.Text) = lsNivel And Trim(Right(CboOperacion.Text, 5)) = lsOpe And Trim(Right(cboagencia, 5)) = lsAge Then
                ValidaControles = False
                MsgBox "Datos ya existen", vbInformation, "Aviso"
                Exit Function
            End If
        Next
        
        ' Validar Montos
        Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
        Dim lsMensaje As String
        Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
            oCapAut.VerificarMontoNivAutRetCan Trim(Right(CboOperacion.Text, 5)), Trim(Right(cboagencia.Text, 5)), TxtMontoS.Text, TxtMontoD.Text, TxtNivel.Text, lsMensaje
        Set oCapAut = Nothing
        If Trim(lsMensaje) <> "" Then
           MsgBox lsMensaje, vbInformation, "Aviso"
           ValidaControles = False
           Exit Function
        End If
    End If
End Function

Public Sub CargaAgencias()
    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
    Dim lrAgenc As ADODB.Recordset
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
        Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = Nothing
    If lrAgenc Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        Me.cboagencia.Clear
        With lrAgenc
            Do While Not .EOF
                cboagencia.AddItem Trim(!cAgeDescripcion) & Space(50) & !cAgeCod
                .MoveNext
            Loop
        End With
    End If
End Sub

Private Sub CboAgencia_Click()
    CargarDatos
End Sub

Private Sub CboAgencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDescripcion.SetFocus
    End If
End Sub

Private Sub CboOperacion_Click()
    CargarDatos
End Sub

Private Sub CboOperacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cboagencia.SetFocus
    End If
End Sub

Private Sub Agregar()
 Dim lista As ListItem
 
 Dim i As Integer
 Set lista = lvwNiveles.ListItems.Add(, , TxtNivel.Text)
 lista.SubItems(1) = CboOperacion.Text
 lista.SubItems(2) = cboagencia.Text
 lista.SubItems(3) = txtDescripcion.Text
 lista.SubItems(4) = TxtMontoD.Text
 lista.SubItems(5) = TxtMontoS.Text
 Limpiar
 
End Sub

Private Sub cmdCancelar_Click()
    TxtNivel.Text = ""
    Limpiar
    'lvwNiveles.ListItems.Clear
    EstadoBotones True, True, False, False
    EstadoControles False, True
End Sub

Private Sub cmdEditar_Click()
    gnOpcion = 2
    EstadoBotones False, False, True, True
    EstadoControles True, False
    
End Sub

Private Sub cmdGrabar_Click()
     Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
     Dim rs As ADODB.Recordset
     Dim i As Integer
     
     If ValidaControles = False Then Exit Sub
     
     
     Set rs = New ADODB.Recordset
     ' crear recordset
     With rs
            'Crear RecordSet
            .fields.Append "sNivCod", adVarChar, 4
            .fields.Append "sOpeTpo", adVarChar, 2
            .fields.Append "sCodage", adVarChar, 2
            .fields.Append "sNivDesc", adVarChar, 250
            .fields.Append "nTpoSol", adCurrency
            .fields.Append "nTpoDol", adCurrency
            .Open
            'Llenar Recordset
           
            .AddNew
            .fields("sNivCod") = Format(TxtNivel, "000")
            .fields("sOpeTpo") = Trim(Right(CboOperacion.Text, 2))
            .fields("sCodage") = Trim(Right(cboagencia, 3))
            .fields("sNivDesc") = Me.txtDescripcion
            .fields("nTpoDol") = Me.TxtMontoD
            .fields("nTpoSol") = Me.TxtMontoS
           
     End With
     
     
     Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
           oCapAut.InsertarNilAutRenCan rs, gnOpcion
           
      'By Capi 21012009
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Nivel"
       'End by

           
     Set oCapAut = Nothing
     cmdCancelar_Click
     CargarDatos
    
End Sub

Private Sub cmdNuevo_Click()
    Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    gnOpcion = 1
    TxtNivel.Text = ""
    Limpiar
    'lvwNiveles.ListItems.Clear
    EstadoBotones False, False, True, True
    'opcion para obtener ultmino nivel insertado
    Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
        TxtNivel.Text = oCapAut.ObtenerNroMaxNivel(Trim(Right(CboOperacion.Text, 2)), Trim(Right(cboagencia.Text, 3)))
    Set oCapAut = Nothing
    EstadoControles True, True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargaAgencias
    CargaOperaciones 'JUEZ 20131210
    EstadoBotones True, True, False, False
    EstadoControles False, True
    cboagencia.ListIndex = 0
    CboOperacion.ListIndex = 0
    CargarDatos
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapNivelesAutoriz
    'End By


End Sub

Private Sub lvwNiveles_DblClick()
    Dim i As Integer
    Dim j As Integer
  If gnOpcion = 2 Then
    If lvwNiveles.ListItems.Count > 0 Then
      i = lvwNiveles.SelectedItem.Index
      TxtNivel.Text = lvwNiveles.SelectedItem.Text
      Call UbicaCombo(Me.CboOperacion, Trim(Right(lvwNiveles.SelectedItem.SubItems(1), 1)), True)
      Call UbicaCombo(Me.cboagencia, Trim(Right(lvwNiveles.SelectedItem.SubItems(2), 3)), True)
      txtDescripcion.Text = lvwNiveles.SelectedItem.SubItems(3)
      TxtMontoS.Text = lvwNiveles.SelectedItem.SubItems(5)
      TxtMontoD.Text = lvwNiveles.SelectedItem.SubItems(4)
      'lvwNiveles.ListItems.Remove (i)
    Else
      MsgBox "No existe ningún crédito que pueda reasignar", vbInformation, "Aviso"
    End If
 End If
End Sub

Private Sub TxtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.TxtMontoD.SetFocus
    End If
End Sub

Private Sub TxtMontoD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         Me.TxtMontoS.SetFocus
    End If
End Sub

Private Sub TxtMontoS_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    End If
End Sub

Public Sub CargarDatos()
    Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    Dim rs As New ADODB.Recordset
    Dim lista As ListItem
    

    Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
        Set rs = oCapAut.ObtenerDatosNivAutRetCan(Trim(Right(CboOperacion.Text, 2)), Trim(Right(cboagencia.Text, 3)))
    Set oCapAut = Nothing
        
    lvwNiveles.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF
            Set lista = lvwNiveles.ListItems.Add(, , Format(rs!cNivCod, "000"))
            lista.SubItems(1) = rs!cOpeTpo
            lista.SubItems(2) = rs!cCodAge
            lista.SubItems(3) = rs!cNivDesc
            lista.SubItems(4) = rs!nTopDol
            lista.SubItems(5) = rs!nTopSol
            rs.MoveNext
        Loop
    End If
    'Me.CboOperacion.SetFocus
    fraNivel.Enabled = False
   
End Sub

'JUEZ 20131210 ****************************************************************
Public Sub CargaOperaciones()
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(2038)
    Set clsGen = Nothing
    
    CboOperacion.Clear
    While Not rsConst.EOF
        CboOperacion.AddItem rsConst.fields(0) & Space(100) & rsConst.fields(1)
        rsConst.MoveNext
    Wend
End Sub
'END JUEZ *********************************************************************
