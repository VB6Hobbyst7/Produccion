VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCapAsigNivelAutRetCan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignación de Niveles  de Autorización Retiros-Cancelación"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraDatos 
      Height          =   1095
      Left            =   0
      TabIndex        =   11
      Top             =   720
      Width           =   7095
      Begin VB.ComboBox CboUsuarios 
         Height          =   315
         ItemData        =   "FrmCapAsigNivelAutRetCan.frx":0000
         Left            =   1560
         List            =   "FrmCapAsigNivelAutRetCan.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   200
         Width           =   5415
      End
      Begin VB.ComboBox CboOperacion 
         Height          =   315
         ItemData        =   "FrmCapAsigNivelAutRetCan.frx":0080
         Left            =   1560
         List            =   "FrmCapAsigNivelAutRetCan.frx":008A
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   2055
      End
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grupos de Usuario:"
         Height          =   195
         Left            =   60
         TabIndex        =   17
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Operacion:"
         Height          =   195
         Left            =   620
         TabIndex        =   15
         Top             =   650
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   3640
         TabIndex        =   14
         Top             =   645
         Width           =   630
      End
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4725
      Width           =   975
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   1320
      TabIndex        =   9
      Top             =   4725
      Width           =   975
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   4725
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   4725
      Width           =   975
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   4725
      Width           =   975
   End
   Begin VB.Frame FraNivel 
      Height          =   735
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7095
      Begin VB.ComboBox CboNivel 
         Height          =   315
         ItemData        =   "FrmCapAsigNivelAutRetCan.frx":0100
         Left            =   1200
         List            =   "FrmCapAsigNivelAutRetCan.frx":010A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   280
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel de Autorización: "
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
         Height          =   435
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FraLista 
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   7095
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "Agregar"
         Height          =   315
         Left            =   5280
         TabIndex        =   2
         Top             =   180
         Width           =   855
      End
      Begin VB.CommandButton CmdQuitar 
         Caption         =   "Quitar"
         Height          =   315
         Left            =   6180
         TabIndex        =   1
         Top             =   180
         Width           =   855
      End
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   2280
         Left            =   60
         TabIndex        =   3
         Top             =   540
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nivel"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Grupo Usuario"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Operación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agencia"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmCapAsigNivelAutRetCan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Limpiar()
    
    CboAgencia.ListIndex = 0
    CboOperacion.ListIndex = 0
    TxtDescripcion.Text = ""
    TxtMontoD.Text = ""
    TxtMontoS.Text = ""
   
End Sub

Public Sub EstadoBotones(ByVal bnuevo As Boolean, ByVal beditar As Boolean, ByVal bcancelar As Boolean, ByVal bgrabar As Boolean, ByVal bagregar As Boolean, ByVal bquitar As Boolean)
    CmdNuevo.Enabled = bnuevo
    CmdEditar.Enabled = beditar
    cmdCancelar.Enabled = bcancelar
    cmdGrabar.Enabled = bgrabar
    CmdAgregar.Enabled = bagregar
    CmdQuitar.Enabled = bquitar
End Sub

Public Sub EstadoControles(ByVal bNivel As Boolean, ByVal bDatos As Boolean)
    FraNivel.Enabled = bNivel
    FraDatos.Enabled = bDatos
End Sub

Public Sub CargaAgencias()
    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
    Dim lrAgenc As ADODB.Recordset
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
        Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = Nothing
    If lrAgenc Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        Me.CboAgencia.Clear
        With lrAgenc
            Do While Not .EOF
                CboAgencia.AddItem Trim(!cAgeDescripcion) & Space(50) & !cAgeCod
                .MoveNext
            Loop
        End With
    End If
End Sub

Public Function ValidaControles() As Boolean
    Dim i As Integer
    Dim lsNivel As String
    Dim lsOpe As String
    Dim lsAge As String
    
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
    
    For i = 1 To lvwNiveles.ListItems.Count
        lsNivel = lvwNiveles.ListItems.Item(i).Text
        lsOpe = Trim(Right(lvwNiveles.ListItems.Item(1).SubItems(1), 5))
        lsAge = Trim(Right(lvwNiveles.ListItems.Item(1).SubItems(2), 5))
        If Trim(TxtNivel.Text) = lsNivel And Trim(Right(CboOperacion.Text, 5)) = lsOpe And Trim(Right(CboAgencia, 5)) = lsAge Then
            ValidaControles = False
            MsgBox "Nivel ya existe", vbInformation, "Aviso"
            Exit Function
        End If
    Next
    
    ' Validar Montos
    Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    Dim lsMensaje As String
    Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
        oCapAut.VerificarMontoNivAutRetCan Trim(Right(CboOperacion.Text, 5)), Trim(Right(CboAgencia.Text, 5)), TxtMontoS.Text, TxtMontoD.Text, TxtNivel.Text, lsMensaje
    Set oCapAut = Nothing
    If Trim(lsMensaje) <> "" Then
       MsgBox lsMensaje, vbInformation, "Aviso"
       ValidaControles = False
       Exit Function
    End If
    ValidaControles = False
End Function

Private Sub CmdGrabar_Click()

End Sub

Private Sub cmdNuevo_Click()
    Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    gsOpcion = "N"
    TxtNivel.Text = ""
    Limpiar
    lvwNiveles.ListItems.Clear
    EstadoBotones False, False, True, True, True, True
    'opcion para obtener ultmino nivel insertado
    Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
        TxtNivel.Text = oCapAut.ObtenerNroMaxNivel()
    Set oCapAut = Nothing
    EstadoControles True, True
End Sub

Private Sub CmdQuitar_Click()
    Dim i As Integer
     i = lvwNiveles.SelectedItem.Index
     lvwNiveles.ListItems.Remove i
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdAgregar_Click()
 Dim lista As ListItem
 Dim i As Integer
 'Validar el ingreso de Datos
 If ValidaControles = False Then Exit Sub
 Set lista = lvwNiveles.ListItems.Add(, , TxtNivel.Text)
 lista.SubItems(1) = CboOperacion.Text
 lista.SubItems(2) = CboAgencia.Text
 lista.SubItems(3) = TxtDescripcion.Text
 lista.SubItems(4) = TxtMontoD.Text
 lista.SubItems(5) = TxtMontoS.Text
 Limpiar
End Sub

Private Sub cmdCancelar_Click()
    TxtNivel.Text = ""
    Limpiar
    lvwNiveles.ListItems.Clear
    EstadoBotones True, True, False, False, False, False
    EstadoControles False, False
End Sub

Private Sub CmdEditar_Click()
    gsOpcion = "E"
    EstadoBotones False, False, True, True, True, True
    EstadoControles True, True
    
End Sub

Private Sub Form_Load()
    CargaAgencias
    EstadoBotones True, True, False, False, False, False
    EstadoControles False, False
End Sub

Private Sub lvwNiveles_DblClick()
    Dim i As Integer
  'If gsOpcion = "E" Then
    If lvwNiveles.ListItems.Count > 0 Then
      i = lvwNiveles.SelectedItem.Index
      Call UbicaCombo(Me.CboOperacion, Trim(Right(lvwNiveles.SelectedItem.SubItems(1), 1)), True)
      Call UbicaCombo(Me.CboAgencia, Trim(Right(lvwNiveles.SelectedItem.SubItems(2), 3)), True)
      TxtDescripcion.Text = lvwNiveles.SelectedItem.SubItems(3)
      TxtMontoS.Text = lvwNiveles.SelectedItem.SubItems(4)
      TxtMontoD.Text = lvwNiveles.SelectedItem.SubItems(5)
      lvwNiveles.ListItems.Remove (i)
    Else
      MsgBox "No existe ningún crédito que pueda reasignar", vbInformation, "Aviso"
    End If
  'End If
End Sub

