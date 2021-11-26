VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCapRegAproAutOtrasOperaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aprobaciòn/Rechazo Operaciones"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12735
   Icon            =   "frmCapRegAproAutOtrasOperaciones.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   12735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   11505
      TabIndex        =   8
      Top             =   5160
      Width           =   1080
   End
   Begin VB.CommandButton CmdAprobar 
      Caption         =   "Aprobar"
      Height          =   375
      Left            =   10320
      TabIndex        =   7
      Top             =   5160
      Width           =   1080
   End
   Begin VB.CommandButton CmdRechazar 
      Caption         =   "Rechazar"
      Height          =   375
      Left            =   9120
      TabIndex        =   6
      Top             =   5160
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12495
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox CboOperacion 
         Height          =   315
         ItemData        =   "frmCapRegAproAutOtrasOperaciones.frx":030A
         Left            =   5160
         List            =   "frmCapRegAproAutOtrasOperaciones.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   315
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Operacion:"
         Height          =   195
         Left            =   4320
         TabIndex        =   3
         Top             =   315
         Width           =   780
      End
   End
   Begin MSComctlLib.ListView lvwNiveles 
      Height          =   4200
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   7408
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Fecha"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Glosa"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Monto"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Moneda"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Persona"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Usuario"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "nMovNro"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmCapRegAproAutOtrasOperaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************
'* Formulario: frmCapRegAproAutOtrasOperaciones
'* Usuario: FRHU '*' Fecha: 20140508 '*' Segun ERS063-2014
'************************
Option Explicit
Private Sub CboOperacion_Click()
    CargaDatos
End Sub

Private Sub CmdAprobar_Click()
    AprobarRechazar gOtraOperacEstAprobado
End Sub

Private Sub cmdRechazar_Click()
    AprobarRechazar gOtraOperacEstRechazado
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CargaAgencias
    CboAgencia.ListIndex = 0
    CargaOperaciones
    CentraForm Me
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

Public Sub CargaOperaciones()
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(2039)
    Set clsGen = Nothing
    
    CboOperacion.Clear
    While Not rsConst.EOF
        CboOperacion.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
End Sub

Public Sub CargaDatos()
    Dim CapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    Dim lista As ListItem
    Dim rs As New ADODB.Recordset

    Set CapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
        Set rs = CapAut.ObtenerDatosOtrasOpeMovAutorizacion(Trim(Right(CboOperacion, 1)), Trim(Right(CboAgencia, 2)), gdFecSis)
    Set CapAut = Nothing
    
    lvwNiveles.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
       lvwNiveles.ListItems.Clear
       Do Until rs.EOF
         Set lista = lvwNiveles.ListItems.Add(, , rs!dFecha)
         lista.SubItems(1) = rs!cGlosa
         lista.SubItems(2) = rs!nMonto
         lista.SubItems(3) = rs("cMoneda")
         lista.SubItems(4) = rs("cPersNombre")
         lista.SubItems(5) = rs("cCodUsu")
         lista.SubItems(6) = rs("nMovNro")
         rs.MoveNext
       Loop
    Else
       MsgBox "No Existen Datos", vbInformation, "Aviso"
    End If
    rs.Close
End Sub

Public Sub AprobarRechazar(ByVal pnEstado As CapNivRetCancEstado)
   Dim CapAut As COMNCaptaGenerales.NCOMCaptAutorizacion
   Dim oMov As COMDMov.DCOMMov
   Dim i As Integer
   Dim lsCtaCod As String, lsOpeTpo As String
   Dim nMovNroOpe As Long
   Dim lnMonto As Double
   Dim ldFecha As Date
   Dim lsmensaje As String
   
   Dim lnNum As Integer
   
   Dim lbAprobado As Boolean
   
   If ValidarNroDatos = False Then Exit Sub
   If MsgBox("Desea continuar con la operacion", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
   lnNum = 0
   For i = 1 To lvwNiveles.ListItems.Count
       If lvwNiveles.ListItems.iTem(i).Checked = True Then
            lsOpeTpo = Trim(Right(CboOperacion.Text, 2))
            lnMonto = CDbl(lvwNiveles.ListItems.iTem(i).SubItems(2))
            nMovNroOpe = CLng(lvwNiveles.ListItems.iTem(i).SubItems(6))
            Set CapAut = New COMNCaptaGenerales.NCOMCaptAutorizacion
            Call CapAut.AprobarAutorizacionOtrasOper(lsOpeTpo, lnMonto, nMovNroOpe, gsCodAge, gsCodUser, gdFecSis, pnEstado)
            
            Set CapAut = Nothing
            lnNum = lnNum + 1
       End If
   Next
   If lnNum = 0 Then
      MsgBox "Seleccione una Solicitud para ser Aprobada", vbInformation, "Aviso"
   Else
      CargaDatos
   End If
End Sub
Public Function ValidarNroDatos() As Boolean
    Dim i As Integer
    Dim c As Integer
    ValidarNroDatos = True
    c = 0
    For i = 1 To lvwNiveles.ListItems.Count
        If lvwNiveles.ListItems.iTem(i).Checked = True Then
            c = c + 1
        End If
    Next
'    If c > 1 Then
'        MsgBox "Solo debe seleccionar una sola Fila de Datos", vbInformation, "Aviso"
'        ValidarNroDatos = False
'    ElseIf c = 0 Then
    If c = 0 Then
        MsgBox "Seleccione al menos una fila", vbInformation, "Aviso"
        ValidarNroDatos = False
    End If
    
End Function
