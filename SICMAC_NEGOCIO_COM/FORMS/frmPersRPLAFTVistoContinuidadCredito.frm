VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersRPLAFTVistoContinuidadCredito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visto de Continuidad del Proceso de Créditos de Personas incluidas en el Registro Preventivo del LAFT"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12270
   Icon            =   "frmPersRPLAFTVistoContinuidadCredito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   3840
      TabIndex        =   9
      Top             =   120
      Width           =   1320
   End
   Begin VB.TextBox txtComentario 
      Height          =   855
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   7080
      Width           =   6375
   End
   Begin VB.ComboBox CboAgencia 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   615
      Left            =   10920
      TabIndex        =   3
      Top             =   7200
      Width           =   1200
   End
   Begin VB.CommandButton CmdAprobar 
      Caption         =   "&Permitir"
      Height          =   615
      Left            =   9480
      TabIndex        =   2
      Top             =   7200
      Width           =   1320
   End
   Begin VB.CommandButton CmdRechazar 
      Caption         =   "&Denegar"
      Height          =   615
      Left            =   8040
      TabIndex        =   1
      Top             =   7200
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   12135
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   5880
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   10372
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Hora"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre "
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Condición"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Relacion"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Operación"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Usuario"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Agencia"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Codigo"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Agencia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Comentario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   7200
      Width           =   960
   End
End
Attribute VB_Name = "frmPersRPLAFTVistoContinuidadCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmPersRPLAFTVistoContinuidadCredito
'** Descripción : Visto de Continuidad del proceso de Créditos a personas incluidas en el Registro Preventivo del LAFT
'** Creación : marg, 2016
'** Referencia : TI-ERS046-2016
'***************************************************************************

Option Explicit
Dim bAut As Boolean

Private Sub cboAgencia_Click()
    CargaDatos
End Sub

Private Sub cmdActualizar_Click()
    CargaDatos
End Sub

Private Sub cmdAprobar_Click()
    EmitirVistoContinuidad True
End Sub

Private Sub cmdRechazar_Click()
    EmitirVistoContinuidad False
End Sub

Sub EmitirVistoContinuidad(ByVal pbAdmision As Boolean)
Dim oPerVisto As comdpersona.DCOMPersonas
Dim lnNum As Integer

Dim nIdVisto As Integer
Dim bAdmision As Boolean

Dim cComentario As String

Dim i As Integer
   
If ValidarNroDatos = False Then Exit Sub
If ValidarDatosVacios = False Then Exit Sub

lnNum = 0

If MsgBox("Se va a proceder a guardar los datos, desea continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If

bAdmision = pbAdmision
cComentario = Me.txtComentario.Text

For i = 1 To lvwNiveles.ListItems.count
    If lvwNiveles.ListItems.Item(i).Checked = True Then
         nIdVisto = CInt(lvwNiveles.ListItems.Item(i).SubItems(8))

         Set oPerVisto = New comdpersona.DCOMPersonas
         oPerVisto.ActualizarPersRPLAFTVistoContinuidad nIdVisto, bAdmision, cComentario, gsCodUser
         Set oPerVisto = Nothing
         lnNum = lnNum + 1
    End If
Next
If lnNum > 0 Then
    MsgBox "Visto de Continuidad registrado correctamente!", vbInformation, "Información"
    Me.txtComentario.Text = ""
    CargaDatos
Else
    MsgBox "Seleccione una Solicitud para emitir el Visto de Continuidad", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
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
        CboAgencia.AddItem Trim("--Todos--") & Space(50) & "%"
        With lrAgenc
            Do While Not .EOF
                CboAgencia.AddItem Trim(!cAgeDescripcion) & Space(50) & !cAgeCod
                .MoveNext
            Loop
        End With
        CboAgencia.ListIndex = 0
    End If
            
End Sub

Public Sub CargaDatos()
    Dim PerAut As comdpersona.DCOMPersonas
    Dim lista As ListItem
    Dim rs As New ADODB.Recordset
    
    Set PerAut = New comdpersona.DCOMPersonas
    Set rs = New ADODB.Recordset
    
    
    Set rs = PerAut.ListarPersRPLAFTVistoContinuidad(Trim(Right(Me.CboAgencia.Text, 2)))
     
    lvwNiveles.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
       lvwNiveles.ListItems.Clear
       Do Until rs.EOF
         Set lista = lvwNiveles.ListItems.Add(, , rs!FechaSolicitud)
         lista.SubItems(1) = IIf(rs!HoraSolicitud = "", "", rs!HoraSolicitud)
         lista.SubItems(2) = rs!Nombre
         lista.SubItems(3) = rs!Condicion
         lista.SubItems(4) = rs!Relacion
         lista.SubItems(5) = rs!Operacion
         lista.SubItems(6) = rs!Usuario
         lista.SubItems(7) = rs!Agencia
         lista.SubItems(8) = rs!nIdVisto
         rs.MoveNext
       Loop
    Else
       MsgBox "No Existen Datos", vbInformation, "Aviso"
    End If
    
    rs.Close
    Set rs = Nothing
    Set PerAut = Nothing
End Sub

Public Function ValidarNroDatos() As Boolean
    Dim i As Integer
    Dim C As Integer
    ValidarNroDatos = True
    C = 0
    For i = 1 To lvwNiveles.ListItems.count
        If lvwNiveles.ListItems.Item(i).Checked = True Then
            C = C + 1
        End If
    Next
    If C > 1 Then
        MsgBox "Solo debe seleccionar una Solicitud", vbInformation, "Aviso"
        ValidarNroDatos = False
    ElseIf C = 0 Then
        MsgBox "Primero seleccione una Solicitud", vbInformation, "Aviso"
        ValidarNroDatos = False
    End If
    
End Function
Public Function ValidarDatosVacios() As Boolean
    Dim sComentario As String
    ValidarDatosVacios = True
    
    sComentario = Me.txtComentario.Text
    If sComentario = "" Then
        MsgBox "Ingrese un comentario", vbInformation, "Aviso"
        Me.txtComentario.SetFocus
        ValidarDatosVacios = False
    End If
End Function

Public Sub Inicio()
    Dim PerAut As comdpersona.DCOMPersonas
    Set PerAut = New comdpersona.DCOMPersonas
    bAut = PerAut.VerificarAutorizacionVistoContinuidad(gsCodUser)
    If bAut = True Then
        CargaAgencias
        Me.Show 1
    Else
        MsgBox "Ud. No cuenta con permisos suficientes para realizar el Visto de Continuidad", vbInformation, "SICMACM"
        Unload Me
    End If
    
End Sub
