VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersAdministrarSesiones 
   Caption         =   "Administrar Sesiones"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   Icon            =   "frmPersAdministrarSesiones.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   8055
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   3840
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6773
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Hora"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Codigo"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Usuario"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Numero"
            Object.Width           =   1764
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8055
      Begin VB.ComboBox cboagencia 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton CmdRechazar 
         Caption         =   "&Desactivar"
         Height          =   375
         Left            =   5400
         TabIndex        =   2
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6600
         TabIndex        =   1
         Top             =   240
         Width           =   1080
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
   End
End
Attribute VB_Name = "frmPersAdministrarSesiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub CargaDatos()
    Dim PerAut As COMDPersona.DCOMPersonas
    Dim lista As ListItem
    Dim rs As New ADODB.Recordset
    
    Set PerAut = New COMDPersona.DCOMPersonas
    Set rs = New ADODB.Recordset
    
    
    Set rs = PerAut.DevuelveRFPorAprobarSession(Right(Me.cboagencia.Text, 2), gdFecSis)
         
            lvwNiveles.ListItems.Clear
            If Not (rs.EOF And rs.BOF) Then
               lvwNiveles.ListItems.Clear
               Do Until rs.EOF
                 Set lista = lvwNiveles.ListItems.Add(, , rs!dFecha)
                 lista.SubItems(1) = rs!Hora
                 lista.SubItems(2) = rs!cPersCod
                 lista.SubItems(3) = rs!cUser
                 lista.SubItems(4) = rs!iPistasId
                 rs.MoveNext
               Loop
            Else
               MsgBox "No Existen Datos", vbInformation, "Aviso"
            End If
            
            rs.Close
    
    Set rs = Nothing
    Set PerAut = Nothing
End Sub

Sub definir_solicitud(ByVal pValor As Integer)
Dim ocapaut As COMDPersona.DCOMPersonas
Set ocapaut = New COMDPersona.DCOMPersonas
Dim lnNum As Integer
Dim fechasol As Date
Dim sNombreCompletox As String
Dim pCodUser As String
Dim valor As Integer
Dim lbAprobado As Boolean
Dim pCodigo As Integer
Dim i As Integer
   
   If ValidarNroDatos = False Then Exit Sub
   lnNum = 0
   valor = pValor
   For i = 1 To lvwNiveles.ListItems.Count
       If lvwNiveles.ListItems.iTem(i).Checked = True Then
            fechasol = CDate(lvwNiveles.ListItems.iTem(i).Text)
            sNombreCompletox = lvwNiveles.ListItems.iTem(i).SubItems(1)
            pCodUser = lvwNiveles.ListItems.iTem(i).SubItems(3)
            pCodigo = CInt(lvwNiveles.ListItems.iTem(i).SubItems(4))
            Set ocapaut = New COMNCaptaGenerales.NCOMCaptAutorizacion
             ocapaut.ModificaPersNegativo_Aprobacion pCodigo, fechasol, gdFecSis, sNombreCompletox, pCodUser, valor
            Set ocapaut = Nothing
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
    If c > 1 Then
        MsgBox "Solo debe seleccionar una sola Fila de Datos", vbInformation, "Aviso"
        ValidarNroDatos = False
    ElseIf c = 0 Then
        MsgBox "Seleccione una sola Fila de Datos", vbInformation, "Aviso"
        ValidarNroDatos = False
    End If
    
End Function

Private Sub CboAgencia_Click()
CargaDatos
End Sub

Private Sub CmdRechazar_Click()
Dim ocapaut As COMDPersona.DCOMPersonas
Set ocapaut = New COMDPersona.DCOMPersonas
Dim pCodigo As Long
Dim lnNum As Long
Dim i As Long
   If ValidarNroDatos = False Then Exit Sub
   lnNum = 0
   For i = 1 To lvwNiveles.ListItems.Count
       If lvwNiveles.ListItems.iTem(i).Checked = True Then
            pCodigo = CLng(lvwNiveles.ListItems.iTem(i).SubItems(4))
            Set ocapaut = New COMNCaptaGenerales.NCOMCaptAutorizacion
             ocapaut.ModificaRFPorAprobarSession pCodigo
            Set ocapaut = Nothing
            lnNum = lnNum + 1
       End If
   Next
   If lnNum = 0 Then
      MsgBox "Seleccione una Solicitud para ser Aprobada", vbInformation, "Aviso"
   Else
      CargaDatos
   End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CargaAgencias
    cboagencia.ListIndex = 0
    'lnNivel = False
End Sub


Public Sub CargaAgencias()
    Dim loCargaAg As COMDConstantes.DCOMAgencias
    Dim lrAgenc As ADODB.Recordset
    Set loCargaAg = New COMDConstantes.DCOMAgencias
        Set lrAgenc = loCargaAg.ObtieneAgenciasIqt()
    Set loCargaAg = Nothing
    If lrAgenc Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        Me.cboagencia.Clear
        With lrAgenc
            Do While Not .EOF
                cboagencia.AddItem Trim(!cConsDescripcion) & Space(50) & !nConsValor
                .MoveNext
            Loop
        End With
    End If
End Sub


