VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPersNegativaAutorizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorizacion de Clientes de Procedimiento Reforzado"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8160
   Icon            =   "frmPersNegativaAutorizacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8055
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6600
         TabIndex        =   7
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton CmdAprobar 
         Caption         =   "&Aprobar"
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   240
         Width           =   1080
      End
      Begin VB.CommandButton CmdRechazar 
         Caption         =   "&Rechazar"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   240
         Width           =   1080
      End
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2655
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
   Begin VB.Frame Frame2 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   8055
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   3840
         Left            =   120
         TabIndex        =   1
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
         NumItems        =   7
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
            Text            =   "Decripcion"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Usuario"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Ocupacion"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmPersNegativaAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gsNivelUsu As String
'Dim gsGrupo As String
Dim gbExisteUnGrupo As Boolean
Dim lnNivel As Boolean
Dim fbPreAut As Boolean 'WIOR 20121123

Private Sub CboAgencia_Click()
CargaDatos
End Sub

Private Sub CmdAprobar_Click()
    definir_solicitud 1
End Sub

Private Sub CmdRechazar_Click()
    definir_solicitud 2
End Sub

Sub definir_solicitud(ByVal pValor As Integer)
Dim ocapaut As COMDPersona.DCOMPersonas
'Set ocapaut = New COMDPersona.DCOMPersonas 'FRHU20160304
Dim lnNum As Integer
Dim fechasol As Date
Dim sNombreCompletox As String
Dim pCodUser, pConcepto As String
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
            sNombreCompletox = lvwNiveles.ListItems.iTem(i).SubItems(2)
            pCodUser = lvwNiveles.ListItems.iTem(i).SubItems(4)
            pConcepto = lvwNiveles.ListItems.iTem(i).SubItems(3)
            pCodigo = CInt(lvwNiveles.ListItems.iTem(i).SubItems(6))
            
            Dim oCont As COMNContabilidad.NCOMContFunciones  'NContFunciones
            Dim sMovNro As String, sOperacion As String
                    
            Set oCont = New COMNContabilidad.NCOMContFunciones
            sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            'Set ocapaut = New COMNCaptaGenerales.NCOMCaptAutorizacion 'FRHU20160304
            Set ocapaut = New COMDPersona.DCOMPersonas 'FRHU20160304
            ocapaut.ModificaPersNegativo_Aprobacion pCodigo, fechasol, gdFecSis, sNombreCompletox, pCodUser, valor, sMovNro, fbPreAut 'WIOR 20121123 AGREGO fbPreAut
            
            Set ocapaut = Nothing
            Set oCont = Nothing
            lnNum = lnNum + 1
       End If
   Next
   If lnNum = 0 Then
      MsgBox "Seleccione una Solicitud para ser Aprobada", vbInformation, "Aviso"
   Else
      CargaDatos
   End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
'WIOR 20121123 COMENTO
'Private Sub Form_Load()
'    CargaAgencias
'    CboAgencia.ListIndex = 0
'    lnNivel = False
'End Sub

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

Public Sub CargaDatos()
    Dim PerAut As COMDPersona.DCOMPersonas
    Dim lista As ListItem
    Dim rs As New ADODB.Recordset
    
    Set PerAut = New COMDPersona.DCOMPersonas
    Set rs = New ADODB.Recordset
    
    'WIOR 20121123 *********************************************************
    If fbPreAut Then
        lnNivel = PerAut.VerificarNivelPreAutorizacionSupJA(gsCodUser)
    Else
        lnNivel = PerAut.VerificarNivelGerencialPEPS(gsCodUser)
    End If
    'WIOR FIN **************************************************************
    
    If lnNivel Then
         Set rs = PerAut.DevuelvePersListaNegativaPorAprobar(Right(Me.CboAgencia.Text, 2), fbPreAut) 'WIOR 20121123 AGREGO fbPreAut
         
            lvwNiveles.ListItems.Clear
            If Not (rs.EOF And rs.BOF) Then
               lvwNiveles.ListItems.Clear
               Do Until rs.EOF
                 Set lista = lvwNiveles.ListItems.Add(, , rs!dfechaSol)
                 lista.SubItems(1) = IIf(rs!dhoraSol = "", "", rs!dhoraSol)
                 lista.SubItems(2) = rs!cNomPers
                 lista.SubItems(3) = rs!cDescripcion
                 lista.SubItems(4) = rs!cUser
                 lista.SubItems(5) = rs!cOcupa
                 lista.SubItems(6) = rs!id
                 rs.MoveNext
               Loop
            Else
               MsgBox "No Existen Datos", vbInformation, "Aviso"
            End If
            
                rs.Close
     Else
        MsgBox "Ud. No cuenta con permisos suficientes para Aprobar/Rechazar PEPS", vbCritical
    End If
    

    Set rs = Nothing
    Set PerAut = Nothing
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
'WIOR 20121123 ************************************************************************************
Public Sub Inicio(Optional ByVal pbPreAut As Boolean = False)
fbPreAut = pbPreAut
CargaAgencias
CboAgencia.ListIndex = 0
lnNivel = False
If fbPreAut Then
    Me.Caption = "Pre-Autorización de Clientes de Procedimiento Reforzado"
Else
    Me.Caption = "Autorización de Clientes de Procedimiento Reforzado"
End If
Me.Show 1
End Sub
'WIOR FIN *****************************************************************************************
