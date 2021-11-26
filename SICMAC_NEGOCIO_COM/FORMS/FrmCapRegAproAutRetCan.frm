VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCapRegAproAutRetCan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Aprobación de Autorización"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13425
   Icon            =   "FrmCapRegAproAutRetCan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdRechazar 
      Caption         =   "&Rechar"
      Height          =   375
      Left            =   9960
      TabIndex        =   9
      Top             =   5400
      Width           =   1080
   End
   Begin VB.CommandButton CmdAprobar 
      Caption         =   "&Aprobar"
      Height          =   375
      Left            =   11160
      TabIndex        =   8
      Top             =   5400
      Width           =   1080
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   12345
      TabIndex        =   3
      Top             =   5400
      Width           =   1080
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   60
      TabIndex        =   6
      Top             =   720
      Width           =   13335
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   4200
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   13095
         _ExtentX        =   23098
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
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cuenta"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Monto  Solicitado"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "cMovNro"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "nNivelMax"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Cliente"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Personeria"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Tipo de Cuenta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "cPersCod"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      Begin VB.ComboBox CboOperacion 
         Height          =   315
         ItemData        =   "FrmCapRegAproAutRetCan.frx":030A
         Left            =   5160
         List            =   "FrmCapRegAproAutRetCan.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Operacion:"
         Height          =   195
         Left            =   4320
         TabIndex        =   5
         Top             =   315
         Width           =   780
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
Attribute VB_Name = "FrmCapRegAproAutRetCan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gsNivelUsu As String
Dim gsGrupo As String
Dim gbExisteUnGrupo As Boolean
Dim gsOpeCodA As String
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by



Public Sub CargaAgencias()
    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
    Dim lrAgenc As ADODB.Recordset
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
        Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = Nothing
    If lrAgenc Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        'Modificado por ANDE 20170706 ERS021-2017
        Me.CboAgencia.Clear
        With lrAgenc
            Do While Not .EOF
                CboAgencia.AddItem Trim(!cAgeDescripcion) & Space(50) & !cAgeCod
                .MoveNext
            Loop
            If gsCodCargo = "006005" Or gsCodCargo = "007026" Then
                Dim i As Integer
                
                For i = 0 To CboAgencia.ListCount - 1
                    If Right(CboAgencia.List(i), 2) = gsCodAge Then
                        CboAgencia.Enabled = False
                        CboAgencia.Text = CboAgencia.List(i)
                    End If
                Next i
            Else
                CboAgencia.ListIndex = 0 'APRI 20170725
            End If
        End With
    End If
    'END ANDE
End Sub

Private Sub CboOperacion_Click()

    CargaDatos
    If Trim(Right(CboOperacion, 2)) = 1 Then
        gsOpeCodA = gOpeAutorizacionRetiro
    ElseIf Trim(Right(CboOperacion, 2)) = 2 Then
        gsOpeCodA = gOpeAutorizacionCancelacion
    Else 'JUEZ 20131210
        gsOpeCodA = gOpeAutorizacionCargoCuenta
    End If
End Sub

Private Sub CmdAprobar_Click()
   AprobarRechazar gCapNivRetCancEstAprobado
End Sub

Private Sub cmdRechazar_Click()
   AprobarRechazar gCapNivRetCancEstRechazado
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lsGrupos As Variant
    Dim i As Integer
    Dim n As Integer
    CargaAgencias
    'CboAgencia.ListIndex = 0 'COMENTADO APRI20170725
    CargaOperaciones 'JUEZ 20131210
    Dim oPers As COMDpersona.UCOMAcceso
    Set oPers = New COMDpersona.UCOMAcceso
       gsGrupo = oPers.cargarUsuarioGrupoAprobacionRechazo(gsCodUser, gsDominio)
    Set oPers = Nothing
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapAprobRechAutoriz
    'End By
    '***Agregado por ELRO el 20121113, según OYP-RFC115-2012
    CentraForm Me
    '***Fin Agregado por ELRO el 20121113*******************
End Sub

Public Sub CargaDatos()
    Dim CapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    Dim lista As ListItem
    Dim rs As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    Dim lnNivel As String
    
    If Trim(Right(CboOperacion, 1)) <> 4 Then 'APRI20170602 ERS033-2017
    gsNivelUsu = ""
    Set CapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
    lnNivel = CapAut.VerificarNivAutRetCan(gsGrupo, Trim(Right(CboOperacion, 1)), gsCodAge, gsCodPersUser) 'RIRO20141107 ERS159
    Set CapAut = Nothing
    
    If Trim(lnNivel) = "" Then
        MsgBox "No existe Nivel ", vbInformation, "Aviso"
        Exit Sub
    End If
    gsNivelUsu = lnNivel
    End If
    'Modify By gitu 27-07-2009 para que se pueda aprobar de retiros de cualquier agencia.
    Set CapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
        Set rs = CapAut.ObtenerDatosMovAutorizacion(Trim(Right(CboOperacion, 1)), Trim(Right(CboAgencia, 2)), gdFecSis, lnNivel)
    Set CapAut = Nothing
    
    lvwNiveles.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
       lvwNiveles.ListItems.Clear
       Do Until rs.EOF
         Set lista = lvwNiveles.ListItems.Add(, , rs!dfecha)
         lista.SubItems(1) = rs!cCtaCod
         lista.SubItems(2) = rs!cUsuAutoriza
         lista.SubItems(3) = rs!nMonto
         lista.SubItems(4) = rs("nMovNro")
         lista.SubItems(5) = rs("nNivelMax")
         '***Modificado por ELRO el 20121109, según OYP-RFC115-2012
         lista.SubItems(6) = rs("cPersNombre")
         lista.SubItems(7) = rs("cPersoneria")
         lista.SubItems(8) = rs("cTipoCuenta")
         lista.SubItems(9) = rs("cPersCod")
         '***Fin Modificado por ELRO*******************************
         rs.MoveNext
       Loop
    Else
       MsgBox "No Existen Datos", vbInformation, "Aviso"
    End If
    rs.Close
End Sub

Public Function ValidarNroDatos() As Boolean
    Dim i As Integer
    Dim c As Integer
    ValidarNroDatos = True
    c = 0
    For i = 1 To lvwNiveles.ListItems.count
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

Public Sub AprobarRechazar(ByVal pnEstado As CapNivRetCancEstado)
   Dim CapAut As COMNCaptaGenerales.NCOMCaptAutorizacion
   Dim oMov As COMDMov.DCOMMov
   Dim i As Integer
   Dim lsCtaCod As String, lsOpeTpo As String
   Dim nMovNroOpe As Long
   Dim sNivelMax As String
   Dim lnMonto As Double
   Dim ldFecha As Date
   Dim lsmensaje As String
   
   Dim lnNum As Integer
   
   Dim lbAprobado As Boolean
   
   If ValidarNroDatos = False Then Exit Sub
   lnNum = 0
   For i = 1 To lvwNiveles.ListItems.count
       If lvwNiveles.ListItems.iTem(i).Checked = True Then
            lsCtaCod = lvwNiveles.ListItems.iTem(i).SubItems(1)
            lsOpeTpo = Trim(Right(CboOperacion.Text, 2))
            lnMonto = CDbl(lvwNiveles.ListItems.iTem(i).SubItems(3))
            nMovNroOpe = CLng(lvwNiveles.ListItems.iTem(i).SubItems(4))
            sNivelMax = lvwNiveles.ListItems.iTem(i).SubItems(5)
            Set CapAut = New COMNCaptaGenerales.NCOMCaptAutorizacion
            Call CapAut.AprobarAutorizacion(lsCtaCod, lsOpeTpo, lnMonto, nMovNroOpe, gsCodAge, gsCodUser, gdFecSis, pnEstado, gsNivelUsu, sNivelMax, lsmensaje)
            'APRI20170603 ERS033-2017
            If pnEstado = 1 Then
                MsgBox "Se ha aprobado la solicitud.", vbInformation, "Aviso"
            Else
                MsgBox "Se ha rechazado la solicitud.", vbInformation, "Aviso"
            End If
            'END APRI20170603 ERS033-2017
            'By Capi 21012009
            If pnEstado = gCapNivRetCancEstAprobado Then
                objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Aprobacion Autorizacion", lsCtaCod, gCodigoCuenta
            Else
                objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Rechazo Autorizacion", lsCtaCod, gCodigoCuenta
            End If
            'End by

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

'***Agregado por ELRO el 20121112, según OYP-RFC115-2012
Private Sub lvwNiveles_DblClick()

If lvwNiveles.ListItems.count > 0 Then
    Dim oForm As New frmPosicionCli
    Dim lista As ListItem
    Dim lsPersCod As String
    Set lista = lvwNiveles.SelectedItem
    lsPersCod = lista.SubItems(9)
    
    oForm.iniciarFormulario lsPersCod
End If
End Sub
'***Fin Agregado por ELRO el 20121112*******************
'JUEZ 20131210 ****************************************************************
Public Sub CargaOperaciones()
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsConst As New ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(2038)
    Set clsGen = Nothing
    
    CboOperacion.Clear
    While Not rsConst.EOF
        CboOperacion.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
End Sub
'END JUEZ *********************************************************************
