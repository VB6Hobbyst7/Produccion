VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmCapRegNivelAutRetCanDet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Niveles de Autorización - por Grupos"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10860
   Icon            =   "FrmCapRegNivelAutRetCanDet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "&Quitar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8760
      TabIndex        =   14
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame fraNivel 
      Height          =   1095
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   10695
      Begin VB.ComboBox CboGrupo 
         Height          =   315
         ItemData        =   "FrmCapRegNivelAutRetCanDet.frx":030A
         Left            =   960
         List            =   "FrmCapRegNivelAutRetCanDet.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   6015
      End
      Begin VB.ComboBox cboNivel 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   640
         Width           =   6015
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nivel:"
         Height          =   195
         Left            =   440
         TabIndex        =   12
         Top             =   675
         Width           =   405
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9780
      TabIndex        =   8
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7755
      TabIndex        =   7
      Top             =   4560
      Width           =   975
   End
   Begin VB.Frame FraDatos 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   10695
      Begin VB.ComboBox CboAgencia 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CboOperacion 
         Height          =   315
         ItemData        =   "FrmCapRegNivelAutRetCanDet.frx":030E
         Left            =   960
         List            =   "FrmCapRegNivelAutRetCanDet.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   3600
         TabIndex        =   6
         Top             =   315
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Operacion:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   315
         Width           =   780
      End
   End
   Begin VB.Frame FraLista 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10695
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   2280
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Operacion"
            Object.Width           =   1571
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agencia"
            Object.Width           =   2806
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Grupo"
            Object.Width           =   8820
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nivel"
            Object.Width           =   4410
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmCapRegNivelAutRetCanDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gnOpcion As Integer '1=nuevo /2=modificar
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by



Private Sub CboAgencia_Click()
    CargaNiveles
    CargarDatos
    Limpiar
End Sub

Private Sub CboOperacion_Click()
    CargaNiveles
    CargarDatos
    Limpiar
End Sub



Private Sub Form_Load()
   CargaGruposUsu
   CargaOperaciones 'JUEZ 20131210
   CboOperacion.ListIndex = 0
   CargaAgencias
   cboagencia.ListIndex = 0
   CargarDatos
   'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapNivGruposAutoriz
   'End By
    '***Agregado por ELRO el 20121113, según OYP-RFC115-2012
    CentraForm Me
    '***Fin Agregado por ELRO el 20121113*******************
End Sub

Public Sub CargaGruposUsu()
   Dim oPers As COMDPersona.UCOMAcceso
   Dim sMatUsu() As Variant
   Dim i As Integer
   Dim n As Integer
   
   Set oPers = New COMDPersona.UCOMAcceso
        oPers.CargaGrupoUsuarios gsDominio, sMatUsu
   Set oPers = Nothing
   n = UBound(sMatUsu)
   For i = 0 To n
      CboGrupo.AddItem sMatUsu(i)
   Next
   
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
        Me.cboagencia.Clear
        With lrAgenc
            Do While Not .EOF
                cboagencia.AddItem Trim(!cAgeDescripcion) & Space(50) & !cAgeCod
                .MoveNext
            Loop
        End With
    End If
 End Sub

 Public Sub CargaNiveles()
    Dim loCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    Dim lrNiv As ADODB.Recordset
    Set loCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
        Set lrNiv = loCapAut.ObtenerNivelesAut(Trim(Right(CboOperacion, 2)), Trim(Right(cboagencia, 3)))
    Set loCapAut = Nothing
    If lrNiv Is Nothing Then
        MsgBox " No se encuentran los Niveles ", vbInformation, " Aviso "
    Else
        Me.cboNivel.Clear
        With lrNiv
            Do While Not .EOF
                cboNivel.AddItem lrNiv!cNivel
                .MoveNext
            Loop
        End With
    End If
 End Sub

Public Sub CargarDatos()
    Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
    Dim rs As New ADODB.Recordset
    Dim lista As ListItem
    
    Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
        Set rs = oCapAut.ObtenerDatosNivAutRetCanDet(Trim(Right(CboOperacion.Text, 2)), Trim(Right(cboagencia.Text, 3)))
    Set oCapAut = Nothing
        
    lvwNiveles.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF
            Set lista = lvwNiveles.ListItems.Add(, , Format(rs!cOpeTpo, "000"))
            lista.SubItems(1) = rs!cCodAge
            lista.SubItems(2) = rs!cGrupoUsu
            lista.SubItems(3) = rs!cNivel
            rs.MoveNext
        Loop
    End If
    'FraNivel.Enabled = False
   
End Sub

Public Sub Limpiar()
    CboGrupo.ListIndex = -1
    cboNivel.ListIndex = -1
End Sub

'Public Sub EstadoBotones(ByVal bnuevo As Boolean, ByVal beditar As Boolean, ByVal bcancelar As Boolean, ByVal bgrabar As Boolean)
'    CmdNuevo.Enabled = bnuevo
'    CmdEditar.Enabled = beditar
'    CmdCancelar.Enabled = bcancelar
'    CmdGrabar.Enabled = bgrabar
'End Sub

Public Sub EstadoControles(ByVal bNivel As Boolean, ByVal bDatos As Boolean)
    fraNivel.Enabled = bNivel
    FraDatos.Enabled = bDatos
End Sub

Public Function ValidaControles() As Boolean
    Dim i As Integer
    Dim lsNivel As String
    Dim lsOpe As String
    Dim lsAge As String
    Dim lsGrupo As String
    Dim lnNum As Integer
    
    
    ValidaControles = True
    
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

     If CboGrupo.ListIndex = -1 Then
        ValidaControles = False
        MsgBox "Seleccione un Grupo", vbInformation, "Aviso"
        Exit Function
    End If
    
    If cboNivel.ListIndex = -1 Then
        ValidaControles = False
        MsgBox "Seleccione un Nivel", vbInformation, "Aviso"
        Exit Function
    End If
    'Datos no se dupliquen
    For i = 1 To lvwNiveles.ListItems.Count
        lsOpe = Trim(Right(lvwNiveles.ListItems.iTem(i).Text, 2))
        lsAge = Trim(Right(lvwNiveles.ListItems.iTem(1).SubItems(1), 5))
        lsGrupo = Trim(lvwNiveles.ListItems.iTem(1).SubItems(2))
        lsNivel = Trim(Left(lvwNiveles.ListItems.iTem(1).SubItems(3), 3))
        
        If Trim(Left(cboNivel, 3)) = lsNivel And Trim(Right(CboOperacion.Text, 2)) = lsOpe And Trim(Right(cboagencia, 5)) = lsAge And Trim(CboGrupo.Text) = lsGrupo Then
            ValidaControles = False
            MsgBox "Datos ya existen", vbInformation, "Aviso"
            Exit Function
        End If
    Next

End Function

'Private Sub cmdNuevo_Click()
'    Dim oCapAut As COMDCaptaGenerales.COMDCaptAutorizacion
'    gnOpcion = 1
'    Limpiar
'    EstadoBotones False, False, True, True
'    EstadoControles True, True
'End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Private Sub CmdEditar_Click()
'    gnOpcion = 2
'    EstadoBotones False, False, True, True
'    EstadoControles True, True
'
'End Sub

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
            .fields.Append "sGrupoUsu", adVarChar, 250
            .Open
            'Llenar Recordset
           
            .AddNew
            .fields("sNivCod") = Format(Trim(Left(cboNivel.Text, 3)), "000")
            .fields("sOpeTpo") = Trim(Right(CboOperacion.Text, 2))
            .fields("sCodage") = Trim(Right(cboagencia.Text, 3))
            .fields("sGrupoUsu") = Trim(CboGrupo.Text)
            
     End With
     
     
     Set oCapAut = New COMDCaptaGenerales.COMDCaptAutorizacion
           oCapAut.InsertarNilAutRenCanDet rs, gnOpcion
        'By Capi 21012009
        objPista.InsertarPista gsOpeCod, GeneraMovNroPistas(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, "Nivel Grupos"
        'End by


     Set oCapAut = Nothing
     cmdCancelar_Click
     CargarDatos
    
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    Me.lvwNiveles.ListItems.Clear
    '***Agregado por ELRO el 20121112, según OYP-RFC115-2012
    cmdQuitar.Enabled = False
    '***Fin Agregado por ELRO el 20121112******************
End Sub


'Private Sub lvwNiveles_DblClick()
'    Dim i As Integer
'    Dim J As Integer
'  If gnOpcion = 2 Then
'    If lvwNiveles.ListItems.Count > 0 Then
'      i = lvwNiveles.SelectedItem.Index
'      Call UbicaCombo(Me.CboOperacion, Trim(Right(lvwNiveles.SelectedItem.Text, 1)), True)
'      Call UbicaCombo(Me.CboAgencia, Trim(Right(lvwNiveles.SelectedItem.SubItems(2), 3)), True)
'
'      Call UbicaCombo(Me.cboNivel, Trim(Left(lvwNiveles.SelectedItem.SubItems(3), 3)), True)
'      lvwNiveles.ListItems.Remove (i)
'    Else
'      MsgBox "No existe ningún crédito que pueda reasignar", vbInformation, "Aviso"
'    End If
' End If
'End Sub

'***Agregado por ELRO el 2121109, según OYP-RFC115-2012
Private Sub lvwNiveles_ItemClick(ByVal iTem As MSComctlLib.ListItem)
    cmdQuitar.Enabled = True
End Sub

Private Sub CboGrupo_GotFocus()
    cmdQuitar.Enabled = False
End Sub

Private Sub cboNivel_GotFocus()
    cmdQuitar.Enabled = False
End Sub

Private Sub CmdGrabar_GotFocus()
    cmdQuitar.Enabled = False
End Sub

Private Sub CmdCancelar_LostFocus()
    cmdQuitar.Enabled = False
End Sub

Private Sub CboOperacion_GotFocus()
    cmdQuitar.Enabled = False
End Sub

Private Sub CboAgencia_GotFocus()
    cmdQuitar.Enabled = False
End Sub

Private Sub cmdQuitar_Click()
    Dim oCOMDCaptAutorizacion As COMDCaptaGenerales.COMDCaptAutorizacion
    Dim lista As ListItem
    Dim lsGrupoUsu As String
    Dim lsOpeTpo As String
    Dim lsNivCod As String
    Dim lsCodAge As String
        
    Set lista = lvwNiveles.SelectedItem
     
    
    lsOpeTpo = Right(lista.Text, 1)
    lsCodAge = Right(lista.SubItems(1), 2)
    lsGrupoUsu = Trim(lista.SubItems(2))
    lsNivCod = Left(Trim(lista.SubItems(3)), 3)
    If MsgBox("¿Desea quitar el grupo " & RTrim(lista.SubItems(2)) & "?", vbYesNo, "!Aviso¡") = vbYes Then
        Set oCOMDCaptAutorizacion = New COMDCaptaGenerales.COMDCaptAutorizacion
        oCOMDCaptAutorizacion.eliminarNivRetiroCancDet lsGrupoUsu, lsNivCod, lsOpeTpo, lsCodAge
        CargarDatos
    End If
End Sub
'***Agregado por ELRO el 2121109***********************

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
