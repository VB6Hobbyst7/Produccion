VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredReprogAprobacion 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   Icon            =   "frmCredReprogAprobacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   11265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBuscar 
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   5895
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1170
   End
   Begin VB.TextBox txtGlosa 
      Height          =   315
      Left            =   2040
      TabIndex        =   4
      Top             =   4080
      Width           =   6015
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "Aprobar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   2
      Top             =   4080
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9960
      TabIndex        =   1
      Top             =   4080
      Width           =   1170
   End
   Begin MSComctlLib.ListView LstCreditos 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   5953
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
      Enabled         =   0   'False
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº"
         Object.Width           =   531
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Nº Credito"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Titular"
         Object.Width           =   6350
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Moneda"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Saldo Cap."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Dias a Reprog."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cuotas Reprog."
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "bSolic"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdRegistrarVB 
      Caption         =   "Registrar V°B°"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   6
      Top             =   4080
      Width           =   1410
   End
   Begin VB.Label lblBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Glosa:"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   4080
      Width           =   615
   End
End
Attribute VB_Name = "frmCredReprogAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************************
'** Nombre : frmCredReprogAprobacion
'** Descripción : Formulario para listar créditos propuestos y aprobar la reprogramación de créditos
'**               según TI-ERS010-2016
'** Creación : JUEZ, 20160314 09:00:00 AM
'*********************************************************************************************************

Option Explicit

Private Enum TipoAcceso
    TipoAprobacion = 0
    TipoVBAdmCred = 1
End Enum

Dim fnTipo As TipoAcceso
Dim oDCred As COMDCredito.DCOMCredito
Dim rs As ADODB.Recordset
Dim sMovNro As String
Dim rsFiltro As ADODB.Recordset 'JOEP ERS039 20170706
Dim nTpPermiso As Integer 'Agrego JOEP20171214 ACTA220-2017

Private Sub cmdActualizar_Click()
    txtBuscar.Text = "" 'Agrego JOEP20171214 ACTA220-2017
    CargaCreditosPropuestos
End Sub

Public Sub Inicio(ByVal pnTipo As Integer)
    fnTipo = pnTipo
    Select Case fnTipo
        Case TipoAprobacion
            cmdAprobar.Visible = True
            cmdRegistrarVB.Visible = False
            Me.Caption = "Aprobación de Propuestas de Reprogramación"
        Case TipoVBAdmCred
            cmdAprobar.Visible = False
            cmdRegistrarVB.Visible = True
            Me.Caption = "Revisión de Reprogramaciones - Administración de Créditos"
    End Select
    ValidarFechaActual
    CargaCreditosPropuestos

If fnTipo = 0 And (nTpPermiso = 0 Or nTpPermiso = 3 Or nTpPermiso = 4) Then 'Agrego JOEP20171214 ACTA220-2017
Else 'Agrego JOEP20171214 ACTA220-2017
    Me.Show 1
End If 'Agrego JOEP20171214 ACTA220-2017

End Sub

Private Sub ValidarFechaActual()
Dim lsFechaValidador As String

    lsFechaValidador = validarFechaSistema
    If lsFechaValidador <> "" Then
        If gdFecSis <> CDate(lsFechaValidador) Then
            MsgBox "La Fecha de tu sesión en el Negocio no coincide con la fecha del Sistema", vbCritical, "Aviso"
            Unload Me
            End
        End If
    End If
End Sub

Private Sub CargaCreditosPropuestos()
Dim nEstado As ColocacReprogEstado
Dim L As ListItem
Dim PermisoAproReprog As COMNCredito.NCOMCredito 'Agrego JOEP20171214 ACTA220-2017

nTpPermiso = 0 'Agrego JOEP20171214 ACTA220-2017

    If fnTipo = TipoAprobacion Then
        nEstado = gEstReprogPropuesto
        
    'Inicio Agrego JOEP20171214 ACTA220-2017
        Set PermisoAproReprog = New COMNCredito.NCOMCredito
        '(1:COORDINADOR DE CRÉDITOS ,2: GERENTE DE RIESGOS)
        nTpPermiso = PermisoAproReprog.ObtieneTipoPermisoReprog(gsCodCargo) ' Obtener el tipo de Permiso, Segun Cargo
        
        If (nTpPermiso = 0 Or nTpPermiso = 3 Or nTpPermiso = 4) Then
            MsgBox "No tiene permiso para Aprobar", vbInformation, "Aviso"
            Exit Sub
        End If
    'Fin Agrego JOEP20171214 ACTA220-2017
        
    ElseIf fnTipo = TipoVBAdmCred Then
        nEstado = gEstReprogReprogramado
    End If
    
    Set oDCred = New COMDCredito.DCOMCredito
        Set rs = oDCred.RecuperaColocacReprogramadoPorAprobar(nEstado, nTpPermiso)
        Set rsFiltro = rs.Clone 'JOEP ERS039 20170706
    Set oDCred = Nothing
    
    LstCreditos.ListItems.Clear
    txtGlosa.Text = ""
    
    If Not rs.EOF And Not rs.BOF Then
        LstCreditos.Enabled = True
        
        Do While Not rs.EOF
            Set L = LstCreditos.ListItems.Add(, , rs.Bookmark)
            L.SubItems(1) = rs!cCtaCod
            L.SubItems(2) = rs!cPersNombre
            L.SubItems(3) = rs!cMoneda
            L.SubItems(4) = Format(rs!nSaldoCap, "#,##0.00")
            L.SubItems(5) = DateDiff("d", CDate(rs!dFecCuotaVenc), CDate(rs!dFecNuevaCuotaVenc))
            L.SubItems(6) = rs!nCuotasReprog
            rs.MoveNext
        Loop
    Else
        LstCreditos.Enabled = False
        MsgBox "No existen créditos pendientes por aprobar", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdAprobar_Click()
Dim oNCred As COMNCredito.NCOMCredito
Dim oVisto As frmVistoElectronico
Dim lbResultadoVisto As Boolean
    
    If LstCreditos.ListItems.count = 0 Then
        MsgBox "No existen datos para aprobar", vbInformation, "Aviso"
        Exit Sub
    End If
    
'JOEP20201008 add Tasa especial reduccion de monto
    If ValidaDatos = True Then
        Exit Sub
    End If
        
    If MsgBox("Se va a aprobar la reprogramación del crédito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oNCred = New COMNCredito.NCOMCredito
    Call oNCred.RegistraReprogramacionEstado(LstCreditos.SelectedItem.SubItems(1), gdFecSis, gEstReprogAprobado, CDbl(LstCreditos.SelectedItem.SubItems(4)), sMovNro, , , , , Trim(txtGlosa.Text))
    Set oNCred = Nothing
    MsgBox "La propuesta fue aprobada. El crédito está listo para ser reprogramado", vbInformation, "Aviso"
    CargaCreditosPropuestos
'JOEP20201008 add Tasa especial reduccion de monto

'JOEP20201008 comento Tasa especial reduccion de monto
'    If Trim(txtGlosa.Text) <> "" Then
'        If MsgBox("Se va a aprobar la reprogramación del crédito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'        sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'        'Comento JOEP20171214 ACTA220-2017
''        Set oVisto = New frmVistoElectronico
''        lbResultadoVisto = False
''        lbResultadoVisto = oVisto.Inicio(2, "")
''        If Not lbResultadoVisto Then
''            Exit Sub
''        End If
'        'Comento JOEP20171214 ACTA220-2017
'
'        Set oNCred = New COMNCredito.NCOMCredito
'            Call oNCred.RegistraReprogramacionEstado(LstCreditos.SelectedItem.SubItems(1), gdFecSis, gEstReprogAprobado, CDbl(LstCreditos.SelectedItem.SubItems(4)), sMovNro, , , , , Trim(txtGlosa.Text))
'        Set oNCred = Nothing
'
'        MsgBox "La propuesta fue aprobada. El crédito está listo para ser reprogramado", vbInformation, "Aviso"
'        CargaCreditosPropuestos
'    Else
'        MsgBox "Debe ingresar una glosa para aprobar", vbInformation, "Aviso"
'        Me.txtGlosa.SetFocus
'    End If
'JOEP20201008 comento Tasa especial reduccion de monto
End Sub

Private Sub cmdRegistrarVB_Click()
Dim oNCred As COMNCredito.NCOMCredito
Dim oVisto As frmVistoElectronico
Dim lbResultadoVisto As Boolean
    
    If LstCreditos.ListItems.count = 0 Then
        MsgBox "No existen datos para revisar", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If Trim(txtGlosa.Text) <> "" Then
        If MsgBox("Se va a registrar la revisión de la reprogramación del crédito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            
        sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        Set oVisto = New frmVistoElectronico
        lbResultadoVisto = False
        lbResultadoVisto = oVisto.Inicio(2, "")
        If Not lbResultadoVisto Then
            Exit Sub
        End If
        
        Set oNCred = New COMNCredito.NCOMCredito
            Call oNCred.RegistraReprogramacionEstado(LstCreditos.SelectedItem.SubItems(1), gdFecSis, gEstReprogVBAdmCred, CDbl(LstCreditos.SelectedItem.SubItems(4)), sMovNro, , , , , Trim(txtGlosa.Text))
        Set oNCred = Nothing
        
        MsgBox "La reprogramación fue revisada.", vbInformation, "Aviso"
        CargaCreditosPropuestos
    Else
        MsgBox "Debe ingresar una glosa para registrar la revisión", vbInformation, "Aviso"
        Me.txtGlosa.SetFocus
    End If
End Sub

'Inicio JOEP ERS039 20170706
Private Sub txtBuscar_Change()

Dim L As ListItem

If txtBuscar.Text <> "" Then
    
        rsFiltro.Filter = "cPersNombre like '*" + txtBuscar.Text + "*'"

        If Not rsFiltro.EOF And Not rsFiltro.BOF Then
            LstCreditos.ListItems.Clear
            LstCreditos.Enabled = True
        
            Do While Not rsFiltro.EOF
                Set L = LstCreditos.ListItems.Add(, , rsFiltro.Bookmark)
                L.SubItems(1) = rsFiltro!cCtaCod
                L.SubItems(2) = rsFiltro!cPersNombre
                L.SubItems(3) = rsFiltro!cMoneda
                L.SubItems(4) = Format(rsFiltro!nSaldoCap, "#,##0.00")
                L.SubItems(5) = DateDiff("d", CDate(rsFiltro!dFecCuotaVenc), CDate(rsFiltro!dFecNuevaCuotaVenc))
                L.SubItems(6) = rsFiltro!nCuotasReprog
                rsFiltro.MoveNext
            Loop
        End If
Else
    Call cmdActualizar_Click
End If
    
End Sub
'Fin JOEP ERS039 20170706

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fnTipo = TipoAprobacion Then
            cmdAprobar.SetFocus
        ElseIf fnTipo = TipoVBAdmCred Then
            cmdRegistrarVB.SetFocus
        End If
    End If
End Sub
'Inicio JOEP ERS039 20170706
Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras2(KeyAscii, True)
End Sub
'Fin JOEP ERS039 20170706

Private Function ValidaDatos()
Dim obValDts As New COMDCredito.DCOMCredito
Dim rsValDts As ADODB.Recordset

Set obValDts = New COMDCredito.DCOMCredito
Set rsValDts = obValDts.ReprogramacionValDtsAprobacion(LstCreditos.SelectedItem.SubItems(1), fnTipo, txtGlosa)

ValidaDatos = False

If Not (rsValDts.BOF And rsValDts.EOF) Then
    If rsValDts!MsgBox <> "" Then
        MsgBox rsValDts!MsgBox, vbInformation, "Aviso"
        ValidaDatos = True
    End If
End If

Set obValDts = Nothing
RSClose rsValDts
End Function
