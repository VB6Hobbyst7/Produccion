VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredExoneraCobertura 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exoneración Cobertura"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10665
   Icon            =   "frmCredExoneraCobertura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   10665
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "A&ctualizar"
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   3495
      Width           =   1050
   End
   Begin VB.TextBox txtTasa 
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
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   6240
      MaxLength       =   15
      TabIndex        =   2
      Text            =   "0.0000"
      Top             =   3525
      Width           =   840
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      Height          =   360
      Left            =   8415
      TabIndex        =   4
      Top             =   3495
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9520
      TabIndex        =   5
      Top             =   3495
      Width           =   1050
   End
   Begin VB.TextBox txtGlosa 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   3525
      Width           =   3015
   End
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "&Autorizar"
      Height          =   360
      Left            =   7320
      TabIndex        =   3
      Top             =   3495
      Width           =   1050
   End
   Begin MSComctlLib.ListView LstSolicitud 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   5741
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
      NumItems        =   18
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   531
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Usuario"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Agencia"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Nro. Crédito"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Titular"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Producto"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Moneda"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Monto"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Comentario Solicitud"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Estado"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Cobertura"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Fecha Autoriza"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Usuario Autoriza"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Comentario Autoriza"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Tipo Producto"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   17
         Text            =   "Ratio Cobertura"
         Object.Width           =   1411
      EndProperty
   End
   Begin VB.Label lblPorcentaje 
      AutoSize        =   -1  'True
      Caption         =   "%"
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
      Left            =   7080
      TabIndex        =   8
      Top             =   3570
      Width           =   150
   End
   Begin VB.Label lblTasaCobertura 
      AutoSize        =   -1  'True
      Caption         =   "Cobertura:"
      Height          =   195
      Left            =   5400
      TabIndex        =   7
      Top             =   3570
      Width           =   735
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Comentario :"
      Height          =   195
      Left            =   1320
      TabIndex        =   6
      Top             =   3570
      Width           =   885
   End
End
Attribute VB_Name = "frmCredExoneraCobertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************
'** Nombre : frmCredExoneraCobertura
'** Descripción : Formulario para autorizar/rechazar solicitudes de exoneraciones coberturas
'** Creación : EJVG, 20150908 09:00:00 AM
'*******************************************************************************************
Option Explicit
Dim fnTasaIni As Double
Dim objProducto As COMDCredito.DCOMCredito '**ARLO20180712 ERS042 - 2018

Private Sub cmdActualizar_Click()
    CargarSolicitudes
End Sub

Private Sub cmdAutorizar_Click()
    AutorizarRechazarSolicitud (True)
End Sub
Private Sub cmdRechazar_Click()
    AutorizarRechazarSolicitud (False)
End Sub
Private Sub AutorizarRechazarSolicitud(ByVal pbAutoriza As Boolean)
    Dim obj As COMDCredito.DCOMCredActBD
    Dim lsMovNro As String
    Dim lsCtaCod As String
    Dim lsTpoProdCod As String
    Dim lnID As Long
    
    On Error GoTo ErrAutorizar
    
    lsCtaCod = LstSolicitud.SelectedItem.SubItems(4)
    lsTpoProdCod = LstSolicitud.SelectedItem.SubItems(16)
    lnID = CLng(LstSolicitud.SelectedItem.SubItems(15))
    
    If Trim(txtGlosa.Text) = "" Then
        MsgBox "Debe ingresar el comentario", vbInformation, "Aviso"
        EnfocaControl txtGlosa
        Exit Sub
    End If
    If pbAutoriza Then
        If Not IsNumeric(txtTasa.Text) Then
            MsgBox "Ud. debe de especificar la Cobertura", vbInformation, "Aviso"
            EnfocaControl txtTasa
            Exit Sub
        Else
            '**ARLO20180712 ERS042 - 2018
            Set objProducto = New COMDCredito.DCOMCredito
            If objProducto.GetResultadoCondicionCatalogo("N0000140", lsTpoProdCod) Then
            'If lsTpoProdCod = "703" Then 'RapiFlash
            '**ARLO20180712 ERS042 - 2018
                If CCur(txtTasa.Text) < 50 Then
                    MsgBox "La Cobertura RapiFlash no puede ser menor a 50.00%", vbInformation, "Aviso"
                    EnfocaControl txtTasa
                    Exit Sub
                End If
                If CCur(txtTasa.Text) > 100 Then
                    MsgBox "La Cobertura RapiFlash no puede ser mayor a 100.00%", vbInformation, "Aviso"
                    EnfocaControl txtTasa
                    Exit Sub
                End If
            Else
                If CCur(txtTasa.Text) < 1 Then
                    MsgBox "La Cobertura no puede ser menor a 1.00", vbInformation, "Aviso"
                    EnfocaControl txtTasa
                    Exit Sub
                End If
                If CCur(txtTasa.Text) >= fnTasaIni Then
                    MsgBox "La Cobertura no puede ser mayor o igual a " & Format(fnTasaIni, "#0.0000"), vbInformation, "Aviso"
                    EnfocaControl txtTasa
                    Exit Sub
                End If
            End If
        End If
    Else
        txtTasa.Text = "0.0000"
    End If
    
    
    If pbAutoriza Then
        If MsgBox("Esta opción autorizará la Cobertura de " & Format(txtTasa.Text, "#0.0000") & " para el Registro de Coberturas del Crédito N° " & lsCtaCod & Chr(13) & Chr(13) & "¿Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Else
        If MsgBox("Esta opción rechazará la Solicitud de Tasa del Crédito N° " & lsCtaCod & Chr(13) & Chr(13) & "¿Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    End If
    
    Set obj = New COMDCredito.DCOMCredActBD
    lsMovNro = obj.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Call obj.dUpdateSolicitudExoneraCobertura(lnID, pbAutoriza, lsMovNro, CCur(txtTasa.Text), Trim(Me.txtGlosa.Text))
    Set obj = Nothing
    
    LstSolicitud.SelectedItem.SubItems(10) = IIf(pbAutoriza, "AUTORIZADA", "RECHAZADA")
    LstSolicitud.SelectedItem.SubItems(11) = txtTasa.Text
    LstSolicitud.SelectedItem.SubItems(12) = Format(fgFechaHoraMovDate(lsMovNro), gsFormatoFechaHoraViewAMPM)
    LstSolicitud.SelectedItem.SubItems(13) = gsCodUser
    LstSolicitud.SelectedItem.SubItems(14) = txtGlosa.Text
    
    LstSolicitud.ListItems.Remove (LstSolicitud.SelectedItem.Index)
    
    SeleccionarRegistro
    
    MsgBox "La solicitud de Tasa para el Crédito N° " & lsCtaCod & " fue " & IIf(pbAutoriza, "autorizada", "rechazada"), vbInformation, "Aviso"
    Exit Sub
ErrAutorizar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    cmdActualizar_Click
End Sub
Private Sub CargarSolicitudes()
    Dim L As ListItem
    Dim obj As New COMDCredito.DCOMCreditos
    Dim rs As New ADODB.Recordset
    
    On Error GoTo ErrCargarSolicitudes
    Screen.MousePointer = 11
    
    Set rs = obj.dListaSolicitudExoneraCobertura(gsCodUser)
    Set obj = Nothing

    LstSolicitud.ListItems.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            Set L = LstSolicitud.ListItems.Add(, , rs.Bookmark)
            L.SubItems(1) = Format(rs!dFechaSolicita, gsFormatoFechaHoraViewAMPM)
            L.SubItems(2) = rs!cUserSolicita
            L.SubItems(3) = rs!cAgeDescripcion
            L.SubItems(4) = rs!cCtaCod
            L.SubItems(5) = Trim(rs!cPersNombre)
            L.SubItems(6) = rs!cTpoProdDesc
            L.SubItems(7) = rs!cMoneda
            L.SubItems(8) = Format(rs!nMonto, "#,##0.00")
            L.SubItems(9) = rs!cComentarioSolicitud
            L.SubItems(10) = rs!cEstado
            L.SubItems(11) = IIf(rs!cEstado = "PENDIENTE", "", rs!nTasa)
            L.SubItems(12) = IIf(rs!cEstado = "PENDIENTE", "", Format(rs!dFechaAutoriza, gsFormatoFechaHoraViewAMPM))
            L.SubItems(13) = IIf(rs!cEstado = "PENDIENTE", "", rs!cUserAutoriza)
            L.SubItems(14) = IIf(rs!cEstado = "PENDIENTE", "", rs!cComentarioAutoriza)
            L.SubItems(15) = rs!nId
            L.SubItems(16) = rs!cTpoProdCod
            L.SubItems(17) = Format(rs!nTasaIni, "#0.0000")
            rs.MoveNext
        Loop
    End If
    RSClose rs
    
    SeleccionarRegistro
    
    Screen.MousePointer = 0
    Exit Sub
ErrCargarSolicitudes:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub LstSolicitud_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SeleccionarRegistro
End Sub
Private Sub SeleccionarRegistro()
    txtGlosa.Text = ""
    lblTasaCobertura.Caption = "Cobertura :"
    txtTasa.Text = "0.0000"
    lblPorcentaje.Caption = ""
    
    '**ARLO20180712 ERS042 - 2018
    Dim lsTpoProdCod As String
    lsTpoProdCod = LstSolicitud.SelectedItem.SubItems(16)
    '**ARLO20180712 ERS042 - 2018
    
    HabilitarControles False
    If LstSolicitud.ListItems.count > 0 Then
        EnfocaControl LstSolicitud
        If LstSolicitud.SelectedItem.SubItems(10) = "PENDIENTE" Then
            HabilitarControles True
            '**ARLO20180712 ERS042 - 2018
            Set objProducto = New COMDCredito.DCOMCredito
            If objProducto.GetResultadoCondicionCatalogo("N0000141", lsTpoProdCod) Then
            'If LstSolicitud.SelectedItem.SubItems(16) = "703" Then
            '**ARLO20180712 ERS042 - 2018
                lblTasaCobertura.Caption = "RapiFlash:"
                lblPorcentaje.Caption = "%"
            End If
            fnTasaIni = CCur(LstSolicitud.SelectedItem.SubItems(17))
            txtTasa.Text = Format(fnTasaIni, "#0.0000")
        End If
    End If
End Sub
Private Sub LstSolicitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtGlosa
    End If
End Sub
Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl txtTasa
    End If
End Sub
Private Sub txtGlosa_LostFocus()
    txtGlosa.Text = Trim(txtGlosa.Text)
End Sub
Private Sub txtTasa_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtTasa, KeyAscii, 15, 4)
    If KeyAscii = 13 Then
        EnfocaControl cmdAutorizar
    End If
End Sub
Private Sub txtTasa_LostFocus()
    txtTasa.Text = Format(txtTasa.Text, "#0.0000")
End Sub
Private Sub HabilitarControles(ByVal pbHabilita As Boolean)
    txtGlosa.Enabled = pbHabilita
    txtTasa.Enabled = pbHabilita
    cmdAutorizar.Enabled = pbHabilita
    cmdRechazar.Enabled = pbHabilita
End Sub
