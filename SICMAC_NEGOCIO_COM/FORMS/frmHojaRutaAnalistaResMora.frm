VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHojaRutaAnalistaResMora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultado de Gestion de Mora"
   ClientHeight    =   8955
   ClientLeft      =   10155
   ClientTop       =   4680
   ClientWidth     =   8220
   Icon            =   "frmHojaRutaAnalistaResMora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   8220
   Begin VB.Frame fraPersona 
      Caption         =   "Cliente"
      Height          =   1575
      Left            =   0
      TabIndex        =   27
      Top             =   240
      Width           =   8055
      Begin VB.TextBox txtCpersCod 
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   480
         Width           =   2655
      End
      Begin VB.CommandButton cmbBuscarPersona 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblNombrePers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   360
         TabIndex        =   30
         Top             =   1080
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmbCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6720
      TabIndex        =   21
      Top             =   8400
      Width           =   1335
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resultado"
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   8055
      Begin VB.Frame frmMotivoNoGestion 
         Caption         =   "Motivo de No Gestión"
         Height          =   4935
         Left            =   8520
         TabIndex        =   22
         Top             =   1080
         Width           =   7455
         Begin VB.TextBox txtGlosaNoGestion 
            Height          =   795
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   1320
            Width           =   6735
         End
         Begin VB.ComboBox cmbMotivoNoGestion 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   600
            Width           =   6735
         End
         Begin VB.Label Label9 
            Caption         =   "Glosa:"
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Motivo:"
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.Frame frmResultadoVisita 
         Caption         =   "Resultado de la Visita"
         Height          =   2295
         Left            =   240
         TabIndex        =   11
         Top             =   3720
         Width           =   7455
         Begin VB.TextBox txtMontoCompromiso 
            Height          =   315
            Left            =   4920
            TabIndex        =   18
            Top             =   1560
            Width           =   2175
         End
         Begin VB.ComboBox cbMoneda 
            Height          =   315
            ItemData        =   "frmHojaRutaAnalistaResMora.frx":030A
            Left            =   2760
            List            =   "frmHojaRutaAnalistaResMora.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1560
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker dtFechaCompromiso 
            Height          =   315
            Left            =   360
            TabIndex        =   14
            Top             =   1560
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            Format          =   125370369
            CurrentDate     =   42171
         End
         Begin VB.ComboBox cmbResultado 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   720
            Width           =   4095
         End
         Begin VB.Label Label7 
            Caption         =   "Monto:"
            Height          =   255
            Left            =   4920
            TabIndex        =   19
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   2760
            TabIndex        =   17
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha de compromiso:"
            Height          =   255
            Left            =   360
            TabIndex        =   15
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Resultado:"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.Frame frmMotivoIncumplimiento 
         Caption         =   "Motivo de Incumplimiento"
         Height          =   975
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   5775
         Begin VB.ComboBox cbMotivoIncumplimiento 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   360
            Width           =   5055
         End
      End
      Begin VB.Frame frmEntrevistaTercero 
         Caption         =   "     Entrevista a tercero"
         Height          =   1335
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   7455
         Begin VB.ComboBox cbTerceroRelac 
            Height          =   315
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txtTerceroNombre 
            Height          =   315
            Left            =   360
            TabIndex        =   5
            Top             =   600
            Width           =   4215
         End
         Begin VB.CheckBox chkEntrevistaTercero 
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   0
            Width           =   255
         End
         Begin VB.Label Label3 
            Caption         =   "Relación:"
            Height          =   255
            Left            =   4800
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.ComboBox cbEstado 
         Height          =   315
         ItemData        =   "frmHojaRutaAnalistaResMora.frx":032F
         Left            =   240
         List            =   "frmHojaRutaAnalistaResMora.frx":0339
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Estado"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmHojaRutaAnalistaResMora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim framLeftini As Integer
Dim cPersCod As String
Dim nLineaRutaId As Integer
Dim dHojaRuta As New DCOMhojaRuta
Dim exito As Boolean
Dim oPersona As COMDPersona.UCOMPersona
Dim bPlaneada As Boolean
Public Function inicio(ByVal pcPersCod As String, ByVal pnLineaRutaId As Integer, ByVal pbPlaneada As Boolean, Optional sNombre As String)
    cPersCod = pcPersCod
    nLineaRutaId = pnLineaRutaId
    bPlaneada = pbPlaneada
    If bPlaneada Then
        cmbBuscarPersona.Enabled = False
        lblNombrePers.Caption = sNombre
        txtCpersCod.Text = pcPersCod
    End If
    Me.Show 1
    inicio = exito
    
End Function

Private Sub cmbAceptar_Click()
    If ValidarForm(cbEstado.ListIndex) Then
        Dim resp As String: resp = MsgBox("¿Está seguro de guardar el resultado de la visita?", vbYesNo, "Confirmar")
        If resp = vbNo Then Exit Sub
        Dim bGestionado As Boolean: bGestionado = (cbEstado.ListIndex = 0)
        Dim bEntrevistaTercero As Boolean: bEntrevistaTercero = (chkEntrevistaTercero.value = 1)
        Dim dFecha As String
        If dHojaRuta.getValCombo(cmbResultado.Text) = 1 Then
            'WIOR 20151125 ****
            If CDate(Format(dtFechaCompromiso.value, "yyyy-mm-dd")) < CDate(Format(gdFecSis, "yyyy-mm-dd")) Then
                MsgBox "Ingrese una Fecha de Compromiso mayor o igual a la actual", vbInformation, "Aviso"
                dtFechaCompromiso.SetFocus
                Exit Sub
            Else
            'WIOR FIN *********
                dFecha = Format(dtFechaCompromiso.value, "YYYY-MM-DD")
            End If 'WIOR FIN *********
        Else
            dFecha = ""
        End If
        Dim nMontoCompromiso As Double
        If txtMontoCompromiso.Text <> "" Then
            nMontoCompromiso = CDbl(txtMontoCompromiso.Text)
        Else
            nMontoCompromiso = -1
        End If
        Dim nMoneda As Integer
        If cbMoneda.ListIndex <> -1 Then
            nMoneda = (cbMoneda.ListIndex + 1)
        Else
            nMoneda = -1
        End If
        
        If bPlaneada Then
            Call dHojaRuta.RegistraVisitaDiariaMora(gsCodUser, cPersCod, Format(gdFecSis, "YYYYMMDD"), bGestionado, bEntrevistaTercero, txtTerceroNombre.Text, dHojaRuta.getValCombo(cbTerceroRelac.Text), txtGlosaNoGestion.Text, dHojaRuta.getValCombo(cmbResultado.Text), dHojaRuta.getValCombo(cbMotivoIncumplimiento.Text), dHojaRuta.getValCombo(cmbResultado.Text) = 1, dFecha, nMoneda, nMontoCompromiso, dHojaRuta.getValCombo(cmbMotivoNoGestion.Text), nLineaRutaId)
            MsgBox "Se ha registrado el resutlado", vbInformation, "Exito"
            exito = True
        Else
            Dim nRes As Integer
            nRes = dHojaRuta.RegistraVisitaDiariaMoraNoPlaneada(gsCodUser, cPersCod, Format(gdFecSis, "YYYYMMDD"), bGestionado, bEntrevistaTercero, txtTerceroNombre.Text, dHojaRuta.getValCombo(cbTerceroRelac.Text), txtGlosaNoGestion.Text, dHojaRuta.getValCombo(cmbResultado.Text), dHojaRuta.getValCombo(cbMotivoIncumplimiento.Text), dHojaRuta.getValCombo(cmbResultado.Text) = 1, dFecha, nMoneda, nMontoCompromiso, dHojaRuta.getValCombo(cmbMotivoNoGestion.Text), nLineaRutaId)
            If nRes = 1 Then
                MsgBox "Se ha registrado el resutlado", vbInformation, "Exito"
            Else
                MsgBox "No se ha registrado el resultado porque la persona elegida no cuenta con ningún crédito vencido", vbInformation, "Alerta"
                Exit Sub
            End If
            exito = False
        End If
        
        Unload Me
    End If
End Sub

Private Sub cmbBuscarPersona_Click()
    Set oPersona = frmBuscaPersona.inicio
    If Not oPersona Is Nothing Then
        lblNombrePers.Caption = oPersona.sPersNombre
        txtCpersCod.Text = oPersona.sPersCod
        cPersCod = oPersona.sPersCod
    End If
End Sub

Private Sub cmbCerrar_Click()
    Dim resp As String: resp = MsgBox("¿Está seguro de salir sin completar el resultado?", vbYesNo, "Confirmar")
    exito = False
    If resp = vbYes Then Unload Me
End Sub

Private Sub Form_Load()
    exito = False
    framLeftini = frmEntrevistaTercero.Left
    elegirModalidad True
    cbEstado.ListIndex = 0
    txtTerceroNombre.Enabled = False
    cbTerceroRelac.Enabled = False
    dtFechaCompromiso.Enabled = False
    cbMoneda.Enabled = False
    txtMontoCompromiso.Enabled = False
    dtFechaCompromiso.value = gdFecSis 'WIOR 20151125
    
    'llenado de combos
    LlenarComboRelacion
    LlenarComboMotivoIncumplimiento
    LlenarComboResultado
    LlenarComboMotivoNoGestion
End Sub

Public Function elegirModalidad(ByVal bGestion As Boolean)
    If bGestion Then
        frmEntrevistaTercero.Visible = True
        frmMotivoIncumplimiento.Visible = True
        frmResultadoVisita.Visible = True
        frmMotivoNoGestion.Visible = False
    Else
        frmEntrevistaTercero.Visible = False
        frmMotivoIncumplimiento.Visible = False
        frmResultadoVisita.Visible = False
        frmMotivoNoGestion.Visible = True
        frmMotivoNoGestion.Left = framLeftini
    End If
End Function
Private Sub cbEstado_Click()
    elegirModalidad cbEstado.ListIndex = 0
End Sub

Private Sub chkEntrevistaTercero_Click()
    If chkEntrevistaTercero.value = 1 Then
        txtTerceroNombre.Enabled = True
        cbTerceroRelac.Enabled = True
    Else
        txtTerceroNombre.Enabled = False
        cbTerceroRelac.Enabled = False
    End If
End Sub

Private Sub LlenarComboRelacion()
Dim CollCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset
    On Error GoTo error
    Set CollCons = New COMDConstantes.DCOMConstantes
    Set R = CollCons.RecuperaConstantes(gPersRelacion)
    Do While Not R.EOF
            cbTerceroRelac.AddItem Trim(R!cConsDescripcion) & Space(150) & R!nConsValor
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set CollCons = Nothing

    Exit Sub

error:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub LlenarComboMotivoIncumplimiento()
Dim CollCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset
    On Error GoTo error
    Set CollCons = New COMDConstantes.DCOMConstantes
    Set R = CollCons.RecuperaConstantes(10070)
    Do While Not R.EOF
            cbMotivoIncumplimiento.AddItem Trim(R!cConsDescripcion) & Space(150) & R!nConsValor
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set CollCons = Nothing
    Exit Sub
error:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub LlenarComboResultado()
Dim CollCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset
    On Error GoTo error
    Set CollCons = New COMDConstantes.DCOMConstantes
    Set R = CollCons.RecuperaConstantes(10071)
    Do While Not R.EOF
            cmbResultado.AddItem Trim(R!cConsDescripcion) & Space(150) & R!nConsValor
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set CollCons = Nothing
    Exit Sub
error:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub LlenarComboMotivoNoGestion()
Dim CollCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset
    On Error GoTo error
    Set CollCons = New COMDConstantes.DCOMConstantes
    Set R = CollCons.RecuperaConstantes(10072)
    Do While Not R.EOF
            cmbMotivoNoGestion.AddItem Trim(R!cConsDescripcion) & Space(150) & R!nConsValor
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set CollCons = Nothing
    Exit Sub
error:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Sub cmbResultado_Click()
    Dim nConsValor As Integer
    nConsValor = CInt(Trim(Right(cmbResultado.Text, 5)))
    If nConsValor = 1 Then
        dtFechaCompromiso.Enabled = True
        cbMoneda.Enabled = True
        txtMontoCompromiso.Enabled = True
    Else
        dtFechaCompromiso.Enabled = False
        cbMoneda.Enabled = False
        txtMontoCompromiso.Enabled = False
    End If
    
End Sub

Private Function ValidarForm(ByVal pNoGestion As Integer)
    Dim ret As Boolean: ret = True
    Dim mensaje As String
    
     If Not bPlaneada Then
        If txtCpersCod.Text = "" Then
            ret = False
            mensaje = mensaje & vbCrLf & ">Debe elegir un cliente, para la visita no planeada"
        End If
    End If
    
    If pNoGestion = 0 Then
        If chkEntrevistaTercero.value = 1 Then
            If txtTerceroNombre.Text = "" Then
                ret = False
                mensaje = mensaje & vbCrLf & ">Parece que no ingresó el nombre del tercero"
            End If
            If cbTerceroRelac.ListIndex = -1 Then
                ret = False
                mensaje = mensaje & vbCrLf & ">Debe elegir la relación del tercero"
            End If
        End If
        
        If cbMotivoIncumplimiento.ListIndex = -1 Then
            ret = False
            mensaje = mensaje & vbCrLf & ">Debe elegir un motivo de incumplimiento"
        End If
        
        If cmbResultado.ListIndex = -1 Then
            ret = False
            mensaje = mensaje & vbCrLf & ">Debe elegir un resultado"
        Else
            If dHojaRuta.getValCombo(cmbResultado.Text) = 1 Then
                If cbMoneda.ListIndex = -1 Then
                    ret = False
                    mensaje = mensaje & vbCrLf & ">Debe elegir una moneda para el monto del compromiso"
                End If
                If txtMontoCompromiso.Text = "" Then
                    ret = False
                    mensaje = mensaje & vbCrLf & ">Parece que no ingresó el monto del compromiso"
                End If
            End If
        End If
        
    Else
        If cmbMotivoNoGestion.ListIndex = -1 Then
            ret = False
            mensaje = mensaje & vbCrLf & ">Debe elegir el motivo por el cual no se gestionó la visita"
        End If
        If txtGlosaNoGestion.Text = "" Then
            ret = False
            mensaje = mensaje & vbCrLf & ">Parece que no ingresó la glosa"
        End If
    End If
    If Not ret Then
            MsgBox mensaje
    End If
    
    ValidarForm = ret
End Function
Private Sub txtMontoCompromiso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMontoCompromiso.SetFocus
    Else
        KeyAscii = NumerosDecimales(txtMontoCompromiso, KeyAscii, , , False)
    End If
End Sub
