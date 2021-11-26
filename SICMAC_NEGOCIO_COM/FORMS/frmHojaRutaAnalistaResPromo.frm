VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmHojaRutaAnalistaResPromo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultado de Promociones"
   ClientHeight    =   9450
   ClientLeft      =   11805
   ClientTop       =   4080
   ClientWidth     =   6390
   Icon            =   "frmHojaRutaAnalistaResPromo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   6390
   Begin VB.Frame fraPersona 
      Caption         =   "Cliente"
      Height          =   1575
      Left            =   120
      TabIndex        =   24
      Top             =   120
      Width           =   6135
      Begin VB.CommandButton cmbBuscarPersona 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtCpersCod 
         Height          =   315
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblNombrePers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   360
         TabIndex        =   25
         Top             =   1080
         Width           =   5415
      End
   End
   Begin VB.CommandButton cbmCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5160
      TabIndex        =   18
      Top             =   9000
      Width           =   1095
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resultado"
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   6135
      Begin VB.Frame framNogestionado 
         Height          =   5655
         Left            =   6840
         TabIndex        =   19
         Top             =   1080
         Width           =   5535
         Begin VB.TextBox txtGlosaNoGestion 
            Height          =   735
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   1320
            Width           =   4815
         End
         Begin VB.ComboBox cbMotivoNoGestion 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label Label8 
            Caption         =   "Glosa:"
            Height          =   255
            Left            =   360
            TabIndex        =   23
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label7 
            Caption         =   "Motivo:"
            Height          =   255
            Left            =   360
            TabIndex        =   21
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox txtGlosa 
         Height          =   735
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   6000
         Width           =   5535
      End
      Begin VB.Frame framResultadoVisita 
         Caption         =   "Resultado de la Visita"
         Height          =   2175
         Left            =   360
         TabIndex        =   9
         Top             =   3360
         Width           =   5535
         Begin MSComCtl2.DTPicker dtHoraVisita 
            Height          =   315
            Left            =   3600
            TabIndex        =   14
            Top             =   1440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   125370370
            CurrentDate     =   42171
         End
         Begin MSComCtl2.DTPicker dtFechaVisita 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   1440
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            Format          =   125370369
            CurrentDate     =   42171
         End
         Begin VB.ComboBox cbResultadoVisita 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label Label5 
            Caption         =   "Fecha de la Próxima Visita:"
            Height          =   255
            Left            =   360
            TabIndex        =   13
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label4 
            Caption         =   "Resultado:"
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.CheckBox chkEntrevistaTercero 
         Caption         =   "Entrevista a Tercero"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Frame framEntrevistaTercero 
         Height          =   2055
         Left            =   360
         TabIndex        =   3
         Top             =   1080
         Width           =   5535
         Begin VB.ComboBox cbTerceroRelac 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txtNombreTercero 
            Height          =   315
            Left            =   360
            TabIndex        =   5
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label Label3 
            Caption         =   "Relación:"
            Height          =   255
            Left            =   360
            TabIndex        =   8
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.ComboBox cbEstado 
         Height          =   315
         ItemData        =   "frmHojaRutaAnalistaResPromo.frx":030A
         Left            =   360
         List            =   "frmHojaRutaAnalistaResPromo.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   5760
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmHojaRutaAnalistaResPromo"
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


Private Sub cmbBuscarPersona_Click()
    Set oPersona = frmBuscaPersona.inicio
    If Not oPersona Is Nothing Then
        lblNombrePers.Caption = oPersona.sPersNombre
        txtCpersCod.Text = oPersona.sPersCod
        cPersCod = oPersona.sPersCod
    End If
End Sub
Private Sub cbmCerrar_Click()
    Dim resp As String: resp = MsgBox("¿Está seguro de salir sin completar el resultado?", vbYesNo, "Confirmar")
    exito = False
    If resp = vbYes Then Unload Me
End Sub

Private Sub cmbAceptar_Click()
    If ValidarForm(cbEstado.ListIndex) Then
    
        Dim resp As String: resp = MsgBox("¿Está seguro de guardar el resultado de la visita?", vbYesNo, "Confirmar")
        If resp = vbNo Then Exit Sub
        Dim bGestionado As Boolean: bGestionado = (cbEstado.ListIndex = 0)
        Dim bEntrevistaTercero As Boolean: bEntrevistaTercero = (chkEntrevistaTercero.value = 1)
        Dim dFechaVisita As String
        If dHojaRuta.getValCombo(cbResultadoVisita.Text) = 2 Then
            'WIOR 20151125 ****
            If CDate(Format(dtFechaVisita.value, "yyyy-mm-dd")) <= CDate(Format(gdFecSis, "yyyy-mm-dd")) Then
                MsgBox "Ingrese una Fecha de Visita mayor a la actual", vbInformation, "Aviso"
                dtFechaVisita.SetFocus
                Exit Sub
            Else
            'WIOR FIN *********
                dFechaVisita = Format(dtFechaVisita.value, "YYYY-MM-DD") & " " & Format(dtHoraVisita.value, "HH:mm:ss")
            End If 'WIOR 20151125 ****
        Else
            dFechaVisita = ""
        End If
        
        If bPlaneada Then
            Call dHojaRuta.RegistraVisitaDiariaPromocion(gsCodUser, cPersCod, Format(gdFecSis, "YYYYMMDD"), bGestionado, bEntrevistaTercero, txtNombreTercero.Text, dHojaRuta.getValCombo(cbTerceroRelac.Text), txtGlosaNoGestion.Text, dHojaRuta.getValCombo(cbResultadoVisita), dFechaVisita, txtGlosa.Text, dHojaRuta.getValCombo(cbMotivoNoGestion.Text), nLineaRutaId)
            MsgBox "Se ha registrado el resutlado", vbInformation, "Exito"
            exito = True
        Else
            Call dHojaRuta.RegistraVisitaDiariaPromocionNoPlaneada(gsCodUser, cPersCod, Format(gdFecSis, "YYYYMMDD"), bGestionado, bEntrevistaTercero, txtNombreTercero.Text, dHojaRuta.getValCombo(cbTerceroRelac.Text), txtGlosaNoGestion.Text, dHojaRuta.getValCombo(cbResultadoVisita), dFechaVisita, txtGlosa.Text, dHojaRuta.getValCombo(cbMotivoNoGestion.Text), nLineaRutaId)
            MsgBox "Se ha registrado el resutlado", vbInformation, "Exito"
            exito = False
        End If
        
        Unload Me
    End If
End Sub
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

Private Sub cbResultadoVisita_Click()
    Dim nConsValor As Integer
    nConsValor = CInt(Trim(Right(cbResultadoVisita.Text, 5)))
    If nConsValor = 2 Then
        dtFechaVisita.Enabled = True
        dtHoraVisita.Enabled = True
    Else
        dtFechaVisita.Enabled = False
        dtHoraVisita.Enabled = False
    End If
End Sub

Private Sub chkEntrevistaTercero_Click()
    If chkEntrevistaTercero.value = 1 Then
        framEntrevistaTercero.Enabled = True
    Else
        framEntrevistaTercero.Enabled = False
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
            If txtNombreTercero.Text = "" Then
                ret = False
                mensaje = mensaje & vbCrLf & ">Parece que no ingresó el nombre del tercero"
            End If
            If cbTerceroRelac.ListIndex = -1 Then
                ret = False
                mensaje = mensaje & vbCrLf & ">Debe elegir la relación del tercero"
            End If
        End If
        If cbResultadoVisita.ListIndex = -1 Then
            ret = False
            mensaje = mensaje & vbCrLf & ">Debe elegir un resultado para la visita"
        End If
        If txtGlosa.Text = "" Then
            ret = False
            mensaje = mensaje & vbCrLf & ">Parece que no ingresó la glosa"
        End If
    Else
        If cbMotivoNoGestion.ListIndex = -1 Then
            ret = False
            mensaje = mensaje & vbCrLf & ">Parece que no eligió el motivo por el cual no se gestionó"
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



Private Sub Form_Load()
    txtCpersCod.Enabled = False
    exito = False
    framLeftini = framEntrevistaTercero.Left
    elegirModalidad True
    cbEstado.ListIndex = 0
    dtFechaVisita.Enabled = False
    dtHoraVisita.Enabled = False
    framEntrevistaTercero.Enabled = False
    dtFechaVisita.value = gdFecSis 'WIOR 20151125
    
    'llenar combos
    LlenarComboRelacion
    LlenarComboMotivoNoGestion
    LlenarComboResultado
End Sub
Public Function elegirModalidad(ByVal bGestion As Boolean)
    If bGestion Then
        framEntrevistaTercero.Visible = True
        framResultadoVisita.Visible = True
        txtGlosa.Visible = True
        framNogestionado.Visible = False
    Else
        framEntrevistaTercero.Visible = False
        framResultadoVisita.Visible = False
        txtGlosa.Visible = False
        framNogestionado.Visible = True
        framNogestionado.Left = framLeftini
    End If
End Function
Private Sub cbEstado_Click()
    elegirModalidad cbEstado.ListIndex = 0
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

Private Sub LlenarComboResultado()
Dim CollCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset
    On Error GoTo error
    Set CollCons = New COMDConstantes.DCOMConstantes
    Set R = CollCons.RecuperaConstantes(10073)
    Do While Not R.EOF
            cbResultadoVisita.AddItem Trim(R!cConsDescripcion) & Space(150) & R!nConsValor
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
            cbMotivoNoGestion.AddItem Trim(R!cConsDescripcion) & Space(150) & R!nConsValor
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set CollCons = Nothing

    Exit Sub

error:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

