VERSION 5.00
Begin VB.Form frmSolicitudContinuidadProcesoCredRPLAFT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de Continuidad"
   ClientHeight    =   2580
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7440
   Icon            =   "frmSolicitudContinuidadProcesoCredRPLAFT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Solicitud:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtCodigo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   8
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtNombre 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3600
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5760
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Enviar"
         Height          =   375
         Left            =   4440
         TabIndex        =   2
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox cboRelacion 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Relacion del Cliente con el Credito"
         Top             =   840
         Width           =   2625
      End
      Begin VB.Label Label4 
         Caption         =   "Relación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   390
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSolicitudContinuidadProcesoCredRPLAFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oPersona As UPersona_Cli
Dim nExito As Integer
Dim cPersCodCliente As String
Dim nCondicion As Integer
Dim cCondicion As String
Dim cOpeCod As String
Dim cUserSolicitud As String
Dim cAgeCod As String

Private Sub cmdAceptar_Click()
    Me.cmdAceptar.Enabled = False
    If ValidarDatos = True Then
        Dim oPerSolicitudContinuidad As New comdpersona.DCOMPersonas
        nExito = oPerSolicitudContinuidad.InsertarPersRPLAFTVistoContinuidad(cPersCodCliente, nCondicion, cOpeCod, cUserSolicitud, cAgeCod, CInt(Trim(Right(cboRelacion.Text, 2))))
        If nExito = 0 Then
            MsgBox "La Solicitud no fue satisfactoria, inténtalo nuevamente", vbExclamation, "Aviso"
            Me.cmdAceptar.Enabled = True
        End If
        If nExito > 0 Then
            Call EnviarSolicitud
            MsgBox "Solicitud Enviada con éxito! " & Chr(13) & vbNewLine & "Esperando Visto de Continuidad del Proceso del Crédito", vbInformation, "Aviso"
            Unload Me
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    nExito = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Call CargarDatos
End Sub

Public Function Inicio(pcPersCodCliente As String, pnCondicion As Integer, pcCondicion As String, pcOpeCod As String, pcUserSolicitud As String, pcAgeCod As String) As Integer
    cPersCodCliente = pcPersCodCliente
    nCondicion = pnCondicion
    cCondicion = pcCondicion
    cOpeCod = pcOpeCod
    cUserSolicitud = pcUserSolicitud
    cAgeCod = pcAgeCod
    Me.Show 1
    Inicio = nExito
End Function

Private Sub CargarDatos()
    Set oPersona = New UPersona_Cli
    oPersona.RecuperaPersona (Trim(cPersCodCliente))
    Me.txtCodigo.Text = cPersCodCliente
    Me.txtNombre.Text = oPersona.NombreCompleto
    nExito = 0
    Call CargarCombo
End Sub

Private Sub CargarCombo()
Dim CollCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

    On Error GoTo ErrorCargaControlRelaCred
    
    Set CollCons = New COMDConstantes.DCOMConstantes
    Set R = CollCons.RecuperaConstantes(gColocRelacPers)
    Do While Not R.EOF
        If R!nConsValor >= 20 And R!nConsValor <= 25 Then
            cboRelacion.AddItem Trim(R!cConsDescripcion) & Space(80) & R!nConsValor
        End If
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Set CollCons = Nothing

    Exit Sub

ErrorCargaControlRelaCred:
        MsgBox Err.Description, vbCritical, "Aviso"
End Sub

Private Function ValidarDatos() As Boolean
    Dim Validado As Boolean
    Validado = True
    If Me.cboRelacion.Text = "" Then
        MsgBox "Debe seleccionar el tipo de relación", vbInformation, "Aviso"
        Validado = False
        Me.cboRelacion.SetFocus
        Me.cmdAceptar.Enabled = True
    End If
    ValidarDatos = Validado
End Function

Private Sub EnviarSolicitud()
    Dim oConst As COMDConstSistema.NCOMConstSistema
    Dim lsCorreoDestino As String, lsContenido As String
    Dim oPerOperacion As New comdpersona.DCOMPersonas
    
    
    Set oConst = New COMDConstSistema.NCOMConstSistema
    gsCorreoHost = oConst.LeeConstSistema(90)
    gsCorreoEnvia = oConst.LeeConstSistema(91)
    lsCorreoDestino = oPerOperacion.GetCorreoDestinoSolicitudVistoContinuidad
    
    lsContenido = "El usuario " & gsCodUser & " envió una solicitud de visto de continuidad " & _
                  "" & "<p><p>" & _
                  "<b>Nombre Usuario:</b> " & gsNomPersUser & "<br>" & _
                  "<b>Agencia:</b> " & gsNomAge & "<br>" & _
                  "<b>Cargo:</b> " & gsNomCargo & _
                  "" & "<p><p>" & _
                  "<b>Nombre Cliente:</b> " & oPersona.NombreCompleto & "<br>" & _
                  "<b>Condición:</b> " & cCondicion & "<br>" & _
                  "<b>Relación:</b> " & Trim(Left(cboRelacion.Text, Len(cboRelacion.Text) - 2)) & "<br>" & _
                  "<b>Operación:</b> " & oPerOperacion.GetOperacionSolicitudVistoContinuidad(cOpeCod) & "<br>"

    EnviarMail gsCorreoHost, gsCorreoEnvia, lsCorreoDestino, "SOLICITUD DE VISTO DE CONTINUIDAD DE PROCESO DE CREDITO", lsContenido
End Sub
