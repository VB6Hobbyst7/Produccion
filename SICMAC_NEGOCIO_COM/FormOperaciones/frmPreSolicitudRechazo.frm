VERSION 5.00
Begin VB.Form frmPreSolicitudRechazo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rechazar Pre Solicitud de Credito"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "frmPreSolicitudRechazo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   28
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
      Height          =   375
      Left            =   6960
      TabIndex        =   27
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Rechazo de Pre Solicitud"
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Width           =   9735
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   840
         Width           =   7935
      End
      Begin VB.ComboBox cmbMotivo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmPreSolicitudRechazo.frx":030A
         Left            =   4080
         List            =   "frmPreSolicitudRechazo.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   5505
      End
      Begin VB.TextBox txtAnalista 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Comentario:"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Motivo:"
         Height          =   375
         Left            =   3360
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Analista:"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox txtMoneda 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      TabIndex        =   19
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtMontoSolicitado 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      TabIndex        =   17
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtDestino 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   3360
      Width           =   4095
   End
   Begin VB.TextBox txtSubProducto 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2880
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Pre Solicitud"
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   9735
      Begin VB.TextBox txtProducto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "Moneda:"
         Height          =   375
         Left            =   6600
         TabIndex        =   18
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Monto Solicitado:"
         Height          =   375
         Left            =   6000
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Destino:"
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Producto:"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Sub Producto:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   8
      Top             =   1440
      Width           =   6375
   End
   Begin VB.TextBox txtDOI 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   9735
      Begin VB.Label Label4 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "DOI:"
         Height          =   375
         Left            =   5160
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo:"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCodigoPresolicitud 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Código Presolicitud:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmPreSolicitudRechazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nPresolicitudId As Integer
Dim nExito As Integer
Dim RPresolicitud As ADODB.Recordset

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdRechazar_Click()
Me.cmdRechazar.Enabled = False
    If ValidarDatos = True Then
        Dim oHojaRuta As COMDCredito.DCOMhojaRuta
        Set oHojaRuta = New COMDCredito.DCOMhojaRuta
        
        nExito = oHojaRuta.InsertarPresolicitudRechazo(nPresolicitudId, CInt(Trim(Right(Me.cmbMotivo.Text, 5))), Me.txtComentario.Text, gsCodUser)
        If nExito = 0 Then
            MsgBox "La acción no fue satisfactorio, inténtalo nuevamente", vbExclamation, "Aviso"
            Me.cmdRechazar.Enabled = True
        End If
        If nExito > 0 Then
            MsgBox "La accion se realizó con éxito! ", vbInformation, "Aviso"
            Unload Me
        End If
    End If
End Sub

Public Function Inicio(ByVal pnPresolicitudId As Integer) As Integer
    nPresolicitudId = pnPresolicitudId
    Call cargarDatos
    Me.Show 1
    Inicio = nPresolicitudId
End Function
Private Sub cargarDatos()
    Dim oHojaRuta As COMDCredito.DCOMhojaRuta
    Set oHojaRuta = New COMDCredito.DCOMhojaRuta
    On Error GoTo ERRORCargaGrid
    Set RPresolicitud = oHojaRuta.ObtenerPreSolicitudesXid(nPresolicitudId)
    
    Me.txtCodigoPresolicitud.Text = RPresolicitud!cCodPresolicitud
    Me.txtCodigo.Text = RPresolicitud!cPersCod
    Me.txtDOI.Text = RPresolicitud!cPersIDnro
    Me.txtNombre.Text = RPresolicitud!cPersNombre
    Me.txtProducto.Text = RPresolicitud!cProducto
    Me.txtSubProducto.Text = RPresolicitud!cSubProducto
    Me.txtDestino.Text = RPresolicitud!cDestino
    Me.txtMontoSolicitado.Text = RPresolicitud!nMonto
    Me.txtMoneda.Text = RPresolicitud!cMoneda
    Me.txtAnalista.Text = RPresolicitud!cUserAnalista
    Call CargarCombo
    Set oHojaRuta = Nothing
    Exit Sub
ERRORCargaGrid:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
Private Sub CargarCombo()
Dim CollCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

On Error GoTo ErrorCargaControlRelaCred
    
    Set CollCons = New COMDConstantes.DCOMConstantes
    Set R = CollCons.RecuperaConstantes(10700)
    Do While Not R.EOF
        cmbMotivo.AddItem Trim(R!cConsDescripcion) & Space(80) & R!nConsValor
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
    If Me.cmbMotivo.Text = "" Then
        MsgBox "Debe seleccionar el Motivo del rechazo", vbInformation, "Aviso"
        Validado = False
        Me.cmbMotivo.SetFocus
        Me.cmdRechazar.Enabled = True
    ElseIf Trim(Me.txtComentario.Text) = "" Then
        MsgBox "Debe ingresar Comentario", vbInformation, "Aviso"
        Validado = False
        Me.txtComentario.Text = ""
        Me.txtComentario.SetFocus
        Me.cmdRechazar.Enabled = True
    End If
    ValidarDatos = Validado
End Function
