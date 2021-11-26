VERSION 5.00
Begin VB.Form frmCapMotivoAutorizacion 
   Caption         =   "Motivo de Autorización"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   Icon            =   "frmCapMotivoAutorizacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.EditMoney txtMonto 
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtMotivo 
      Height          =   1695
      Left            =   120
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label lblMoneda 
      Caption         =   "[moneda]"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblMonto 
      Caption         =   "Monto"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Glosa:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCapMotivoAutorizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'***Nombre      : frmCapMotivoAutorizacion ----SUBIDO DESDE LA 60
'***Descripción : Formulario para registrar el motivo de autorización de operaciones sin tarjeta
'***Creación    : MARG el 20171201, según TI-ERS 065-2017
'************************************************************************************************
Option Explicit
Dim nExito As Integer
Dim cUserSolicitud As String, cCodAgeSolicitud As String, cMotivoAutorizacion As String, cCtaCodCliente As String, cPersCodCliente As String, cOpecod As String

Private Sub cmdCancelar_Click()
    nExito = 0
    Unload Me
End Sub

Private Sub cmdEnviar_Click()
     Me.cmdEnviar.Enabled = False
    
    cMotivoAutorizacion = Me.txtMotivo.Text
    'GIPO 20180802 MEMO 1809-2018-GM-DI/CMACM
    Dim nMonto As Currency
    nMonto = txtMonto.value
    
    If ValidarDatos = True Then
        Dim oSolicitud As New COMDCaptaGenerales.DCOMCaptaGenerales
        nExito = oSolicitud.InsertarCapAutSinTarjetaVisto(cUserSolicitud, cCodAgeSolicitud, cMotivoAutorizacion, cCtaCodCliente, cPersCodCliente, cOpecod, nMonto)
        If nExito = 0 Then
            MsgBox "La Solicitud no fue satisfactoria, inténtalo nuevamente", vbExclamation, "Aviso"
            Me.cmdEnviar.Enabled = True
        End If
        If nExito > 0 Then
            'MsgBox "La solicitud de atención sin tarjeta fue enviada. " & Chr(13) & vbNewLine & "Comuníquese con el Coordinador o Jefe de Operaciones para la aprobación o rechazo de la misma", vbInformation, "Aviso"
            Unload Me
        End If
    End If
End Sub

Public Function Inicio(pcUserSolicitud As String, pcCodAgeSolicitud As String, pcCtaCodCliente As String, _
        pcPersCodCliente As String, pcOpeCod) As Integer
    cUserSolicitud = pcUserSolicitud
    cCodAgeSolicitud = pcCodAgeSolicitud
    cCtaCodCliente = pcCtaCodCliente
    cPersCodCliente = pcPersCodCliente
    cOpecod = pcOpeCod
    
    'GIPO 20180802 MEMO 1809-2018-GM-DI/CMACM
    If pcOpeCod = gAhoRetEfec Or pcOpeCod = gCTSRetEfec Then
        txtMonto.Visible = True
        lblMonto.Visible = True
        lblMoneda.Visible = True
        
        Dim digitoMoneda As String
        digitoMoneda = Mid(pcCtaCodCliente, 9, 1)
        lblMoneda.Caption = IIf(digitoMoneda = "1", gcPEN_PLURAL, "DÓLARES")
        
    Else
        txtMonto.Visible = False
        lblMonto.Visible = False
        lblMoneda.Visible = False
    End If
    'END GIPO
    
    Me.Show 1
    Inicio = nExito
End Function

Private Function ValidarDatos() As Boolean
    Dim Validado As Boolean
    Dim camposACompletar As String
    camposACompletar = ""
    Validado = True
    
    
    If Me.txtMotivo.Text = "" Then
       camposACompletar = "- Falta ingresar el Motivo." & vbCrLf
       Validado = False
    End If
    
    If cOpecod = gAhoRetEfec Or cOpecod = gCTSRetEfec Then
        If txtMonto.value <= 0 Then
            camposACompletar = camposACompletar & "- El Monto debe ser superior a cero"
            Validado = False
        End If
    End If
    
    If Not Validado Then
        MsgBox "Hay observaciones." & vbCrLf & camposACompletar, vbInformation, "Aviso"
        Validado = False
        'Me.txtMotivo.SetFocus
        Me.cmdEnviar.Enabled = True
    End If
        
    ValidarDatos = Validado
End Function


Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdEnviar.SetFocus
    End If
End Sub
