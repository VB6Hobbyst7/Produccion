VERSION 5.00
Begin VB.Form frmDesbloqueoClienteOrdPago 
   Caption         =   "Desbloqueo de Cliente"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   Icon            =   "frmDesbloqueoClienteOrdPago.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSelCliente 
      Caption         =   "Seleccionar Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Width           =   5295
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   1750
         Width           =   855
      End
      Begin VB.CommandButton cmdDesbloquear 
         Caption         =   "Desbloquear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   1750
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   ". . ."
         Height          =   375
         Left            =   3120
         TabIndex        =   9
         Top             =   300
         Width           =   615
      End
      Begin VB.TextBox txtEstCliente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtDocCliente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtNomCliente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   840
         Width           =   3975
      End
      Begin VB.TextBox txtCodCliente 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblExtCliente 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDocCliente 
         Caption         =   "Documento:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblNomCliente 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblCodCliente 
         Caption         =   "Codigo:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmDesbloqueoClienteOrdPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************************
'*** Nombre : frmDesbloqueoClienteOrdPago
'*** Descripción : Formulario para desbloquear cliente por sobre giro de orden de pago.
'*** Creación : MIOL el 20120928, según OYP-RFC089-2012 Objetivo D
'***************************************************************************************
Option Explicit
Dim oPersEstSobreGiro As COMDPersona.DCOMPersona
Dim oPersDesbloquear As COMNPersona.NCOMPersona
Dim rsEstadoCliente As Recordset
Dim pcPersCod As String

Private Sub cmdBuscar_Click()
Dim oPersona As COMDPersona.UCOMPersona
Dim nEstCliente As Integer
    Set oPersona = frmBuscaPersona.Inicio
    If Not oPersona Is Nothing Then
        Me.txtCodCliente.Text = oPersona.sPersCod
        Me.txtNomCliente.Text = oPersona.sPersNombre
        Me.txtDocCliente.Text = Trim(oPersona.sPersIdnroDNI)
        
        Set oPersEstSobreGiro = New COMDPersona.DCOMPersona
        Set rsEstadoCliente = New ADODB.Recordset
        
        Set rsEstadoCliente = oPersEstSobreGiro.RecuperaEstadoClienteSobreGiro(oPersona.sPersCod)
        If rsEstadoCliente.RecordCount > 0 Then
            If rsEstadoCliente!nEstado = 1 Then
                Me.txtEstCliente.Text = "Bloqueado"
                Me.cmdDesbloquear.Enabled = True
            End If
            If rsEstadoCliente!nEstado = 0 Then
                Me.txtEstCliente.Text = "Desbloqueado"
                Me.cmdDesbloquear.Enabled = False
            End If
        Else
            Me.txtEstCliente.Text = "Desbloqueado"
            Me.cmdDesbloquear.Enabled = False
        End If
    Else
        Exit Sub
    End If
    pcPersCod = oPersona.sPersCod
    Set oPersona = Nothing
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdDesbloquear_Click()
Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
Set oPersDesbloquear = New COMNPersona.NCOMPersona
Dim lcMovNro As String
Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    lcMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    Call oPersDesbloquear.updClienteBloqueadoxSobreGiro(pcPersCod, lcMovNro)
    MsgBox "La actualización se realizo de forma correcta"
    Me.cmdDesbloquear.Enabled = False
End Sub

Private Sub Form_Load()
Me.cmdDesbloquear.Enabled = False
End Sub
