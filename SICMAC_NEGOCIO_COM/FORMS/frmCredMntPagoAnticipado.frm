VERSION 5.00
Begin VB.Form frmCredMntPagoAnticipado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reestructuración de Calendario"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "frmCredMntPagoAnticipado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1080
      TabIndex        =   2
      Top             =   2040
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   3
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4335
      Begin VB.OptionButton OptTipoCuota 
         Caption         =   "Reducción del monto de las cuotas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   1080
         Value           =   -1  'True
         Width           =   2910
      End
      Begin VB.OptionButton OptTipoCuota 
         Caption         =   "Reducción del número de cuotas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   $"frmCredMntPagoAnticipado.frx":030A
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmCredMntPagoAnticipado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredMntPagoAnticipado
'** Descripción : Formulario para elegir forma de reestructuración del nuevo calendario en
'**               base a un pago anticipado creado segun TI-ERS008-2015
'** Creación : JUEZ, 20150415 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim sCtaCod As String
Dim nCuotasPend As Integer
Dim bRegistra As Boolean

Public Function Registrar(ByVal psCtaCod As String, ByVal pnCuotasPend As Integer) As Boolean
sCtaCod = psCtaCod
nCuotasPend = pnCuotasPend
'JUEZ 20150626 ***************
If nCuotasPend = 1 Then
    OptTipoCuota(1).value = 0
    OptTipoCuota(1).Enabled = False
Else
    OptTipoCuota(1).Enabled = True
End If
'END JUEZ ********************
Me.Show 1
Registrar = bRegistra
End Function

Private Sub cmdAceptar_Click()
Dim oCred As COMDCredito.DCOMCredActBD
Dim objPista As COMManejador.Pista
Dim sMovNro As String

    Set oCred = New COMDCredito.DCOMCredActBD
    Call oCred.dUpdateColocacCred(sCtaCod, , , , , , , , , , , , 1)
     
    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oCred.dInsertCredMantPrepago sCtaCod, sMovNro, 1, IIf(OptTipoCuota(0).value, 1, 0), 0, nCuotasPend
    Set oCred = Nothing
    
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gCredAdminPrepagosNormales, sMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Cambio a dinamico.", sCtaCod, gCodigoCuenta
    Set objPista = Nothing
    
    bRegistra = True
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    bRegistra = False
End Sub
