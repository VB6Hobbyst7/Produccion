VERSION 5.00
Begin VB.Form frmInicioDia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inicio de Dia"
   ClientHeight    =   3735
   ClientLeft      =   2925
   ClientTop       =   2685
   ClientWidth     =   6825
   ControlBox      =   0   'False
   Icon            =   "frmInicioDia.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3270
      TabIndex        =   0
      Top             =   3120
      Width           =   1680
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00A56E3A&
      Height          =   2835
      Left            =   120
      ScaleHeight     =   2775
      ScaleWidth      =   6495
      TabIndex        =   2
      Top             =   90
      Width           =   6555
      Begin VB.Label Label4 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Para empezar a efectuar operaciones deberá ingresar el tipo de cambio del día"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   600
         TabIndex        =   6
         Top             =   2040
         Width           =   5760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SE VA A REALIZAR EL INICIO DE DIA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   960
         TabIndex        =   5
         Top             =   120
         Width           =   4875
      End
      Begin VB.Image imgAlerta 
         Height          =   480
         Left            =   240
         Picture         =   "frmInicioDia.frx":030A
         Top             =   0
         Width           =   480
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Recuerde que el Inicio de Dia es para toda las Agencias Informe Cuando el Proceso haya Finalizado."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   600
         TabIndex        =   4
         Top             =   1320
         Width           =   5760
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "No Realize Ninguna Operacion mientras el proceso de Inicio de Dia No haya Culminado."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   510
         Left            =   600
         TabIndex        =   3
         Top             =   600
         Width           =   5760
      End
   End
   Begin VB.CommandButton CmdInicioDia 
      Caption         =   "Inicio de Dia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1545
      TabIndex        =   1
      Top             =   3120
      Width           =   1680
   End
End
Attribute VB_Name = "frmInicioDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub InicioDia()
Dim lsMov As String
Dim oGen As COMDConstSistema.DCOMGeneral
Dim oMov As COMDMov.DCOMMov

Set oMov = New COMDMov.DCOMMov
    lsMov = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
Set oMov = Nothing


Set oGen = New COMDConstSistema.DCOMGeneral
    Call oGen.InicioDia(gsCodUser, gsCodAge, gdFecSis, lsMov)
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
Set oGen = Nothing


'Dim cmd As New ADODB.Command
'Dim prm As New ADODB.Parameter
'Dim sSql As String
'Dim PObjConec As COMConecta.DCOMConecta
'
'    Set PObjConec = New COMConecta.DCOMConecta
'    PObjConec.AbreConexion
'    PObjConec.ConexionActiva.BeginTrans
'    gdFecSis = gdFecSis + 1
'
'    cmd.CommandText = "CapValorizaCheque"
'    cmd.CommandType = adCmdStoredProc
'    cmd.Name = "CapValorizaCheque"
'    Set prm = cmd.CreateParameter("FechaTran", adDate, adParamInput)
'    cmd.Parameters.Append prm
'    Set prm = cmd.CreateParameter("Usuario", adChar, adParamInput, 4)
'    cmd.Parameters.Append prm
'    Set prm = cmd.CreateParameter("Agencia", adChar, adParamInput, 2)
'    cmd.Parameters.Append prm
'    Set cmd.ActiveConnection = PObjConec.ConexionActiva
'    cmd.CommandTimeout = 720
'    cmd.Parameters.Refresh
'    PObjConec.ConexionActiva.CapValorizaCheque Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss"), gsCodUser, Right(gsCodAge, 2)
'
'    Set cmd = Nothing
'    'ejecutamos el calculo de cobro de notificaciones
'    cmd.CommandText = "sp_ComisionNotificaciones"
'    cmd.CommandType = adCmdStoredProc
'    cmd.Name = "ComisionNotificaciones"
'    Set prm = cmd.CreateParameter("dFecCierre", adDate, adParamInput)
'    cmd.Parameters.Append prm
'    Set cmd.ActiveConnection = PObjConec.ConexionActiva
'    cmd.CommandTimeout = 720
'    cmd.Parameters.Refresh
'    PObjConec.ConexionActiva.ComisionNotificaciones Format(gdFecSis & " " & GetHoraServer(), "dd/mm/yyyy hh:mm:ss")
'
'    'Actualiza Fecha de Inicio de Dia
'    sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 15"
'    PObjConec.ConexionActiva.Execute sSql
'    'Actualiza Fecha de Inicio del Sistema
'    sSql = " UPDATE ConstSistema SET nConsSisValor = '" & Format(gdFecSis, "dd/mm/yyyy") & "' WHERE nConsSisCod = 16"
'    PObjConec.ConexionActiva.Execute sSql
'
'    'Actualiza Fecha de Inicio del Sistema
'    sSql = " Delete MovDiario "
'    PObjConec.ConexionActiva.Execute sSql
'
'    sSql = "UPDATE ConstSistema set nConsSisValor =0 WHERE nConsSisCod = 2"
'    PObjConec.ConexionActiva.Execute sSql
'
'    Me.Icon = LoadPicture(App.path & gsRutaIcono)
'    PObjConec.ConexionActiva.CommitTrans
'    PObjConec.CierraConexion
'    'Set PObjConec = Nothing
'    Set cmd = Nothing
'    Set prm = Nothing

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Public Sub CmdInicioDia_Click()
    If MsgBox("Se va a Realizar el Inicio de Dia, Desea Continuar?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    Screen.MousePointer = 11
    Call InicioDia
    Screen.MousePointer = 0
    MsgBox "Proceso de Inicio de Dia a Finalizado", vbInformation, "Aviso"
    Unload Me
End Sub

