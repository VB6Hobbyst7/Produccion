VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackUp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2925
   ClientLeft      =   2880
   ClientTop       =   2265
   ClientWidth     =   5715
   Icon            =   "frmBackUp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmBackUp.frx":030A
   ScaleHeight     =   2925
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   465
      Left            =   1020
      TabIndex        =   7
      Top             =   1275
      Width           =   4155
      Begin VB.Label lblProgreso 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   105
         TabIndex        =   8
         Top             =   165
         Width           =   3900
      End
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   135
      TabIndex        =   6
      Top             =   1200
      Width           =   5370
   End
   Begin VB.Frame Frame1 
      Height          =   450
      Left            =   1005
      TabIndex        =   4
      Top             =   1725
      Width           =   4170
      Begin MSComctlLib.ProgressBar barra 
         Height          =   285
         Left            =   45
         TabIndex        =   5
         Top             =   120
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
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
      Height          =   465
      Left            =   3240
      TabIndex        =   2
      Top             =   2265
      Width           =   1665
   End
   Begin VB.CommandButton cmdBackUp 
      Caption         =   "Copia de &Seguridad"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   585
      TabIndex        =   0
      Top             =   2280
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Avance"
      Height          =   195
      Left            =   225
      TabIndex        =   9
      Top             =   1425
      Width           =   555
   End
   Begin VB.Label lblPorc 
      AutoSize        =   -1  'True
      Caption         =   "Progreso :"
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   1875
      Width           =   720
   End
   Begin VB.Image imgAlerta 
      Height          =   645
      Left            =   225
      Picture         =   "frmBackUp.frx":0614
      Top             =   225
      Width           =   630
   End
   Begin VB.Label lblMsg 
      Caption         =   $"frmBackUp.frx":1BD6
      Height          =   975
      Left            =   1140
      TabIndex        =   1
      Top             =   105
      Width           =   3990
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oSQLServer As SQLDMO.SQLServer
Dim DataBase As SQLDMO.DataBase
Private WithEvents Backup As SQLDMO.Backup
Attribute Backup.VB_VarHelpID = -1
'Dim Backup As SQLDMO.Backup
Private Function ConectaServer(lsNomServer As String, lsUID As String, lsPWID As String) As Boolean
On Error GoTo ErrorConexion
Set oSQLServer = New SQLDMO.SQLServer
ConectaServer = True
oSQLServer.ApplicationName = "SQL-DMO Explorer"
    
oSQLServer.Connect lsNomServer, lsUID, lsPWID
oSQLServer.Application.GroupRegistrationServer = ""
oSQLServer.Configuration.ShowAdvancedOptions = True

Exit Function
ErrorConexion:
    ConectaServer = False
    MsgBox "Error N°[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
End Function
Private Function DesConectaServer()
oSQLServer.Disconnect
Set oSQLServer = Nothing
End Function
Private Function CreaBackup(lsNomDevice As String, lsNombreFisico As String) As Boolean
Dim Device As New SQLDMO.BackupDevice
Dim lbNomDevice As Boolean
lbNomDevice = False

On Error GoTo ErrorBackup
CreaBackup = True

For Each Device In oSQLServer.BackupDevices
    If lsNomDevice = Trim(Device.Name) Then
        lbNomDevice = True
    End If
Next

Set Backup = New SQLDMO.Backup
If lbNomDevice = False Then
    CreaDevice lsNomDevice, lsNombreFisico, gsDBName, gsServerName
End If

Backup.Initialize = True
Backup.Action = SQLDMOBackup_Database
Backup.DataBase = gsDBName
Backup.Devices = lsNomDevice
Backup.BackupSetName = lsNomDevice
Backup.BackupSetDescription = "Backup de Contabilidad del dia " & gdFecSis & " a las " & Time
Backup.SQLBackup oSQLServer

Exit Function
ErrorBackup:
    CreaBackup = False
    MsgBox "Error N° [" & Err.Number & "] " & Err.Description, vbInformation, "Error en Generacion de Backup"
    
End Function
Private Sub CreaDevice(lsNomDevice As String, lsNomFisico As String, lsNomDatabase As String, lsNomServer As String)
Dim ConBackUp As New ADODB.Connection
Dim lsConBackUp As String
Dim cmd As New ADODB.Command
Dim prm As New ADODB.Parameter
Dim sql As String


       lsConBackUp = "PROVIDER=SQLOLEDB;UID=sa;PWD=dba;DataBase=" & lsNomDatabase & ";Server=" & lsNomServer
       ConBackUp.CommandTimeout = 100
       ConBackUp.Open lsConBackUp
       
       cmd.CommandText = "spBackUp"
       cmd.CommandType = adCmdStoredProc
       cmd.Name = "spBUp"
       Set prm = cmd.CreateParameter("logicalname", adVarChar, adParamInput, 100)
       cmd.Parameters.Append prm
       Set prm = cmd.CreateParameter("physicalname", adVarChar, adParamInput, 250)
       cmd.Parameters.Append prm
       Set prm = cmd.CreateParameter("DataBase", adVarChar, adParamInput, 100)
       cmd.Parameters.Append prm
       Set cmd.ActiveConnection = ConBackUp
       cmd.CommandTimeout = 720
       cmd.Parameters.Refresh
       ConBackUp.spBUp lsNomDevice, lsNomFisico & ".BAK", lsNomDatabase
       Set cmd = Nothing
       ConBackUp.Close
       Set ConBackUp = Nothing
       
End Sub
Private Sub Backup_PercentComplete(ByVal Message As String, ByVal Percent As Long)
    Me.lblProgreso = Percent & "% Generado"
    Me.barra.Value = Percent
    DoEvents
End Sub
Private Sub Backup_Complete(ByVal Message As String)
    MousePointer = vbArrow
    MsgBox "Proceso de Copia de Seguridad Finalizado... ", vbInformation, "Aviso"
End Sub
Private Sub CmdBackup_Click()
Dim lsNomDev As String
Dim lsBackUp As String

    cmdBackUp.Enabled = False
    MousePointer = vbHourglass
    'gsUID = "sa"
    'gsPWD = ""
    If ConectaServer(gsServerName, "sa", "dba") Then
        Me.lblProgreso = "Copia de Seguridad en Proceso...                     "
        Me.lblProgreso.Refresh
        
        lsNomDev = GetNomBacKUpCont(Trim(Format(gdFecSis & " " & GetHoraServer(), "yyyymmddhhmmss")))
        lsBackUp = gsDirBackup & "\" & lsNomDev
        If CreaBackup(lsNomDev, lsBackUp) Then
            MousePointer = vbArrow
            lblProgreso.Refresh
        Else
            MousePointer = vbArrow
            MsgBox "Hubo un error al realizar la copia de seguridad. Consulte con el Area de Sistemas", vbExclamation, "Error"
            lblProgreso = "Error al realizar la copia de seguridad. Consulte con el Area de Sistemas"
            lblProgreso.Refresh
        End If
        DesConectaServer
    Else
       MousePointer = vbArrow
       MsgBox "Error al efectuar la conexión con el servidor, el proceso se cancela", vbExclamation, "Aviso"
       Exit Sub
    End If
    Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Copia de Seguridad"
End Sub

Private Sub Form_Unload(Cancel As Integer)
CierraConexion
End Sub

Private Function GetNomBacKUpCont(pFecha As String) As String
    GetNomBacKUpCont = "BkUp" & gsDBName & Right(gsCodAge, 2) & "_" & pFecha
End Function


