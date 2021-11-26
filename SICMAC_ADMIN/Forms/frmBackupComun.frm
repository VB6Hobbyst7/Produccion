VERSION 5.00
Begin VB.Form frmBackupComun 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Bases Comunes"
   ClientHeight    =   3165
   ClientLeft      =   3255
   ClientTop       =   2130
   ClientWidth     =   5760
   Icon            =   "frmBackupComun.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   150
      TabIndex        =   4
      Top             =   870
      Width           =   5535
      Begin VB.Label Label2 
         Caption         =   "DBPersona"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   7
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "DBComunes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2220
         TabIndex        =   6
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label4 
         Caption         =   "DBImagenes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3660
         TabIndex        =   5
         Top             =   180
         Width           =   1155
      End
      Begin VB.Image ImgPersona 
         Height          =   480
         Left            =   870
         Picture         =   "frmBackupComun.frx":030A
         Top             =   480
         Width           =   480
      End
      Begin VB.Image ImgComunes 
         Height          =   480
         Left            =   2520
         Picture         =   "frmBackupComun.frx":074C
         Top             =   480
         Width           =   480
      End
      Begin VB.Image ImgImagen 
         Height          =   480
         Left            =   3930
         Picture         =   "frmBackupComun.frx":0B8E
         Top             =   480
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1245
      Left            =   120
      TabIndex        =   2
      Top             =   1860
      Width           =   5565
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   435
         Left            =   2880
         TabIndex        =   3
         Top             =   450
         Width           =   2265
      End
      Begin VB.CommandButton CmdBackup 
         Caption         =   "&Realizar Backup "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   330
         TabIndex        =   0
         Top             =   450
         Width           =   2265
      End
   End
   Begin VB.Image ImgRojo 
      Height          =   480
      Left            =   5610
      Picture         =   "frmBackupComun.frx":0FD0
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgVerde 
      Height          =   480
      Left            =   5640
      Picture         =   "frmBackupComun.frx":1412
      Top             =   1230
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgAmarillo 
      Height          =   480
      Left            =   5610
      Picture         =   "frmBackupComun.frx":1854
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgAlerta 
      Height          =   645
      Left            =   360
      Picture         =   "frmBackupComun.frx":1C96
      Top             =   0
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Este proceso realiza el Backup de las Bases da Datos Personas, Comunes, y Firmas. El Proceso puede tardar unos 10 minutos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1410
      TabIndex        =   1
      Top             =   120
      Width           =   3945
   End
End
Attribute VB_Name = "frmBackupComun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sNomDevice As String
Dim sRuta As String

Private Function ObtenerRuta() As String
Dim sSql As String
Dim R As ADODB.Recordset

   sSql = "Select * from Varsistema Where cCodProd = 'ADM' And cNomVar = 'cBakupCom'"
   Set R = New ADODB.Recordset
   R.Open sSql, dbCmact, adOpenStatic, adLockReadOnly, adCmdText
   If Not R.BOF And Not R.EOF Then
      ObtenerRuta = Trim(R!cDescVar) & "\" & sNomDevice & ".Bak"
   Else
      ObtenerRuta = "c:\" & sNomDevice & ".Bak"
   End If
   R.Close
   Set R = Nothing
End Function
Private Function ObtenerDevice(ByVal sBase As String) As String
   ObtenerDevice = "BkUp" & sBase & Right(gsCodAge, 2) & "_" & Format(gdFecSis + Time, "yyyymmddhhmmss")
End Function
Private Sub CmdBackup_Click()
Dim sSql As String
Screen.MousePointer = 11
ImgPersona.Picture = ImgRojo.Picture
ImgComunes.Picture = ImgRojo.Picture
ImgImagen.Picture = ImgRojo.Picture

'Realizando el Backup de Persona
ImgPersona.Picture = ImgAmarillo.Picture
'dbCmact.CommandTimeout = 7200
'sNomDevice = ObtenerDevice("Persona")
'sRuta = ObtenerRuta
'sSql = "EXEC sp_addumpdevice 'disk','" & sNomDevice & "','" & sRuta & "'"
'dbCmact.Execute sSql
''Backup full database Persona.
'sSql = "BACKUP DATABASE DBPersona TO " & sNomDevice
'dbCmact.Execute sSql
ImgPersona.Picture = ImgVerde.Picture
Screen.MousePointer = 0
'MsgBox "Backup de la Base DBPersona Realizado con Exito", vbInformation, "Aviso"
DoEvents
'Realizando el Backup de Comunes
Screen.MousePointer = 11
ImgComunes.Picture = ImgAmarillo.Picture
sNomDevice = ObtenerDevice("Comunes")
sRuta = ObtenerRuta
sSql = "EXEC sp_addumpdevice 'disk','" & sNomDevice & "','" & sRuta & "'"
dbCmact.Execute sSql
'Backup full database Persona.
sSql = "BACKUP DATABASE DBComunes TO " & sNomDevice
dbCmact.CommandTimeout = 1500
dbCmact.Execute sSql
ImgComunes.Picture = ImgVerde.Picture
Screen.MousePointer = 0
MsgBox "Backup de la Base DBComunes Realizado con Exito", vbInformation, "Aviso"

'Realizando el Backup de Imagenes
Screen.MousePointer = 11
ImgImagen.Picture = ImgAmarillo.Picture
'sNomDevice = ObtenerDevice("Imagenes")
'sRuta = ObtenerRuta
'sSql = "EXEC sp_addumpdevice 'disk','" & sNomDevice & "','" & sRuta & "'"
'dbCmact.Execute sSql
''Backup full database Persona.
'sSql = "BACKUP DATABASE DBImagenes TO " & sNomDevice
'dbCmact.Execute sSql
'dbCmact.CommandTimeout = 30
ImgImagen.Picture = ImgVerde.Picture
Screen.MousePointer = 0
'MsgBox "Bakcup de la Base DBImagenes Realizado con Exito", vbInformation, "Aviso"
MsgBox "Proceso de Backups Finalizado", vbInformation, "Aviso"
Unload Me
End Sub

Private Sub CmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   AbreConexion
End Sub

Private Sub Form_Unload(Cancel As Integer)
   CierraConexion
End Sub

