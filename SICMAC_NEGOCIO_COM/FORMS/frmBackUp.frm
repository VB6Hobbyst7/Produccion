VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBackUp 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3060
   ClientLeft      =   5160
   ClientTop       =   4335
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   465
      Left            =   1260
      TabIndex        =   6
      Top             =   1275
      Width           =   4155
      Begin VB.Label lblProgreso 
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   105
         TabIndex        =   7
         Top             =   150
         Width           =   3900
      End
   End
   Begin VB.Frame Frame2 
      Height          =   45
      Left            =   135
      TabIndex        =   5
      Top             =   1200
      Width           =   5610
   End
   Begin VB.Frame Frame1 
      Height          =   450
      Left            =   1245
      TabIndex        =   3
      Top             =   1725
      Width           =   4170
      Begin MSComctlLib.ProgressBar barra 
         Height          =   285
         Left            =   45
         TabIndex        =   4
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
      Height          =   465
      Left            =   3105
      TabIndex        =   1
      Top             =   2415
      Width           =   2100
   End
   Begin VB.CommandButton cmdBackUp 
      Caption         =   "Copia de &Seguridad"
      Default         =   -1  'True
      Height          =   465
      Left            =   870
      TabIndex        =   0
      Top             =   2415
      Width           =   2100
   End
   Begin VB.PictureBox Pic2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   120
      ScaleHeight     =   870
      ScaleWidth      =   5745
      TabIndex        =   9
      Top             =   120
      Width           =   5775
      Begin VB.Image imgAlerta 
         Height          =   480
         Left            =   120
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblBase 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   720
         TabIndex        =   11
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BACKUP DE LA BASE DE DATOS"
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
         Left            =   840
         TabIndex        =   10
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Avance"
      Height          =   195
      Left            =   465
      TabIndex        =   8
      Top             =   1425
      Width           =   555
   End
   Begin VB.Label lblPorc 
      AutoSize        =   -1  'True
      Caption         =   "Progreso :"
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   1875
      Width           =   720
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function FindeMes() As Boolean
Dim nMes As Integer
Dim nDia As Integer
    nMes = Month(gdFecSis)
    Select Case nMes
        Case 1, 3, 5, 7, 8, 10, 12
            nDia = 31
        Case 4, 6, 9, 11
            nDia = 30
        Case 2
            If (Year(gdFecSis) Mod 4) = 0 Then
                nDia = 29
            Else
                nDia = 28
            End If
    End Select
    If Day(gdFecSis) = nDia Then
        FindeMes = True
    Else
        FindeMes = False
    End If
End Function

Private Sub cmdBackUp_Click()
Dim lsNomDev As String
Dim lsBkFullDiff As String
Dim lsBackUp As String
Dim SSql As String
Dim PObjConec As COMConecta.DCOMConecta
Dim cn As ADODB.Connection

    '****** Modificado
    lblBase.Caption = gsDirBackup & "BkUp" & Trim(Format$(gdFecSis & " " & Time, "yyyymmddhhmmss")) & gsCodUser & ".bak"
    Me.lblProgreso = "Copia de Seguridad en Proceso...                     "
    Me.lblProgreso.Refresh
    cmdBackUp.Enabled = False
    MousePointer = vbHourglass
    barra.value = barra.Max / 3
    'If BackupFull Then
        'lsBkFullDiff = gsDirBackup & "BkUpFULL" & "_" & Trim(Format$(gdFecSis & " " & Time, "yyyymmddhhmmss")) & ".bak"
        lsBkFullDiff = gsDirBackup & "BkUp" & Trim(Format$(gdFecSis & " " & Time, "yyyymmddhhmmss")) & gsCodUser & ".bak"
        SSql = "BACKUP DATABASE " & gsDBName & " TO DISK = '" & lsBkFullDiff & "'"
    'Else
    '    lsBkFullDiff = gsDirBackup & "\" & "BkUpDIFF" & Right(gsCodAge, 2) & "_" & Trim(Format$(gdFecSis & " " & Time, "yyyymmddhhmmss")) & ".bak"
    '    sSql = "BACKUP DATABASE " & gsDBName & " TO DISK = '" & lsBkFullDiff & "' WITH DIFFERENTIAL "
    'End If
    'Set PObjConec = New COMConecta.DCOMConecta
    Set cn = New ADODB.Connection
    PObjConec.AbreConexion
    
    cn.ConnectionString = PObjConec.CadenaConexion
    
    cn.CommandTimeout = 10000
    cn.Open
    
    cn.Execute SSql
    barra.value = (barra.Max / 2) + (barra.Max / 3)
    cn.Close
    PObjConec.CierraConexion
    Set cn = Nothing
    'Set PObjConec = Nothing
    barra.value = barra.Max
    Me.lblProgreso = "Copia de Seguridad Finalizada."
    MousePointer = 0
    MsgBox "Proceso de Copia de Seguridad Finalizado ... ", vbInformation, "Aviso"
    lblProgreso.Refresh
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraSdi Me
Me.Caption = "Copia de Seguridad"
lblBase.Caption = gsDirBackup & "BkUp" & Trim(Format$(gdFecSis & " " & Time, "yyyymmddhhmmss")) & gsCodUser & ".bak"
End Sub
