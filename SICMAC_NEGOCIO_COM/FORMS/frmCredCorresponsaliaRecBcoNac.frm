VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredCorresponsaliaRecBcoNac 
   Caption         =   "Recuperaciones de Créditos por Corresponsalia"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Grabar Datos de Sugerencia"
      Top             =   2280
      Width           =   1080
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   3480
      TabIndex        =   0
      Top             =   2280
      Width           =   1080
   End
   Begin MSMask.MaskEdBox TxtFecha 
      Height          =   315
      Left            =   3240
      TabIndex        =   2
      Top             =   1080
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   $"frmCredCorresponsaliaRecBcoNac.frx":0000
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Proceso :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   1140
      Width           =   1680
   End
End
Attribute VB_Name = "frmCredCorresponsaliaRecBcoNac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdProcesar_Click()
Dim loNCredito As COMNCredito.NCOMCredito
Dim fso As Scripting.FileSystemObject
Dim lsCadena As String
Dim lsmensaje As String
Dim ts As TextStream
Dim lsFile As String
Dim lsCodCliCMACMaynas As String


Set loNCredito = New COMNCredito.NCOMCredito
    lsCadena = loNCredito.GenerarArchivoRecuperacionesPorCorresponsalia(CDate(TxtFecha.Text), lsmensaje, lsCodCliCMACMaynas)
Set loNCredito = Nothing
    If lsmensaje <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
    Else
        lsFile = App.path & "\Spooler\R" & Format(CDate(TxtFecha.Text), "yymmdd") & "01-" & Right(lsCodCliCMACMaynas, 2) & ".ING"
        Set fso = New Scripting.FileSystemObject
        If fso.FileExists(lsFile) Then
            If MsgBox("El archivo ya existe, desea reemplazarlo", vbYesNo + vbInformation, "Aviso") = vbNo Then
                Set fso = Nothing
                Exit Sub
            End If
        End If
        Set ts = fso.CreateTextFile(lsFile, True)
            ts.Write (lsCadena)
            MsgBox "El archivo se generó satisfactoriamente", vbInformation, "Aviso"
            ts.Close
        Set fso = Nothing
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    TxtFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
End Sub
