VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPigEvaluacionMensualClientes 
   Caption         =   "Evaluación Mensual Pignoraticia"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmPigEvaluacionMensualClientes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   4095
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese la fecha del Proceso:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
End
Attribute VB_Name = "frmPigEvaluacionMensualClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdProcesar_Click()
Dim loRegPig As DPigActualizaBD
Dim mbTrans As Boolean
Dim x As Integer

On Error GoTo ErrorRegPig

Set loRegPig = New DPigActualizaBD

loRegPig.dBeginTrans
mbTrans = True

   
Call loRegPig.dEvalPigno(txtFecha)
  
loRegPig.dCommitTrans
Set loRegPig = Nothing
mbTrans = False

MsgBox "El Proceso de Evaluación Pignoraticia ha terminado.", vbInformation, "Aviso"
    
Exit Sub
    
ErrorRegPig:
    If mbTrans Then
        loRegPig.dRollbackTrans
        mbTrans = False
    End If
    Err.Raise vbObjectError + 100, "Error Evaluación Pignoraticia", "Error en Evaluación Pignoraticia"

End Sub

Private Sub Form_Load()
txtFecha = Format$(gdFecSis, "dd/mm/yyyy")
End Sub
