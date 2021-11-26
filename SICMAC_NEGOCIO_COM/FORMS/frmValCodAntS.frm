VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmValCodAntS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EQUIVALENCIAS DE CODIGOS"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmValCodAntS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2070
      TabIndex        =   4
      Top             =   1395
      Width           =   1035
   End
   Begin VB.Frame fraDato 
      Caption         =   "Dato"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1155
      Left            =   165
      TabIndex        =   1
      Top             =   60
      Width           =   5190
      Begin VB.OptionButton Option1 
         Caption         =   "Cod. Nuevo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   225
         TabIndex        =   6
         Top             =   660
         Width           =   1605
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cod. Antiguo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   1605
      End
      Begin MSMask.MaskEdBox txtCuenta 
         Height          =   375
         Left            =   2115
         TabIndex        =   2
         Top             =   240
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   23
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-###-##-#-#######-##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtTarjeta 
         Height          =   345
         Left            =   2115
         TabIndex        =   3
         Top             =   675
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   609
         _Version        =   393216
         MaxLength       =   23
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-##-###-#-########-#"
         PromptChar      =   "_"
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3210
      TabIndex        =   0
      Top             =   1395
      Width           =   1035
   End
End
Attribute VB_Name = "frmValCodAntS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
