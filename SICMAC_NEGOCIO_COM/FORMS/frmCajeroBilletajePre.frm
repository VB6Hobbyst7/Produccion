VERSION 5.00
Begin VB.Form frmCajeroBilletajePre 
   Caption         =   "Opciones de Efectivo"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.CommandButton Salir 
         Caption         =   "Salir"
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
         Left            =   1560
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdPreCuadre 
         Caption         =   "Aceptar"
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
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdregistro 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Aceptar"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "PreCuadre Operaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Registro de Efectivo "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmCajeroBilletajePre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPreCuadre_Click()
    frmCajeroCorte.Show 1
    Unload Me
End Sub

Private Sub cmdregistro_Click()
    If Not ValidaDevTarjetas Then
        Exit Sub
    End If
    
    frmCajaGenEfectivo.RegistroEfectivo True, gOpeHabCajRegEfect
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub Salir_Click()
    Unload Me
End Sub
