VERSION 5.00
Begin VB.Form frmSolicitudTasasEspeciales 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SOLICITUD DE APROBACION DE TASAS DE INTERES ESPECIALES"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmSolicitudTasasEspeciales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTasa 
      Height          =   330
      Left            =   6465
      TabIndex        =   10
      Top             =   2010
      Width           =   900
   End
   Begin VB.ComboBox cboProducto 
      Height          =   315
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2040
      Width           =   3195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   435
      Left            =   6360
      TabIndex        =   6
      Top             =   2775
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Grabar"
      Height          =   435
      Left            =   4995
      TabIndex        =   5
      Top             =   2790
      Width           =   1050
   End
   Begin VB.CommandButton cmdBuscar 
      Height          =   480
      Left            =   7470
      Picture         =   "frmSolicitudTasasEspeciales.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   300
      Width           =   480
   End
   Begin VB.Label lbldi 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1290
      TabIndex        =   12
      Top             =   945
      Width           =   4665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "DNI / RUC:"
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
      Left            =   195
      TabIndex        =   11
      Top             =   1035
      Width           =   1005
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tasa Solicitada:"
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
      Left            =   5025
      TabIndex        =   9
      Top             =   2130
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Producto:"
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
      Left            =   195
      TabIndex        =   7
      Top             =   2130
      Width           =   840
   End
   Begin VB.Label lblDireccion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1290
      TabIndex        =   4
      Top             =   1545
      Width           =   6105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
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
      Left            =   180
      TabIndex        =   3
      Top             =   1635
      Width           =   885
   End
   Begin VB.Label lblCliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1290
      TabIndex        =   2
      Top             =   405
      Width           =   6105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
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
      Left            =   195
      TabIndex        =   0
      Top             =   480
      Width           =   660
   End
End
Attribute VB_Name = "frmSolicitudTasasEspeciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As DColPContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As UProdPersona

On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    
    lblCliente.Caption = lsPersNombre
    lblCliente.Tag = lsPersCod
    lblDireccion.Caption = loPers.sPersDireccDomicilio
    lbldi.Caption=iif(lopers.sPersPersoneria=1,
Set loPers = Nothing

' Selecciona Estados




Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub Form_Load()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Label1_Click()

End Sub
