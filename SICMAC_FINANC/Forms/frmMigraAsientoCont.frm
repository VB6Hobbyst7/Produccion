VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmContabMigraAsientoCont 
   Caption         =   "Migración Asientos de Agencias SIAFC"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6060
   Icon            =   "frmMigraAsientoCont.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4920
      TabIndex        =   3
      Top             =   1200
      Width           =   1005
   End
   Begin VB.CommandButton cmdAceptar 
      Cancel          =   -1  'True
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3840
      TabIndex        =   2
      Top             =   1200
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   4935
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Archivo"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label LblMsg 
      AutoSize        =   -1  'True
      Caption         =   "Migrando Datos ...Por Favor Espere un Momento"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   3465
   End
End
Attribute VB_Name = "frmContabMigraAsientoCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsmov As DMov
Dim lsDts As String
Dim lsDtsa As String
Dim lsDtsb As String
Dim lsDtsc As String
Dim s As Variant
Dim x As Variant
Dim y As Variant
Dim z As Variant



Private Sub cmdAceptar_Click()

MousePointer = 11
Me.Enabled = False
LblMsg.Visible = True

'Inserta en la tabla Financ_Asientos
 clsmov.InsertaAsientoRuta (Text1.Text)

'LLAMAR EL DTS migracion local
'dtsrun /Sserver_name /Uuser_nName /Ppassword /Npackage_name /Mpackage_password
'lsDts = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NDTS_IMPORTAFINAC /M"
's = Shell(lsDts, vbMaximizedFocus)
'Migra Asientos
'lsDtsa = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_Asientos_18022004 /M"
'x = Shell(lsDtsa, vbMaximizedFocus)
'Migra Cuentas
'lsDtsb = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_Cuentas_18022004 /M"
'y = Shell(lsDtsb, vbMaximizedFocus)
'Migra Saldos
'lsDtsc = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_saldos_18022004 /M"
'z = Shell(lsDtsc, vbMaximizedFocus)

'Llamar DTS Migracion Red

'dtsrun /Sserver_name /Uuser_nName /Ppassword /Npackage_name /Mpackage_password
lsDts = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NDTS_IMPORTAFINAC /M"
s = Shell(lsDts, vbMaximizedFocus)
'Migra Asientos
lsDtsa = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_Asientos_28022004 /M"
x = Shell(lsDtsa, vbMaximizedFocus)
'Migra Cuentas
lsDtsb = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_Cuentas_28022004 /M"
y = Shell(lsDtsb, vbMaximizedFocus)
'Migra Saldos
lsDtsc = "dtsrun /S01SRVSICMAC01 /Usa /Pcmacica /NFINANC_Migra_saldos_28022004 /M"
z = Shell(lsDtsc, vbMaximizedFocus)




LblMsg.Visible = False
MousePointer = 0
Me.Enabled = True

MsgBox "Proceso Terminado...!", vbInformation, "Aviso"

End Sub

Private Sub cmdBuscar_Click()
CommonDialog1.DialogTitle = "Seleccione la Ruta "
CommonDialog1.ShowOpen
sRuta = CommonDialog1.FileName
Text1.Text = sRuta
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set clsmov = New DMov
CentraForm Me
End Sub
