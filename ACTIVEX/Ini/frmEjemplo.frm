VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Generador de Archivo Ini"
   ClientHeight    =   2775
   ClientLeft      =   1725
   ClientTop       =   2745
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   6300
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "&Mostrar"
      Height          =   405
      Left            =   540
      TabIndex        =   3
      Top             =   2325
      Width           =   1485
   End
   Begin VB.ListBox lstVariables 
      Height          =   1425
      Left            =   375
      TabIndex        =   2
      Top             =   825
      Width           =   4980
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "Crear Ini"
      Height          =   450
      Left            =   345
      TabIndex        =   1
      Top             =   150
      Width           =   1695
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar Cadena de Conexion"
      Height          =   450
      Left            =   2175
      TabIndex        =   0
      Top             =   150
      Width           =   2985
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lsCadena As String
Dim oIni As clsIni.ClasIni
Private Sub cmdCrear_Click()
'oIni.ArchivoIni = App.Path & "\Sicmact.ini"
oIni.CrearArchivoIni
End Sub

Private Sub cmdGenerar_Click()
Dim lsCon As String
Dim lsServerCom As String
lsCon = oIni.CadenaConexion
MsgBox lsCon & vbCrLf
End Sub

Private Sub cmdMostrar_Click()
lstVariables.Clear
lstVariables.AddItem "Base Personas :" & oIni.BasePersonas
lstVariables.AddItem "Base Imagenes :" & oIni.BaseImagenes
lstVariables.AddItem "Base Comunes  :" & oIni.BaseComunes
lstVariables.AddItem "Base Negocio  :" & oIni.BaseNegocio
lstVariables.AddItem "Base Administrativa :" & oIni.BaseAdministracion

End Sub

Private Sub Form_Load()
Set oIni = New clsIni.ClasIni
End Sub
