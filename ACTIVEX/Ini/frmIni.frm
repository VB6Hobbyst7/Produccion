VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmIni 
   Caption         =   "Generador de Archivo Ini"
   ClientHeight    =   5385
   ClientLeft      =   1920
   ClientTop       =   1890
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmdGuardar 
      Left            =   1485
      Top             =   4830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Conexion"
      Height          =   2145
      Left            =   180
      TabIndex        =   21
      Top             =   2580
      Width           =   5460
      Begin VB.TextBox txtpersonas 
         Height          =   315
         Left            =   1050
         TabIndex        =   7
         Top             =   270
         Width           =   3315
      End
      Begin VB.TextBox txtComunes 
         Height          =   315
         Left            =   1050
         TabIndex        =   9
         Top             =   945
         Width           =   3315
      End
      Begin VB.TextBox txtImagenes 
         Height          =   315
         Left            =   1050
         TabIndex        =   8
         Top             =   600
         Width           =   3315
      End
      Begin VB.TextBox txtAdministrativa 
         Height          =   315
         Left            =   1050
         TabIndex        =   11
         Top             =   1650
         Width           =   3315
      End
      Begin VB.TextBox txtNegocio 
         Height          =   315
         Left            =   1050
         TabIndex        =   10
         Top             =   1290
         Width           =   3315
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Personas:"
         Height          =   195
         Left            =   210
         TabIndex        =   26
         Top             =   315
         Width           =   705
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Comunes :"
         Height          =   195
         Left            =   210
         TabIndex        =   25
         Top             =   1005
         Width           =   750
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Imagenes:"
         Height          =   195
         Left            =   210
         TabIndex        =   24
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Administ. :"
         Height          =   195
         Left            =   210
         TabIndex        =   23
         Top             =   1710
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Negocio :"
         Height          =   195
         Left            =   210
         TabIndex        =   22
         Top             =   1365
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   4080
      TabIndex        =   13
      Top             =   4890
      Width           =   1470
   End
   Begin VB.CommandButton cmdgenerar 
      Caption         =   "&Generar Archivo"
      Height          =   390
      Left            =   2625
      TabIndex        =   12
      Top             =   4890
      Width           =   1470
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Conexion"
      Height          =   2460
      Left            =   180
      TabIndex        =   14
      Top             =   75
      Width           =   5460
      Begin VB.ComboBox cboAplicacion 
         Height          =   315
         ItemData        =   "frmIni.frx":0000
         Left            =   1065
         List            =   "frmIni.frx":001C
         TabIndex        =   27
         Top             =   255
         Width           =   2775
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test Conexion"
         Height          =   390
         Left            =   3915
         TabIndex        =   6
         Top             =   1950
         Width           =   1470
      End
      Begin VB.TextBox txtDataBase 
         Height          =   315
         Left            =   1050
         TabIndex        =   3
         Top             =   1290
         Width           =   2760
      End
      Begin VB.TextBox txtPassWord 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1050
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   2010
         Width           =   2040
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Left            =   1050
         TabIndex        =   4
         Top             =   1650
         Width           =   2760
      End
      Begin VB.TextBox txtProvider 
         Height          =   315
         Left            =   1050
         TabIndex        =   1
         Text            =   "SQLOLEDB"
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txtServer 
         Height          =   315
         Left            =   1050
         TabIndex        =   2
         Top             =   945
         Width           =   2760
      End
      Begin VB.TextBox txtAplicacion 
         Height          =   300
         Left            =   1050
         TabIndex        =   0
         Text            =   "SICMACT"
         Top             =   270
         Visible         =   0   'False
         Width           =   3690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DataBase :"
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   1365
         Width           =   795
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   2040
         Width           =   690
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "User :"
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   1710
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Provider :"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   645
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servidor :"
         Height          =   195
         Left            =   210
         TabIndex        =   16
         Top             =   1005
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Aplicacion"
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbOk As Boolean
Public lsArchivoIni As String
Dim lsAplicacion As String
Dim lsArchIni As String
Public Sub Inicio(psAplicacion As String, psArchIni As String)
lsAplicacion = psAplicacion
lsArchivoIni = psArchIni
Me.Show 1
End Sub
Private Sub cmdCancelar_Click()
Me.Hide
End Sub
Private Sub cmdgenerar_Click()
'JIPR20190626 INICIO
Dim Apl As String 'JIPR20190626
'mdGuardar.CancelError = True JIPR20190626
'On Error GoTo ErrHandler JIPR20190626
'cmdGuardar.Flags = cdlOFNHideReadOnly JIPR20190626
' Establecer los filtros
'cmdGuardar.DialogTitle = "Guardar Achivo Ini como...." JIPR20190626

    cmdGuardar.CancelError = True
    On Error GoTo ErrHandler
    cmdGuardar.Flags = cdlOFNHideReadOnly
    cmdGuardar.DialogTitle = "Guardar Achivo Ini como...."

    If Me.cboAplicacion.ListIndex = 0 Or Me.cboAplicacion.ListIndex = 1 Or Me.cboAplicacion.ListIndex = 2 Then
       cmdGuardar.FileName = "SicmactN.Ini"
       Apl = "SICMACT"
    ElseIf Me.cboAplicacion.ListIndex = 3 Then
       cmdGuardar.FileName = "SicmactF.Ini"
       Apl = "SICMACT"
    ElseIf Me.cboAplicacion.ListIndex = 4 Then
       cmdGuardar.FileName = "SicmactA.Ini"
       Apl = "SICMACT"
    ElseIf Me.cboAplicacion.ListIndex = 5 Then
       cmdGuardar.FileName = "TarjAdm.Ini"
       Apl = "TarjAdm"
    ElseIf Me.cboAplicacion.ListIndex = 6 Then
        If Me.txtDataBase = "DBCmacMaynas" Then
         cmdGuardar.FileName = "SicmactN.Ini"
         Apl = "SICMACT"
        Else
        cmdGuardar.FileName = "AutCC.Ini"
        Apl = "AutCC"
        End If
    Else
       cmdGuardar.FileName = "TarjAdm.Ini"
        Apl = "TXRXProcess"
    End If

'cmdGuardar.FileName = "Sicmact.Ini" ' JIPR20190626

cmdGuardar.InitDir = App.Path
cmdGuardar.Filter = "Archivos de Inicio (*.Ini)|*.Ini"
' Especificar el filtro predeterminado
cmdGuardar.FilterIndex = 1
cmdGuardar.ShowSave
 
    EscribirIni Encripta(Apl), Encripta("Provider"), Encripta(txtProvider), cmdGuardar.FileName
    EscribirIni Encripta(Apl), Encripta("Server"), Encripta(txtServer), cmdGuardar.FileName
    EscribirIni Encripta(Apl), Encripta("DataBase"), Encripta(txtDataBase), cmdGuardar.FileName
    EscribirIni Encripta(Apl), Encripta("User"), Encripta(txtUser), cmdGuardar.FileName
    EscribirIni Encripta(Apl), Encripta("Password"), Encripta(txtPassWord), cmdGuardar.FileName

    EscribirIni Encripta("BASE COMUNES"), Encripta("dbPersonas"), Encripta(txtpersonas), cmdGuardar.FileName
    EscribirIni Encripta("BASE COMUNES"), Encripta("dbImagenes"), Encripta(txtImagenes), cmdGuardar.FileName
    EscribirIni Encripta("BASE COMUNES"), Encripta("dbComunes"), Encripta(txtComunes), cmdGuardar.FileName
    EscribirIni Encripta("BASE COMUNES"), Encripta("dbNegocio"), Encripta(txtNegocio), cmdGuardar.FileName
    EscribirIni Encripta("BASE COMUNES"), Encripta("dbAdmin"), Encripta(txtAdministrativa), cmdGuardar.FileName
    
    lsArchivoIni = cmdGuardar.FileName

    MsgBox "Archivo " & cmdGuardar.FileName & vbCrLf & " Generado con Exito ", vbInformation, "Aviso"
    cmdgenerar.Enabled = False

'JIPR20190626
'EscribirIni Encripta(txtAplicacion), Encripta("Provider"), Encripta(txtProvider), cmdGuardar.FileName
'EscribirIni Encripta(txtAplicacion), Encripta("Server"), Encripta(txtServer), cmdGuardar.FileName
'EscribirIni Encripta(txtAplicacion), Encripta("DataBase"), Encripta(txtDataBase), cmdGuardar.FileName
'EscribirIni Encripta(txtAplicacion), Encripta("User"), Encripta(txtUser), cmdGuardar.FileName
'EscribirIni Encripta(txtAplicacion), Encripta("Password"), Encripta(txtPassWord), cmdGuardar.FileName
'
'EscribirIni Encripta("BASE COMUNES"), Encripta("dbPersonas"), Encripta(txtpersonas), cmdGuardar.FileName
'EscribirIni Encripta("BASE COMUNES"), Encripta("dbImagenes"), Encripta(txtImagenes), cmdGuardar.FileName
'EscribirIni Encripta("BASE COMUNES"), Encripta("dbComunes"), Encripta(txtComunes), cmdGuardar.FileName
'EscribirIni Encripta("BASE COMUNES"), Encripta("dbNegocio"), Encripta(txtNegocio), cmdGuardar.FileName
'EscribirIni Encripta("BASE COMUNES"), Encripta("dbAdmin"), Encripta(txtAdministrativa), cmdGuardar.FileName
'
'
'
'
'lsArchivoIni = cmdGuardar.FileName
'
'MsgBox "Archivo " & cmdGuardar.FileName & vbCrLf & " Generado con Exito ", vbInformation, "Aviso"
'cmdgenerar.Enabled = False
'JIPR20190626 FIN

Exit Sub
ErrHandler:
    MsgBox "Error N° [" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
 
End Sub
Private Function Valida() As Boolean
Valida = True
If Len(Trim(txtAplicacion)) = 0 Then
    MsgBox "Nombre de aplicacion no válida", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If
If Len(Trim(txtProvider)) = 0 Then
    MsgBox "Provider no válido", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If
If Len(Trim(txtServer)) = 0 Then
    MsgBox "Servidor no válido", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If
If Len(Trim(Me.txtDataBase)) = 0 Then
    MsgBox "Base de Datos no válida", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If
If Len(Trim(txtUser)) = 0 Then
    MsgBox "User no válido", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If
If Len(Trim(txtPassWord)) = 0 Then
    MsgBox "Password no válido", vbInformation, "Aviso"
    Valida = False
    Exit Function
End If

End Function

Private Sub cmdTest_Click()
Dim Con As ADODB.Connection
Set Con = New ADODB.Connection
Dim lsCadena As String
If Valida = False Then Exit Sub
On Error GoTo ErrorCon
'cmdgenerar.Enabled = False
lsCadena = "PROVIDER=" & Trim(txtProvider) & ";uid=" & Trim(txtUser) & ";pwd=" & Trim(txtPassWord) & ";DATABASE=" & Trim(txtDataBase) & ";SERVER=" & Trim(txtServer) & ""
Con.Open lsCadena
Con.Close
Set Con = Nothing
MsgBox "Test de Conexion realizado con éxito", vbInformation, "Aviso"
'cmdgenerar.Enabled = True
'cmdgenerar.SetFocus
Exit Sub
ErrorCon:
    If Not Con Is Nothing Then
        If Con.State = adStateOpen Then
            Con.Close
            Set Con = Nothing
        End If
    End If
    MsgBox "Error en Conexion con Servidor " & vbCrLf & Err.Description, vbInformation, "Test de conexion"
End Sub
Private Sub Form_Load()
'JIPRJIPR20190626 COMENTÓ
'txtAplicacion = lsAplicacion
'txtProvider = Encripta(LeerArchivoIni(Encripta(lsAplicacion), Encripta("Provider"), lsArchivoIni), False)
'If txtProvider = "" Then txtProvider = "SQLOLEDB"
'txtServer = Encripta(LeerArchivoIni(Encripta(lsAplicacion), Encripta("Server"), lsArchivoIni), False)
'txtUser = Encripta(LeerArchivoIni(Encripta(lsAplicacion), Encripta("User"), lsArchivoIni), False)
'txtPassWord = Encripta(LeerArchivoIni(Encripta(lsAplicacion), Encripta("Password"), lsArchivoIni), False)
'txtDataBase = Encripta(LeerArchivoIni(Encripta(lsAplicacion), Encripta("DataBase"), lsArchivoIni), False)
'
'
'txtpersonas = Encripta(LeerArchivoIni(Encripta("BASE COMUNES"), Encripta("dbPersonas"), lsArchivoIni), False)
'txtImagenes = Encripta(LeerArchivoIni(Encripta("BASE COMUNES"), Encripta("dbImagenes"), lsArchivoIni), False)
'txtComunes = Encripta(LeerArchivoIni(Encripta("BASE COMUNES"), Encripta("dbComunes"), lsArchivoIni), False)
'txtNegocio = Encripta(LeerArchivoIni(Encripta("BASE COMUNES"), Encripta("dbNegocio"), lsArchivoIni), False)
'txtAdministrativa = Encripta(LeerArchivoIni(Encripta("BASE COMUNES"), Encripta("dbAdmin"), lsArchivoIni), False)

End Sub

Private Sub txtAdministrativa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtAplicacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtProvider.SetFocus
End If
End Sub

Private Sub txtComunes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtDataBase_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtUser.SetFocus
End If

End Sub

Private Sub txtImagenes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtNegocio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtPassWord_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdTest.SetFocus
End If
End Sub

Private Sub txtpersonas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
End If
End Sub

Private Sub txtProvider_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtServer.SetFocus
End If
End Sub
Private Sub txtServer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtDataBase.SetFocus
End If
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPassWord.SetFocus
End If
End Sub
