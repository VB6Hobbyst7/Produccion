VERSION 5.00
Begin VB.Form frmProveedorBuscaxAcumulado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5370
   Icon            =   "frmProveedorBuscaxAcumulado.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   5370
   StartUpPosition =   2  'CenterScreen
   Begin Sicmact.TxtBuscar txtPersonaCod 
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TipoBusqueda    =   3
      sTitulo         =   ""
      TipoBusPers     =   1
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   320
      Left            =   3110
      TabIndex        =   9
      Top             =   2470
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Total Acumulado Mes (Recibos por Honorario)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   80
      TabIndex        =   6
      Top             =   1440
      Width           =   5250
      Begin VB.Label lblTotalAcumulado 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   840
         TabIndex        =   8
         Top             =   420
         Width           =   1515
      End
      Begin VB.Label Label6 
         Caption         =   "S/. :"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   320
      Left            =   4215
      TabIndex        =   2
      Top             =   2470
      Width           =   1095
   End
   Begin VB.Frame fraProveedor 
      Caption         =   "Datos del Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1335
      Left            =   80
      TabIndex        =   0
      Top             =   120
      Width           =   5250
      Begin VB.Label Label3 
         Caption         =   "RUC :"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblDOI 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   840
         TabIndex        =   4
         Top             =   840
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   705
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   840
         TabIndex        =   1
         Top             =   480
         Width           =   4275
      End
   End
End
Attribute VB_Name = "frmProveedorBuscaxAcumulado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************
'** Nombre : frmReciboHonorarioProvBusca
'** Descripción : Formulario para buscar el acumulado de los recibo por Honorarios del Proveedor
'** Creación : EJVG, 20140721 09:00:00 AM
'***********************************************************************************************
Option Explicit

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdLimpiar_Click()
    Limpiar
    EnfocaControl txtPersonaCod
End Sub
Private Sub Form_Load()
    cmdLimpiar_Click
End Sub
Private Sub Limpiar()
    txtPersonaCod.Text = ""
    lblNombre.Caption = ""
    lblDOI.Caption = ""
    lblTotalAcumulado.Caption = "0.00"
End Sub
Private Sub txtPersonaCod_EmiteDatos()
    If txtPersonaCod.Text <> "" Then
        If txtPersonaCod.PersPersoneria = gPersonaNat Then
            If Not cargar_datos(txtPersonaCod.Text) Then
                MsgBox "No se encontraron datos de la persona seleccionada", vbInformation, "Aviso"
                txtPersonaCod.Text = ""
                Exit Sub
            End If
        Else
            MsgBox "Ud. debe seleccionar una Persona Natural", vbInformation, "Aviso"
            txtPersonaCod.Text = ""
            Exit Sub
        End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            Dim lsArendir As String
            'gsOpeCod = LogPistaCierreDiarioCont
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Consulto Datos del Proveedor Nombre : " & lblNombre & " |RUC : " & lblDOI _
            & " |Monto Acumlado al Mes : " & lblTotalAcumulado, 3
            Set objPista = Nothing
            '*******
        
    Else
        Limpiar
    End If
End Sub
Private Function cargar_datos(ByVal psPersCod As String) As Boolean
    Dim oPersona As New DPersonas
    Dim rsPersona As New ADODB.Recordset
    
    On Error GoTo ErrorCargar_datos
    Screen.MousePointer = 11
    Set rsPersona = oPersona.ObtieneDatosProveedorRetencSistPens(psPersCod, gdFecSis)
    If Not rsPersona.EOF Then
        lblNombre.Caption = rsPersona!cPersNombre
        lblDOI.Caption = rsPersona!cDOI
        lblTotalAcumulado.Caption = Format(rsPersona!nMontoAcumulado, gsFormatoNumeroView)
        cargar_datos = True
    Else
        cargar_datos = False
    End If
    RSClose rsPersona
    Set oPersona = Nothing
    Screen.MousePointer = 0
    Exit Function
ErrorCargar_datos:
    cargar_datos = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
