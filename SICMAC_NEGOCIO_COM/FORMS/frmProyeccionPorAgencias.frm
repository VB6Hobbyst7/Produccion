VERSION 5.00
Begin VB.Form frmProyeccionPorAgencias 
   Caption         =   "Proyeccion Semanal"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5535
   Icon            =   "frmProyeccionPorAgencias.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesembMN 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1275
      Width           =   1500
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   4320
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Height          =   375
      Left            =   3120
      TabIndex        =   10
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox txtCarteraME 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3915
      TabIndex        =   9
      Top             =   2955
      Width           =   1500
   End
   Begin VB.TextBox txtCrecimME 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3915
      TabIndex        =   7
      Top             =   2115
      Width           =   1500
   End
   Begin VB.TextBox txtDesembME 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3915
      TabIndex        =   5
      Top             =   1275
      Width           =   1500
   End
   Begin VB.TextBox txtCarteraMN 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   2955
      Width           =   1500
   End
   Begin VB.TextBox txtCrecimMN 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   2115
      Width           =   1500
   End
   Begin VB.Label Label13 
      Caption         =   "F. Hasta:"
      Height          =   255
      Left            =   4200
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "F. Desde:"
      Height          =   255
      Left            =   2880
      TabIndex        =   22
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblFechaHasta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   4200
      TabIndex        =   21
      Top             =   555
      Width           =   1215
   End
   Begin VB.Label lblFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   270
      Left            =   2880
      TabIndex        =   20
      Top             =   555
      Width           =   1215
   End
   Begin VB.Label lblAgencia 
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
      Height          =   300
      Left            =   285
      TabIndex        =   19
      Top             =   555
      Width           =   2295
   End
   Begin VB.Label Label11 
      Caption         =   "* Datos al Cierre Semanal"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label10 
      Caption         =   "M.E.:"
      Height          =   255
      Left            =   3465
      TabIndex        =   16
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "M.E.:"
      Height          =   255
      Left            =   3465
      TabIndex        =   15
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "Cartera Atrasada (valores absolutos)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label7 
      Caption         =   "Crecimiento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "M.N.:"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   3000
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "M.N.:"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "M.E.:"
      Height          =   255
      Left            =   3465
      TabIndex        =   3
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "M.N.:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Desembolsos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Agencia:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "frmProyeccionPorAgencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************************************************************************************************************
'* NOMBRE         : "frmProyeccionPorAgencias"
'* DESCRIPCION    : Formulario creado para registrar las proyecciones semanales
'* CREACION       : FRHU, 20140415 10:00 AM
'*********************************************************************************************************************************************************
Option Explicit
Dim dFecha As Date
Dim sAgeCod As String
Dim nMostrarVentana As Integer
Dim nSalir As Integer
Dim nTipo As Integer

Public Sub Inicio(ByVal psAgecod As String, ByVal pdFecha As Date)
    Dim objAgencia As New COMDConstantes.DCOMAgencias
    dFecha = pdFecha
    sAgeCod = psAgecod
    Me.lblAgencia.Caption = objAgencia.NombreAgencia(psAgecod)
    Call LimpiarPantalla
    Call CargarDatos
    If nMostrarVentana = 1 Then
        Me.Show 1
    End If
End Sub
Private Sub CmdGrabar_Click()
    Dim objDCred As New COMDCredito.DCOMCredito
    Dim loContFunct As New COMNContabilidad.NCOMContFunciones
    Dim lsMovNro As String
    If Not ValidarCampos Then
        Exit Sub
    End If
    If MsgBox("Se van a Guardar los Datos. Desea Continuar?", vbYesNo) = vbNo Then
        Exit Sub
    End If
    'Generar el MovNro
    lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    'Registra
    Call objDCred.InsertarProyeccionColocAge(sAgeCod, lblFecha.Caption, txtDesembMN.Text, txtDesembME.Text, txtCrecimMN.Text, txtCrecimME.Text, txtCarteraMN.Text, txtCarteraME.Text, lsMovNro)
    Call CargarDatos
    Call LimpiarPantalla
    If nMostrarVentana = 0 Then
        Unload Me
    End If
End Sub
Private Function ValidarCampos() As Boolean
    If CStr(Me.txtDesembMN.Text) = "" Or CStr(Me.txtDesembME.Text) = "" Or CStr(Me.txtCarteraMN.Text) = "" Or CStr(Me.txtCarteraME.Text) = "" Or CStr(Me.txtCrecimMN.Text) = "" Or CStr(Me.txtCrecimME.Text) = "" Then
       MsgBox "Debe llenar todos los campos", vbInformation
       ValidarCampos = False
    Else
       ValidarCampos = True
    End If
End Function
Private Sub cmdsalir_Click()
    End
End Sub
Private Sub CargarDatos()
    Dim objRS As ADODB.Recordset
    Dim objDCred As New COMDCredito.DCOMCredito
    
    Set objRS = objDCred.GetFechaSemanaProyectadoAge(dFecha, sAgeCod)
    If Not objRS.EOF And Not objRS.BOF Then
        Me.lblFecha.Caption = CStr(objRS!FechaInicioSemana)
        Me.lblFechaHasta.Caption = CStr(objRS!FechaFinSemana)
        nMostrarVentana = objRS!mostrarVentana
    End If
    Set objDCred = Nothing
End Sub
Private Sub LimpiarPantalla()
Me.txtDesembMN.Text = ""
Me.txtDesembME.Text = ""
Me.txtCrecimMN.Text = ""
Me.txtCrecimME.Text = ""
Me.txtCarteraMN.Text = ""
Me.txtCarteraME.Text = ""
End Sub

Public Sub MostrarProyeccionSemanal(ByVal psCodAge As String, ByVal psAgencia As String, ByVal pdFechaIni As Date, ByVal pdFechaFin As Date, ByVal pnTipo As Integer)
'Esto procedmiento se usa en el formulario frmProyeccionSeguimientoPorAgencia
    Dim oDCred As New COMDCredito.DCOMCredito
    Dim oRs As ADODB.Recordset
    nTipo = pnTipo
    nMostrarVentana = 0 'No muestra la ventana al inicio del sicma
    Me.lblAgencia.Caption = psAgencia
    Me.lblFecha.Caption = pdFechaIni
    Me.lblFechaHasta.Caption = pdFechaFin
    Set oRs = oDCred.ObtenerProyeccionColocAge(psCodAge, pdFechaIni)
    If Not oRs.EOF And Not oRs.BOF Then
        Me.txtDesembMN.Text = Format(oRs!nDesembMN, "#,#0.00")
        Me.txtDesembME.Text = Format(oRs!nDesembME, "#,#0.00")
        Me.txtCrecimMN.Text = Format(oRs!nCrecimMN, "#,#0.00")
        Me.txtCrecimME.Text = Format(oRs!nCrecimME, "#,#0.00")
        Me.txtCarteraMN.Text = Format(oRs!nCarteraMN, "#,#0.00")
        Me.txtCarteraME.Text = Format(oRs!nCarteraME, "#,#0.00")
        frmProyeccionPorAgencias.Height = 4200
        Call txtSoloLectura
        Me.Show 1
    Else
        MsgBox "No hay datos que mostrar", vbInformation
        Unload Me
    End If
End Sub
Private Sub txtSoloLectura()
    txtDesembMN.Locked = True
    txtDesembME.Locked = True
    txtCrecimMN.Locked = True
    txtCrecimME.Locked = True
    txtCarteraMN.Locked = True
    txtCarteraME.Locked = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If nMostrarVentana = 0 Then
        Unload Me
    Else
        End
    End If
End Sub
Private Sub txtDesembMN_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtDesembMN, KeyAscii, 13, 2)
    If KeyAscii = 13 Then
        Me.txtDesembME.SetFocus
    End If
End Sub
Private Sub txtDesembMN_LostFocus()
    If Trim(txtDesembMN.Text) = "" Then
        txtDesembMN.Text = "0.00"
    Else
        txtDesembMN.Text = Format(txtDesembMN.Text, "#,#0.00")
    End If
End Sub
Private Sub txtDesembME_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtDesembME, KeyAscii, 13, 2)
    If KeyAscii = 13 Then
        Me.txtCrecimMN.SetFocus
    End If
End Sub
Private Sub txtDesembME_LostFocus()
    If Trim(txtDesembME.Text) = "" Then
        txtDesembME.Text = "0.00"
    Else
        txtDesembME.Text = Format(txtDesembME.Text, "#,#0.00")
    End If
End Sub
Private Sub txtCrecimMN_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCrecimMN, KeyAscii, 13, 2)
    If KeyAscii = 13 Then
        Me.txtCrecimME.SetFocus
    End If
End Sub
Private Sub txtCrecimMN_LostFocus()
    If Trim(txtCrecimMN.Text) = "" Then
        txtCrecimMN.Text = "0.00"
    Else
        txtCrecimMN.Text = Format(txtCrecimMN.Text, "#,#0.00")
    End If
End Sub
Private Sub txtCrecimME_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCrecimME, KeyAscii, 13, 2)
    If KeyAscii = 13 Then
        Me.txtCarteraMN.SetFocus
    End If
End Sub
Private Sub txtCrecimME_LostFocus()
    If Trim(txtCrecimME.Text) = "" Then
        txtCrecimME.Text = "0.00"
    Else
        txtCrecimME.Text = Format(txtCrecimME.Text, "#,#0.00")
    End If
End Sub
Private Sub txtCarteraMN_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCarteraMN, KeyAscii, 13, 2)
    If KeyAscii = 13 Then
        Me.txtCarteraME.SetFocus
    End If
End Sub
Private Sub txtCarteraMN_LostFocus()
    If Trim(txtCarteraMN.Text) = "" Then
        txtCarteraMN.Text = "0.00"
    Else
        txtCarteraMN.Text = Format(txtCarteraMN.Text, "#,#0.00")
    End If
End Sub
Private Sub txtCarteraME_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtCarteraME, KeyAscii, 13, 2)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    End If
End Sub
Private Sub txtCarteraME_LostFocus()
    If Trim(txtCarteraME.Text) = "" Then
        txtCarteraME.Text = "0.00"
    Else
        txtCarteraME.Text = Format(txtCarteraME.Text, "#,#0.00")
    End If
End Sub









