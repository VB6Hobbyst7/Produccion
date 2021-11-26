VERSION 5.00
Begin VB.Form frmActInfContacto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualización de Información de Contacto"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   Icon            =   "frmActInfContacto.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Actualización de Datos"
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
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   6840
         TabIndex        =   18
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos de Contacto"
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   7935
         Begin VB.TextBox txtCel3 
            Height          =   285
            Left            =   6360
            MaxLength       =   12
            TabIndex        =   17
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtCel2 
            Height          =   285
            Left            =   4800
            MaxLength       =   12
            TabIndex        =   15
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtCel 
            Height          =   285
            Left            =   3240
            MaxLength       =   12
            TabIndex        =   13
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtTrabajo 
            Height          =   285
            Left            =   1680
            MaxLength       =   12
            TabIndex        =   11
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtDomicilio 
            Height          =   285
            Left            =   120
            MaxLength       =   12
            TabIndex        =   9
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Telf. Celular 3:"
            Height          =   255
            Left            =   6360
            TabIndex        =   16
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Telf. Celular 2:"
            Height          =   255
            Left            =   4800
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Telf. Celular 1:"
            Height          =   255
            Left            =   3240
            TabIndex        =   12
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Telf. Trabajo:"
            Height          =   255
            Left            =   1680
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Telf. Domicilio:"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Cliente"
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7935
         Begin VB.TextBox txtDOI 
            Enabled         =   0   'False
            Height          =   285
            Left            =   6480
            TabIndex        =   7
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtTipoDOI 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4920
            TabIndex        =   6
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtNombre 
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            TabIndex        =   4
            Top             =   240
            Width           =   3375
         End
         Begin VB.Label Label2 
            Caption         =   "DOI:"
            Height          =   255
            Left            =   4440
            TabIndex        =   5
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre:"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmActInfContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'** Nombre : frmActInfContacto
'** Descripción : Esta ventana se mostrara en diferentes procesos para registrar los datos faltantes del contacto.
'** Creación : AMDO, 20130702 05:00:00 AM
'** Referencia : TI-ERS063-2013
'***************************************************************************
Option Explicit
Dim cPersCod As String

Private fTelDomicilio As String
Private fTelTrabajo As String
Private fTelCel As String
Private fTelCel2 As String
Private fTelCel3 As String

Dim oPersona As UPersona_Cli

Public Sub Inicio(ByVal cPersCodigo As String)
     Set oPersona = New UPersona_Cli
     cPersCod = cPersCodigo
     oPersona.RecuperaPersona (cPersCod)
    
     Dim rsDOI As New ADODB.Recordset
     Dim lnTipoDoc As Integer
     Dim lsTipoDoc As String
     Dim lsNumDoc As String
     Call oPersona.ObtenerDatosDocumentoxPos(0, lnTipoDoc, lsTipoDoc, lsNumDoc)
     
     fTelDomicilio = oPersona.Telefonos
     fTelTrabajo = oPersona.Telefonos2
     fTelCel = oPersona.Celular
     fTelCel2 = oPersona.Celular2
     fTelCel3 = oPersona.Celular3
     
     txtTipoDOI.Text = lsTipoDoc
     txtDOI.Text = lsNumDoc
     txtNombre.Text = oPersona.NombreCompleto
     txtDomicilio.Text = oPersona.Telefonos
     txtTrabajo.Text = oPersona.Telefonos2
     txtCel.Text = oPersona.Celular
     txtCel2.Text = oPersona.Celular2
     txtCel3.Text = oPersona.Celular3
     Me.Show 1
End Sub

Private Sub cmdActualizar_Click()
Dim oPersona As COMDPersona.DCOMPersona
Set oPersona = New COMDPersona.DCOMPersona

    If Not ValidaDatos Then
        Exit Sub
    End If

    Call oPersona.ActualizarInfoContacto(cPersCod, gsCodUser, gsCodAge, Format(gdFecSis, "YYYY-MM-DD") & " " & Format(Time(), "HH:MM:SS"), txtDomicilio.Text, _
                                    txtTrabajo.Text, txtCel.Text, txtCel2.Text, txtCel3.Text)

    MsgBox "La Actualización de Información del contacto fue registrada con Exito", vbInformation, "Aviso"

Unload Me
End Sub

Private Function ValidaDatos() As Boolean
    ValidaDatos = True
    If fTelDomicilio = txtDomicilio.Text And fTelTrabajo = txtTrabajo.Text And fTelCel = txtCel.Text _
       And fTelCel2 = txtCel2.Text And fTelCel3 = txtCel3.Text Then
        MsgBox "Es necesario actualizar uno o mas campos para continuar con la operación."
        txtDomicilio.SetFocus
        ValidaDatos = False
        Exit Function
    End If
    
    If txtDomicilio.Text = "" And txtTrabajo.Text = "" And txtCel.Text = "" _
       And txtCel2.Text = "" And txtCel3.Text = "" Then
        MsgBox "Todos los campos de los numeros no pueden quedar vacios."
        txtDomicilio.SetFocus
        ValidaDatos = False
        Exit Function
    End If
End Function

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtTrabajo.SetFocus
    End If
End Sub

Private Sub txtTrabajo_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        txtCel.SetFocus
    End If
End Sub

Private Sub txtCel_KeyPress(KeyAscii As Integer)
    If txtCel.Text = "" Then
        If DigitoRPM(KeyAscii) = False Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                txtCel2.SetFocus
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            txtCel2.SetFocus
        End If
    End If
End Sub

Private Sub txtCel2_KeyPress(KeyAscii As Integer)
    If txtCel2.Text = "" Then
        If DigitoRPM(KeyAscii) = False Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                txtCel3.SetFocus
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            txtCel3.SetFocus
        End If
    End If
End Sub

Private Sub txtCel3_KeyPress(KeyAscii As Integer)
    If txtCel3.Text = "" Then
        If DigitoRPM(KeyAscii) = False Then
            KeyAscii = NumerosEnteros(KeyAscii)
            If KeyAscii = 13 Then
                cmdActualizar.SetFocus
            End If
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
        If KeyAscii = 13 Then
            cmdActualizar.SetFocus
        End If
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set oPersona = Nothing
End Sub
