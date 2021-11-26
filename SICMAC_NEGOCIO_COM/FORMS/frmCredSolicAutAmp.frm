VERSION 5.00
Begin VB.Form frmCredSolicAutAmp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorización de Solicitud de Ampliación Excepcional"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11295
   Icon            =   "frmCredSolicAutAmp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "Autorizar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7320
      TabIndex        =   3
      Top             =   3840
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9960
      TabIndex        =   2
      Top             =   3840
      Width           =   1170
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   1
      Top             =   3840
      Width           =   1170
   End
   Begin VB.TextBox txtGlosa 
      Height          =   315
      Left            =   720
      TabIndex        =   0
      Top             =   3850
      Width           =   6375
   End
   Begin SICMACT.FlexEdit FECred 
      Height          =   3585
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   6324
      Cols0           =   13
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "-N° Credito-Agencia-Titular-Moneda-Monto Solic.-Saldo Cap. Amp.-Dias Atraso-Prom. Dias-Calificación-Comentario-cComent-Condición"
      EncabezadosAnchos=   "300-2000-2000-3000-1000-1350-1350-1000-900-1200-1000-0-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   65535
      BackColorControl=   65535
      BackColorControl=   65535
      EncabezadosAlineacion=   "C-C-L-L-C-R-R-C-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Glosa:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3900
      Width           =   615
   End
End
Attribute VB_Name = "frmCredSolicAutAmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredSolicAutAmp
'** Descripción : Formulario para autorizar/rechazar solicitudes de ampliaciones excepcionales
'**               creado segun TI-ERS030-2016
'** Creación : JUEZ, 20160510 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset

Public Sub Inicia()
    CargarDatos
    Me.Show 1
End Sub

Private Sub CargarDatos()
    Set oDCred = New COMDCredito.DCOMCredito
        Set R = oDCred.RecuperaSolicitudAutorizacionAmpliacion(, 0)
    Set oDCred = Nothing
    
    LimpiaFlex FECred
    If Not (R.EOF And R.BOF) Then
        Do While Not R.EOF
            FECred.AdicionaFila
            FECred.TextMatrix(FECred.row, 1) = R!cCtaCod
            FECred.TextMatrix(FECred.row, 2) = R!cAgeDescripcion
            FECred.TextMatrix(FECred.row, 3) = R!cPersNombre
            FECred.TextMatrix(FECred.row, 4) = R!cMoneda
            FECred.TextMatrix(FECred.row, 5) = Format(R!nMontoSol, "#,##0.00")
            FECred.TextMatrix(FECred.row, 6) = Format(R!nSaldoCapAmp, "#,##0.00")
            FECred.TextMatrix(FECred.row, 7) = R!nDiasAtraso
            FECred.TextMatrix(FECred.row, 8) = R!nDiasAtrasoProm
            FECred.TextMatrix(FECred.row, 9) = R!cCalif
            FECred.TextMatrix(FECred.row, 10) = "..."
            FECred.TextMatrix(FECred.row, 11) = R!cComent
            FECred.TextMatrix(FECred.row, 12) = R!cCondicion 'ARLO2018126
            R.MoveNext
        Loop
    Else
        MsgBox "No existen créditos solicitados para la autorización", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdAutorizar_Click()
    RegistrarResultadoSolicitudAmpliacion 1
End Sub

Private Sub cmdRechazar_Click()
    RegistrarResultadoSolicitudAmpliacion 2
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub FECred_Click()
    If FECred.Col = 10 Then
        frmCredListaDatos.InicioTextBox "Comentario", FECred.TextMatrix(FECred.row, 11)
    End If
End Sub

Private Sub RegistrarResultadoSolicitudAmpliacion(ByVal pnResultado As Integer)
    
    If Trim(FECred.TextMatrix(FECred.row, 1)) = "" Then
        MsgBox "Debe seleccionar un crédito", vbInformation, "Aviso"
        Exit Sub
    End If
    If Trim(txtGlosa.Text) = "" Then
        MsgBox "Debe ingresar una glosa", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Se va a " & IIf(pnResultado = 1, "Autorizar", "Rechazar") & " la solicitud de ampliación excepcional del crédito " & FECred.TextMatrix(FECred.row, 1) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set oDCred = New COMDCredito.DCOMCredito
        Call oDCred.RegistraResultadoSolicitudAutorizacionAmpliacion(FECred.TextMatrix(FECred.row, 1), pnResultado, Trim(txtGlosa.Text), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser))
    Set oDCred = Nothing
    
    MsgBox "La solicitud fue " & IIf(pnResultado = 1, "autorizada", "rechazada"), vbInformation, "Aviso"
    txtGlosa.Text = ""
    CargarDatos
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAutorizar.SetFocus
    End If
End Sub
