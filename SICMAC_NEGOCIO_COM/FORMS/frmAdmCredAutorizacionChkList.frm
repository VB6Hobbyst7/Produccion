VERSION 5.00
Begin VB.Form frmAdmCredAutorizacionChkList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorización de Mantenimiento de checkList de Requisitos"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "frmAdmCredAutorizacionChkList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "Autorizar"
      Height          =   300
      Left            =   4200
      TabIndex        =   20
      Top             =   3690
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   300
      Left            =   5400
      TabIndex        =   19
      Top             =   3690
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   " Autorización "
      ForeColor       =   &H00FF0000&
      Height          =   1660
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   6495
      Begin VB.TextBox txtGlosa 
         Height          =   285
         Left            =   960
         TabIndex        =   21
         Top             =   1230
         Width           =   5390
      End
      Begin SICMACT.TxtBuscar TxtBuscar 
         Height          =   270
         Left            =   960
         TabIndex        =   15
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   476
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
         sTitulo         =   ""
      End
      Begin VB.Label Label16 
         Caption         =   "Glosa:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblCargo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   960
         TabIndex        =   17
         Top             =   950
         Width           =   5385
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   960
         TabIndex        =   16
         Top             =   660
         Width           =   5380
      End
      Begin VB.Label Label12 
         Caption         =   "Cargo:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Crédito"
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   6495
      Begin VB.Label lblProducto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   4860
         TabIndex        =   10
         Top             =   680
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   4080
         TabIndex        =   9
         Top             =   690
         Width           =   735
      End
      Begin VB.Label lblFecha 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   2880
         TabIndex        =   8
         Top             =   680
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha:"
         Height          =   255
         Left            =   2340
         TabIndex        =   7
         Top             =   690
         Width           =   615
      End
      Begin VB.Label lblMonto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1050
         TabIndex        =   6
         Top             =   680
         Width           =   1215
      End
      Begin VB.Label lblMoneda 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   660
         TabIndex        =   5
         Top             =   680
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   690
         Width           =   615
      End
      Begin VB.Label lblTitular 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   660
         TabIndex        =   3
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "Titular:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   380
         Width           =   615
      End
   End
   Begin SICMACT.ActXCodCta_New ActXCodCta 
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1296
      Texto           =   "Crédito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmAdmCredAutorizacionChkList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmAdmCredAutorizacionChkList
'** Descripción : Formulario que permite autorizar el mantenimiento de checkList de requisitos
'** Creación    : RECO, 20150421 - ERS010-2015
'**********************************************************************************************
Option Explicit
Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call CargarDatosCred
    End If
End Sub

Private Sub cmdAutorizar_Click()
    Dim obj As New COMNCredito.NCOMCredito
    If ValidaDatos = "" Then
        Call obj.CredAdmRegistraAutorizacionChekList(ActXCodCta.NroCuenta, Trim(TxtBuscar.Text), txtGlosa.Text, Format(gdFecSis, "yyyy/MM/dd"), 1)
        Call LimpiarFormulario
        MsgBox "La autorización se registró de forma satisfactoria.", vbInformation, "Alerta"
    Else
        MsgBox ValidaDatos, vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub CargarDatosCred()
    Dim obj As New COMNCredito.NCOMCredito
    Dim rs As New ADODB.Recordset
    
    Set rs = obj.ObtieneDatosAutorizaChekList(ActXCodCta.NroCuenta)
    If Not (rs.EOF And rs.BOF) Then
        lblTitular.Caption = rs!cPersNombre
        lblMoneda.Caption = rs!nmoneda
        lblMonto.Caption = Format(rs!nMontoCol, gcFormView)
        lblFecha.Caption = rs!dVigencia
        lblProducto.Caption = rs!cConsDescripcion
    End If
End Sub

Private Sub Form_Load()
    Call LimpiarFormulario
End Sub

Private Sub TxtBuscar_EmiteDatos()
    Dim o As New COMDPersona.UCOMPersona
    Dim objRH As New COMDPersona.UCOMAcceso
    Dim rs As New ADODB.Recordset
    
    Set o = frmBuscaPersona.Inicio(True)
    TxtBuscar.Text = o.sPersCod
    If TxtBuscar.Text <> "" Then
        TxtBuscar.Enabled = False
        txtGlosa.SetFocus
    End If
    lblNombre.Caption = o.sPersNombre
    Set rs = objRH.ObtieneDatosPersonaRRHH("", Trim(TxtBuscar.Text))
    If Not (rs.BOF And rs.EOF) Then
        lblCargo.Caption = rs!cCargo
        cmdAutorizar.Enabled = True
    End If
End Sub

Private Sub LimpiarFormulario()
    ActXCodCta.NroCuenta = ""
    lblTitular.Caption = ""
    lblMoneda.Caption = ""
    lblMonto.Caption = ""
    lblFecha.Caption = "__/__/____"
    lblProducto.Caption = ""
    TxtBuscar.Text = ""
    lblNombre.Caption = ""
    lblCargo.Caption = ""
    txtGlosa.Text = ""
    cmdAutorizar.Enabled = False
    TxtBuscar.Enabled = True
End Sub

Private Function ValidaDatos() As String
    If TxtBuscar.Text = "" Then
        ValidaDatos = "Debe seleccionar un usuario para realizar la autorización."
        Exit Function
    End If
    If txtGlosa.Text = "" Then
        ValidaDatos = "Debe ingresar una glosa."
    End If
End Function
